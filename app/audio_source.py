"""NDS — turn a spoken recording into a timed transcript and per-slide clips.

Option A flow: upload audio → Whisper transcript → Claude drafts slides →
slice the *original* recording to each slide (no TTS).
"""
from __future__ import annotations

import json
import re
import subprocess
from difflib import SequenceMatcher
from pathlib import Path
from typing import List, Optional, Tuple

AUDIO_SUFFIXES = {".mp3", ".m4a", ".wav", ".mpeg", ".mpga", ".webm", ".ogg"}


class AudioSourceError(RuntimeError):
    pass


def is_audio(path: Path) -> bool:
    return path.suffix.lower() in AUDIO_SUFFIXES


def transcribe(path: Path, api_key: str, language: str = "") -> dict:
    """Whisper verbose transcription with segment timestamps.

    Returns ``{"text": str, "segments": [{"start", "end", "text"}, ...], "duration": float}``.
    """
    from openai import OpenAI

    if not path.exists():
        raise AudioSourceError(f"Audio file not found: {path}")
    client = OpenAI(api_key=api_key)
    # Whisper accepts mp3/mp4/mpeg/mpga/m4a/wav/webm
    kwargs = {
        "model": "whisper-1",
        "file": path.open("rb"),
        "response_format": "verbose_json",
    }
    # Optional language hint (ISO-639-1). English/French/… from the UI.
    code = _lang_code(language)
    if code:
        kwargs["language"] = code
    try:
        try:
            result = client.audio.transcriptions.create(
                **kwargs, timestamp_granularities=["segment"])
        except TypeError:
            result = client.audio.transcriptions.create(**kwargs)
        except Exception:
            # Older accounts / models may reject timestamp_granularities
            result = client.audio.transcriptions.create(**kwargs)
    except Exception as exc:
        raise AudioSourceError(f"Transcription failed: {exc}") from exc
    finally:
        kwargs["file"].close()

    # SDK may return object or dict-like
    if hasattr(result, "model_dump"):
        data = result.model_dump()
    elif isinstance(result, dict):
        data = result
    else:
        data = {
            "text": getattr(result, "text", "") or "",
            "segments": [
                {"start": getattr(s, "start", 0), "end": getattr(s, "end", 0),
                 "text": getattr(s, "text", "")}
                for s in (getattr(result, "segments", None) or [])
            ],
            "duration": getattr(result, "duration", None),
        }

    segments = []
    for s in data.get("segments") or []:
        text = (s.get("text") or "").strip()
        if not text:
            continue
        segments.append({
            "start": float(s.get("start") or 0),
            "end": float(s.get("end") or 0),
            "text": text,
        })
    text = (data.get("text") or "").strip()
    if not text and segments:
        text = " ".join(s["text"] for s in segments)
    if not text:
        raise AudioSourceError("No speech detected in the audio file.")

    duration = data.get("duration")
    if duration is None and segments:
        duration = segments[-1]["end"]
    return {"text": text, "segments": segments, "duration": float(duration or 0)}


def save_transcript(job_dir: Path, transcript: dict) -> Path:
    path = job_dir / "transcript.json"
    path.write_text(json.dumps(transcript, indent=2), encoding="utf-8")
    (job_dir / "source_text.txt").write_text(transcript["text"], encoding="utf-8")
    return path


def load_transcript(job_dir: Path) -> Optional[dict]:
    path = job_dir / "transcript.json"
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def align_slides(
    narrations: List[str],
    segments: List[dict],
    duration: float = 0.0,
) -> List[Tuple[float, float]]:
    """Map each slide's narration to a [start, end] range on the original audio.

    Uses sequential fuzzy matching against Whisper segments, falling back to
    word-weighted proportional split across the talk.
    """
    n = len(narrations)
    if n == 0:
        return []
    if not segments:
        # Equal split across duration
        dur = duration or 8.0 * n
        step = dur / n
        return [(i * step, (i + 1) * step) for i in range(n)]

    talk_start = float(segments[0]["start"])
    talk_end = float(segments[-1]["end"])
    if duration and duration > talk_end:
        talk_end = float(duration)

    # Try sequential text alignment first
    ranges = _align_by_matching(narrations, segments, talk_start, talk_end)
    if ranges and _ranges_sane(ranges, talk_start, talk_end):
        return ranges
    return _align_proportional(narrations, talk_start, talk_end)


def slice_original(
    source_audio: Path,
    ranges: List[Tuple[float, float]],
    out_dir: Path,
) -> List[Path]:
    """Cut one mp3 clip per range from the original recording."""
    out_dir.mkdir(parents=True, exist_ok=True)
    clips = []
    for i, (start, end) in enumerate(ranges):
        out = out_dir / f"original_{i:02d}.mp3"
        _ffmpeg_slice(source_audio, start, end, out)
        clips.append(out)
    return clips


def _ffmpeg_slice(src: Path, start: float, end: float, out: Path) -> None:
    import imageio_ffmpeg

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    start = max(0.0, float(start))
    end = max(start + 0.4, float(end))  # minimum ~0.4s clip
    # Re-encode to mp3 so PowerPoint embedding is reliable across source codecs.
    result = subprocess.run(
        [ffmpeg, "-y", "-ss", f"{start:.3f}", "-to", f"{end:.3f}",
         "-i", str(src), "-vn", "-acodec", "libmp3lame", "-b:a", "160k",
         "-ar", "44100", "-ac", "2", str(out)],
        capture_output=True, timeout=300,
    )
    if result.returncode != 0 or not out.exists() or out.stat().st_size < 200:
        detail = (result.stderr or b"").decode(errors="replace")[-400:]
        raise AudioSourceError(f"Could not slice audio for a slide: {detail}")


def _norm(s: str) -> str:
    s = s.lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _align_by_matching(
    narrations: List[str],
    segments: List[dict],
    talk_start: float,
    talk_end: float,
) -> Optional[List[Tuple[float, float]]]:
    """Greedy forward match: each narration claims the next best-matching window."""
    seg_texts = [_norm(s["text"]) for s in segments]
    cursor = 0
    ranges: List[Tuple[float, float]] = []
    n_seg = len(segments)
    if n_seg == 0:
        return None

    for i, narr in enumerate(narrations):
        remaining_slides = len(narrations) - i
        target = _norm(narr)

        if cursor >= n_seg:
            # Past the end — carve remaining time evenly
            left = max(0.5, talk_end - (ranges[-1][1] if ranges else talk_start))
            piece = left / remaining_slides
            start = ranges[-1][1] if ranges else talk_start
            end = talk_end if i == len(narrations) - 1 else start + piece
            ranges.append((start, max(start + 0.4, end)))
            continue

        if not target:
            s0 = float(segments[cursor]["start"])
            s1 = min(float(segments[cursor]["end"]), s0 + 1.0)
            ranges.append((s0, s1))
            cursor = min(cursor + 1, n_seg)
            continue

        best_score, best_j, best_k = -1.0, cursor, cursor
        for j in range(cursor, n_seg):
            acc = ""
            for k in range(j, min(n_seg, j + 12)):
                acc = (acc + " " + seg_texts[k]).strip()
                score = SequenceMatcher(None, target[:400], acc[:400]).ratio()
                len_pen = min(len(acc), len(target)) / max(len(acc), len(target), 1)
                score = score * 0.7 + len_pen * 0.3
                if score > best_score:
                    best_score, best_j, best_k = score, j, k
                if len(acc) > len(target) * 1.6:
                    break

        if best_score < 0.22:
            words = max(1, len(target.split()))
            taken = 0
            j = cursor
            k = cursor
            while k < n_seg and taken < words:
                taken += max(1, len(seg_texts[k].split()))
                k += 1
            best_j, best_k = j, max(j, k - 1)

        best_j = min(max(best_j, 0), n_seg - 1)
        best_k = min(max(best_k, best_j), n_seg - 1)
        start = float(segments[best_j]["start"])
        end = float(segments[best_k]["end"])
        if i == len(narrations) - 1:
            end = talk_end
        if end <= start:
            end = start + 0.8
        ranges.append((start, end))
        cursor = best_k + 1

    return ranges


def _align_proportional(
    narrations: List[str],
    talk_start: float,
    talk_end: float,
) -> List[Tuple[float, float]]:
    weights = [max(1, len((n or "").split())) for n in narrations]
    total = sum(weights) or len(weights)
    span = max(0.8, talk_end - talk_start)
    ranges = []
    t = talk_start
    for i, w in enumerate(weights):
        piece = span * (w / total)
        end = talk_end if i == len(weights) - 1 else t + piece
        ranges.append((t, end))
        t = end
    return ranges


def _ranges_sane(ranges: List[Tuple[float, float]], talk_start: float, talk_end: float) -> bool:
    if not ranges:
        return False
    prev = talk_start - 0.01
    for start, end in ranges:
        if end < start or start < talk_start - 1 or end > talk_end + 2:
            return False
        if start + 0.05 < prev:  # large backward jump
            return False
        prev = start
    return True


def _lang_code(language: str) -> str:
    m = {
        "english": "en", "french": "fr", "german": "de",
        "spanish": "es", "italian": "it",
    }
    return m.get((language or "").strip().lower(), "")
