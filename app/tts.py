"""NDS — text-to-speech providers.

Providers implement synthesize(text, voice, out_path) and return the path of the
audio file they wrote. The provider layer is deliberately tiny so ElevenLabs /
Azure can be added later without touching the rest of the app.
"""
import subprocess
from pathlib import Path


class OpenAITTS:
    """Cloud voices via the OpenAI speech API (model gpt-4o-mini-tts, mp3 output)."""

    name = "openai"
    VOICES = [
        "alloy", "ash", "ballad", "coral", "echo",
        "fable", "nova", "onyx", "sage", "shimmer",
        "verse", "marin", "cedar",
    ]

    def __init__(self, api_key: str):
        from openai import OpenAI

        self.client = OpenAI(api_key=api_key)

    def synthesize(self, text: str, voice: str, out_path: Path) -> Path:
        out_path = out_path.with_suffix(".mp3")
        with self.client.audio.speech.with_streaming_response.create(
            model="gpt-4o-mini-tts",
            voice=voice,
            input=text,
            response_format="mp3",
        ) as response:
            response.stream_to_file(str(out_path))
        return out_path


class ElevenLabsTTS:
    """Cloud voices via the ElevenLabs API, including cloned / custom voices.
    Voices are addressed by voice ID (set in Settings or ELEVENLABS_VOICE_ID)."""

    name = "elevenlabs"
    # Built-in voice IDs, always offered when an ElevenLabs key is set; IDs
    # saved in Settings are added after these.
    VOICES = ["GZ4PpFJV8ikEGUtBrjK7", "3C1zYzXNXNzrB66ON8rj", "VwC51uc4PUblWEJSPzeo"]
    API = "https://api.elevenlabs.io/v1"
    MODEL = "eleven_multilingual_v2"

    def __init__(self, api_key: str):
        self.api_key = api_key

    def _request(self, path: str, body: dict | None = None, timeout: int = 180):
        import json
        import urllib.request

        req = urllib.request.Request(
            f"{self.API}{path}",
            data=json.dumps(body).encode() if body is not None else None,
            headers={"xi-api-key": self.api_key, "Content-Type": "application/json"},
            method="POST" if body is not None else "GET",
        )
        return urllib.request.urlopen(req, timeout=timeout)

    def synthesize(self, text: str, voice: str, out_path: Path) -> Path:
        import shutil
        import urllib.error
        from urllib.parse import quote

        out_path = out_path.with_suffix(".mp3")
        try:
            with self._request(
                f"/text-to-speech/{quote(voice, safe='')}?output_format=mp3_44100_128",
                {"text": text, "model_id": self.MODEL},
            ) as r, open(out_path, "wb") as f:
                shutil.copyfileobj(r, f)
        except urllib.error.HTTPError as exc:
            out_path.unlink(missing_ok=True)
            detail = exc.read().decode(errors="replace")[:300]
            raise RuntimeError(f"ElevenLabs TTS failed ({exc.code}): {detail}") from exc
        except urllib.error.URLError as exc:
            out_path.unlink(missing_ok=True)
            raise RuntimeError(f"Could not reach ElevenLabs: {exc.reason}") from exc
        return out_path

    def voice_name(self, voice_id: str) -> str:
        """Display name of a voice ID; falls back to the ID itself."""
        import json
        from urllib.parse import quote

        try:
            with self._request(f"/voices/{quote(voice_id, safe='')}", timeout=10) as r:
                return json.loads(r.read()).get("name") or voice_id
        except Exception:
            return voice_id


class MacSayTTS:
    """Offline fallback using the built-in macOS voice — for demos and testing
    without an OpenAI key. Produces .m4a (AAC), which PowerPoint embeds fine."""

    name = "macos"
    VOICES = ["Samantha", "Daniel", "Karen", "Moira", "Rishi"]

    def synthesize(self, text: str, voice: str, out_path: Path) -> Path:
        aiff = out_path.with_suffix(".aiff")
        m4a = out_path.with_suffix(".m4a")
        subprocess.run(["say", "-v", voice, "-o", str(aiff), text], check=True)
        subprocess.run(
            ["afconvert", "-f", "m4af", "-d", "aac", str(aiff), str(m4a)],
            check=True, capture_output=True,
        )
        aiff.unlink(missing_ok=True)
        return m4a


import re

_TAG_RE = re.compile(r"([A-Z][\w.'\-]{1,15})\s*:\s+")
_SENT_END = ".!?…\"'’”)]"


def _tag_positions(text: str):
    """Speaker-tag candidates: 'Name: ' at the text start, at a line start, or
    inline right after sentence-ending punctuation. The drafter is told to put
    one turn per line but often writes the whole conversation in one paragraph —
    line-anchored parsing alone silently breaks those slides."""
    tags = []
    for m in _TAG_RE.finditer(text):
        j = m.start() - 1
        while j >= 0 and text[j] in " \t":
            j -= 1
        if j < 0 or text[j] == "\n" or text[j] in _SENT_END:
            tags.append(m)
    return tags


def _is_dialogue(text: str, tags) -> bool:
    """Only treat the text as a two-voice script when the tags genuinely look
    like it: known host names (Alex/Sam), a repeating name, or the classic
    one-turn-per-line layout. Guards against prose like 'Remember: …. Second: …'
    being split apart."""
    if len(tags) < 2:
        return False
    names = [m.group(1).lower() for m in tags]
    if len(set(names)) > 3:
        return False
    if set(names) <= {"alex", "sam"}:
        return True
    if len(set(names)) < len(names):        # a speaker takes a second turn
        return True
    # every tag sits at a line start → the original line-based format
    return all(m.start() == 0 or text[: m.start()].rstrip(" \t").endswith("\n")
               for m in tags)


def split_dialogue(text: str):
    """Parse 'Alex: …' / 'Sam: …' narration into [(speaker_slot, text)] where
    slot is 0 for the first distinct speaker, 1 for the second. Handles both
    one-turn-per-line and everything-in-one-paragraph formats. Non-dialogue
    text comes back as a single slot-0 turn, untouched."""
    tags = _tag_positions(text)
    if not _is_dialogue(text, tags):
        clean = " ".join(text.split())
        return [(0, clean)] if clean else []
    turns, order = [], []
    lead = text[: tags[0].start()].strip()
    if lead:
        turns.append((0, " ".join(lead.split())))
    for i, m in enumerate(tags):
        name = m.group(1).lower()
        if name not in order:
            order.append(name)
        slot = min(order.index(name), 1)
        end = tags[i + 1].start() if i + 1 < len(tags) else len(text)
        seg = " ".join(text[m.end(): end].split())
        if seg:
            turns.append((slot, seg))
    # merge consecutive same-speaker turns to cut per-call overhead
    merged = []
    for slot, t in turns:
        if merged and merged[-1][0] == slot:
            merged[-1] = (slot, merged[-1][1] + " " + t)
        else:
            merged.append((slot, t))
    return merged


def synthesize_narration(engine, text: str, voices, out_base: Path, style: str = "single") -> Path:
    """Voice one slide's narration. voices is (voice_a, voice_b); in 'single'
    style everything uses voice_a. Conversation output is stitched to one .m4a."""
    voice_a, voice_b = voices
    turns = split_dialogue(text) if style == "conversation" else []
    if style != "conversation" or len({s for s, _ in turns}) < 2:
        clean = " ".join(t for _, t in turns) if turns else text
        return engine.synthesize(clean, voice_a, out_base)

    seg_paths = []
    for i, (slot, turn_text) in enumerate(turns):
        seg = engine.synthesize(turn_text, voice_a if slot == 0 else voice_b,
                                out_base.parent / f"{out_base.name}_seg{i:02d}")
        seg_paths.append(seg)
    out = _concat_audio(seg_paths, out_base.with_suffix(".m4a"))
    for p in seg_paths:
        p.unlink(missing_ok=True)
    return out


def _concat_audio(parts, out_path: Path) -> Path:
    import imageio_ffmpeg

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    list_file = out_path.with_suffix(".txt")
    list_file.write_text("".join(f"file '{p.name}'\n" for p in parts))
    result = subprocess.run(
        [ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", str(list_file),
         "-c:a", "aac", "-b:a", "160k", "-ar", "44100", "-ac", "2", str(out_path)],
        capture_output=True, timeout=300, cwd=str(out_path.parent),
    )
    list_file.unlink(missing_ok=True)
    if result.returncode != 0:
        raise RuntimeError(f"Audio stitching failed: {result.stderr.decode()[-300:]}")
    return out_path


def audio_duration_seconds(path: Path) -> float:
    """Length of an mp3/m4a file, used to time slide auto-advance."""
    import mutagen

    f = mutagen.File(str(path))
    if f is None or f.info is None:
        return 10.0
    return float(f.info.length)
