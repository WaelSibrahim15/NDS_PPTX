"""NDS — Narrated Deck Studio, local web app (Phase 2).

Flow: upload → draft (Claude) → REVIEW & AMEND in the browser → build
(voice only what changed → assemble PPTX).

Run with:  uvicorn app.main:app --port 8765
On Railway:  uvicorn app.main:app --host 0.0.0.0 --port $PORT
"""
import hashlib
import json
import os
import platform
import secrets
import shutil
import threading
import uuid
from pathlib import Path

from io import BytesIO

from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import PlainTextResponse

from . import extract, tts, audio_source
from .drafter import (draft_deck, draft_narration_for_existing,
                      regenerate_slide)
from .models import DeckPlan, Slide
from .pptx_builder import build_deck, narrate_existing_pptx
from .renderer import render_slide_preview
from .video import export_mp4


class JobCancelled(Exception):
    """Raised when the user cancels a drafting/building/rendering job."""

ROOT = Path(__file__).resolve().parent.parent
JOBS_DIR = Path(os.environ.get("JOBS_DIR", str(ROOT / "jobs"))).expanduser()
SAMPLES_DIR = ROOT / "voice_samples"
CONFIG_PATH = ROOT / "config.json"
HISTORY_LIMIT = 4  # only keep / show the most recent decks
ENV_KEY_MAP = {
    "anthropic_api_key": "ANTHROPIC_API_KEY",
    "openai_api_key": "OPENAI_API_KEY",
    "elevenlabs_api_key": "ELEVENLABS_API_KEY",
    "elevenlabs_voice_ids": "ELEVENLABS_VOICE_ID",
}

AUDIO_GUIDANCE = (
    "SOURCE IS A SPOKEN RECORDING TRANSCRIPT. Keep slide order chronological "
    "with the talk. Write each slide's narration so it closely follows what the "
    "speaker said in that section (paraphrase lightly for clarity, do not invent "
    "new topics). The original audio will be split onto the slides — do not write "
    "stage directions about audio."
)

SAMPLE_TEXT = {
    "openai": "Hello! This is how I sound. I will be the narrator of your presentation.",
    "macos": "Hello! This is how I sound. I will be the narrator of your presentation.",
    "elevenlabs": "Hello! This is how I sound. I will be the narrator of your presentation.",
}

app = FastAPI(title="NDS — Narrated Deck Studio")

_jobs: dict = {}  # job_id -> live state (mirrored to jobs/<id>/state.json)
_lock = threading.Lock()

JOBS_DIR.mkdir(parents=True, exist_ok=True)


def _macos_tts_available() -> bool:
    if os.environ.get("DISABLE_MACOS_TTS", "").strip() in ("1", "true", "yes"):
        return False
    return platform.system() == "Darwin"


def _env_managed_keys() -> list:
    """Config keys whose value comes from an environment variable. Only these
    are locked in the Settings form; the rest stay editable."""
    return [key for key, name in ENV_KEY_MAP.items() if os.environ.get(name, "").strip()]


def _app_password() -> str:
    return os.environ.get("APP_PASSWORD", "").strip()


class BasicAuthMiddleware(BaseHTTPMiddleware):
    """Gate the whole app when APP_PASSWORD is set (Railway / public URL)."""

    async def dispatch(self, request: Request, call_next):
        password = _app_password()
        if not password or request.url.path == "/health":
            return await call_next(request)
        auth = request.headers.get("Authorization", "")
        ok = False
        if auth.startswith("Basic "):
            import base64
            try:
                decoded = base64.b64decode(auth.split(" ", 1)[1]).decode("utf-8")
                user, _, pwd = decoded.partition(":")
                ok = secrets.compare_digest(pwd, password) and user in ("", "nds")
            except Exception:
                ok = False
        if not ok:
            return PlainTextResponse(
                "Authentication required",
                status_code=401,
                headers={"WWW-Authenticate": 'Basic realm="NDS"'},
            )
        return await call_next(request)


if _app_password():
    app.add_middleware(BasicAuthMiddleware)


# ------------------------------------------------------------------ settings

def load_config() -> dict:
    cfg = json.loads(CONFIG_PATH.read_text()) if CONFIG_PATH.exists() else {}
    # Local Settings UI writes config.json. On Railway / production, set these
    # env vars instead — they override the file when present.
    for key, env_name in ENV_KEY_MAP.items():
        env_val = os.environ.get(env_name, "").strip()
        if env_val:
            cfg[key] = env_val
        else:
            cfg.setdefault(key, "")
    return {k: v for k, v in cfg.items() if v}


def save_config(cfg: dict):
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2))
    CONFIG_PATH.chmod(0o600)  # keys stay private to this user account


@app.get("/health")
def health():
    return {"ok": True}


_voice_names: dict = {}  # ElevenLabs voice ID -> display name


def _elevenlabs_voice_ids(cfg: dict) -> list:
    raw = cfg.get("elevenlabs_voice_ids", "")
    ids = []
    for part in raw.replace("\n", ",").replace(";", ",").split(","):
        vid = part.strip().strip("\"'").strip()  # tolerate quotes pasted into the value
        if vid and vid not in ids:
            ids.append(vid)
    return ids


@app.get("/api/config")
def get_config():
    cfg = load_config()
    voices = {"openai": tts.OpenAITTS.VOICES}
    voice_labels = {}
    el_ids = _elevenlabs_voice_ids(cfg)
    if cfg.get("elevenlabs_api_key") and el_ids:
        voices["elevenlabs"] = el_ids
        for vid in el_ids:
            if vid not in _voice_names:
                _voice_names[vid] = tts.ElevenLabsTTS(cfg["elevenlabs_api_key"]).voice_name(vid)
            voice_labels[vid] = _voice_names[vid]
    if _macos_tts_available():
        voices["macos"] = tts.MacSayTTS.VOICES
    return {
        "anthropic_key_set": bool(cfg.get("anthropic_api_key")),
        "openai_key_set": bool(cfg.get("openai_api_key")),
        "elevenlabs_key_set": bool(cfg.get("elevenlabs_api_key")),
        "elevenlabs_voice_ids": ", ".join(el_ids),
        "voices": voices,
        "voice_labels": voice_labels,
        "env_keys": _env_managed_keys(),
        "settings_locked": set(_env_managed_keys()) == set(ENV_KEY_MAP),
        "macos_tts": _macos_tts_available(),
    }


@app.post("/api/settings")
def set_settings(anthropic_api_key: str = Form(""), openai_api_key: str = Form(""),
                 elevenlabs_api_key: str = Form(""), elevenlabs_voice_ids: str = Form("")):
    submitted = {
        "anthropic_api_key": anthropic_api_key.strip(),
        "openai_api_key": openai_api_key.strip(),
        "elevenlabs_api_key": elevenlabs_api_key.strip(),
        "elevenlabs_voice_ids": elevenlabs_voice_ids.strip(),
    }
    submitted = {k: v for k, v in submitted.items() if v}
    env_keys = _env_managed_keys()
    blocked = [k for k in submitted if k in env_keys]
    if blocked and len(blocked) == len(submitted):
        raise HTTPException(
            403,
            "These keys are managed by server environment variables — edit them there, not here.",
        )
    cfg = json.loads(CONFIG_PATH.read_text()) if CONFIG_PATH.exists() else {}
    cfg.update({k: v for k, v in submitted.items() if k not in env_keys})
    save_config(cfg)
    return {"ok": True}


# ------------------------------------------------------------ job persistence

def _job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def _persist(job_id: str):
    state = {k: v for k, v in _jobs[job_id].items() if k != "deck"}
    (_job_dir(job_id) / "state.json").write_text(json.dumps(state, indent=2))


def _get_job(job_id: str) -> dict:
    with _lock:
        if job_id in _jobs:
            return _jobs[job_id]
        state_path = _job_dir(job_id) / "state.json"
        if state_path.exists():  # survive a server restart
            job = json.loads(state_path.read_text())
            deck_path = _job_dir(job_id) / "deck.json"
            job["deck"] = json.loads(deck_path.read_text()) if deck_path.exists() else None
            _jobs[job_id] = job
            return job
    raise HTTPException(404, "Unknown job")


def _load_plan(job_id: str) -> DeckPlan:
    deck_path = _job_dir(job_id) / "deck.json"
    if not deck_path.exists():
        raise HTTPException(400, "This job has no draft yet.")
    return DeckPlan.model_validate_json(deck_path.read_text())


def _save_plan(job_id: str, plan: DeckPlan):
    (_job_dir(job_id) / "deck.json").write_text(plan.model_dump_json(indent=2))
    _jobs[job_id]["deck"] = json.loads(plan.model_dump_json())


def _engine(provider: str, cfg: dict):
    if provider == "macos":
        if not _macos_tts_available():
            raise RuntimeError("macOS offline voice is not available on this host — use OpenAI.")
        return tts.MacSayTTS()
    if provider == "elevenlabs":
        if not cfg.get("elevenlabs_api_key"):
            raise RuntimeError("No ElevenLabs API key saved — add it under Settings.")
        return tts.ElevenLabsTTS(api_key=cfg["elevenlabs_api_key"])
    if not cfg.get("openai_api_key"):
        raise RuntimeError("No OpenAI API key saved — add it under Settings, or pick the offline macOS voice.")
    return tts.OpenAITTS(api_key=cfg["openai_api_key"])


def _audio_for(job_id: str, narration: str, opts: dict, cfg: dict) -> Path:
    """Synthesize (or reuse cached) audio for one narration text."""
    provider = opts["provider"]
    voice = opts["voice"]
    voice2 = opts.get("voice2") or voice
    style = opts.get("narration_style", "single")
    cache = _job_dir(job_id) / "audio"
    cache.mkdir(exist_ok=True)
    key = hashlib.sha1(f"{provider}|{voice}|{voice2}|{style}|{narration}".encode()).hexdigest()[:16]
    for suffix in (".mp3", ".m4a"):
        hit = cache / (key + suffix)
        if hit.exists():
            return hit
    return tts.synthesize_narration(_engine(provider, cfg), narration,
                                    (voice, voice2), cache / key, style=style)


def _slide_secs(opts: dict) -> float:
    """Rough per-slide voicing time for the ETA display."""
    base = 1.5 if opts.get("provider") == "macos" else 4.0
    return base * (1.7 if opts.get("narration_style") == "conversation" else 1.0)


def _tts_key_error(provider: str, cfg: dict) -> str | None:
    if provider == "openai" and not cfg.get("openai_api_key"):
        return "No OpenAI API key saved — add it under Settings, or pick the offline macOS voice."
    if provider == "elevenlabs" and not cfg.get("elevenlabs_api_key"):
        return "No ElevenLabs API key saved — add it under Settings."
    return None


def _check_cancel(job_id: str):
    job = _jobs.get(job_id)
    if job and job.get("cancel_requested"):
        raise JobCancelled()


# ------------------------------------------------------------------ drafting

@app.post("/api/draft")
def start_draft(
    file: UploadFile | None = File(None),
    source_text: str = Form(""),
    language: str = Form("English"),
    guidance: str = Form(""),
    template: str = Form("niq"),
    provider: str = Form("openai"),
    voice: str = Form("nova"),
    voice2: str = Form(""),
    narration_style: str = Form("single"),
    mode: str = Form("generate"),
    design_notes: str = Form(""),
):
    cfg = load_config()
    if not cfg.get("anthropic_api_key"):
        raise HTTPException(400, "No Anthropic API key saved yet — add it under Settings.")

    # Enhance mode retired — same outcome as narrate (scripts from slide content).
    if mode == "enhance":
        mode = "narrate"
    if mode not in ("generate", "narrate"):
        raise HTTPException(400, f"Unknown mode '{mode}' — choose NDS design or narration.")

    pasted = (source_text or "").strip()
    has_file = bool(file and file.filename)
    suffix = Path(file.filename).suffix.lower() if has_file else ""
    is_audio = has_file and suffix in audio_source.AUDIO_SUFFIXES

    if mode == "narrate":
        if not has_file or suffix != ".pptx":
            raise HTTPException(400, "Create narration for each slide needs a .pptx file — upload the PowerPoint itself.")
    elif is_audio:
        if mode != "generate":
            raise HTTPException(400, "Audio upload works with NDS design.")
        if not cfg.get("openai_api_key"):
            raise HTTPException(400, "Audio sources need an OpenAI API key (Whisper transcription) — add it under Settings.")
    elif not has_file and not pasted:
        raise HTTPException(400, "Upload a document, audio file, or paste source text.")

    job_id = uuid.uuid4().hex[:12]
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)

    # Prefer an uploaded file when both are present; otherwise write pasted text as .txt.
    if has_file:
        src_path = job_dir / ("source" + (suffix or ".txt"))
        with src_path.open("wb") as f:
            shutil.copyfileobj(file.file, f)
        source_name = file.filename
    else:
        src_path = job_dir / "source.txt"
        src_path.write_text(pasted, encoding="utf-8")
        source_name = "pasted-text.txt"

    _jobs[job_id] = {
        "status": "queued", "step": "Queued", "progress": 0, "error": None, "deck": None,
        "options": {"language": language, "guidance": guidance, "template": template,
                    "provider": provider, "voice": voice, "voice2": voice2,
                    "narration_style": narration_style, "mode": mode,
                    "design_notes": design_notes.strip(),
                    "from_audio": is_audio},
        "source_name": source_name,
    }
    _persist(job_id)
    threading.Thread(target=_run_draft, args=(job_id, src_path, cfg), daemon=True).start()
    _prune_old_jobs()
    return {"job_id": job_id}


def _run_draft(job_id: str, src_path: Path, cfg: dict):
    job = _jobs[job_id]
    job_dir = src_path.parent
    opts = job["options"]
    # User's design requirements ride along with the guidance for whichever
    # engine designs the deck (NDS drafter).
    guidance = opts.get("guidance", "")
    if opts.get("design_notes"):
        guidance = (guidance + "\nDesign requirements: " + opts["design_notes"]).strip()
    try:
        _check_cancel(job_id)
        if opts.get("mode") == "enhance":
            opts["mode"] = "narrate"
        job.update(status="drafting", step="Reading the document…", progress=10, eta_seconds=45)

        # Audio source → Whisper transcript, then continue as text (keep original file for slicing).
        if audio_source.is_audio(src_path):
            if not cfg.get("openai_api_key"):
                raise RuntimeError("OpenAI API key required to transcribe audio.")
            job.update(step="Transcribing the recording (Whisper)…", progress=12, eta_seconds=90)
            transcript = audio_source.transcribe(
                src_path, cfg["openai_api_key"], language=opts.get("language", ""))
            audio_source.save_transcript(job_dir, transcript)
            opts["from_audio"] = True
            opts["audio_file"] = src_path.name
            # Draft from a .txt sibling so extract_text sees plain text.
            text_path = job_dir / "source_from_audio.txt"
            text_path.write_text(transcript["text"], encoding="utf-8")
            src_path = text_path
            guidance = ((guidance + "\n" if guidance else "") + AUDIO_GUIDANCE).strip()
            job.update(step="Transcript ready — drafting slides…", progress=22, eta_seconds=40)

        _check_cancel(job_id)
        plan = _produce_plan(job_id, src_path, cfg, guidance)
        _check_cancel(job_id)
        _save_plan(job_id, plan)
        job.update(status="review", step="Draft ready — review and amend below",
                   progress=100, eta_seconds=None, cancel_requested=False)
    except JobCancelled:
        job.update(status="cancelled", step="Cancelled", progress=0, error="Cancelled by user",
                   eta_seconds=None, cancel_requested=False)
    except Exception as exc:
        job.update(status="error", step="Failed", error=str(exc))
    _persist(job_id)


def _produce_plan(job_id: str, src_path: Path, cfg: dict, guidance: str) -> DeckPlan:
    """Run the Claude drafting step for a job: narration scripts for an uploaded
    deck (narrate mode) or a full deck plan from the extracted text (generate mode)."""
    job = _jobs[job_id]
    job_dir = _job_dir(job_id)
    opts = job["options"]
    if opts.get("mode") == "narrate":
        slides = extract.extract_slides(src_path)
        (job_dir / "source_text.txt").write_text(
            "\n\n".join(f"--- Slide {s['index']} ---\n" + "\n".join(s["texts"]) for s in slides)
        )
        job.update(step="Writing narration scripts from each slide…", progress=30,
                   eta_seconds=20 + 2 * len(slides))
        narrations = draft_narration_for_existing(
            slides, api_key=cfg["anthropic_api_key"], language=opts["language"],
            guidance=guidance or opts.get("guidance", ""),
            narration_style=opts.get("narration_style", "single"),
        )
        first_title = (slides[0]["texts"][0].splitlines()[0][:60]
                       if slides and slides[0]["texts"] else "")
        return DeckPlan(
            deck_title=first_title or Path(job.get("source_name") or "Narrated deck").stem,
            subtitle="Narration scripts from slide content",
            slides=[
                Slide(layout="content",
                      title=(s["texts"][0].splitlines()[0][:80] if s["texts"] else f"Slide {s['index']}"),
                      bullets=[t[:120] for t in s["texts"][1:5]],
                      narration=narrations[i])
                for i, s in enumerate(slides)
            ],
        )
    text = extract.extract_text(src_path)
    (job_dir / "source_text.txt").write_text(text)
    job.update(step="Drafting slides and narration with Claude…", progress=30, eta_seconds=35)
    return draft_deck(text, api_key=cfg["anthropic_api_key"],
                      language=opts["language"], guidance=guidance,
                      narration_style=opts.get("narration_style", "single"))


def _redraft_source(job_id: str) -> Path:
    """The file the drafter should read again for a whole-deck redraft."""
    job_dir = _job_dir(job_id)
    opts = _jobs[job_id]["options"]
    if opts.get("from_audio"):
        return job_dir / "source_from_audio.txt"
    if opts.get("mode") == "narrate":
        return job_dir / "source.pptx"
    for p in sorted(job_dir.iterdir()):
        if p.name.startswith("source.") and p.is_file():
            return p
    raise HTTPException(400, "This job's source file is missing — start a new draft instead.")


@app.post("/api/jobs/{job_id}/redraft")
def start_redraft(job_id: str, payload: dict = Body(default={})):
    """Throw the current draft away and ask Claude for a fresh one from the same
    source. The previous deck is kept in deck_previous.json for a one-step undo."""
    job = _get_job(job_id)
    if job["status"] in ("drafting", "building", "rendering", "queued"):
        raise HTTPException(400, "Job is busy.")
    cfg = load_config()
    if not cfg.get("anthropic_api_key"):
        raise HTTPException(400, "No Anthropic API key saved.")
    src_path = _redraft_source(job_id)
    if not src_path.exists():
        raise HTTPException(400, "This job's source file is missing — start a new draft instead.")
    instruction = (payload.get("instruction") or "").strip()
    job.update(status="drafting", step="Redrafting the whole deck with Claude…",
               progress=5, error=None, eta_seconds=40, cancel_requested=False)
    _persist(job_id)
    threading.Thread(target=_run_redraft, args=(job_id, src_path, cfg, instruction),
                     daemon=True).start()
    return {"ok": True}


def _run_redraft(job_id: str, src_path: Path, cfg: dict, instruction: str):
    job = _jobs[job_id]
    job_dir = _job_dir(job_id)
    opts = job["options"]
    guidance = opts.get("guidance", "")
    if opts.get("design_notes"):
        guidance = (guidance + "\nDesign requirements: " + opts["design_notes"]).strip()
    if opts.get("from_audio"):
        guidance = ((guidance + "\n" if guidance else "") + AUDIO_GUIDANCE).strip()
    if instruction:
        guidance = ((guidance + "\n" if guidance else "")
                    + "REDRAFT INSTRUCTION (takes priority over everything above): "
                    + instruction).strip()
    try:
        _check_cancel(job_id)
        plan = _produce_plan(job_id, src_path, cfg, guidance)
        _check_cancel(job_id)
        deck_path = job_dir / "deck.json"
        if deck_path.exists():
            shutil.copyfile(deck_path, job_dir / "deck_previous.json")
        _save_plan(job_id, plan)
        job.update(status="review", step="Fresh draft ready — review and amend below",
                   progress=100, eta_seconds=None, cancel_requested=False,
                   can_undo_redraft=True)
    except JobCancelled:
        job.update(status="review", step="Redraft cancelled — your previous draft is unchanged",
                   progress=100, error=None, eta_seconds=None, cancel_requested=False)
    except Exception as exc:
        job.update(status="review", step="Redraft failed — your previous draft is unchanged",
                   progress=100, error=str(exc), eta_seconds=None, cancel_requested=False)
    _persist(job_id)


@app.post("/api/jobs/{job_id}/redraft/undo")
def undo_redraft(job_id: str):
    """Restore the deck as it was before the last whole-deck redraft."""
    job = _get_job(job_id)
    if job["status"] in ("drafting", "building", "rendering", "queued"):
        raise HTTPException(400, "Job is busy.")
    prev = _job_dir(job_id) / "deck_previous.json"
    if not prev.exists():
        raise HTTPException(400, "Nothing to undo.")
    plan = DeckPlan.model_validate_json(prev.read_text())
    _save_plan(job_id, plan)
    prev.unlink()
    job.update(status="review", step="Previous draft restored", progress=100,
               error=None, can_undo_redraft=False)
    _persist(job_id)
    return json.loads(plan.model_dump_json())


# ------------------------------------------------------------ review & amend

@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    return _get_job(job_id)


@app.put("/api/jobs/{job_id}/deck")
def save_deck(job_id: str, deck: dict = Body(...)):
    job = _get_job(job_id)
    try:
        plan = DeckPlan.model_validate(deck)
    except Exception as exc:
        raise HTTPException(422, f"Invalid deck: {exc}")
    _save_plan(job_id, plan)
    # Narration edits invalidate original-audio slices — rebuild on next preview/build.
    if job.get("options", {}).get("from_audio"):
        clips = _job_dir(job_id) / "audio_original"
        if clips.exists():
            shutil.rmtree(clips, ignore_errors=True)
    _persist(job_id)
    return {"ok": True, "slides": len(plan.slides)}


@app.post("/api/jobs/{job_id}/slides/{index}/regenerate")
def regen_slide(job_id: str, index: int, payload: dict = Body(default={})):
    job = _get_job(job_id)
    cfg = load_config()
    if not cfg.get("anthropic_api_key"):
        raise HTTPException(400, "No Anthropic API key saved.")
    plan = _load_plan(job_id)
    if not 0 <= index < len(plan.slides):
        raise HTTPException(400, "No such slide.")
    source_text = (_job_dir(job_id) / "source_text.txt").read_text()
    opts = job["options"]
    new_slide = regenerate_slide(
        source_text, plan, index, payload.get("instruction", ""),
        api_key=cfg["anthropic_api_key"], language=opts["language"],
        narration_only=opts.get("mode") == "narrate",
        narration_style=opts.get("narration_style", "single"),
    )
    if opts.get("mode") == "narrate":  # design locked — only narration may change
        old = plan.slides[index]
        old.narration = new_slide.narration
        new_slide = old
    plan.slides[index] = new_slide
    _save_plan(job_id, plan)
    return json.loads(new_slide.model_dump_json())


@app.get("/api/jobs/{job_id}/audio/{index}")
def preview_audio(job_id: str, index: int, provider: str = "openai", voice: str = "nova",
                  voice2: str = "", narration_style: str = "single"):
    job = _get_job(job_id)
    plan = _load_plan(job_id)
    if not 0 <= index < len(plan.slides):
        raise HTTPException(400, "No such slide.")
    cfg = load_config()
    opts = dict(job["options"], provider=provider, voice=voice,
                voice2=voice2 or voice, narration_style=narration_style)
    try:
        if opts.get("from_audio"):
            path = _original_clip_for(job_id, plan, index)
        else:
            path = _audio_for(job_id, plan.slides[index].narration, opts, cfg)
    except RuntimeError as exc:
        raise HTTPException(400, str(exc))
    return FileResponse(path, media_type="audio/mpeg" if path.suffix == ".mp3" else "audio/mp4")


def _original_clip_for(job_id: str, plan: DeckPlan, index: int) -> Path:
    """Ensure a sliced original-audio clip exists for slide ``index`` and return it."""
    job = _jobs[job_id]
    job_dir = _job_dir(job_id)
    opts = job["options"]
    clips_dir = job_dir / "audio_original"
    clip = clips_dir / f"original_{index:02d}.mp3"
    if clip.exists() and clip.stat().st_size > 200:
        return clip
    audio_name = opts.get("audio_file") or ""
    src_audio = job_dir / audio_name if audio_name else None
    if not src_audio or not src_audio.exists():
        src_audio = next(
            (p for p in job_dir.iterdir()
             if p.name.startswith("source") and audio_source.is_audio(p)),
            None,
        )
    if not src_audio:
        raise RuntimeError("Original audio file is missing from this job.")
    transcript = audio_source.load_transcript(job_dir) or {"segments": [], "duration": 0}
    ranges = audio_source.align_slides(
        [s.narration for s in plan.slides],
        transcript.get("segments") or [],
        float(transcript.get("duration") or 0),
    )
    clips = audio_source.slice_original(src_audio, ranges, clips_dir)
    return clips[index]


@app.get("/api/voices/{provider}/{voice}/sample")
def voice_sample(provider: str, voice: str):
    cfg = load_config()
    SAMPLES_DIR.mkdir(exist_ok=True)
    base = SAMPLES_DIR / f"{provider}_{voice}"
    for suffix in (".mp3", ".m4a"):
        hit = base.with_suffix(suffix)
        if hit.exists():
            return FileResponse(hit, media_type="audio/mpeg" if suffix == ".mp3" else "audio/mp4")
    try:
        path = _engine(provider, cfg).synthesize(SAMPLE_TEXT.get(provider, SAMPLE_TEXT["openai"]), voice, base)
    except RuntimeError as exc:
        raise HTTPException(400, str(exc))
    return FileResponse(path, media_type="audio/mpeg" if path.suffix == ".mp3" else "audio/mp4")


# ------------------------------------------------------------------ building

@app.post("/api/jobs/{job_id}/build")
def start_build(job_id: str, payload: dict = Body(default={})):
    job = _get_job(job_id)
    if job["status"] in ("drafting", "building", "rendering", "queued"):
        raise HTTPException(400, "Job is busy.")
    cfg = load_config()
    opts = job["options"]
    for k in ("provider", "voice", "voice2", "template", "narration_style"):
        if payload.get(k) is not None and payload.get(k) != "":
            opts[k] = payload[k]
    with_voice = bool(payload.get("with_voice", True))
    opts["with_voice"] = with_voice
    needs_tts = with_voice and not opts.get("from_audio")
    if needs_tts and (err := _tts_key_error(opts["provider"], cfg)):
        raise HTTPException(400, err)
    job.update(status="building",
               step=("Slicing original audio…" if (with_voice and opts.get("from_audio"))
                     else "Preparing narration…" if with_voice else "Preparing PPTX draft…"),
               progress=0, error=None)
    _persist(job_id)
    threading.Thread(target=_run_build, args=(job_id, cfg), daemon=True).start()
    return {"ok": True}


def _run_build(job_id: str, cfg: dict):
    job = _jobs[job_id]
    opts = job["options"]
    with_voice = opts.get("with_voice", True)
    try:
        plan = _load_plan(job_id)
        n = len(plan.slides)
        per = _slide_secs(opts)
        audio_files = []
        if with_voice and opts.get("from_audio"):
            job_dir = _job_dir(job_id)
            audio_name = opts.get("audio_file") or ""
            # Prefer stored name; else first matching source.* audio in job dir.
            src_audio = job_dir / audio_name if audio_name else None
            if not src_audio or not src_audio.exists():
                src_audio = next(
                    (p for p in job_dir.iterdir()
                     if p.name.startswith("source") and audio_source.is_audio(p)),
                    None,
                )
            if not src_audio or not src_audio.exists():
                raise RuntimeError("Original audio file is missing from this job.")
            transcript = audio_source.load_transcript(job_dir) or {"segments": [], "duration": 0}
            job.update(step="Aligning original audio to slides…", progress=15, eta_seconds=20)
            ranges = audio_source.align_slides(
                [s.narration for s in plan.slides],
                transcript.get("segments") or [],
                float(transcript.get("duration") or 0),
            )
            (job_dir / "audio_alignment.json").write_text(
                json.dumps([{"start": a, "end": b} for a, b in ranges], indent=2))
            clips_dir = job_dir / "audio_original"
            for i in range(n):
                _check_cancel(job_id)
                job.update(step=f"Slicing original audio for slide {i + 1} of {n}…",
                           progress=20 + int(60 * i / max(n, 1)),
                           eta_seconds=int((n - i) * 1.5) + 6)
            audio_files = audio_source.slice_original(src_audio, ranges, clips_dir)
        elif with_voice:
            for i, slide in enumerate(plan.slides):
                _check_cancel(job_id)
                job.update(step=f"Voicing slide {i + 1} of {n}…", progress=int(80 * i / n),
                           eta_seconds=int((n - i) * per) + 6)
                audio_files.append(_audio_for(job_id, slide.narration, opts, cfg))
        else:
            audio_files = [None] * n
            job.update(step="Assembling the PowerPoint draft (no voice)…", progress=40,
                       eta_seconds=8)

        safe_title = "".join(ch for ch in plan.deck_title if ch.isalnum() or ch in " -_")[:60].strip() or "NDS Deck"
        enhancing = False  # Enhance/polish mode retired
        narrate = opts.get("mode") == "narrate"
        src_deck = _job_dir(job_id) / "source.pptx"

        _check_cancel(job_id)
        job.update(step="Assembling the PowerPoint…", progress=90, eta_seconds=6)
        if with_voice:
            if opts.get("from_audio"):
                base_name = f"{safe_title} (original audio).pptx"
            else:
                base_name = f"{safe_title} (narrated).pptx" if narrate else f"{safe_title}.pptx"
        else:
            base_name = f"{safe_title} (draft).pptx"
        if narrate:
            out = narrate_existing_pptx(
                src_deck, [s.narration for s in plan.slides],
                audio_files, _job_dir(job_id) / base_name,
                animate=False)
        else:
            out = build_deck(plan, audio_files, _job_dir(job_id) / base_name,
                             template=opts["template"])

        job.update(status="done",
                   step="Done — download below" if with_voice else "Draft ready — download below (or add narration)",
                   progress=100, file=out.name, extra_files=[], eta_seconds=None,
                   cancel_requested=False, last_build_voiced=with_voice)
    except JobCancelled:
        job.update(status="cancelled", step="Cancelled", progress=0, error="Cancelled by user",
                   eta_seconds=None, cancel_requested=False)
    except Exception as exc:
        job.update(status="error", step="Failed", error=str(exc))
    _persist(job_id)


@app.post("/api/jobs/{job_id}/export-video")
def start_video(job_id: str, payload: dict = Body(default={})):
    job = _get_job(job_id)
    if job["status"] in ("drafting", "building", "rendering", "queued"):
        raise HTTPException(400, "Job is busy.")
    if job["options"].get("mode") == "narrate":
        raise HTTPException(400, "MP4 export isn't available for narrated existing decks yet — "
                                 "PowerPoint itself can export those (File → Export → Create a Video).")
    cfg = load_config()
    opts = job["options"]
    for k in ("provider", "voice", "voice2", "template", "narration_style"):
        if payload.get(k) is not None and payload.get(k) != "":
            opts[k] = payload[k]
    if err := _tts_key_error(opts["provider"], cfg):
        raise HTTPException(400, err)
    job.update(status="rendering", step="Preparing video…", progress=0, error=None)
    _persist(job_id)
    threading.Thread(target=_run_video, args=(job_id, cfg), daemon=True).start()
    return {"ok": True}


def _run_video(job_id: str, cfg: dict):
    job = _jobs[job_id]
    opts = job["options"]
    try:
        plan = _load_plan(job_id)
        n = len(plan.slides)
        per = _slide_secs(opts)
        video_est = n * 2 + 12
        audio_files = []
        for i, slide in enumerate(plan.slides):
            _check_cancel(job_id)
            job.update(step=f"Voicing slide {i + 1} of {n}…", progress=int(30 * i / n),
                       eta_seconds=int((n - i) * per) + video_est)
            audio_files.append(_audio_for(job_id, slide.narration, opts, cfg))

        def on_progress(pct, msg):
            _check_cancel(job_id)
            job.update(step=msg, progress=30 + int(pct * 0.7),
                       eta_seconds=max(2, int(video_est * (1 - pct / 100))))

        safe_title = "".join(ch for ch in plan.deck_title if ch.isalnum() or ch in " -_")[:60].strip() or "NDS Deck"
        out = export_mp4(plan, audio_files, _job_dir(job_id) / f"{safe_title}.mp4",
                         template=opts["template"], progress=on_progress)
        job.update(status="done", step="Video ready — download below", progress=100, video=out.name,
                   cancel_requested=False)
    except JobCancelled:
        job.update(status="cancelled", step="Cancelled", progress=0, error="Cancelled by user",
                   eta_seconds=None, cancel_requested=False)
    except Exception as exc:
        job.update(status="error", step="Failed", error=str(exc))
    _persist(job_id)


@app.get("/api/jobs/{job_id}/video")
def job_video(job_id: str):
    job = _get_job(job_id)
    if not job.get("video"):
        raise HTTPException(404, "Video not ready")
    path = _job_dir(job_id) / job["video"]
    return FileResponse(path, filename=path.name, media_type="video/mp4")


@app.get("/api/jobs")
def job_history():
    _prune_old_jobs()
    return _list_job_summaries()[:HISTORY_LIMIT]


def _list_job_summaries() -> list:
    items = []
    if not JOBS_DIR.exists():
        return items
    for state_path in JOBS_DIR.glob("*/state.json"):
        try:
            state = json.loads(state_path.read_text())
        except Exception:
            continue
        deck_path = state_path.parent / "deck.json"
        title = None
        if deck_path.exists():
            try:
                title = json.loads(deck_path.read_text()).get("deck_title")
            except Exception:
                pass
        opts = state.get("options") or {}
        mode = opts.get("mode")
        items.append({
            "id": state_path.parent.name,
            "title": title or state.get("source_name") or state_path.parent.name,
            "status": state.get("status"),
            "mode": mode,
            "has_pptx": bool(state.get("file")),
            "has_video": bool(state.get("video")),
            "updated": state_path.stat().st_mtime,
        })
    items.sort(key=lambda x: x["updated"], reverse=True)
    return items


def _prune_old_jobs():
    """Delete job folders beyond HISTORY_LIMIT (newest kept). Skip busy jobs."""
    items = _list_job_summaries()
    busy = {"queued", "drafting", "building", "rendering"}
    for old in items[HISTORY_LIMIT:]:
        if old["status"] in busy:
            continue
        jid = old["id"]
        with _lock:
            _jobs.pop(jid, None)
        job_dir = _job_dir(jid)
        if job_dir.exists():
            shutil.rmtree(job_dir, ignore_errors=True)


@app.post("/api/jobs/{job_id}/cancel")
def cancel_job(job_id: str):
    job = _get_job(job_id)
    if job["status"] not in ("queued", "drafting", "building", "rendering"):
        raise HTTPException(400, "Nothing to cancel — this job isn’t running.")
    job["cancel_requested"] = True
    job["step"] = "Cancelling…"
    _persist(job_id)
    return {"ok": True}


@app.delete("/api/jobs/{job_id}")
def delete_job(job_id: str):
    job = None
    try:
        job = _get_job(job_id)
    except HTTPException:
        pass
    if job and job["status"] in ("queued", "drafting", "building", "rendering"):
        job["cancel_requested"] = True
        _persist(job_id)
        raise HTTPException(400, "Job is still running — cancel it first, then delete.")
    with _lock:
        _jobs.pop(job_id, None)
    job_dir = _job_dir(job_id)
    if job_dir.exists():
        shutil.rmtree(job_dir)
    return {"ok": True}


@app.get("/api/jobs/{job_id}/slides/{index}/preview.png")
def slide_preview(job_id: str, index: int):
    """PNG thumbnail of one NDS-designed slide (generate mode only)."""
    job = _get_job(job_id)
    if job.get("options", {}).get("mode") == "narrate":
        raise HTTPException(400, "Preview is only available for NDS-designed decks.")
    plan = _load_plan(job_id)
    if not 0 <= index < len(plan.slides):
        raise HTTPException(400, "No such slide.")
    template = (job.get("options") or {}).get("template") or "niq"
    try:
        img = render_slide_preview(plan, index, template=template, width=640)
    except Exception as exc:
        raise HTTPException(500, f"Preview failed: {exc}")
    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return Response(content=buf.getvalue(), media_type="image/png",
                    headers={"Cache-Control": "no-store"})


@app.get("/api/jobs/{job_id}/download")
def job_download(job_id: str, name: str = ""):
    job = _get_job(job_id)
    if job.get("status") != "done" or not job.get("file"):
        raise HTTPException(404, "Deck not ready")
    # Only serve the primary deck or registered extra files — never an
    # arbitrary path from the query string.
    allowed = {job["file"]} | {e["name"] for e in (job.get("extra_files") or [])}
    fname = name if (name and name in allowed) else job["file"]
    path = _job_dir(job_id) / fname
    if not path.exists():
        raise HTTPException(404, "File not found")
    return FileResponse(
        path, filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


app.mount("/", StaticFiles(directory=ROOT / "static", html=True), name="static")
