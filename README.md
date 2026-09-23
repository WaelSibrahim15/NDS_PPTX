# NDS — Narrated Deck Studio

A local web app that turns content into **narrated PowerPoint presentations**: slides
with embedded per-slide audio that auto-plays and auto-advances, the narration in the
speaker notes, and optional MP4 export. Runs entirely on this Mac; the only network
calls are to the AI APIs you configure (Anthropic, OpenAI).

---

## Quick start

1. Double-click **Start NDS.command** (the first run creates `.venv` and installs
   dependencies). The app opens at <http://localhost:8765> — leave the terminal
   window open.
2. Open **Settings — API keys** in the app and paste:

   | Key | Used for | Required? |
   |---|---|---|
   | Anthropic | drafting slides, narration, image concepts | yes |
   | OpenAI | narration voices (TTS) + transcription of audio uploads | recommended — without it the offline macOS voice or ElevenLabs still work |
   | ElevenLabs key + voice ID(s) | custom / cloned narration voices | optional — adds an **ElevenLabs** voice provider; several voice IDs can be comma-separated |

   Keys live only in `config.json` in this folder (chmod 600). That file is
   gitignored — never commit it. For Railway / hosting, set the same values as
   environment variables instead: `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`,
   `ELEVENLABS_API_KEY`, `ELEVENLABS_VOICE_ID`
   (see `config.example.json`).

---

## Deploy on Railway (public URL)

Repo: push this `NDS/` folder as the GitHub repository root (e.g.
[WaelSibrahim15/NDS_PPTX](https://github.com/WaelSibrahim15/NDS_PPTX)).

1. In [Railway](https://railway.app): **New Project → Deploy from GitHub** → select
   the repo. Railway reads `railway.toml` and starts:
   `uvicorn app.main:app --host 0.0.0.0 --port $PORT`.
2. **Variables** (service → Variables):

   | Variable | Required | Purpose |
   |---|---|---|
   | `ANTHROPIC_API_KEY` | yes | drafting |
   | `OPENAI_API_KEY` | yes on Railway | cloud TTS (macOS `say` is unavailable) |
   | `ELEVENLABS_API_KEY` | optional | ElevenLabs voices |
   | `ELEVENLABS_VOICE_ID` | with the key above | ElevenLabs voice ID (comma-separate several) |
   | `APP_PASSWORD` | strongly recommended | HTTP Basic Auth password (username `nds` or blank) |
   | `JOBS_DIR` | recommended | set to `/data/jobs` when using a volume |

3. **Volume:** add a volume mounted at `/data`, with `JOBS_DIR=/data/jobs`, so
   past decks survive redeploys.
4. **Resources:** ~2 GB RAM, **one replica** (jobs use in-process threads).
5. Generate a Railway domain. Open the URL, enter the password, confirm Settings
   shows keys as saved (from env — the form is locked).

Health check: `GET /health` → `{"ok": true}`.

---

## The two modes

| Mode | Input | What NDS does |
|---|---|---|
| **Create (NDS design)** | `.docx` `.pdf` `.pptx` `.txt` `.md` | Claude drafts a deck plan (layouts + narration); NDS renders slides itself in the NIQ design system (from Claude Design) |
| **Narrate my PowerPoint** | a finished `.pptx` | Design stays byte-identical; NDS writes the narration (polishing existing speaker notes where present), voices it, embeds audio |

All modes share the same flow: **1 · Source & options → 2 · Review & amend →
3 · Progress & downloads** (the progress bar and download buttons appear directly
under the *Build* button). Every job is resumable via its URL
(`/?job=<id>`) and listed in **Past decks**.

### Options available in step 1

- **Narration style** — single narrator, or a **two-voice conversation** (Alex/Sam
  podcast style, voiced with two different voices and stitched per slide).
- **Language** — English, French, German, Spanish, Italian (slides + narration).
- **Guidance** — free-text instructions to the drafter (length, tone, audience).
- **NIQ design requirements** — extra design asks passed to the drafter (hidden in
  Narrate mode where the design is locked).

### The review step

Each slide is an editable card: title/bullets (or read-only slide content in
narrate mode, where the design is locked), the narration text, per-slide
**Preview audio**, and **Regenerate…** with an optional instruction. Voice provider
(OpenAI / ElevenLabs cloud, macOS offline), voice A/B and sample players sit in the toolbar.
Edits auto-save; rebuilds only re-voice slides whose narration changed.

Below the slide cards, **Redraft whole deck** throws the current draft away and asks
Claude for a fresh one from the same source, with an optional instruction (e.g. "fewer
slides, more card layouts"). The previous draft is kept in `deck_previous.json` so
**Undo last redraft** can restore it once.

---

## Architecture

```
NDS/
├── Start NDS.command       # zsh launcher: venv bootstrap + uvicorn on :8765
├── config.json             # API keys (0600)
├── requirements.txt
├── app/
│   ├── main.py             # FastAPI app: endpoints, job lifecycle, build orchestration
│   ├── models.py           # Pydantic deck model (DeckPlan / Slide / Card / Stat / CompareSide)
│   ├── extract.py          # text + per-slide extraction from docx/pdf/pptx/txt/md
│   ├── drafter.py          # all Claude calls (drafting, narration, slide regeneration)
│   ├── audio_source.py     # audio uploads: Whisper transcript + slicing the original recording
│   ├── tts.py              # TTS providers (OpenAI, ElevenLabs, macOS), dialogue splitting, stitching
│   ├── design.py           # slide layouts (NIQ design system), shared by the two backends below
│   ├── pptx_builder.py     # PPTX assembly: draws design.py layouts, audio embed, timing XML
│   ├── renderer.py         # Pillow renderer: draws design.py layouts as 1920x1080 PNG frames
│   ├── video.py            # MP4 export using the bundled imageio-ffmpeg binary
│   └── assets/             # NIQ logos, brand symbols, fallback fonts
├── static/index.html       # the whole UI (vanilla JS, no build step)
└── jobs/<12-hex-id>/       # one folder per deck (see "Job anatomy")
```

**No system dependencies.** Everything runs from the venv: slide frames are drawn by
Pillow (not LibreOffice), video is encoded by the ffmpeg binary that ships inside
`imageio-ffmpeg`, offline TTS uses the macOS `say`/`afconvert` built-ins.

### Job anatomy (`jobs/<id>/`)

| File | Contents |
|---|---|
| `state.json` | status, step, progress, options, output filenames — the job survives server restarts |
| `source.*` | the uploaded file |
| `source_text.txt` | extracted text handed to Claude |
| `deck.json` | the editable `DeckPlan` (layouts, titles, bullets, narration per slide) |
| `audio/<sha1>.mp3/.m4a` | per-narration audio cache |
| `*.pptx` / `*.mp4` | the deliverables (narrated deck, video) |

### HTTP API

| Endpoint | Purpose |
|---|---|
| `GET /api/config` · `POST /api/settings` | key status / save keys |
| `POST /api/draft` | upload + options → returns `job_id`, drafts in a background thread |
| `GET /api/jobs/{id}` | live job state (poll target) |
| `PUT /api/jobs/{id}/deck` | save the edited deck plan |
| `POST /api/jobs/{id}/slides/{i}/regenerate` | redraft one slide (narration-only when design is locked) |
| `GET /api/jobs/{id}/audio/{i}` | preview one slide's audio (cached) |
| `GET /api/voices/{provider}/{voice}/sample` | voice sample |
| `POST /api/jobs/{id}/build` | voice + assemble the PPTX |
| `POST /api/jobs/{id}/export-video` | render the MP4 (NDS-designed decks only) |
| `GET /api/jobs/{id}/download?name=` | primary deck, or a registered extra file by name (allowlisted) |
| `GET /api/jobs` | Past-decks history |

Long work happens in daemon threads; the UI polls `GET /api/jobs/{id}` every 1.5 s.

---

## Technical details by subsystem

### Drafting (`drafter.py`, model `claude-opus-4-8`)

All calls use `messages.parse` with **structured outputs** (Pydantic schemas) and
adaptive thinking — no JSON parsing, invalid outputs are retried at the API layer.

- `draft_deck` — document text → `DeckPlan`. The system prompt encodes the **NIQ
  PowerPoint compliance checklist**: stay within the official NIQ visual system
  (no invented masters/templates), NIQ blues + white dominant, secondary colors
  sparingly (not as backgrounds; grey tints OK), Arial/Georgia with restrained
  mixing, **sentence case** (no all-caps), strong contrast and white space, icons
  sparingly, figures on `stats`/`compare` so they dominate, NIQ data-viz palette
  (never PowerPoint default chart colors), insight-led takeaway titles, layout
  variety (`content`, `cards`, `stats`, `compare`, plus `title`/`section`/`closing`),
  and brand limits (neutral measurer; no isometric/bespoke art; GfK / Tech &
  Durables co-brand only when source requires it).
- `draft_narration_for_existing` — one narration per slide of an uploaded deck,
  based on speaker notes when present.
- `regenerate_slide` — single-slide redraft with the deck outline as context;
  when the design is locked only the `narration` field may change.

### Voices (`tts.py`)

- **OpenAITTS** — model `gpt-4o-mini-tts`, 10 voices, mp3.
- **MacSayTTS** — offline `say` → `afconvert` → m4a; zero-key demo path.
- **Two-voice conversations**: `split_dialogue` parses `Alex:`/`Sam:` scripts.
  Tags are recognised at line starts **and inline after sentence punctuation**
  (the drafter often writes the whole exchange in one paragraph). An
  `_is_dialogue` guard (known host names, a repeating speaker, or the
  line-per-turn layout) prevents prose like "Remember: check the register.
  Second: …" from being split. Turns are synthesized per speaker with voice A/B
  and stitched into one m4a with ffmpeg concat.
- **Audio cache**: `audio/<sha1(provider|voiceA|voiceB|style|narration)[:16]>.*`.
  Rebuilds only voice changed narrations. (Note: the key covers only those inputs —
  after a *code* change to voicing, delete `jobs/<id>/audio/` to force re-synthesis.)

### PPTX assembly (`pptx_builder.py`)

- **Design** lives in `design.py`, taken from the NIQ design system in Claude
  Design (tokens: deep blue `060A45`, bright blue `2D6DF6`, light blue `31D1FF`,
  orange `EF5F17`, blue tint `B4CBF9`, grey panel `F2F2F2`, ink `555555`). Each
  layout returns a list of shapes (rects, circles, arcs, lines, images, text) in
  inches; `pptx_builder.py` turns them into PowerPoint shapes and `renderer.py`
  draws the same list with Pillow for previews and MP4 frames, so they match.
- **Layouts**: White cover with the circle motif and NIQ mark; section dividers
  alternating Blue and Dark grounds with the arc motif; sidebar-plus-content
  (Georgia statement title); feature cards (rounded Bright Blue header bar + NIQ
  brand symbol, chosen by the drafter from `app/assets/symbols/`); one metric
  callout plus supporting figures; Gray/Blue half panels for compare; Dark
  closing. Content slides share the footer: NIQ mark, legal line, page number
  above a hairline rule. A "neutral professional" template uses the same layouts
  without NIQ branding.
- **Assets**: `app/assets/logos` (NIQ mark PNGs), `app/assets/symbols` (24 NIQ
  brand symbols, Bright Blue variant), `app/assets/fonts` (Liberation Sans/Serif,
  SIL OFL, used by the renderer when Arial/Georgia are missing, e.g. on Railway).
- **Audio embed**: each slide gets the audio as a media shape plus a hand-built
  `<p:timing>` tree that fires `playFrom(0.0)` on slide entry, and a
  `<p:transition advTm>` fade so the show auto-advances after narration + 1 s.
  Slide-element order (`cSld → clrMapOvr → transition → timing`) is enforced.
- **`narrate_existing_pptx`** copies the uploaded deck and only adds notes,
  audio, timing — plus a monkeypatch for python-pptx ≤ 1.0.x
  (`_MediaParts._find_by_sha1` crashes on decks that already contain media).

### Video export (`video.py` + `renderer.py`)

NDS-designed decks only: `renderer.py` draws every `design.py` layout as 1920×1080
PNGs with Pillow (the same layouts the PPTX uses, so they always match), then ffmpeg concats per-slide segments
(frame + narration + 1 s tail) into one MP4. Uploaded decks have no frame
source, so MP4 export is disabled for them.

### Build orchestration (`main.py::_run_build`)

1. Voice every slide (cache-aware), or slice the original recording for audio uploads.
2. Assemble the narrated PPTX.

---

## Known limits & notes

- **MP4 / video export** is limited to NDS-designed decks (no LibreOffice on this
  machine, by design).
- The audio cache does not know about code changes — delete `jobs/<id>/audio/` to
  force re-voicing after modifying the TTS pipeline.
- Uploaded decks keep their own slide size; NDS-designed decks are 13.333×7.5 in.
- Conversation narration relies on the `Alex:`/`Sam:` names the drafter is
  instructed to use; custom speaker names work when they repeat or sit one per line.
