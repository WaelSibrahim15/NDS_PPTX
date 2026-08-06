# NDS — Narrated Deck Studio

A local web app that turns content into **narrated PowerPoint presentations**: slides
with embedded per-slide audio that auto-plays and auto-advances, the narration in the
speaker notes, and optional MP4 export. Runs entirely on this Mac; the only network
calls are to the AI APIs you configure (Anthropic, OpenAI, Gamma).

---

## Quick start

1. Double-click **Start NDS.command** (the first run creates `.venv` and installs
   dependencies). The app opens at <http://localhost:8765> — leave the terminal
   window open.
2. Open **Settings — API keys** in the app and paste:

   | Key | Used for | Required? |
   |---|---|---|
   | Anthropic | drafting slides, narration, image concepts | yes |
   | OpenAI | narration voices (TTS) + generated imagery (Enhance mode) | recommended — without it the offline macOS voice still works, images are skipped |
   | Gamma | "Design with Gamma" mode | only for that mode |

   Keys live only in `config.json` in this folder (chmod 600). That file is
   gitignored — never commit it. For Railway / hosting, set the same values as
   environment variables instead: `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`,
   `GAMMA_API_KEY` (see `config.example.json`).

---

## Deploy on Railway (public URL)

Repo: push this `NDS/` folder as the GitHub repository root (e.g.
[WaelSibrahim15/NDS_PPTX](https://github.com/WaelSibrahim15/NDS_PPTX)).

1. In [Railway](https://railway.app): **New Project → Deploy from GitHub** → select
   the repo. Railway picks up `Procfile` / `railway.toml` and starts:
   `uvicorn app.main:app --host 0.0.0.0 --port $PORT`.
2. **Variables** (service → Variables):

   | Variable | Required | Purpose |
   |---|---|---|
   | `ANTHROPIC_API_KEY` | yes | drafting |
   | `OPENAI_API_KEY` | yes on Railway | cloud TTS (macOS `say` is unavailable) |
   | `GAMMA_API_KEY` | only for Gamma mode | Gamma Generations API |
   | `APP_PASSWORD` | strongly recommended | HTTP Basic Auth password (username `nds` or blank) |
   | `JOBS_DIR` | recommended | set to `/data/jobs` when using a volume |

3. **Volume:** add a volume mounted at `/data`, with `JOBS_DIR=/data/jobs`, so
   past decks survive redeploys.
4. **Resources:** ~2 GB RAM, **one replica** (jobs use in-process threads).
5. Generate a Railway domain. Open the URL, enter the password, confirm Settings
   shows keys as saved (from env — the form is locked).

Health check: `GET /health` → `{"ok": true}`.

---

## The four modes

| Mode | Input | What NDS does |
|---|---|---|
| **Create (NDS design)** | `.docx` `.pdf` `.pptx` `.txt` `.md` | Claude drafts a deck plan (layouts + narration); NDS renders slides itself in the NIQ 2026 design language |
| **Design with Gamma** | any document | Gamma's Generations API designs the deck (theme, layouts, imagery) and exports a `.pptx`; NDS then narrates it like an uploaded deck |
| **Narrate my PowerPoint** | a finished `.pptx` | Design stays byte-identical; NDS writes the narration (polishing existing speaker notes where present), voices it, embeds audio |
| **Enhance my PowerPoint** | a designed `.pptx` | Design is kept, and NDS adds: AI-generated imagery (only into detected empty regions), build-in motion + slide fades, and full narration/voice |

All modes share the same flow: **1 · Source & options → 2 · Review & amend →
3 · Progress & downloads** (the progress bar and download buttons appear directly
under the *Build* button). Every job is resumable via its URL
(`/?job=<id>`) and listed in **Past decks**.

### Options available in step 1

- **Narration style** — single narrator, or a **two-voice conversation** (Alex/Sam
  podcast style, voiced with two different voices and stitched per slide).
- **Language** — English, French, German, Spanish, Italian (slides + narration).
- **Guidance** — free-text instructions to the drafter (length, tone, audience).
- **NIQ design requirements** — extra design asks passed to whichever engine designs
  or enriches visuals (hidden in Narrate mode where the design is locked).

### The review step

Each slide is an editable card: title/bullets (or read-only slide content in
narrate/enhance modes, where the design is locked), the narration text, per-slide
**Preview audio**, and **Regenerate…** with an optional instruction. Voice provider
(OpenAI cloud / macOS offline), voice A/B and a sample player sit in the toolbar.
Edits auto-save; rebuilds only re-voice slides whose narration changed.

---

## Architecture

```
NDS/
├── Start NDS.command       # zsh launcher: venv bootstrap + uvicorn on :8765
├── config.json             # API keys + optional gamma_theme / gamma_brand (0600)
├── requirements.txt
├── app/
│   ├── main.py             # FastAPI app: endpoints, job lifecycle, build orchestration
│   ├── models.py           # Pydantic deck model (DeckPlan / Slide / Card / Stat / CompareSide)
│   ├── extract.py          # text + per-slide extraction from docx/pdf/pptx/txt/md
│   ├── drafter.py          # all Claude calls (drafting, narration, enhance plan, translation)
│   ├── gamma.py            # Gamma Generations API client + NIQ brand brief
│   ├── enhance.py          # Enhance mode: empty-region detection + OpenAI image generation
│   ├── tts.py              # TTS providers, dialogue splitting, audio stitching
│   ├── pptx_builder.py     # PPTX assembly: layouts, audio embed, timing XML
│   ├── renderer.py         # Pillow renderer: 1920x1080 PNG frames (mirrors pptx_builder)
│   └── video.py            # MP4 export using the bundled imageio-ffmpeg binary
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
| `source.*` | the uploaded file (Gamma mode replaces it with Gamma's designed `source.pptx`) |
| `source_text.txt` | extracted text handed to Claude |
| `deck.json` | the editable `DeckPlan` (layouts, titles, bullets, narration per slide) |
| `enhance.json` | Enhance mode: per-slide image prompts + detected empty rectangles |
| `enhanced.pptx` | Enhance mode: the design + inserted images, pre-narration |
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
- `draft_enhancements` — Enhance mode: narrations **plus** a per-slide image prompt
  (or `null`), only for slides the analyzer flagged as having a usable empty slot.
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

- **NIQ 2026 palette** (exact brand theme hexes): navy `060A45`, bright blue
  `2C6DF6`, cyan `31D1FF`, orange `EF5F17`, green `59AD00`, pink `EF5890`, amber
  `FFB500`, body grey `555555`; Arial + Georgia-italic accents; signature circle
  motif. A "neutral professional" template exists alongside.
- **Layouts**: dark title/closing, blue section divider, and content-style slides
  (bulleted content, numbered card grid, dark stat band, two-panel compare), all
  sharing header chrome (title, kicker, page number, circle motif).
- **Audio embed**: each slide gets the audio as a media shape plus a hand-built
  `<p:timing>` tree that fires `playFrom(0.0)` on slide entry, and a
  `<p:transition advTm>` fade so the show auto-advances after narration + 1 s.
  Slide-element order (`cSld → clrMapOvr → transition → timing`) is enforced.
- **Movements** (Enhance mode): the timing tree can carry per-shape **fade
  entrance builds as siblings of one auto-firing group node** — each with an
  absolute delay spread across the narration, so they run without clicks under
  auto-advance. (An earlier design used click-gated main-sequence steps, which
  blanked auto-playing shows — do not regress this.)
- **`narrate_existing_pptx`** copies the uploaded deck and only adds notes,
  audio, timing — plus a monkeypatch for python-pptx ≤ 1.0.x
  (`_MediaParts._find_by_sha1` crashes on decks that already contain media).

### Enhance mode (`enhance.py`)

- `analyze_slides` rasterises every shape's bounding box onto a 64×36 grid and
  finds the **largest empty rectangle** (histogram method). An image slot exists
  only if the rect is ≥ 20 % of slide width × 25 % of height **and** the slide has
  no picture yet — an added image can never overlap the original design; on dense
  decks NDS correctly adds nothing.
- `add_images` calls the OpenAI Images API (`gpt-image-1`, falling back to
  `dall-e-3`), size picked from the rect's aspect ratio, appends a consistent
  corporate-photography style suffix (no text/logos/charts), contain-fits the
  result into the rect, capped at 6 images per deck. Image failures degrade to
  "skip images", never fail the build.
- The result is saved as `enhanced.pptx` and then flows through the normal
  narrate pipeline with movements enabled.

### Gamma engine (`gamma.py`)

- `POST /v1.0/generations` (`exportAs: "pptx"`), polled until `completed`; the
  deck downloads from `exportUrl`. v0.2 of the API is sunset; Cloudflare rejects
  Python's default user agent, so a browser-style UA is sent.
- Every generation is prefixed with the **NIQ PowerPoint compliance brief**
  (official-template behaviour, logo clear space, blues/white dominant, sentence
  case, Arial/Georgia, chart recoloring, data neutrality, co-brand rules) inside
  `additionalInstructions`; a saved Gamma theme can be pinned via `gamma_theme`
  in `config.json` (`gamma_brand: false` disables the brief). Generation costs
  Gamma credits (~45 s).
- Gamma **cannot edit an existing pptx** — that's what Enhance mode is for.

### Video export (`video.py` + `renderer.py`)

NDS-designed decks only: `renderer.py` re-draws every layout as 1920×1080 PNGs with
Pillow (it deliberately mirrors `pptx_builder`'s geometry and palette — **keep the
two in sync when changing layouts**), then ffmpeg concats per-slide segments
(frame + narration + 1 s tail) into one MP4. Uploaded/Gamma decks have no frame
source, so MP4 export is disabled for them.

### Build orchestration (`main.py::_run_build`)

1. Voice every slide (cache-aware).
2. Enhance mode: generate + insert images → `enhanced.pptx`.
3. Assemble the narrated PPTX (movements on for Enhance).

---

## Known limits & notes

- **MP4 / video export** is limited to NDS-designed decks (no LibreOffice on this
  machine, by design).
- **Entrance animations** (Enhance mode) are structurally valid and validator-clean,
  but PowerPoint playback can't be verified on this machine — if content ever
  appears blank during a show, disable the `animate=` flag in the enhance path.
- The audio cache does not know about code changes — delete `jobs/<id>/audio/` to
  force re-voicing after modifying the TTS pipeline.
- Uploaded decks keep their own slide size; NDS-designed decks are 13.333×7.5 in.
- Conversation narration relies on the `Alex:`/`Sam:` names the drafter is
  instructed to use; custom speaker names work when they repeat or sit one per line.
- HeyGen avatar export existed in an earlier version and has been **removed**.
