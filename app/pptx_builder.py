"""NDS — assemble the narrated .pptx from a deck plan + per-slide audio files.

Styling follows the NIQ 2026 design language: Deep Blue / Bright Blue / orange
accent, Arial headings, Georgia italic accents, circle motif.
"""
from pathlib import Path
from typing import Dict, List, Optional

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .models import DeckPlan, Slide
from .tts import audio_duration_seconds

# python-pptx (≤1.0.x) crashes when adding media to a deck that already contains
# media/images loaded as plain Parts (no .sha1). Same guard the library already
# uses for images (pptx/package.py _ImageParts) applied to media parts.
from pptx.package import _MediaParts


def _find_media_by_sha1_safe(self, sha1):
    for media_part in self:
        if not hasattr(media_part, "sha1"):
            continue
        if media_part.sha1 == sha1:
            return media_part
    return None


_MediaParts._find_by_sha1 = _find_media_by_sha1_safe

# Official NIQ 2026 brand palette (from the brand deck theme: dk2 + accent1-6).
DEEP_BLUE = RGBColor(0x06, 0x0A, 0x45)   # dk2  — Deep Navy (title/section/closing bg)
BRIGHT_BLUE = RGBColor(0x2C, 0x6D, 0xF6)  # accent1 — Bright Blue (primary accent)
CYAN = RGBColor(0x31, 0xD1, 0xFF)         # accent2 — Cyan (cool accent, best on dark)
ORANGE = RGBColor(0xEF, 0x5F, 0x17)       # accent3 — Orange (sparing warm highlight)
GREEN = RGBColor(0x59, 0xAD, 0x00)        # accent4
PINK = RGBColor(0xEF, 0x58, 0x90)         # accent5
AMBER = RGBColor(0xFF, 0xB5, 0x00)        # accent6
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GREY = RGBColor(0x55, 0x55, 0x55)    # dk1 — brand body/text grey
LIGHT_BLUE_TINT = RGBColor(0xF0, 0xF4, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

NEUTRAL = {
    "bg_dark": RGBColor(0x1F, 0x2A, 0x37),
    "accent": RGBColor(0x2F, 0x6F, 0xED),
    "accent2": RGBColor(0xD9, 0x77, 0x06),
}


def _card_entrance_anims(card_groups, duration_s: float):
    """Give each card a fade-in delay so the cards build in as the voice-over
    progresses. Reveals are spread across the first ~65% of the narration, with
    the first card appearing almost immediately. Returns [(shape_id, delay_ms)]."""
    n = len(card_groups)
    if n == 0:
        return []
    window = max(1.0, duration_s * 0.65)
    step = window / n
    anims = []
    for i, group in enumerate(card_groups):
        delay_ms = int((0.25 + i * step) * 1000)
        for shape_id in group:
            anims.append((shape_id, delay_ms))
    return anims


def build_deck(
    plan: DeckPlan,
    audio_files: List[Optional[Path]],
    out_path: Path,
    template: str = "niq",
    animate: bool = False,
) -> Path:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    colors = _palette(template)
    content_count = sum(1 for s in plan.slides if s.layout == "content")
    content_seen = 0

    for idx, spec in enumerate(plan.slides):
        slide = prs.slides.add_slide(blank)
        card_groups = None
        if spec.layout == "title":
            _title_slide(slide, plan, spec, colors)
        elif spec.layout == "section":
            content_seen += 0
            _section_slide(slide, spec, colors)
        elif spec.layout == "closing":
            _closing_slide(slide, spec, colors)
        elif spec.layout == "cards" and spec.cards:
            card_groups = _cards_slide(slide, spec, colors, idx + 1, len(plan.slides))
        elif spec.layout == "stats" and spec.stats:
            _stats_slide(slide, spec, colors, idx + 1, len(plan.slides))
        elif spec.layout == "compare" and (spec.compare_left or spec.compare_right):
            _compare_slide(slide, spec, colors, idx + 1, len(plan.slides))
        else:
            content_seen += 1
            _content_slide(slide, spec, colors, idx + 1, len(plan.slides))

        # narration into speaker notes so a human presenter can reuse the deck
        slide.notes_slide.notes_text_frame.text = spec.narration

        audio = audio_files[idx] if idx < len(audio_files) else None
        if audio is not None and Path(audio).exists():
            duration = audio_duration_seconds(Path(audio))
            anims = _card_entrance_anims(card_groups, duration) if (animate and card_groups) else None
            _embed_autoplay_audio(slide, Path(audio), anims=anims)
            _set_auto_advance(slide, duration + 1.0)
        # Silent drafts: no auto-advance — presenter (or reviewer) advances manually.

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def _palette(template: str) -> Dict[str, RGBColor]:
    if template == "neutral":
        return {
            "dark": NEUTRAL["bg_dark"], "accent": NEUTRAL["accent"],
            "accent2": NEUTRAL["accent2"], "cyan": RGBColor(0x38, 0xB2, 0xC4),
            "green": RGBColor(0x2F, 0x9E, 0x44), "pink": RGBColor(0xC2, 0x41, 0x7A),
            "amber": RGBColor(0xD9, 0x9E, 0x06), "body": DARK_GREY,
            "tint": RGBColor(0xEE, 0xF2, 0xF8),
        }
    return {
        "dark": DEEP_BLUE, "accent": BRIGHT_BLUE, "accent2": ORANGE,
        "cyan": CYAN, "green": GREEN, "pink": PINK, "amber": AMBER,
        "body": DARK_GREY, "tint": LIGHT_BLUE_TINT,
    }


# ---------------------------------------------------------------- slide styles

def _fill(slide, color: RGBColor):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _textbox(slide, left, top, width, height):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    return box, tf


def _run(para, text, *, font="Arial", size=18, bold=False, italic=False, color=DARK_GREY):
    r = para.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return r


def _circle(slide, left, top, diameter, color: RGBColor, outline_only=False, line_w=2.5):
    from pptx.enum.shapes import MSO_SHAPE

    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, diameter, diameter)
    if outline_only:
        shp.fill.background()
        shp.line.color.rgb = color
        shp.line.width = Pt(line_w)
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = color
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _title_slide(slide, plan: DeckPlan, spec: Slide, c):
    _fill(slide, c["dark"])
    # signature circle motif, right side
    _circle(slide, Inches(9.2), Inches(1.4), Inches(4.6), c["accent"], outline_only=True, line_w=3)
    _circle(slide, Inches(10.6), Inches(4.6), Inches(1.5), c["cyan"])
    _circle(slide, Inches(8.7), Inches(5.4), Inches(0.7), WHITE, outline_only=True, line_w=2)

    _, tf = _textbox(slide, Inches(0.9), Inches(2.3), Inches(8.0), Inches(2.2))
    p = tf.paragraphs[0]
    _run(p, spec.title or plan.deck_title, size=40, bold=True, color=WHITE)

    _, tf2 = _textbox(slide, Inches(0.9), Inches(4.3), Inches(7.6), Inches(1.0))
    p2 = tf2.paragraphs[0]
    _run(p2, spec.subtitle or plan.subtitle, font="Georgia", size=20, italic=True, color=c["cyan"])

    _, tf3 = _textbox(slide, Inches(0.9), Inches(6.7), Inches(7.0), Inches(0.4))
    _run(tf3.paragraphs[0], "Narrated presentation · generated with NDS", size=11,
         color=RGBColor(0xB0, 0xB0, 0xD8))


def _section_slide(slide, spec: Slide, c):
    _fill(slide, c["accent"])
    _circle(slide, Inches(-1.6), Inches(4.4), Inches(5.2), c["dark"])
    _circle(slide, Inches(11.6), Inches(-1.2), Inches(3.4), c["accent2"], outline_only=True, line_w=3)

    _, tf = _textbox(slide, Inches(1.0), Inches(2.9), Inches(11.0), Inches(1.8))
    p = tf.paragraphs[0]
    _run(p, spec.title, size=36, bold=True, color=WHITE)
    if spec.subtitle:
        _, tf2 = _textbox(slide, Inches(1.0), Inches(4.5), Inches(10.0), Inches(0.8))
        _run(tf2.paragraphs[0], spec.subtitle, font="Georgia", size=18, italic=True, color=WHITE)


GREY_TXT = RGBColor(0x55, 0x55, 0x55)
CARD_TINT = RGBColor(0xF4, 0xF6, 0xFE)
PAGE_GREY = RGBColor(0x9A, 0x9A, 0x9A)


def _header(slide, spec: Slide, c, page_no: int, total: int):
    """Shared chrome for content-style slides: title, kicker, footer, motif."""
    _fill(slide, WHITE)
    _circle(slide, Inches(12.35), Inches(6.75), Inches(0.45), c["accent"], outline_only=True, line_w=1.75)

    _, tf = _textbox(slide, Inches(0.65), Inches(0.42), Inches(11.9), Inches(0.75))
    _run(tf.paragraphs[0], spec.title, size=24, bold=True, color=c["dark"])
    if spec.subtitle:
        _, kick = _textbox(slide, Inches(0.65), Inches(1.05), Inches(11.9), Inches(0.4))
        _run(kick.paragraphs[0], spec.subtitle, size=13, color=GREY_TXT)

    _, pn = _textbox(slide, Inches(0.65), Inches(6.98), Inches(3.0), Inches(0.35))
    _run(pn.paragraphs[0], f"{page_no} / {total}", size=10, color=PAGE_GREY)


def _panel(slide, left, top, width, height, fill_color, rounded=True):
    from pptx.enum.shapes import MSO_SHAPE

    shape = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    p = slide.shapes.add_shape(shape, left, top, width, height)
    p.fill.solid()
    p.fill.fore_color.rgb = fill_color
    p.line.fill.background()
    p.shadow.inherit = False
    if rounded:
        try:
            p.adjustments[0] = 0.06
        except Exception:
            pass
    return p


def _content_slide(slide, spec: Slide, c, page_no: int, total: int):
    _header(slide, spec, c, page_no, total)
    _, body_tf = _textbox(slide, Inches(0.65), Inches(1.85), Inches(11.6), Inches(4.9))
    first = True
    for b in spec.bullets:
        p = body_tf.paragraphs[0] if first else body_tf.add_paragraph()
        first = False
        p.space_after = Pt(16)
        _run(p, "●  ", size=11, color=c["accent"])
        _run(p, b, size=16, color=c["body"])


def _cards_slide(slide, spec: Slide, c, page_no: int, total: int):
    """Draw the numbered card grid. Returns one list of shape-ids per card so the
    caller can give each card its own timed entrance animation."""
    _header(slide, spec, c, page_no, total)
    cards = spec.cards[:4]
    n = len(cards)
    accents = [c["accent"], c["accent2"], c["dark"], c["accent"]]
    gap = Inches(0.3)
    total_w = Inches(12.0)
    card_w = int((total_w - gap * (n - 1)) / n)
    top, card_h = Inches(1.9), Inches(4.3)
    groups = []
    for i, card in enumerate(cards):
        left = Inches(0.65) + i * (card_w + gap)
        panel = _panel(slide, left, top, card_w, card_h, CARD_TINT)
        pad = Inches(0.28)
        num_box, num = _textbox(slide, left + pad, top + Inches(0.25), card_w - pad * 2, Inches(0.6))
        _run(num.paragraphs[0], f"{i + 1:02d}", size=28, bold=True, color=accents[i % len(accents)])
        tt_box, tt = _textbox(slide, left + pad, top + Inches(1.0), card_w - pad * 2, Inches(1.0))
        _run(tt.paragraphs[0], card.title, size=16, bold=True, color=c["dark"])
        dd_box, dd = _textbox(slide, left + pad, top + Inches(1.95), card_w - pad * 2, card_h - Inches(2.2))
        _run(dd.paragraphs[0], card.desc, size=11.5, color=GREY_TXT)
        groups.append([panel.shape_id, num_box.shape_id, tt_box.shape_id, dd_box.shape_id])
    return groups


def _stats_slide(slide, spec: Slide, c, page_no: int, total: int):
    _header(slide, spec, c, page_no, total)
    stats = spec.stats[:5]
    n = len(stats)
    panel_top, panel_h = Inches(2.3), Inches(3.2)
    _panel(slide, Inches(0.65), panel_top, Inches(12.0), panel_h, c["dark"])
    col_w = int(Inches(12.0) / n)
    for i, st in enumerate(stats):
        left = Inches(0.65) + i * col_w
        _, val = _textbox(slide, left + Inches(0.25), panel_top + Inches(0.85), col_w - Inches(0.5), Inches(0.9))
        _run(val.paragraphs[0], st.value, size=30, bold=True,
             color=WHITE if i % 2 == 0 else c["cyan"])
        _, lab = _textbox(slide, left + Inches(0.25), panel_top + Inches(1.85), col_w - Inches(0.5), Inches(0.9))
        _run(lab.paragraphs[0], st.label, size=11, color=RGBColor(0xC6, 0xD0, 0xFF))


def _compare_slide(slide, spec: Slide, c, page_no: int, total: int):
    _header(slide, spec, c, page_no, total)
    panels = [(spec.compare_left, c["accent"]), (spec.compare_right, c["accent2"])]
    top, head_h, body_h = Inches(1.9), Inches(0.6), Inches(4.0)
    width = Inches(5.85)
    for i, (side, accent) in enumerate(panels):
        if side is None:
            continue
        left = Inches(0.65) + i * (width + Inches(0.3))
        head = _panel(slide, left, top, width, head_h, accent, rounded=False)
        htf = head.text_frame
        htf.margin_left = Inches(0.25)
        htf.word_wrap = True
        htf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _run(htf.paragraphs[0], side.heading, size=14, bold=True, color=WHITE)
        _panel(slide, left, top + head_h, width, body_h, CARD_TINT, rounded=False)
        _, btf = _textbox(slide, left + Inches(0.25), top + head_h + Inches(0.2),
                          width - Inches(0.5), body_h - Inches(0.4))
        first = True
        for item in side.items:
            p = btf.paragraphs[0] if first else btf.add_paragraph()
            first = False
            p.space_after = Pt(10)
            _run(p, "●  ", size=10, color=accent)
            _run(p, item, size=13, color=c["body"])


def _closing_slide(slide, spec: Slide, c):
    _fill(slide, c["dark"])
    _circle(slide, Inches(10.4), Inches(-1.8), Inches(4.4), c["accent"], outline_only=True, line_w=3)
    _circle(slide, Inches(0.4), Inches(5.9), Inches(1.1), c["cyan"])

    _, tf = _textbox(slide, Inches(0.9), Inches(2.2), Inches(10.5), Inches(1.4))
    _run(tf.paragraphs[0], spec.title, size=34, bold=True, color=WHITE)

    if spec.bullets:
        _, tfb = _textbox(slide, Inches(0.9), Inches(3.8), Inches(10.5), Inches(2.4))
        first = True
        for b in spec.bullets:
            p = tfb.paragraphs[0] if first else tfb.add_paragraph()
            first = False
            p.space_after = Pt(10)
            _run(p, "●  ", size=12, color=c["accent2"])
            _run(p, b, size=16, color=WHITE)


def narrate_existing_pptx(
    src_path: Path,
    narrations: List[str],
    audio_files: List[Optional[Path]],
    out_path: Path,
    animate: bool = False,
) -> Path:
    """Embed narration audio into a COPY of an existing deck, leaving its design
    untouched: per-slide auto-play audio, auto-advance, narration in the notes.
    When ``animate`` is set, the slide's existing content shapes fade in one after
    another as the voice-over plays (a generic build for designed/Gamma decks)."""
    prs = Presentation(str(src_path))
    for idx, slide in enumerate(prs.slides):
        # capture the existing design shapes BEFORE we add audio so we only
        # animate the deck's own content, never our overlays.
        design_ids = [sh.shape_id for sh in slide.shapes] if animate else []
        if idx < len(narrations) and narrations[idx].strip():
            slide.notes_slide.notes_text_frame.text = narrations[idx]
        audio = audio_files[idx] if idx < len(audio_files) else None
        if audio is not None and Path(audio).exists():
            duration = audio_duration_seconds(Path(audio))
            anims = _sequential_anims(design_ids, duration) if (animate and design_ids) else None
            _embed_autoplay_audio(slide, Path(audio), anims=anims)
            _set_auto_advance(slide, duration + 1.0)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def _sequential_anims(shape_ids, duration_s: float, cap: int = 10):
    """Sequential fade-in for an existing deck's shapes, spread across the first
    ~70% of the narration. Caps the count so a busy slide doesn't flicker."""
    ids = shape_ids[:cap]
    n = len(ids)
    if n == 0:
        return []
    window = max(1.0, duration_s * 0.7)
    step = window / n
    return [(sid, int((0.2 + i * step) * 1000)) for i, sid in enumerate(ids)]


# ------------------------------------------------------- audio embed + timing

_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Slide timing tree. The media playback AND every entrance build live as siblings
# inside ONE group time-node that begins when the slide appears (onBegin of the
# main sequence). Each sibling starts at its own absolute delay, so under an
# auto-advancing show (no clicks) they ALL fire automatically. (An earlier version
# put each entrance as a separate top-level main-sequence step, which are
# click-gated — so under auto-advance they never fired and the shapes stayed
# hidden, blanking the slide.)
_TIMING_HEAD = f"""<p:timing xmlns:p="{_P}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <p:tnLst>
  <p:par>
   <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
    <p:childTnLst>
     <p:seq concurrent="1" nextAc="seek">
      <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
       <p:childTnLst>
        <p:par>
         <p:cTn id="3" fill="hold">
          <p:stCondLst>
           <p:cond delay="indefinite"/>
           <p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>
          </p:stCondLst>
          <p:childTnLst>"""

# Media playback child (starts immediately, delay 0). ids {i0}-{i2}.
_TIMING_MEDIA = """           <p:par>
            <p:cTn id="{i0}" fill="hold">
             <p:stCondLst><p:cond delay="0"/></p:stCondLst>
             <p:childTnLst>
              <p:par>
               <p:cTn id="{i1}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">
                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                <p:childTnLst>
                 <p:cmd type="call" cmd="playFrom(0.0)">
                  <p:cBhvr>
                   <p:cTn id="{i2}" dur="{dur}" fill="hold"/>
                   <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                  </p:cBhvr>
                 </p:cmd>
                </p:childTnLst>
               </p:cTn>
              </p:par>
             </p:childTnLst>
            </p:cTn>
           </p:par>"""

# One timed fade-in entrance child. Starts at its own {delay}. ids {i0}-{i3}.
_TIMING_ENTRANCE = """           <p:par>
            <p:cTn id="{i0}" fill="hold">
             <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
             <p:childTnLst>
              <p:par>
               <p:cTn id="{i1}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="afterEffect">
                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                <p:childTnLst>
                 <p:set>
                  <p:cBhvr>
                   <p:cTn id="{i2}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                   <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                   <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                  </p:cBhvr>
                  <p:to><p:strVal val="visible"/></p:to>
                 </p:set>
                 <p:animEffect transition="in" filter="fade">
                  <p:cBhvr>
                   <p:cTn id="{i3}" dur="400"/>
                   <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                  </p:cBhvr>
                 </p:animEffect>
                </p:childTnLst>
               </p:cTn>
              </p:par>
             </p:childTnLst>
            </p:cTn>
           </p:par>"""

_TIMING_TAIL = """          </p:childTnLst>
         </p:cTn>
        </p:par>
       </p:childTnLst>
       <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
       <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
      </p:cTn>
     </p:seq>
    </p:childTnLst>
   </p:cTn>
  </p:par>
 </p:tnLst>
</p:timing>"""

_MIME_BY_SUFFIX = {".mp3": "audio/mpeg", ".m4a": "audio/mp4", ".wav": "audio/wav"}


def _build_timing_xml(spid: int, dur_ms: int, anims=None) -> str:
    """Assemble the slide timing tree: media autoplay + optional per-shape fade
    entrances, all as auto-firing siblings of one group node. `anims` is a list of
    (shape_id, delay_ms)."""
    # ids 1-3 are the header (tmRoot, mainSeq, group). Media child uses 4-6.
    pars = _TIMING_MEDIA.format(i0=4, i1=5, i2=6, spid=spid, dur=dur_ms)
    nid = 7
    for shape_id, delay_ms in (anims or []):
        pars += "\n" + _TIMING_ENTRANCE.format(
            i0=nid, i1=nid + 1, i2=nid + 2, i3=nid + 3,
            delay=int(delay_ms), spid=shape_id,
        )
        nid += 4
    return _TIMING_HEAD + "\n" + pars + "\n" + _TIMING_TAIL


def _embed_autoplay_audio(slide, audio_path: Path, anims=None):
    mime = _MIME_BY_SUFFIX.get(audio_path.suffix.lower(), "audio/mpeg")
    movie = slide.shapes.add_movie(
        str(audio_path),
        Inches(12.7), Inches(0.15), Inches(0.45), Inches(0.45),
        mime_type=mime,
    )
    spid = movie.shape_id
    duration_ms = int(audio_duration_seconds(audio_path) * 1000)
    try:
        timing = etree.fromstring(_build_timing_xml(spid, duration_ms, anims).encode())
    except Exception:
        # Never let an animation glitch corrupt the deck — fall back to plain audio.
        timing = etree.fromstring(_build_timing_xml(spid, duration_ms, None).encode())
    _insert_slide_child(slide, timing)


def _set_auto_advance(slide, seconds: float):
    """Advance to the next slide automatically once the narration finishes."""
    adv_ms = int(seconds * 1000)
    transition = etree.fromstring(
        f'<p:transition xmlns:p="{_P}" spd="med" advClick="1" advTm="{adv_ms}"><p:fade/></p:transition>'.encode()
    )
    _insert_slide_child(slide, transition)


# CT_Slide child order (the ones we touch): cSld, clrMapOvr, transition, timing
_ORDER = [qn("p:cSld"), qn("p:clrMapOvr"), qn("p:transition"), qn("p:timing")]


def _insert_slide_child(slide, element):
    root = slide.element
    existing = root.find(element.tag)
    if existing is not None:
        root.remove(existing)
    rank = _ORDER.index(element.tag)
    for child in root:
        if child.tag in _ORDER and _ORDER.index(child.tag) > rank:
            child.addprevious(element)
            return
    root.append(element)
