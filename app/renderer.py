"""NDS — render deck slides as 1920x1080 PNG frames with Pillow.

Mirrors the pptx_builder layouts (NIQ 2026 / neutral) so the MP4 export looks
like the deck without needing LibreOffice or PowerPoint installed.

Scale: the deck is 13.333in wide -> 1920px, so 1in = 144px and 1pt = 2px.
"""
from pathlib import Path
from typing import List

from PIL import Image, ImageDraw, ImageFont

from .models import DeckPlan, Slide

W, H = 1920, 1080
IN = 144          # pixels per inch
PT = 2            # pixels per font point

FONT_DIR = Path("/System/Library/Fonts/Supplemental")
FONTS = {
    ("arial", False): "Arial.ttf",
    ("arial", True): "Arial Bold.ttf",
    ("georgia_italic", False): "Georgia Italic.ttf",
}

# Mirrors pptx_builder._palette — official NIQ 2026 brand hexes.
PALETTES = {
    "niq": {
        "dark": (0x06, 0x0A, 0x45), "accent": (0x2C, 0x6D, 0xF6), "accent2": (0xEF, 0x5F, 0x17),
        "cyan": (0x31, 0xD1, 0xFF), "green": (0x59, 0xAD, 0x00), "pink": (0xEF, 0x58, 0x90),
        "amber": (0xFF, 0xB5, 0x00),
        "body": (0x55, 0x55, 0x55), "white": (255, 255, 255), "grey": (154, 154, 154),
        "footer": (176, 176, 216),
    },
    "neutral": {
        "dark": (31, 42, 55), "accent": (47, 111, 237), "accent2": (217, 119, 6),
        "cyan": (56, 178, 196), "green": (47, 158, 68), "pink": (194, 65, 122),
        "amber": (217, 158, 6),
        "body": (64, 64, 64), "white": (255, 255, 255), "grey": (154, 154, 154),
        "footer": (170, 180, 195),
    },
}


def _font(kind: str, bold: bool, size_pt: int) -> ImageFont.FreeTypeFont:
    name = FONTS.get((kind, bold), "Arial.ttf")
    path = FONT_DIR / name
    try:
        return ImageFont.truetype(str(path), size_pt * PT)
    except OSError:
        return ImageFont.load_default(size_pt * PT)


def _wrap(draw: ImageDraw.ImageDraw, text: str, font, max_w: int) -> List[str]:
    lines, line = [], ""
    for word in text.split():
        trial = (line + " " + word).strip()
        if draw.textlength(trial, font=font) <= max_w:
            line = trial
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines or [""]


def _text(draw, xy, text, *, kind="arial", bold=False, size=18, color, max_w=None, line_gap=0.35):
    font = _font(kind, bold, size)
    x, y = xy
    lines = _wrap(draw, text, font, max_w) if max_w else [text]
    lh = int(size * PT * (1 + line_gap))
    for ln in lines:
        draw.text((x, y), ln, font=font, fill=color)
        y += lh
    return y


def _circle(draw, cx, cy, r, color, width=0):
    box = [cx - r, cy - r, cx + r, cy + r]
    if width:
        draw.ellipse(box, outline=color, width=width)
    else:
        draw.ellipse(box, fill=color)


def render_slide(plan: DeckPlan, index: int, template: str = "niq") -> Image.Image:
    """Render one slide at full 1920×1080 resolution."""
    if not 0 <= index < len(plan.slides):
        raise IndexError(f"slide index {index} out of range")
    c = PALETTES.get(template, PALETTES["niq"])
    total = len(plan.slides)
    spec = plan.slides[index]
    img = Image.new("RGB", (W, H), c["white"])
    draw = ImageDraw.Draw(img)
    if spec.layout == "title":
        _title(draw, plan, spec, c)
    elif spec.layout == "section":
        _section(draw, spec, c)
    elif spec.layout == "closing":
        _closing(draw, spec, c)
    elif spec.layout == "cards" and spec.cards:
        _cards(draw, spec, c, index + 1, total)
    elif spec.layout == "stats" and spec.stats:
        _stats(draw, spec, c, index + 1, total)
    elif spec.layout == "compare" and (spec.compare_left or spec.compare_right):
        _compare(draw, spec, c, index + 1, total)
    else:
        _content(draw, spec, c, index + 1, total)
    return img


def render_slide_preview(plan: DeckPlan, index: int, template: str = "niq",
                         width: int = 640) -> Image.Image:
    """Thumbnail for the review UI — same layout as the final deck, scaled down."""
    img = render_slide(plan, index, template=template)
    h = max(1, int(round(img.height * (width / img.width))))
    return img.resize((width, h), Image.Resampling.LANCZOS)


def render_frames(plan: DeckPlan, out_dir: Path, template: str = "niq") -> List[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    frames = []
    for i in range(len(plan.slides)):
        path = out_dir / f"frame_{i + 1:03d}.png"
        render_slide(plan, i, template=template).save(str(path))
        frames.append(path)
    return frames


def _title(draw, plan: DeckPlan, spec: Slide, c):
    draw.rectangle([0, 0, W, H], fill=c["dark"])
    _circle(draw, int(9.2 * IN + 2.3 * IN), int(1.4 * IN + 2.3 * IN), int(2.3 * IN), c["accent"], width=6)
    _circle(draw, int(10.6 * IN + 0.75 * IN), int(4.6 * IN + 0.75 * IN), int(0.75 * IN), c["cyan"])
    _circle(draw, int(8.7 * IN + 0.35 * IN), int(5.4 * IN + 0.35 * IN), int(0.35 * IN), c["white"], width=4)
    _text(draw, (int(0.9 * IN), int(2.3 * IN)), spec.title or plan.deck_title,
          bold=True, size=40, color=c["white"], max_w=int(8.0 * IN))
    _text(draw, (int(0.9 * IN), int(4.3 * IN)), spec.subtitle or plan.subtitle,
          kind="georgia_italic", size=20, color=c["cyan"], max_w=int(7.6 * IN))
    _text(draw, (int(0.9 * IN), int(6.7 * IN)), "Narrated presentation · generated with NDS",
          size=11, color=c["footer"])


def _section(draw, spec: Slide, c):
    draw.rectangle([0, 0, W, H], fill=c["accent"])
    _circle(draw, int(-1.6 * IN + 2.6 * IN), int(4.4 * IN + 2.6 * IN), int(2.6 * IN), c["dark"])
    _circle(draw, int(11.6 * IN + 1.7 * IN), int(-1.2 * IN + 1.7 * IN), int(1.7 * IN), c["accent2"], width=6)
    _text(draw, (int(1.0 * IN), int(2.9 * IN)), spec.title, bold=True, size=36,
          color=c["white"], max_w=int(11.0 * IN))
    if spec.subtitle:
        _text(draw, (int(1.0 * IN), int(4.5 * IN)), spec.subtitle, kind="georgia_italic",
              size=18, color=c["white"], max_w=int(10.0 * IN))


GREY_TXT = (85, 85, 85)
CARD_TINT = (244, 246, 254)
LABEL_TINT = (198, 208, 255)


def _rounded(draw, box, color, radius=28):
    draw.rounded_rectangle(box, radius=radius, fill=color)


def _chrome(draw, spec: Slide, c, page_no: int, total: int):
    """Shared header/footer for content-style slides (mirrors pptx_builder._header)."""
    _circle(draw, int(12.35 * IN + 0.225 * IN), int(6.75 * IN + 0.225 * IN), int(0.225 * IN), c["accent"], width=4)
    _text(draw, (int(0.65 * IN), int(0.42 * IN)), spec.title, bold=True, size=24,
          color=c["dark"], max_w=int(11.9 * IN))
    if spec.subtitle:
        _text(draw, (int(0.65 * IN), int(1.05 * IN)), spec.subtitle, size=13,
              color=GREY_TXT, max_w=int(11.9 * IN))
    _text(draw, (int(0.65 * IN), int(6.98 * IN)), f"{page_no} / {total}", size=10, color=c["grey"])


def _content(draw, spec: Slide, c, page_no: int, total: int):
    _chrome(draw, spec, c, page_no, total)
    y = int(1.85 * IN)
    for b in spec.bullets:
        _circle(draw, int(0.65 * IN) + 12, y + 18, 8, c["accent"])
        y = _text(draw, (int(0.65 * IN) + 40, y), b, size=16, color=c["body"], max_w=int(11.2 * IN))
        y += 32
    _text(draw, (int(0.65 * IN), int(6.98 * IN)), "", size=10, color=c["grey"])


def _cards(draw, spec: Slide, c, page_no: int, total: int):
    _chrome(draw, spec, c, page_no, total)
    cards = spec.cards[:4]
    n = len(cards)
    accents = [c["accent"], c["accent2"], c["dark"], c["accent"]]
    gap = int(0.3 * IN)
    total_w = int(12.0 * IN)
    card_w = (total_w - gap * (n - 1)) // n
    top, card_h = int(1.9 * IN), int(4.3 * IN)
    for i, card in enumerate(cards):
        left = int(0.65 * IN) + i * (card_w + gap)
        _rounded(draw, [left, top, left + card_w, top + card_h], CARD_TINT)
        pad = int(0.28 * IN)
        _text(draw, (left + pad, top + int(0.25 * IN)), f"{i + 1:02d}", bold=True, size=28,
              color=accents[i % len(accents)])
        _text(draw, (left + pad, top + int(1.0 * IN)), card.title, bold=True, size=16,
              color=c["dark"], max_w=card_w - pad * 2)
        _text(draw, (left + pad, top + int(1.95 * IN)), card.desc, size=11,
              color=GREY_TXT, max_w=card_w - pad * 2)


def _stats(draw, spec: Slide, c, page_no: int, total: int):
    _chrome(draw, spec, c, page_no, total)
    stats = spec.stats[:5]
    n = len(stats)
    panel_top, panel_h = int(2.3 * IN), int(3.2 * IN)
    left0 = int(0.65 * IN)
    _rounded(draw, [left0, panel_top, left0 + int(12.0 * IN), panel_top + panel_h], c["dark"])
    col_w = int(12.0 * IN) // n
    for i, st in enumerate(stats):
        left = left0 + i * col_w
        _text(draw, (left + int(0.25 * IN), panel_top + int(0.85 * IN)), st.value, bold=True,
              size=30, color=c["white"] if i % 2 == 0 else c["cyan"], max_w=col_w - int(0.5 * IN))
        _text(draw, (left + int(0.25 * IN), panel_top + int(1.85 * IN)), st.label, size=11,
              color=LABEL_TINT, max_w=col_w - int(0.5 * IN))


def _compare(draw, spec: Slide, c, page_no: int, total: int):
    _chrome(draw, spec, c, page_no, total)
    panels = [(spec.compare_left, c["accent"]), (spec.compare_right, c["accent2"])]
    top, head_h, body_h = int(1.9 * IN), int(0.6 * IN), int(4.0 * IN)
    width = int(5.85 * IN)
    for i, (side, accent) in enumerate(panels):
        if side is None:
            continue
        left = int(0.65 * IN) + i * (width + int(0.3 * IN))
        draw.rectangle([left, top, left + width, top + head_h], fill=accent)
        _text(draw, (left + int(0.25 * IN), top + int(0.13 * IN)), side.heading, bold=True,
              size=14, color=c["white"], max_w=width - int(0.5 * IN))
        draw.rectangle([left, top + head_h, left + width, top + head_h + body_h], fill=CARD_TINT)
        y = top + head_h + int(0.2 * IN)
        for item in side.items:
            _circle(draw, left + int(0.25 * IN) + 8, y + 14, 7, accent)
            y = _text(draw, (left + int(0.25 * IN) + 32, y), item, size=13, color=c["body"],
                      max_w=width - int(0.75 * IN))
            y += 20


def _closing(draw, spec: Slide, c):
    draw.rectangle([0, 0, W, H], fill=c["dark"])
    _circle(draw, int(10.4 * IN + 2.2 * IN), int(-1.8 * IN + 2.2 * IN), int(2.2 * IN), c["accent"], width=6)
    _circle(draw, int(0.4 * IN + 0.55 * IN), int(5.9 * IN + 0.55 * IN), int(0.55 * IN), c["cyan"])
    _text(draw, (int(0.9 * IN), int(2.2 * IN)), spec.title, bold=True, size=34,
          color=c["white"], max_w=int(10.5 * IN))
    y = int(3.8 * IN)
    for b in spec.bullets:
        _circle(draw, int(0.9 * IN) + 12, y + 16, 8, c["accent2"])
        y = _text(draw, (int(0.9 * IN) + 40, y), b, size=16, color=c["white"], max_w=int(10.0 * IN))
        y += 20
