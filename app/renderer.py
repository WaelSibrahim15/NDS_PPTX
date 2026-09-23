"""NDS — render deck slides as 1920x1080 PNG frames with Pillow.

Draws the same layouts as the PPTX builder (both read app/design.py), so the
review thumbnails and the MP4 export match the deck without needing
LibreOffice or PowerPoint installed.

Scale: the deck is 13.333in wide -> 1920px, so 1in = 144px and 1pt = 2px.
Shapes are drawn at 2x and scaled down for smooth edges.
"""
from functools import lru_cache
from pathlib import Path
from typing import List

from PIL import Image, ImageDraw, ImageFont

from . import design
from .models import DeckPlan

W, H = 1920, 1080
SS = 2                    # supersampling factor
IN = 144 * SS             # pixels per inch while drawing
PT = 2 * SS               # pixels per point while drawing

# Real Arial / Georgia when present (macOS), else the bundled metric-compatible
# Liberation fonts (SIL OFL), so Linux hosts such as Railway render properly.
_MAC = Path("/System/Library/Fonts/Supplemental")
_BUNDLED = design.ASSETS / "fonts"
FONT_CANDIDATES = {
    ("sans", False, False): [_MAC / "Arial.ttf", _BUNDLED / "LiberationSans-Regular.ttf"],
    ("sans", True, False): [_MAC / "Arial Bold.ttf", _BUNDLED / "LiberationSans-Bold.ttf"],
    ("serif", False, False): [_MAC / "Georgia.ttf", _BUNDLED / "LiberationSerif-Regular.ttf"],
    ("serif", False, True): [_MAC / "Georgia Italic.ttf", _BUNDLED / "LiberationSerif-Italic.ttf"],
}


@lru_cache(maxsize=128)
def _font(family: str, bold: bool, italic: bool, px: int) -> ImageFont.FreeTypeFont:
    key = (family, bold, italic)
    if key not in FONT_CANDIDATES:
        key = (family, bold, False) if (family, bold, False) in FONT_CANDIDATES else (family, False, False)
    for path in FONT_CANDIDATES.get(key, FONT_CANDIDATES[("sans", False, False)]):
        try:
            return ImageFont.truetype(str(path), px)
        except OSError:
            continue
    return ImageFont.load_default(px)


def _rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _px(v: float) -> int:
    return int(round(v * IN))


# ------------------------------------------------------------------ text

def _layout_para(draw, para: "design.Para", max_w: int):
    """Wrap a paragraph of styled runs. Returns [(line_height, [(x, text, font, color)])]."""
    words = []  # (text, font, color, trailing_space)
    for run in para.runs:
        font = _font(run.font, run.bold, run.italic, max(1, int(run.size * PT)))
        parts = run.text.split(" ")
        for i, w in enumerate(parts):
            if w == "" and i not in (0, len(parts) - 1):
                continue
            words.append((w, font, _rgb(run.color), i < len(parts) - 1, run.size))
    lines, cur, x = [], [], 0
    size_max = 0
    space_cache = {}
    for text, font, color, space, size in words:
        sp = space_cache.setdefault(id(font), draw.textlength(" ", font=font))
        w = draw.textlength(text, font=font)
        if cur and x + w > max_w and text:
            lines.append((size_max, cur))
            cur, x, size_max = [], 0, 0
        if text or cur:
            cur.append((x, text, font, color))
        x += w + (sp if space else 0)
        size_max = max(size_max, size)
    if cur:
        lines.append((size_max, cur))
    return [(int(size * PT * 1.2 * para.line_spacing), items) for size, items in lines]


def _draw_text(draw, el: "design.Text"):
    x0, y0, bw, bh = _px(el.x), _px(el.y), _px(el.w), _px(el.h)
    blocks = []
    total = 0
    for para in el.paras:
        lines = _layout_para(draw, para, bw)
        gap = int(para.space_after * PT)
        blocks.append((para, lines, gap))
        total += sum(h for h, _ in lines) + gap
    if blocks:
        total -= blocks[-1][2]
    y = y0 + {"top": 0, "middle": (bh - total) // 2, "bottom": bh - total}[el.anchor]
    for para, lines, gap in blocks:
        for lh, items in lines:
            line_w = (items[-1][0] + draw.textlength(items[-1][1], font=items[-1][2])) if items else 0
            dx = {"left": 0, "center": (bw - line_w) / 2, "right": bw - line_w}[para.align]
            for x, text, font, color in items:
                ascent = font.getmetrics()[0]
                # baseline sits ~80% down the line box, like PowerPoint's single spacing
                draw.text((x0 + dx + x, y + int(lh * 0.8) - ascent), text, font=font, fill=color)
            y += lh
        y += gap


# ------------------------------------------------------------------ shapes

def _draw_element(img: Image.Image, draw: ImageDraw.ImageDraw, el):
    if isinstance(el, design.Rect):
        box = [_px(el.x), _px(el.y), _px(el.x + el.w), _px(el.y + el.h)]
        if el.radius:
            draw.rounded_rectangle(box, radius=_px(el.radius), fill=_rgb(el.fill))
        else:
            draw.rectangle(box, fill=_rgb(el.fill))
    elif isinstance(el, design.Oval):
        box = [_px(el.cx - el.r), _px(el.cy - el.r), _px(el.cx + el.r), _px(el.cy + el.r)]
        if el.fill:
            draw.ellipse(box, fill=_rgb(el.fill))
        if el.line:
            draw.ellipse(box, outline=_rgb(el.line), width=max(1, int(el.line_w * PT)))
    elif isinstance(el, design.Arc):
        box = [_px(el.cx - el.r), _px(el.cy - el.r), _px(el.cx + el.r), _px(el.cy + el.r)]
        draw.arc(box, el.start, el.end, fill=_rgb(el.color), width=max(1, int(el.line_w * PT)))
    elif isinstance(el, design.Line):
        draw.line([_px(el.x1), _px(el.y1), _px(el.x2), _px(el.y2)], fill=_rgb(el.color),
                  width=max(1, int(el.line_w * PT)))
    elif isinstance(el, design.Image):
        pic = _load_image(str(el.path), _px(el.w), _px(el.h))
        img.paste(pic, (_px(el.x), _px(el.y)), pic)
    elif isinstance(el, design.Text):
        _draw_text(draw, el)


@lru_cache(maxsize=64)
def _load_image(path: str, w: int, h: int) -> Image.Image:
    return Image.open(path).convert("RGBA").resize((max(1, w), max(1, h)), Image.Resampling.LANCZOS)


# ------------------------------------------------------------------ public API

def render_slide(plan: DeckPlan, index: int, template: str = "niq") -> Image.Image:
    """Render one slide at full 1920×1080 resolution."""
    if not 0 <= index < len(plan.slides):
        raise IndexError(f"slide index {index} out of range")
    scene = design.layout_slide(plan, index, template)
    img = Image.new("RGB", (W * SS, H * SS), _rgb(scene.background))
    draw = ImageDraw.Draw(img)
    for el in scene.elements:
        _draw_element(img, draw, el)
    return img.resize((W, H), Image.Resampling.LANCZOS)


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
