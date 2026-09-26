"""NDS — slide design, defined once and drawn twice.

Each layout returns a list of simple drawing elements (rectangles, circles,
arcs, lines, images, text) in inches on the 13.333 x 7.5in slide. The PPTX
builder turns them into PowerPoint shapes and the renderer draws them with
Pillow for previews and MP4 frames, so both always show the same design.

The layouts follow the NIQ design system published in Claude Design (rounded
rectangles and circles, White / Blue / Dark grounds, logo plus legal line in
the footer, brand symbols on feature cards). Colours, fonts, logos, footer and
symbols come from the template (themes.py): the built-in "niq" and "neutral",
or one imported from a Claude Design export.
"""
from dataclasses import dataclass, field
import zlib
from functools import lru_cache
from pathlib import Path
from typing import List, Optional

from . import themes
from .models import DeckPlan, Slide

ASSETS = themes.ASSETS

W, H = 13.333, 7.5
M = 0.39          # grid-margin (37.5px on the 1280px canvas)


# ------------------------------------------------------------------ elements

@dataclass
class Run:
    text: str
    size: float                  # points
    color: str                   # "#RRGGBB"
    bold: bool = False
    italic: bool = False
    font: str = "sans"           # "sans" (Arial) or "serif" (Georgia)


@dataclass
class Para:
    runs: List[Run]
    align: str = "left"          # left | center | right
    space_after: float = 0       # points
    line_spacing: float = 1.0    # multiple of single spacing


@dataclass
class Text:
    x: float
    y: float
    w: float
    h: float
    paras: List[Para]
    anchor: str = "top"          # top | middle | bottom


@dataclass
class Rect:
    x: float
    y: float
    w: float
    h: float
    fill: str
    radius: float = 0.0          # inches


@dataclass
class Oval:
    cx: float
    cy: float
    r: float
    fill: Optional[str] = None
    line: Optional[str] = None
    line_w: float = 0.0          # points


@dataclass
class Arc:
    cx: float
    cy: float
    r: float
    start: float                 # degrees, clockwise from 3 o'clock
    end: float
    color: str
    line_w: float                # points


@dataclass
class Line:
    x1: float
    y1: float
    x2: float
    y2: float
    color: str
    line_w: float = 0.75


@dataclass
class Image:
    path: Path
    x: float
    y: float
    w: float
    h: float


@dataclass
class Scene:
    background: str
    elements: list = field(default_factory=list)
    fonts: dict = field(default_factory=lambda: {"sans": "Arial", "serif": "Georgia"})
    font_files: dict = field(default_factory=dict)   # sans / sans_bold / serif / serif_italic → Path

    def add(self, *els):
        self.elements.extend(els)


# ------------------------------------------------------------------ helpers

@lru_cache(maxsize=64)
def _aspect(path: Path) -> float:
    from PIL import Image as PILImage
    with PILImage.open(path) as im:
        return im.width / max(1, im.height)


def _logo(path: Optional[Path], x: float, y: float, h: float, max_w: float):
    """Logo image at (x, y), height h (shrunk to fit max_w); returns (element, width)."""
    if not path:
        return None, 0.0
    w = h * _aspect(path)
    if w > max_w:
        h, w = h * max_w / w, max_w
    return Image(path, x, y, w, h), w


def _p(text, size, color, **kw) -> Para:
    run_kw = {k: kw.pop(k) for k in ("bold", "italic", "font") if k in kw}
    return Para([Run(text, size, color, **run_kw)], **kw)


def _fit(text: str, sizes) -> float:
    """Pick a font size by text length: sizes is [(max_chars, pt), ...], last wins."""
    for max_chars, pt in sizes:
        if len(text) <= max_chars:
            return pt
    return sizes[-1][1]


# ------------------------------------------------------------------ chrome

def _footer(sc: Scene, c: dict, ground: str, page: Optional[str]):
    """NIQ mark bottom-left, legal line, page number, above a hairline rule."""
    on_color = ground != "light"
    text_col = c["white"] if on_color else c["ink"]
    rule_col = {"light": c["hairline"], "blue": c["tint"], "dark": c["bright"]}[ground]
    sc.add(Line(M, 6.98, W - M, 6.98, rule_col, 0.75))
    x = M
    logo, lw = _logo(c["logo_dark"] if on_color else c["logo_light"], M, 7.1, 0.17, 1.6)
    if logo:
        sc.add(logo)
        x = M + lw + 0.18
    if c["footer"]:
        sc.add(Text(x, 7.08, 6.0, 0.22, [_p(c["footer"], 7, text_col)], anchor="middle"))
    if page:
        sc.add(Text(W - M - 1.5, 7.08, 1.5, 0.22, [_p(page, 7, text_col, align="right")],
                    anchor="middle"))


def _header(sc: Scene, spec: Slide, c: dict):
    """Title (Arial bold) + kicker (Georgia italic) for full-width content slides."""
    size = _fit(spec.title, [(60, 26), (95, 22), (999, 19)])
    sc.add(Text(M, 0.45, W - 2 * M, 1.0,
                [_p(spec.title, size, c["deep"], bold=True, line_spacing=0.9)], anchor="bottom"))
    if spec.subtitle:
        sc.add(Text(M, 1.55, W - 2 * M, 0.5,
                    [_p(spec.subtitle, 16, c["bright"], italic=True, font="serif")]))


# ------------------------------------------------------------------ layouts

def _title_layout(sc: Scene, plan: DeckPlan, spec: Slide, c: dict):
    """Cover: title bottom-left on White, the circle motif on the right."""
    al = 7.2
    sc.add(
        Rect(al + 0.5, 0, 2.4, H, c["bright"]),
        Rect(al, 5.35, 0.5, H - 5.35, c["tint"]),
        Oval(al + 3.9, 3.75, 2.85, line=c["light"], line_w=7),
        Oval(al + 3.9, 3.75, 2.35, fill=c["deep"]),
        Arc(al + 3.9, 3.75, 1.2, 270, 360, c["tint"], 7),
        Arc(al + 3.9, 3.75, 1.65, 270, 360, c["tint"], 7),
        Oval(al + 1.7, 1.45, 0.6, fill=c["light"]),
        Oval(al + 1.7, 6.2, 0.28, fill=c["orange"]),
    )
    logo, _ = _logo(c["logo_light"], M, 0.5, 0.34, 2.6)
    if logo:
        sc.add(logo)
    title = spec.title or plan.deck_title
    size = _fit(title, [(28, 54), (55, 44), (90, 36), (999, 30)])
    sc.add(Text(M, 1.5, 6.4, 4.0, [_p(title, size, c["deep"], bold=True, line_spacing=0.9)],
                anchor="bottom"))
    sub = spec.subtitle or plan.subtitle
    if sub:
        sc.add(Text(M, 5.7, 6.3, 1.1, [_p(sub, 18, c["ink"], line_spacing=1.1)]))
    if c["footer"]:
        sc.add(Text(M, 7.08, 6.0, 0.22, [_p(c["footer"], 7, c["ink"])], anchor="middle"))


def _section_layout(sc: Scene, spec: Slide, c: dict, number: int, page: str):
    """Section divider on a Blue or Dark ground (alternating), arc motif top-right."""
    ground = "blue" if number % 2 == 1 else "dark"
    sc.background = c["bright"] if ground == "blue" else c["deep"]
    alt = c["deep"] if ground == "blue" else c["bright"]
    sc.add(Oval(W - 0.4, 0.3, 2.9, fill=alt),
           Oval(W - 0.4, 0.3, 3.45, line=c["light"] if ground == "dark" else c["tint"], line_w=5))
    sc.add(Text(M, 0.45, 2.0, 0.4, [_p(f"{number:02d}", 14, c["white"], bold=True)]))
    size = _fit(spec.title, [(30, 44), (60, 36), (999, 30)])
    sc.add(Text(M, 1.9, 8.0, 3.0, [_p(spec.title, size, c["white"], bold=True, line_spacing=0.9)],
                anchor="bottom"))
    if spec.subtitle:
        sc.add(Text(M, 5.1, 8.0, 1.0, [_p(spec.subtitle, 18, c["light"] if ground == "dark"
                                          else c["white"], italic=True, font="serif")]))
    _footer(sc, c, ground, page)


def _content_layout(sc: Scene, spec: Slide, c: dict, page: str):
    """Sidebar plus content: Georgia statement title on grey, bullets as rows."""
    side_w = 4.45
    sc.add(Rect(0, 0, side_w, H, c["panel"]))
    size = _fit(spec.title, [(45, 32), (80, 26), (999, 22)])
    sc.add(Text(M, 0.7, side_w - M - 0.4, 3.5,
                [_p(spec.title, size, c["deep"], font="serif", line_spacing=0.95)],
                anchor="bottom"))
    sc.add(Rect(M, 4.4, 0.6, 0.06, c["bright"]))
    if spec.subtitle:
        sc.add(Text(M, 4.65, side_w - M - 0.4, 1.9, [_p(spec.subtitle, 16, c["ink"], line_spacing=1.1)]))

    bullets = spec.bullets[:6]
    if not bullets:
        _footer(sc, c, "light", page)
        return
    x0, x1 = side_w + 0.6, W - M
    row_h = min(1.05, 5.8 / len(bullets))
    y = 0.6 + (5.9 - row_h * len(bullets)) / 2
    for i, b in enumerate(bullets):
        mid = y + row_h / 2
        sc.add(Oval(x0 + 0.11, mid, 0.11, fill=c["bright"]))
        sc.add(Text(x0 + 0.45, y, x1 - x0 - 0.45, row_h, [_p(b, 18, c["ink"], line_spacing=1.1)],
                    anchor="middle"))
        if i < len(bullets) - 1:
            sc.add(Line(x0, y + row_h, x1, y + row_h, c["hairline"], 0.5))
        y += row_h
    _footer(sc, c, "light", page)


def _cards_layout(sc: Scene, spec: Slide, c: dict, page: str):
    """Feature cards: rounded Bright Blue header bar with label and NIQ symbol."""
    _header(sc, spec, c)
    cards = spec.cards[:4]
    n = len(cards)
    gap = 0.3
    top = 2.45
    cw = (W - 2 * M - gap * (n - 1)) / n
    head_h = 0.7
    label_size = 16 if n <= 3 else 14
    for i, card in enumerate(cards):
        x = M + i * (cw + gap)
        sc.add(Rect(x, top, cw, head_h, c["bright"], radius=0.11))
        icon = (c["symbols_dir"] / f"{card.icon}.png"
                if c["symbols_dir"] and getattr(card, "icon", None) else None)
        has_icon = bool(icon and icon.exists())
        label_w = cw - 0.2 - (0.7 if has_icon else 0.2)
        sc.add(Text(x + 0.2, top, label_w, head_h,
                    [_p(card.title, label_size, c["white"], bold=True, line_spacing=0.95)],
                    anchor="middle"))
        if has_icon:
            d = 0.52
            cx, cy = x + cw - 0.14 - d / 2, top + head_h / 2
            sc.add(Oval(cx, cy, d / 2, fill=c["white"]))
            s = d * 0.8
            sc.add(Image(icon, cx - s / 2, cy - s / 2, s, s))
        sc.add(Text(x, top + head_h + 0.25, cw, 3.4,
                    [_p(card.desc, 16 if n <= 3 else 14, c["ink"], line_spacing=1.15)]))
    _footer(sc, c, "light", page)


def _stats_layout(sc: Scene, spec: Slide, c: dict, page: str):
    """One metric callout carries the message; the other figures sit beside it."""
    _header(sc, spec, c)
    stats = spec.stats[:5]
    top, box_h, box_w = 2.45, 4.1, 4.6
    lead = stats[0]
    sc.add(Rect(M, top, box_w, box_h, c["bright"], radius=0.24))
    sc.add(Text(M + 0.4, top + 0.35, box_w - 0.8, 2.0,
                [_p(lead.value, _fit(lead.value, [(5, 60), (8, 48), (999, 36)]), c["white"],
                    bold=True, line_spacing=0.9)], anchor="bottom"))
    sc.add(Text(M + 0.4, top + 2.55, box_w - 0.8, 1.3, [_p(lead.label, 18, c["white"], bold=True)]))

    others = stats[1:]
    if not others:
        _footer(sc, c, "light", page)
        return
    x0 = M + box_w + 0.5
    area_w = W - M - x0
    cols = 2 if len(others) > 2 else len(others)
    rows = (len(others) + cols - 1) // cols
    cell_w, cell_h = area_w / cols, box_h / rows
    for i, st in enumerate(others):
        x = x0 + (i % cols) * cell_w
        y = top + (i // cols) * cell_h
        sc.add(Line(x, y, x + cell_w - 0.3, y, c["hairline"], 0.75))
        sc.add(Text(x, y + 0.2, cell_w - 0.3, 0.95,
                    [_p(st.value, _fit(st.value, [(7, 40), (12, 32), (999, 26)]), c["bright"],
                        bold=True)]))
        sc.add(Text(x, y + 1.1, cell_w - 0.3, cell_h - 1.2, [_p(st.label, 16, c["ink"])]))
    _footer(sc, c, "light", page)


def _bullet_paras(items, size, text_col, dot_col, space_after=10):
    return [Para([Run("●  ", size * 0.7, dot_col), Run(it, size, text_col)],
                 space_after=space_after, line_spacing=1.1) for it in items]


def _compare_layout(sc: Scene, spec: Slide, c: dict, page: str):
    """Half-and-half panels: Gray on the left, Blue on the right."""
    _header(sc, spec, c)
    top, h = 2.45, 4.25
    pw = (W - 2 * M - 0.3) / 2
    sides = [(spec.compare_left, c["panel"], c["deep"], c["ink"], c["bright"]),
             (spec.compare_right, c["bright"], c["white"], c["white"], c["white"])]
    sides = [sd for sd in sides if sd[0] is not None]
    if len(sides) == 1:
        pw = W - 2 * M
    for i, (side, fill, head_col, text_col, dot_col) in enumerate(sides):
        x = M + i * (pw + 0.3)
        sc.add(Rect(x, top, pw, h, fill, radius=0.24))
        sc.add(Text(x + 0.4, top + 0.35, pw - 0.8, 0.55,
                    [_p(side.heading, 22, head_col, bold=True)], anchor="middle"))
        sc.add(Text(x + 0.4, top + 1.15, pw - 0.8, h - 1.4,
                    _bullet_paras(side.items[:6], 16, text_col, dot_col)))
    _footer(sc, c, "light", page)


def _closing_layout(sc: Scene, spec: Slide, c: dict):
    """Dark ground, circle motif top-right, Georgia italic sign-off."""
    sc.background = c["deep"]
    sc.add(Oval(W - 1.3, 0.9, 2.3, fill=c["bright"]),
           Oval(W - 1.3, 0.9, 2.85, line=c["light"], line_w=6),
           Oval(W - 4.35, 3.55, 0.2, fill=c["orange"]))
    size = _fit(spec.title, [(30, 44), (60, 36), (999, 30)])
    sc.add(Text(M, 1.2, 8.2, 2.4, [_p(spec.title, size, c["white"], bold=True, line_spacing=0.9)],
                anchor="bottom"))
    y = 3.85
    if spec.subtitle:
        sc.add(Text(M, y, 8.2, 0.9, [_p(spec.subtitle, 22, c["light"], italic=True, font="serif")]))
        y += 1.0
    if spec.bullets:
        sc.add(Text(M, y, 8.2, 6.7 - y, _bullet_paras(spec.bullets[:4], 18, c["white"], c["light"])))
    _footer(sc, c, "dark", None)


# ------------------------------------------------------------------ variants
# Each layout has a few visual versions. NDS picks one per slide so a layout used
# twice looks different each time, seeded by the deck title so two decks start on
# different versions. A slide's own "variant" (set by Claude or the user) wins.

VARIANTS = {"title": 3, "section": 2, "content": 3, "cards": 3, "stats": 3, "compare": 2,
            "closing": 2, "quote": 2, "statement": 2, "agenda": 1, "timeline": 2, "chart": 2}


def pick_variant(plan: DeckPlan, index: int) -> int:
    spec = plan.slides[index]
    n = VARIANTS.get(spec.layout, 1)
    if spec.variant is not None:
        return spec.variant % n
    seed = zlib.crc32((plan.deck_title or "").encode("utf-8"))
    if spec.layout in ("title", "closing"):
        return seed % n
    seen = sum(1 for s in plan.slides[:index] if s.layout == spec.layout)
    return (seen + seed) % n


def _title_dark(sc: Scene, plan: DeckPlan, spec: Slide, c: dict):
    """Cover on the dark ground, big circle bottom-right."""
    sc.background = c["deep"]
    sc.add(Oval(W - 2.3, H - 1.0, 3.3, fill=c["bright"]),
           Oval(W - 2.3, H - 1.0, 3.85, line=c["light"], line_w=6),
           Arc(W - 2.3, H - 1.0, 2.1, 180, 270, c["tint"], 6),
           Oval(W - 6.3, 2.2, 0.22, fill=c["orange"]))
    logo, _ = _logo(c["logo_dark"], M, 0.5, 0.34, 2.6)
    if logo:
        sc.add(logo)
    title = spec.title or plan.deck_title
    size = _fit(title, [(28, 54), (55, 44), (90, 36), (999, 30)])
    sc.add(Text(M, 1.4, 7.4, 3.6, [_p(title, size, c["white"], bold=True, line_spacing=0.9)],
                anchor="bottom"))
    sub = spec.subtitle or plan.subtitle
    if sub:
        sc.add(Text(M, 5.2, 7.0, 1.2, [_p(sub, 20, c["light"], italic=True, font="serif", line_spacing=1.1)]))
    if c["footer"]:
        sc.add(Text(M, 7.08, 6.0, 0.22, [_p(c["footer"], 7, c["white"])], anchor="middle"))


def _title_split(sc: Scene, plan: DeckPlan, spec: Slide, c: dict):
    """Cover split: title on a Bright Blue panel, rings on white."""
    pw = 7.0
    sc.add(Rect(0, 0, pw, H, c["bright"]))
    sc.add(Oval(pw + 3.2, 3.75, 2.05, fill=c["deep"]),
           Oval(pw + 3.2, 3.75, 2.6, line=c["light"], line_w=7),
           Arc(pw + 3.2, 3.75, 1.05, 90, 180, c["tint"], 7),
           Oval(pw + 5.5, 1.25, 0.45, fill=c["tint"]),
           Oval(pw + 1.0, 6.4, 0.24, fill=c["orange"]))
    logo, _ = _logo(c["logo_dark"], M, 0.5, 0.34, 2.6)
    if logo:
        sc.add(logo)
    title = spec.title or plan.deck_title
    size = _fit(title, [(28, 50), (55, 42), (90, 34), (999, 28)])
    sc.add(Text(M, 1.3, pw - 2 * M, 3.8, [_p(title, size, c["white"], bold=True, line_spacing=0.9)],
                anchor="bottom"))
    sub = spec.subtitle or plan.subtitle
    if sub:
        sc.add(Text(M, 5.35, pw - 2 * M, 1.2, [_p(sub, 18, c["white"], line_spacing=1.1)]))
    if c["footer"]:
        sc.add(Text(M, 7.08, 6.0, 0.22, [_p(c["footer"], 7, c["white"])], anchor="middle"))


def _section_number(sc: Scene, spec: Slide, c: dict, number: int, page: str):
    """Section divider with a big chapter number on the left."""
    ground = "blue" if number % 2 == 1 else "dark"
    sc.background = c["bright"] if ground == "blue" else c["deep"]
    num_col = c["tint"] if ground == "blue" else c["bright"]
    sc.add(Oval(1.6, H + 0.2, 2.1, fill=c["deep"] if ground == "blue" else c["bright"]),
           Oval(1.6, H + 0.2, 2.6, line=c["light"] if ground == "dark" else c["tint"], line_w=5))
    sc.add(Text(M, 0.9, 4.6, 3.2, [_p(f"{number:02d}", 150, num_col, bold=True, line_spacing=0.85)],
                anchor="bottom"))
    size = _fit(spec.title, [(30, 42), (60, 34), (999, 28)])
    sc.add(Rect(5.2, 2.2, 0.08, 2.6, c["light"] if ground == "dark" else c["white"]))
    sc.add(Text(5.6, 1.6, W - 5.6 - M, 2.6, [_p(spec.title, size, c["white"], bold=True, line_spacing=0.92)],
                anchor="bottom"))
    if spec.subtitle:
        sc.add(Text(5.6, 4.35, W - 5.6 - M, 1.2, [_p(spec.subtitle, 18, c["light"] if ground == "dark"
                                                    else c["white"], italic=True, font="serif")]))
    _footer(sc, c, ground, page)


def _content_numbered(sc: Scene, spec: Slide, c: dict, page: str):
    """Numbered points in one or two columns under a full-width title."""
    _header(sc, spec, c)
    items = spec.bullets[:6]
    if not items:
        _footer(sc, c, "light", page)
        return
    cols = 2 if len(items) > 3 else 1
    rows = (len(items) + cols - 1) // cols
    top, area_h = 2.5, 4.2
    col_w = (W - 2 * M - 0.5 * (cols - 1)) / cols
    row_h = min(1.4, area_h / rows)
    for i, text in enumerate(items):
        col, row = divmod(i, rows)
        x, y = M + col * (col_w + 0.5), top + row * row_h
        sc.add(Oval(x + 0.34, y + 0.34, 0.34, line=c["bright"], line_w=2.25))
        sc.add(Text(x, y, 0.68, 0.68, [_p(f"{i + 1:02d}", 14, c["bright"], bold=True, align="center")],
                    anchor="middle"))
        sc.add(Text(x + 0.95, y, col_w - 0.95, row_h - 0.1, [_p(text, 17, c["ink"], line_spacing=1.1)],
                    anchor="top"))
    _footer(sc, c, "light", page)


def _content_panels(sc: Scene, spec: Slide, c: dict, page: str):
    """Each point in its own soft panel with a bright edge."""
    _header(sc, spec, c)
    items = spec.bullets[:6]
    if not items:
        _footer(sc, c, "light", page)
        return
    cols = 2 if len(items) > 4 else 1
    rows = (len(items) + cols - 1) // cols
    top, area_h, gap = 2.45, 4.25, 0.18
    col_w = (W - 2 * M - 0.3 * (cols - 1)) / cols
    row_h = min(1.0, (area_h - gap * (rows - 1)) / rows)
    for i, text in enumerate(items):
        col, row = divmod(i, rows)
        x, y = M + col * (col_w + 0.3), top + row * (row_h + gap)
        sc.add(Rect(x, y, col_w, row_h, c["panel"], radius=0.14))
        sc.add(Rect(x, y + 0.18, 0.07, row_h - 0.36, c["bright"]))
        sc.add(Text(x + 0.35, y, col_w - 0.6, row_h, [_p(text, 17, c["deep"], line_spacing=1.1)],
                    anchor="middle"))
    _footer(sc, c, "light", page)


def _card_icon(sc: Scene, card, c: dict, cx: float, cy: float, d: float, number: int, fill: str, fg: str):
    """Brand symbol in a white circle when the card has one, else its number."""
    icon = (c["symbols_dir"] / f"{card.icon}.png"
            if c["symbols_dir"] and getattr(card, "icon", None) else None)
    if icon and icon.exists():
        sc.add(Oval(cx, cy, d / 2, fill=c["white"]))
        s = d * 0.8
        sc.add(Image(icon, cx - s / 2, cy - s / 2, s, s))
    else:
        sc.add(Oval(cx, cy, d / 2, fill=fill))
        sc.add(Text(cx - d / 2, cy - d / 2, d, d, [_p(f"{number:02d}", d * 22, fg, bold=True, align="center")],
                    anchor="middle"))


def _cards_grid(sc: Scene, spec: Slide, c: dict, page: str):
    """Cards as soft panels in a grid, symbol or number top-left."""
    _header(sc, spec, c)
    cards = spec.cards[:4]
    n = len(cards)
    cols = 2 if n == 4 else n
    rows = (n + cols - 1) // cols
    top, area_h, gap = 2.45, 4.25, 0.25
    cw = (W - 2 * M - gap * (cols - 1)) / cols
    ch = (area_h - gap * (rows - 1)) / rows
    for i, card in enumerate(cards):
        x, y = M + (i % cols) * (cw + gap), top + (i // cols) * (ch + gap)
        sc.add(Rect(x, y, cw, ch, c["panel"], radius=0.24))
        _card_icon(sc, card, c, x + 0.6, y + 0.6, 0.62, i + 1, c["bright"], c["white"])
        if rows == 1:
            sc.add(Text(x + 0.3, y + 1.15, cw - 0.6, 0.8, [_p(card.title, 18, c["deep"], bold=True, line_spacing=0.95)]))
            sc.add(Text(x + 0.3, y + 1.95, cw - 0.6, ch - 2.1, [_p(card.desc, 15, c["ink"], line_spacing=1.15)]))
        else:
            sc.add(Text(x + 1.15, y + 0.25, cw - 1.45, 0.7, [_p(card.title, 17, c["deep"], bold=True, line_spacing=0.95)],
                        anchor="middle"))
            sc.add(Text(x + 1.15, y + 1.0, cw - 1.45, ch - 1.1, [_p(card.desc, 14, c["ink"], line_spacing=1.15)]))
    _footer(sc, c, "light", page)


def _cards_columns(sc: Scene, spec: Slide, c: dict, page: str):
    """Open columns split by hairlines, a dark numbered circle on top of each."""
    _header(sc, spec, c)
    cards = spec.cards[:4]
    n = len(cards)
    top = 2.55
    cw = (W - 2 * M) / n
    for i, card in enumerate(cards):
        x = M + i * cw
        if i:
            sc.add(Line(x, top, x, 6.6, c["hairline"], 0.75))
        pad = 0.3 if i else 0
        _card_icon(sc, card, c, x + pad + 0.4, top + 0.4, 0.8, i + 1, c["deep"], c["white"])
        sc.add(Text(x + pad, top + 1.05, cw - pad - 0.3, 0.9, [_p(card.title, 18, c["bright"], bold=True, line_spacing=0.95)]))
        sc.add(Text(x + pad, top + 1.95, cw - pad - 0.3, 2.3, [_p(card.desc, 15, c["ink"], line_spacing=1.15)]))
    _footer(sc, c, "light", page)


def _stats_row(sc: Scene, spec: Slide, c: dict, page: str):
    """All figures side by side, each over a short bright bar."""
    _header(sc, spec, c)
    stats = spec.stats[:5]
    n = len(stats)
    top = 2.9
    cw = (W - 2 * M) / n
    for i, st in enumerate(stats):
        x = M + i * cw
        sc.add(Rect(x, top, 0.9, 0.08, c["orange"] if i == 0 else c["bright"]))
        sc.add(Text(x, top + 0.3, cw - 0.3, 1.5,
                    [_p(st.value, _fit(st.value, [(5, 54), (8, 44), (12, 34), (999, 28)]), c["deep"], bold=True)],
                    anchor="bottom"))
        sc.add(Text(x, top + 1.95, cw - 0.3, 1.7, [_p(st.label, 16, c["ink"], line_spacing=1.15)]))
    _footer(sc, c, "light", page)


def _stats_dark(sc: Scene, spec: Slide, c: dict, page: str):
    """Dark ground: the lead figure huge on the left, the rest listed on the right."""
    sc.background = c["deep"]
    size = _fit(spec.title, [(60, 26), (95, 22), (999, 19)])
    sc.add(Text(M, 0.45, W - 2 * M, 1.0, [_p(spec.title, size, c["white"], bold=True, line_spacing=0.9)],
                anchor="bottom"))
    if spec.subtitle:
        sc.add(Text(M, 1.55, W - 2 * M, 0.5, [_p(spec.subtitle, 16, c["light"], italic=True, font="serif")]))
    stats = spec.stats[:5]
    lead, others = stats[0], stats[1:]
    sc.add(Oval(3.0, 4.45, 2.05, line=c["bright"], line_w=6))
    sc.add(Text(M, 3.1, 5.2, 1.6, [_p(lead.value, _fit(lead.value, [(5, 66), (8, 50), (999, 38)]), c["light"],
                                      bold=True, align="center")], anchor="bottom"))
    sc.add(Text(1.2, 4.75, 3.6, 1.1, [_p(lead.label, 16, c["white"], align="center", line_spacing=1.1)]))
    if others:
        x0, top = 6.3, 2.5
        row_h = 4.2 / len(others)
        for i, st in enumerate(others):
            y = top + i * row_h
            if i:
                sc.add(Line(x0, y, W - M, y, c["bright"], 0.75))
            sc.add(Text(x0, y, 2.6, row_h, [_p(st.value, _fit(st.value, [(7, 32), (12, 26), (999, 22)]),
                                              c["white"], bold=True)], anchor="middle"))
            sc.add(Text(x0 + 2.8, y, W - M - x0 - 2.8, row_h, [_p(st.label, 16, c["tint"])], anchor="middle"))
    _footer(sc, c, "dark", page)


def _compare_split(sc: Scene, spec: Slide, c: dict, page: str):
    """Two open columns split by a rule, marked hollow (left) and solid (right)."""
    _header(sc, spec, c)
    top = 2.55
    half = (W - 2 * M) / 2
    sc.add(Line(M + half, top, M + half, 6.6, c["hairline"], 1))
    sides = [(spec.compare_left, c["deep"], None), (spec.compare_right, c["bright"], c["bright"])]
    for i, (side, head_col, dot_fill) in enumerate(sides):
        if side is None:
            continue
        x = M + i * half + (0.45 if i else 0)
        w = half - 0.45
        sc.add(Oval(x + 0.22, top + 0.3, 0.2, fill=dot_fill, line=None if dot_fill else c["deep"],
                    line_w=0 if dot_fill else 2))
        sc.add(Text(x + 0.6, top, w - 0.6, 0.6, [_p(side.heading, 22, head_col, bold=True)], anchor="middle"))
        sc.add(Text(x, top + 0.9, w - 0.3, 3.2, _bullet_paras(side.items[:6], 16, c["ink"], head_col)))
    _footer(sc, c, "light", page)


def _closing_blue(sc: Scene, spec: Slide, c: dict):
    """Blue ground, rings bottom-left, sign-off on the right half."""
    sc.background = c["bright"]
    sc.add(Oval(1.2, H - 0.4, 2.6, fill=c["deep"]),
           Oval(1.2, H - 0.4, 3.15, line=c["tint"], line_w=6),
           Oval(4.9, 2.1, 0.24, fill=c["orange"]))
    x = 5.4
    size = _fit(spec.title, [(30, 42), (60, 34), (999, 28)])
    sc.add(Text(x, 0.9, W - x - M, 2.6, [_p(spec.title, size, c["white"], bold=True, line_spacing=0.9)],
                anchor="bottom"))
    y = 3.75
    if spec.subtitle:
        sc.add(Text(x, y, W - x - M, 0.9, [_p(spec.subtitle, 20, c["white"], italic=True, font="serif")]))
        y += 1.0
    if spec.bullets:
        sc.add(Text(x, y, W - x - M, 6.7 - y, _bullet_paras(spec.bullets[:4], 17, c["white"], c["deep"])))
    _footer(sc, c, "blue", None)


# ------------------------------------------------------------------ new layouts

def _quote_layout(sc: Scene, spec: Slide, c: dict, page: str, dark: bool):
    """A short quotation (title) with its attribution (subtitle)."""
    ground = "dark" if dark else "light"
    if dark:
        sc.background = c["deep"]
    text_col = c["white"] if dark else c["deep"]
    sc.add(Oval(W - 1.2, 1.1, 1.6, fill=c["bright"] if dark else c["tint"]),
           Oval(W - 1.2, 1.1, 2.05, line=c["light"], line_w=4))
    sc.add(Text(M + 0.2, 0.55, 2.0, 2.0, [_p("“", 150, c["light"] if dark else c["bright"], font="serif")]))
    size = _fit(spec.title, [(80, 34), (150, 28), (240, 23), (999, 20)])
    sc.add(Text(1.5, 1.9, 9.4, 3.4, [_p(spec.title, size, text_col, italic=True, font="serif", line_spacing=1.12)],
                anchor="middle"))
    if spec.subtitle:
        sc.add(Rect(1.5, 5.6, 0.6, 0.06, c["orange"]))
        sc.add(Text(1.5, 5.75, 9.0, 0.6, [_p(spec.subtitle, 16, c["light"] if dark else c["ink"], bold=True)]))
    _footer(sc, c, ground, page)


def _statement_layout(sc: Scene, spec: Slide, c: dict, page: str, light: bool):
    """One big message on its own slide."""
    ground = "light" if light else "blue"
    if not light:
        sc.background = c["bright"]
    sc.add(Oval(W + 0.3, 3.35, 2.8, fill=c["tint"] if light else c["deep"]),
           Oval(W + 0.3, 3.35, 3.3, line=c["bright"] if light else c["light"], line_w=6))
    size = _fit(spec.title, [(40, 50), (80, 42), (130, 34), (999, 28)])
    sc.add(Text(M, 1.1, 9.0, 4.0, [_p(spec.title, size, c["deep"] if light else c["white"], bold=True,
                                     line_spacing=0.95)], anchor="bottom"))
    if light:
        sc.add(Rect(M, 5.3, 1.1, 0.08, c["bright"]), Oval(M + 1.4, 5.34, 0.09, fill=c["orange"]))
    if spec.subtitle:
        sc.add(Text(M, 5.6, 9.0, 1.1, [_p(spec.subtitle, 20, c["bright"] if light else c["white"],
                                         italic=True, font="serif")]))
    _footer(sc, c, ground, page)


def _agenda_layout(sc: Scene, spec: Slide, c: dict, page: str):
    """Dark panel with the title on the left, numbered agenda items on the right."""
    pw = 4.6
    sc.add(Rect(0, 0, pw, 6.85, c["deep"]))  # stops above the footer so it stays readable
    sc.add(Oval(pw - 1.1, H - 1.55, 0.7, line=c["bright"], line_w=5), Oval(1.0, 1.0, 0.16, fill=c["orange"]))
    size = _fit(spec.title, [(20, 40), (45, 32), (999, 26)])
    sc.add(Text(M, 1.4, pw - 2 * M, 3.0, [_p(spec.title, size, c["white"], bold=True, line_spacing=0.92)],
                anchor="bottom"))
    if spec.subtitle:
        sc.add(Text(M, 4.6, pw - 2 * M, 1.5, [_p(spec.subtitle, 16, c["light"], italic=True, font="serif")]))
    items = spec.bullets[:7]
    x0, x1 = pw + 0.7, W - M
    if items:
        row_h = min(0.95, 5.7 / len(items))
        y = 0.7 + (5.8 - row_h * len(items)) / 2
        for i, it in enumerate(items):
            sc.add(Text(x0, y, 0.9, row_h, [_p(f"{i + 1:02d}", 20, c["bright"], bold=True)], anchor="middle"))
            sc.add(Text(x0 + 1.0, y, x1 - x0 - 1.0, row_h, [_p(it, 19, c["deep"])], anchor="middle"))
            if i < len(items) - 1:
                sc.add(Line(x0, y + row_h, x1, y + row_h, c["hairline"], 0.5))
            y += row_h
    _footer(sc, c, "light", page)


def _timeline_layout(sc: Scene, spec: Slide, c: dict, page: str, zigzag: bool):
    """Steps in order along a line, numbered circles on the line."""
    _header(sc, spec, c)
    steps = (spec.steps or [])[:6]
    n = len(steps)
    if not n:
        _footer(sc, c, "light", page)
        return
    cw = (W - 2 * M) / n
    line_y = 4.25 if zigzag else 3.05
    sc.add(Line(M + cw / 2, line_y, W - M - cw / 2, line_y, c["tint"], 3))
    small = n > 4
    for i, st in enumerate(steps):
        cx = M + cw * i + cw / 2
        last = i == n - 1
        sc.add(Oval(cx, line_y, 0.34, fill=c["orange"] if last else c["bright"]))
        sc.add(Text(cx - 0.34, line_y - 0.34, 0.68, 0.68, [_p(f"{i + 1:02d}", 13, c["white"], bold=True, align="center")],
                    anchor="middle"))
        up = zigzag and i % 2 == 0
        ty = line_y - 0.55 - 1.6 if up else line_y + 0.55
        paras = [_p(st.title, 15 if small else 17, c["deep"], bold=True, align="center", space_after=4, line_spacing=0.95),
                 _p(st.desc, 12 if small else 14, c["ink"], align="center", line_spacing=1.1)]
        sc.add(Text(cx - cw / 2 + 0.1, ty, cw - 0.2, 1.6 if up else 2.4, paras, anchor="bottom" if up else "top"))
    _footer(sc, c, "light", page)


def _fmt_value(v: float, unit: Optional[str]) -> str:
    num = f"{v:,.0f}" if float(v).is_integer() else f"{v:,.1f}"
    unit = (unit or "").strip()
    if unit[:1] in ("$", "€", "£"):
        return unit[:1] + num + unit[1:]
    return num + ("" if not unit else unit if unit == "%" else " " + unit)


def _chart_layout(sc: Scene, spec: Slide, c: dict, page: str, columns: bool):
    """Bar (or column) chart drawn from the figures; the biggest value stands out.
    Bullets, if any, sit beside the chart as the takeaway."""
    _header(sc, spec, c)
    pts = (spec.chart or [])[:8]
    if not pts:
        _footer(sc, c, "light", page)
        return
    notes = spec.bullets[:3]
    right = W - M - (3.9 if notes else 0)
    top, bottom = 2.5, 6.55
    peak = max(p.value for p in pts) or 1
    lead = max(range(len(pts)), key=lambda i: pts[i].value)
    if columns:
        n = len(pts)
        slot = (right - M) / n
        bw = min(1.1, slot * 0.58)
        base = bottom - 0.45
        sc.add(Line(M, base, right, base, c["hairline"], 0.75))
        for i, pt in enumerate(pts):
            cx = M + slot * i + slot / 2
            h = max(0.04, (base - top - 0.5) * max(pt.value, 0) / peak)
            sc.add(Rect(cx - bw / 2, base - h, bw, h, c["bright"] if i == lead else c["tint"], radius=0.06))
            sc.add(Text(cx - slot / 2, base - h - 0.45, slot, 0.4, [_p(_fmt_value(pt.value, spec.chart_unit), 14,
                                                                        c["deep"], bold=True, align="center")],
                        anchor="bottom"))
            sc.add(Text(cx - slot / 2, base + 0.08, slot, 0.4, [_p(pt.label, 12, c["ink"], align="center")]))
    else:
        n = len(pts)
        row = (bottom - top) / n
        bh = min(0.42, row * 0.62)
        label_w = 2.2
        x0 = M + label_w
        span = right - x0 - 1.4
        for i, pt in enumerate(pts):
            cy = top + row * i + row / 2
            w = max(0.04, span * max(pt.value, 0) / peak)
            sc.add(Text(M, cy - row / 2, label_w - 0.2, row, [_p(pt.label, 14, c["ink"], align="right")], anchor="middle"))
            sc.add(Rect(x0, cy - bh / 2, w, bh, c["bright"] if i == lead else c["tint"], radius=0.06))
            sc.add(Text(x0 + w + 0.15, cy - row / 2, 1.3, row, [_p(_fmt_value(pt.value, spec.chart_unit), 14,
                                                                   c["deep"], bold=True)], anchor="middle"))
    if notes:
        x = W - M - 3.6
        sc.add(Rect(x, top, 3.6, bottom - top, c["panel"], radius=0.24))
        sc.add(Text(x + 0.3, top + 0.3, 3.0, bottom - top - 0.6, _bullet_paras(notes, 15, c["deep"], c["bright"])))
    _footer(sc, c, "light", page)


# ------------------------------------------------------------------ entry point

def layout_slide(plan: DeckPlan, index: int, template="niq") -> Scene:
    """template: a template id, or a theme dict already resolved by themes."""
    c = themes.load(template)
    spec = plan.slides[index]
    total = len(plan.slides)
    page = f"{index + 1} / {total}"
    sc = Scene(background=c["white"], fonts=dict(c["fonts"]), font_files=dict(c["font_files"]))
    v = pick_variant(plan, index)
    lay = spec.layout
    if lay == "title":
        [_title_layout, _title_dark, _title_split][v](sc, plan, spec, c)
    elif lay == "section":
        number = sum(1 for s in plan.slides[: index + 1] if s.layout == "section")
        [_section_layout, _section_number][v](sc, spec, c, number, page)
    elif lay == "closing":
        [_closing_layout, _closing_blue][v](sc, spec, c)
    elif lay == "cards" and spec.cards:
        [_cards_layout, _cards_grid, _cards_columns][v](sc, spec, c, page)
    elif lay == "stats" and spec.stats:
        [_stats_layout, _stats_row, _stats_dark][v](sc, spec, c, page)
    elif lay == "compare" and (spec.compare_left or spec.compare_right):
        [_compare_layout, _compare_split][v](sc, spec, c, page)
    elif lay == "quote":
        _quote_layout(sc, spec, c, page, dark=v == 1)
    elif lay == "statement":
        _statement_layout(sc, spec, c, page, light=v == 1)
    elif lay == "agenda" and spec.bullets:
        _agenda_layout(sc, spec, c, page)
    elif lay == "timeline" and spec.steps:
        _timeline_layout(sc, spec, c, page, zigzag=v == 1)
    elif lay == "chart" and spec.chart:
        _chart_layout(sc, spec, c, page, columns=v == 1)
    else:
        [_content_layout, _content_numbered, _content_panels][v if lay == "content" else 0](sc, spec, c, page)
    return sc
