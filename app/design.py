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


# ------------------------------------------------------------------ entry point

def layout_slide(plan: DeckPlan, index: int, template="niq") -> Scene:
    """template: a template id, or a theme dict already resolved by themes."""
    c = themes.load(template)
    spec = plan.slides[index]
    total = len(plan.slides)
    page = f"{index + 1} / {total}"
    sc = Scene(background=c["white"], fonts=dict(c["fonts"]), font_files=dict(c["font_files"]))
    if spec.layout == "title":
        _title_layout(sc, plan, spec, c)
    elif spec.layout == "section":
        number = sum(1 for s in plan.slides[: index + 1] if s.layout == "section")
        _section_layout(sc, spec, c, number, page)
    elif spec.layout == "closing":
        _closing_layout(sc, spec, c)
    elif spec.layout == "cards" and spec.cards:
        _cards_layout(sc, spec, c, page)
    elif spec.layout == "stats" and spec.stats:
        _stats_layout(sc, spec, c, page)
    elif spec.layout == "compare" and (spec.compare_left or spec.compare_right):
        _compare_layout(sc, spec, c, page)
    else:
        _content_layout(sc, spec, c, page)
    return sc
