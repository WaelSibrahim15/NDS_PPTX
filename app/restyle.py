"""NDS — match an existing .pptx to a template without touching its structure.

Only the look changes: theme colours and fonts, colours written directly on
shapes and text, slide backgrounds, footer text, and a logo. Shapes keep their
ids and positions, and the timing (animations) and transition XML is never
edited, so every animation and effect still plays as before.

Charts, SmartArt and pictures keep their own colours: they live in separate
parts (or inside the picture) that this module does not rewrite.
"""
import colorsys
import io
import math
import re
from collections import Counter
from pathlib import Path
from typing import Optional

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from . import themes

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

LOGO_NAME = "NDS Template Logo"
FOOTER_NAME = "NDS Template Footer"

# Fonts that carry icons or symbols: swapping them turns glyphs into letters.
SYMBOL_FONTS = {"wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
                "marlett", "segoe ui symbol", "segoe mdl2 assets", "font awesome"}
SERIF_HINTS = ("georgia", "times", "cambria", "garamond", "serif", "palatino", "baskerville",
               "book antiqua", "didot", "bodoni", "caslon", "merriweather", "playfair")
# Colour roles a shape or text colour may be mapped to.
PALETTE_ROLES = ["white", "deep", "bright", "light", "orange", "tint", "panel", "ink", "hairline"]


# ------------------------------------------------------------------ colour maths

def _rgb(hex_color: str) -> tuple:
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _lab(rgb: tuple) -> tuple:
    def lin(c):
        c /= 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (lin(c) for c in rgb)
    x = (0.4124 * r + 0.3576 * g + 0.1805 * b) / 0.95047
    y = 0.2126 * r + 0.7152 * g + 0.0722 * b
    z = (0.0193 * r + 0.1192 * g + 0.9505 * b) / 1.08883

    def f(t):
        return t ** (1 / 3) if t > 0.008856 else 7.787 * t + 16 / 116
    fx, fy, fz = f(x), f(y), f(z)
    return 116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz)


def _luminance(rgb: tuple) -> float:
    return _lab(rgb)[0] / 100


def _grey(rgb: tuple) -> bool:
    _, sat, val = colorsys.rgb_to_hsv(*(c / 255 for c in rgb))
    return sat < 0.12 or val < 0.12


def _hue(rgb: tuple) -> float:
    return colorsys.rgb_to_hsv(*(c / 255 for c in rgb))[0] * 360


class ColorMap:
    """Maps the deck's colours onto template roles the way a designer would:
    - the deck's most-used brand colour becomes the template's bright primary;
    - other colours of that family map by lightness (dark → deep, pale → tint);
    - colours far from that hue (accents) become the template accent;
    - greys map by lightness to the template's greys, so neutrals stay neutral."""

    CHROMA_ROLES = ["deep", "bright", "light", "tint"]
    GREY_ROLES = ["white", "panel", "hairline", "ink", "deep"]
    ACCENT_HUE_GAP = 75  # degrees from the primary hue that count as an accent

    def __init__(self, theme: dict, usage: Counter):
        self.theme = theme
        self.hex = {r: theme[r].lstrip("#").upper() for r in PALETTE_ROLES if theme.get(r)}
        self.light = {r: _lab(_rgb(h))[0] for r, h in self.hex.items()}
        chroma = [(n, c) for c, n in usage.items() if not _grey(_rgb(c))]
        self.primary = max(chroma)[1] if chroma else None
        self.primary_hue = _hue(_rgb(self.primary)) if self.primary else None
        self._cache = {}

    def _by_lightness(self, rgb: tuple, roles: list) -> str:
        light = _lab(rgb)[0]
        roles = [r for r in roles if r in self.hex]
        return min(roles, key=lambda r: abs(self.light[r] - light))

    def role(self, hex_color: str) -> str:
        key = hex_color.upper()
        if key not in self._cache:
            rgb = _rgb(key)
            if _grey(rgb):
                role = self._by_lightness(rgb, self.GREY_ROLES)
            elif key == self.primary:
                role = "bright"
            else:
                gap = abs(_hue(rgb) - self.primary_hue) if self.primary_hue is not None else 0
                gap = min(gap, 360 - gap)
                role = "orange" if gap > self.ACCENT_HUE_GAP and "orange" in self.hex \
                    else self._by_lightness(rgb, self.CHROMA_ROLES)
            self._cache[key] = role
        return self._cache[key]

    def nearest(self, hex_color: str) -> str:
        return self.hex[self.role(hex_color)]

    def background(self, hex_color: str) -> str:
        """A background keeps its lightness, so the text on it stays readable:
        greys become the template background or deep; colours the nearest in
        lightness of deep, bright, tint and the background (never the accent)."""
        rgb = _rgb(hex_color)
        if _grey(rgb):
            return self.hex["white" if _luminance(rgb) >= 0.5 else "deep"]
        return self.hex[self._by_lightness(rgb, ["deep", "bright", "tint", "white"])]

    def ground(self, dark: bool) -> str:
        """Background for a picture: the template's dark primary or its background."""
        return self.hex["deep" if dark else "white"]


# ------------------------------------------------------------------ theme part

def _restyle_theme_part(part, theme: dict):
    """Rewrite the colour scheme and font scheme of a theme part in place."""
    root = etree.fromstring(part.blob)
    c = {k: theme[k].lstrip("#").upper() for k in PALETTE_ROLES if theme.get(k)}
    scheme = {
        "dk1": c["deep"], "lt1": c["white"], "dk2": c["deep"], "lt2": c["panel"],
        "accent1": c["bright"], "accent2": c["light"], "accent3": c["orange"],
        "accent4": c["tint"], "accent5": c["deep"], "accent6": c["ink"],
        "hlink": c["bright"], "folHlink": c["deep"],
    }
    clr = root.find(".//a:clrScheme", NS)
    if clr is not None:
        clr.set("name", theme.get("name", "Template"))
        for slot, value in scheme.items():
            el = clr.find(f"a:{slot}", NS)
            if el is None:
                continue
            for child in list(el):
                el.remove(child)
            etree.SubElement(el, qn("a:srgbClr")).set("val", value)
    sans = theme["fonts"]["sans"]
    for tag in ("majorFont", "minorFont"):
        latin = root.find(f".//a:fontScheme/a:{tag}/a:latin", NS)
        if latin is not None:
            latin.set("typeface", sans)
            for attr in ("panose", "pitchFamily", "charset"):
                latin.attrib.pop(attr, None)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ------------------------------------------------------------------ shape colours and fonts

def _inside(el, *tags) -> bool:
    return any(a.tag in tags for a in el.iterancestors())


def _recolor(tree, cmap: ColorMap):
    """Map every explicit colour on shapes, lines and text to the template.
    Shadows and glows (effect lists), picture recolouring and the background
    (handled separately) are left alone."""
    skip = (qn("a:effectLst"), qn("a:blip"), qn("p:bg"), qn("a:effectDag"))
    for el in tree.iter(qn("a:srgbClr")):
        if _inside(el, *skip):
            continue
        val = el.get("val")
        if val and len(val) == 6:
            el.set("val", cmap.nearest(val))


def _swap_fonts(tree, theme: dict):
    sans = theme["fonts"]["sans"]
    serif = theme["fonts"].get("serif") or sans
    for el in tree.iter(qn("a:latin")):
        face = el.get("typeface", "")
        if not face or face.startswith("+") or face.lower() in SYMBOL_FONTS:
            continue
        el.set("typeface", serif if any(h in face.lower() for h in SERIF_HINTS) else sans)
        for attr in ("panose", "pitchFamily", "charset"):
            el.attrib.pop(attr, None)


# ------------------------------------------------------------------ backgrounds

def _image_is_dark(part, rid: str) -> Optional[bool]:
    try:
        blob = part.related_part(rid).blob
        with Image.open(io.BytesIO(blob)) as im:
            small = im.convert("RGB").resize((24, 24))
            px = list(small.getdata())
    except Exception:
        return None
    avg = tuple(sum(p[i] for p in px) / len(px) for i in range(3))
    return _luminance(avg) < 0.5


def _fill_is_dark(fill, part) -> Optional[bool]:
    """Whether a background fill reads as dark; None when it can't be told."""
    if fill is None:
        return None
    blip = fill.find(".//a:blip", NS)
    if blip is not None:
        return _image_is_dark(part, blip.get(qn("r:embed")))
    colors = [e.get("val") for e in fill.iter(qn("a:srgbClr")) if e.get("val")]
    if colors:
        lums = [_luminance(_rgb(v)) for v in colors]
        return sum(lums) / len(lums) < 0.5
    scheme = [e.get("val") for e in fill.iter(qn("a:schemeClr"))]
    if scheme:
        return scheme[0] in ("dk1", "dk2", "tx1", "tx2", "accent1", "accent5")
    return None


def _restyle_background(cSld, part, cmap: ColorMap) -> Optional[bool]:
    """Give a background its template colour. Solid and gradient colours map to
    their role; a picture becomes the template's dark or light ground, matching
    the picture's lightness so the text on it stays readable. Backgrounds built
    from theme colours already follow the new theme and are left as they are.
    Returns whether the background is dark (None when the slide has none)."""
    bg = cSld.find("p:bg", NS)
    if bg is None:
        return None
    fill = bg.find("p:bgPr", NS)
    if fill is None:
        fill = bg.find("p:bgRef", NS)
    if fill is None:
        return None
    blip = fill.find(".//a:blip", NS)
    colors = [e.get("val") for e in fill.iter(qn("a:srgbClr")) if e.get("val")]
    if blip is None and not colors:
        return _fill_is_dark(fill, part)
    if blip is not None:
        new = cmap.ground(bool(_image_is_dark(part, blip.get(qn("r:embed")))))
    else:
        avg = tuple(sum(_rgb(c)[i] for c in colors) // len(colors) for i in range(3))
        new = cmap.background("%02X%02X%02X" % avg)
    for child in list(bg):
        bg.remove(child)
    bg.append(etree.fromstring(
        f'<p:bgPr xmlns:p="{P}" xmlns:a="{A}"><a:solidFill><a:srgbClr val="{new}"/>'
        f'</a:solidFill><a:effectLst/></p:bgPr>'.encode()))
    return _luminance(_rgb(new)) < 0.5


def _usage(trees) -> Counter:
    """How often each explicit colour is used across the deck (shapes, text,
    backgrounds), so the dominant brand colour can be found."""
    skip = (qn("a:effectLst"), qn("a:blip"), qn("a:effectDag"))
    counts = Counter()
    for tree in trees:
        for el in tree.iter(qn("a:srgbClr")):
            val = (el.get("val") or "").upper()
            if len(val) == 6 and not _inside(el, *skip):
                counts[val] += 1
    return counts


# ------------------------------------------------------------------ footer and logo

PAGE_NUMBER = re.compile(r"^\s*(page\s*)?\d+(\s*(/|of)\s*\d+)?\s*$", re.I)


def _replace_text(body, text: str):
    """Set a text body to one line of text, keeping the first run's formatting."""
    paras = body.findall("a:p", NS)
    keep = paras[0] if paras else etree.SubElement(body, qn("a:p"))
    for extra in paras[1:]:
        body.remove(extra)
    runs = keep.findall("a:r", NS)
    rpr = runs[0].find("a:rPr", NS) if runs else None
    for child in list(keep):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            keep.remove(child)
    r = etree.Element(qn("a:r"))
    if rpr is not None:
        r.append(rpr)
    etree.SubElement(r, qn("a:t")).text = text
    end = keep.find("a:endParaRPr", NS)
    if end is not None:
        end.addprevious(r)
    else:
        keep.append(r)


def _set_footer_placeholders(tree, text: str) -> bool:
    """Put the template footer into any footer placeholder. True if one exists."""
    found = False
    for sp in tree.iter(qn("p:sp")):
        ph = sp.find(".//p:nvPr/p:ph", NS)
        if ph is None or ph.get("type") != "ftr":
            continue
        found = True
        body = sp.find("p:txBody", NS)
        if body is not None:
            _replace_text(body, text)
    return found


def _footer_boxes(slide, slide_h: int) -> list:
    """Plain text boxes that act as a footer: short text in the bottom band of
    the slide that isn't a page number (and isn't a placeholder)."""
    found = []
    for shp in slide.shapes:
        if shp.name in (LOGO_NAME, FOOTER_NAME) or shp.is_placeholder or not shp.has_text_frame:
            continue
        text = shp.text_frame.text.strip()
        if not text or len(text) > 120 or PAGE_NUMBER.match(text):
            continue
        if shp.top is not None and shp.top >= slide_h * 0.9:
            found.append(shp)
    return found


def _animated_ids(slide) -> set:
    timing = slide._element.find(qn("p:timing"))
    if timing is None:
        return set()
    return {int(t.get("spid")) for t in timing.iter(qn("p:spTgt")) if (t.get("spid") or "").isdigit()}


def _slide_is_dark(slide, bg_dark: bool, slide_w: int, slide_h: int) -> bool:
    """A shape filling the slide counts as its background."""
    for shp in slide.shapes:
        if shp.left is None or shp.width is None:
            continue
        if shp.width >= slide_w * 0.9 and shp.height >= slide_h * 0.9:
            fill = shp._element.find("p:spPr/a:solidFill/a:srgbClr", NS)
            if fill is not None:
                bg_dark = _luminance(_rgb(fill.get("val"))) < 0.5
    return bg_dark


def _remove_named(slide, name: str):
    for shp in list(slide.shapes):
        if shp.name == name:
            shp._element.getparent().remove(shp._element)


def _add_logo_and_footer(slide, theme: dict, dark: bool, slide_w: int, slide_h: int,
                         footer_in_placeholder: bool):
    """Small logo bottom-left and the footer line beside it, on every slide.
    A footer the deck already has (placeholder or text box) gets the template
    text instead of a second line. Named shapes, so running the restyle again
    replaces instead of stacking."""
    _remove_named(slide, LOGO_NAME)
    _remove_named(slide, FOOTER_NAME)
    scale = slide_h / Emu(6858000)  # sizes are tuned for a 7.5in-high slide
    margin = int(Emu(457200) * scale)
    height = int(Emu(155448) * scale)  # 0.17in
    gap = int(Emu(164592) * scale)
    y = slide_h - int(Emu(365760) * scale)
    x = margin
    footer = theme.get("footer")
    boxes = _footer_boxes(slide, slide_h) if footer else []
    logo = theme.get("logo_dark" if dark else "logo_light")
    if logo and Path(logo).exists():
        pic = slide.shapes.add_picture(str(logo), x, y, height=height)
        if pic.width > int(Emu(1463040) * scale):  # 1.6in cap for wide logos
            ratio = int(Emu(1463040) * scale) / pic.width
            pic.width, pic.height = int(pic.width * ratio), int(pic.height * ratio)
        pic.name = LOGO_NAME
        x += pic.width + gap
        # Make room: an old footer box starting under the logo moves right,
        # unless an animation targets it (then nothing is moved).
        animated = _animated_ids(slide)
        for box in boxes:
            if box.left < x and box.shape_id not in animated:
                shift = x - box.left
                box.left = x
                box.width = max(int(box.width - shift), int(Emu(914400) * scale))
    for box in boxes:
        _replace_text(box.text_frame._txBody, footer)
    if footer and not footer_in_placeholder and not boxes:
        box = slide.shapes.add_textbox(x, y - int(Emu(27432) * scale), int(Emu(5486400) * scale),
                                       int(Emu(201168) * scale))
        box.name = FOOTER_NAME
        tf = box.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        run = tf.paragraphs[0].add_run()
        run.text = footer
        run.font.size = Pt(7 * scale)
        run.font.name = theme["fonts"]["sans"]
        run.font.color.rgb = RGBColor.from_string(
            (theme["white"] if dark else theme["ink"]).lstrip("#"))


# ------------------------------------------------------------------ public

def restyle_pptx(src_path: Path, template, out_path: Path) -> Path:
    """Write a copy of src_path in the template's colours, fonts, backgrounds,
    footer and logo. Animations, transitions, shape ids and layout positions
    are left exactly as they were (an old footer box may shift right to make
    room for the logo)."""
    theme = themes.load(template)
    prs = Presentation(str(src_path))
    footer = theme.get("footer")
    parts = [m._element for m in prs.slide_masters] + \
        [l._element for m in prs.slide_masters for l in m.slide_layouts] + \
        [s._element for s in prs.slides]
    cmap = ColorMap(theme, _usage(parts))

    seen_themes = set()
    master_dark = {}
    for master in prs.slide_masters:
        for rel in master.part.rels.values():
            if rel.reltype == RT.THEME and id(rel.target_part) not in seen_themes:
                seen_themes.add(id(rel.target_part))
                _restyle_theme_part(rel.target_part, theme)
        mtree = master._element
        dark = _restyle_background(mtree.find("p:cSld", NS), master.part, cmap)
        master_dark[id(master)] = bool(dark)
        _recolor(mtree, cmap)
        _swap_fonts(mtree, theme)
        if footer:
            _set_footer_placeholders(mtree, footer)
        for layout in master.slide_layouts:
            ltree = layout._element
            _restyle_background(ltree.find("p:cSld", NS), layout.part, cmap)
            _recolor(ltree, cmap)
            _swap_fonts(ltree, theme)
            if footer:
                _set_footer_placeholders(ltree, footer)

    for slide in prs.slides:
        stree = slide._element
        layout = slide.slide_layout
        dark = _restyle_background(stree.find("p:cSld", NS), slide.part, cmap)
        if dark is None:  # inherits: layout background, else the master's
            lbg = layout._element.find("p:cSld/p:bg", NS)
            dark = _fill_is_dark(lbg[0], layout.part) if lbg is not None and len(lbg) else None
            if dark is None:
                dark = master_dark.get(id(layout.slide_master), False)
        _recolor(stree, cmap)
        _swap_fonts(stree, theme)
        dark = _slide_is_dark(slide, dark, prs.slide_width, prs.slide_height)
        in_ph = _set_footer_placeholders(stree, footer) if footer else False
        _add_logo_and_footer(slide, theme, dark, prs.slide_width, prs.slide_height, in_ph)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path
