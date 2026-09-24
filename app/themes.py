"""NDS — slide templates: the two built-ins plus ones imported from Claude Design.

A template is a theme for the layouts in design.py: nine colour roles, a sans
and a serif font (with font files for previews), light and dark logos, a
footer line, optional brand symbols for feature cards, and brand rules that
go to the drafter. Only the built-in "niq" template carries the NIQ
PowerPoint compliance rules.

Imported templates live in TEMPLATES_DIR/saved/<id>/ (theme.json plus
logos/, fonts/, symbols/). An import first lands in TEMPLATES_DIR/drafts/<id>/
so it can be reviewed before it is saved. Each job keeps a copy of its
template in jobs/<job>/template/, so editing or deleting a template never
changes decks already drafted.
"""
import datetime
import io
import json
import os
import re
import shutil
import uuid
import zipfile
from pathlib import Path
from typing import Optional

ROOT = Path(__file__).resolve().parent.parent
ASSETS = Path(__file__).resolve().parent / "assets"
SYMBOLS_DIR = ASSETS / "symbols"
LOGOS_DIR = ASSETS / "logos"


def _templates_dir() -> Path:
    if os.environ.get("TEMPLATES_DIR", "").strip():
        return Path(os.environ["TEMPLATES_DIR"]).expanduser()
    # Railway sets this when a volume is attached, so templates survive redeploys.
    if os.environ.get("RAILWAY_VOLUME_MOUNT_PATH", "").strip():
        return Path(os.environ["RAILWAY_VOLUME_MOUNT_PATH"]) / "templates"
    return ROOT / "templates"


TEMPLATES_DIR = _templates_dir()
SAVED = TEMPLATES_DIR / "saved"
DRAFTS = TEMPLATES_DIR / "drafts"

ROLES = ["white", "deep", "bright", "light", "orange", "tint", "panel", "ink", "hairline"]
ROLE_LABELS = {
    "white": "Background",
    "deep": "Dark primary (titles, dark slides)",
    "bright": "Bright primary (bars, callouts, blue slides)",
    "light": "Highlight (rings, text on dark)",
    "orange": "Accent dot",
    "tint": "Soft tint",
    "panel": "Panel (sidebars, grey boxes)",
    "ink": "Body text",
    "hairline": "Rules and dividers",
}

NIQ_FOOTER = "© {year} Nielsen Consumer LLC. All Rights Reserved."

BUILTIN = {
    "niq": {
        "name": "NIQ (Claude Design)",
        "colors": {"white": "#FFFFFF", "deep": "#060A45", "bright": "#2D6DF6", "light": "#31D1FF",
                   "orange": "#EF5F17", "tint": "#B4CBF9", "panel": "#F2F2F2", "ink": "#555555",
                   "hairline": "#B3B3B3"},
        "fonts": {"sans": "Arial", "serif": "Georgia"},
    },
    "neutral": {
        "name": "Neutral professional",
        "colors": {"white": "#FFFFFF", "deep": "#1F2A37", "bright": "#2F6FED", "light": "#38B2C4",
                   "orange": "#D97706", "tint": "#C9D6EA", "panel": "#F3F4F6", "ink": "#404040",
                   "hairline": "#C0C4CC"},
        "fonts": {"sans": "Arial", "serif": "Georgia"},
    },
}

# Fonts every PowerPoint install has; no files needed (previews use Liberation).
SAFE_FONTS = {"arial", "helvetica", "helvetica neue", "georgia", "times new roman", "times",
              "calibri", "cambria", "verdana", "tahoma", "trebuchet ms", "segoe ui", "garamond"}
FONT_FILE_KEYS = ["sans", "sans_bold", "serif", "serif_italic"]
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".svg"}
MAX_UPLOAD = 60 * 1024 * 1024
MAX_SYMBOLS = 60


# ------------------------------------------------------------------ runtime themes

def footer_text(footer: Optional[str]) -> Optional[str]:
    if not footer:
        return None
    return footer.replace("{year}", str(datetime.date.today().year))


def _builtin(ref: str) -> dict:
    b = BUILTIN[ref]
    niq = ref == "niq"
    return {
        **b["colors"],
        "id": ref, "name": b["name"], "niq": niq, "builtin": True,
        "fonts": dict(b["fonts"]), "font_files": {},
        "logo_light": LOGOS_DIR / "niq-mark-bright-blue.png" if niq else None,
        "logo_dark": LOGOS_DIR / "niq-mark-white.png" if niq else None,
        "footer": footer_text(NIQ_FOOTER) if niq else None,
        "symbols_dir": SYMBOLS_DIR if niq else None,
        "symbols": sorted(p.stem for p in SYMBOLS_DIR.glob("*.png")) if niq else [],
        "brand_rules": "",
    }


def _runtime(folder: Path, data: dict) -> dict:
    """theme.json (plus its folder) → the dict design.py draws with."""
    def f(rel):
        p = folder / rel if rel else None
        return p if p and p.exists() else None

    colors = {**BUILTIN["neutral"]["colors"], **(data.get("colors") or {})}
    sym_dir = folder / "symbols"
    symbols = sorted(p.stem for p in sym_dir.glob("*.png")) if sym_dir.exists() else []
    fonts = data.get("fonts") or {}
    return {
        **colors,
        "id": data.get("id", folder.name), "name": data.get("name", folder.name),
        "niq": False, "builtin": False,
        "fonts": {"sans": fonts.get("sans") or "Arial", "serif": fonts.get("serif") or "Georgia"},
        "font_files": {k: p for k in FONT_FILE_KEYS if (p := f((data.get("font_files") or {}).get(k)))},
        "logo_light": f(data.get("logo_light")),
        "logo_dark": f(data.get("logo_dark")),
        "footer": footer_text(data.get("footer")),
        "symbols_dir": sym_dir if symbols else None,
        "symbols": symbols,
        "brand_rules": data.get("brand_rules") or "",
    }


def _read(folder: Path) -> dict:
    return json.loads((folder / "theme.json").read_text(encoding="utf-8"))


def _write(folder: Path, data: dict):
    (folder / "theme.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def _safe_id(ref: str) -> str:
    if not re.fullmatch(r"[a-z0-9][a-z0-9-]{0,63}", ref or ""):
        raise KeyError(ref)
    return ref


def folder_for(ref: str) -> Path:
    """Saved template or import draft folder for an id (raises KeyError)."""
    ref = _safe_id(ref)
    for base in (SAVED, DRAFTS):
        if (base / ref / "theme.json").exists():
            return base / ref
    raise KeyError(ref)


def load(ref) -> dict:
    """Runtime theme for a template id; a dict passes through (already resolved)."""
    if isinstance(ref, dict):
        return ref
    ref = ref or "niq"
    if ref in BUILTIN:
        return _builtin(ref)
    try:
        folder = folder_for(ref)
    except KeyError:
        return _builtin("niq")
    return _runtime(folder, _read(folder))


def job_theme(ref: str, job_dir: Path) -> dict:
    """The template a job draws with: its own snapshot, taken the first time the
    job uses a custom template (and again if the job switches template)."""
    ref = ref or "niq"
    if ref in BUILTIN:
        return _builtin(ref)
    snap = job_dir / "template"
    if (snap / "theme.json").exists() and _read(snap).get("id") == ref:
        return _runtime(snap, _read(snap))
    try:
        src = SAVED / _safe_id(ref)
    except KeyError:
        return _builtin("niq")
    if (src / "theme.json").exists():
        if snap.exists():
            shutil.rmtree(snap)
        shutil.copytree(src, snap)
        return _runtime(snap, _read(snap))
    return _builtin("niq")


def list_templates() -> list:
    out = [{"id": k, "name": v["name"], "builtin": True, "colors": v["colors"]}
           for k, v in BUILTIN.items()]
    if SAVED.exists():
        rows = []
        for p in SAVED.glob("*/theme.json"):
            try:
                d = json.loads(p.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                continue
            rows.append({"id": d["id"], "name": d.get("name") or d["id"], "builtin": False,
                         "colors": d.get("colors") or {}, "updated": d.get("updated", "")})
        out += sorted(rows, key=lambda r: r["name"].lower())
    return out


def describe(ref: str) -> dict:
    """theme.json plus what the review screen needs (candidate logos, fonts found)."""
    folder = folder_for(ref)
    data = _read(folder)
    logos = sorted(p.relative_to(folder).as_posix() for p in (folder / "logos").glob("*.png")) \
        if (folder / "logos").exists() else []
    return {**data, "draft": folder.parent == DRAFTS, "logo_options": logos,
            "symbol_count": len(list((folder / "symbols").glob("*.png")))
            if (folder / "symbols").exists() else 0,
            "role_labels": ROLE_LABELS}


# ------------------------------------------------------------------ editing

HEX = re.compile(r"^#[0-9a-fA-F]{6}$")


def _clean_edit(data: dict, folder: Path) -> dict:
    """Validate the fields the review screen can change."""
    out = {}
    if "name" in data:
        name = str(data["name"] or "").strip()[:60]
        if not name:
            raise ValueError("Give the template a name.")
        out["name"] = name
    if "colors" in data:
        colors = {}
        for role in ROLES:
            v = str((data["colors"] or {}).get(role, "")).strip()
            if not HEX.match(v):
                raise ValueError(f"{ROLE_LABELS[role]}: use a colour like #2D6DF6.")
            colors[role] = v.upper()
        out["colors"] = colors
    if "fonts" in data:
        fonts = {}
        for k in ("sans", "serif"):
            v = str((data["fonts"] or {}).get(k, "")).strip().strip("'\"")[:60]
            if not v or re.search(r"[;{}<>\\]", v):
                raise ValueError("Font names can't be empty or contain ; { } < > \\.")
            fonts[k] = v
        out["fonts"] = fonts
    for k in ("logo_light", "logo_dark"):
        if k in data:
            v = (data[k] or "").strip()
            if v and not ((folder / v).resolve().is_relative_to(folder.resolve()) and (folder / v).exists()):
                raise ValueError("Pick a logo from the list or upload one.")
            out[k] = v
    if "footer" in data:
        out["footer"] = str(data["footer"] or "").strip()[:160]
    if "brand_rules" in data:
        out["brand_rules"] = str(data["brand_rules"] or "").strip()[:4000]
    return out


def update(ref: str, data: dict) -> dict:
    folder = folder_for(ref)
    theme = _read(folder)
    theme.update(_clean_edit(data, folder))
    theme["updated"] = _now()
    _write(folder, theme)
    return theme


def save_draft(ref: str, data: dict) -> dict:
    """Apply the review edits to an import draft and move it into saved templates."""
    folder = folder_for(ref)
    if folder.parent != DRAFTS:
        return update(ref, data)
    theme = _read(folder)
    theme.update(_clean_edit(data, folder))
    theme["updated"] = _now()
    _write(folder, theme)
    SAVED.mkdir(parents=True, exist_ok=True)
    shutil.move(str(folder), str(SAVED / ref))
    return theme


def delete(ref: str):
    if ref in BUILTIN:
        raise ValueError("Built-in templates can't be deleted.")
    shutil.rmtree(folder_for(ref))


def set_logo(ref: str, kind: str, filename: str, raw: bytes) -> str:
    """Store an uploaded logo (PNG, JPG, WebP or SVG) as logos/<kind>-upload.png."""
    if kind not in ("light", "dark"):
        raise ValueError("Logo kind must be light or dark.")
    folder = folder_for(ref)
    png = _to_png(raw, Path(filename).suffix.lower(), height=240)
    rel = f"logos/{kind}-upload-{uuid.uuid4().hex[:6]}.png"
    (folder / "logos").mkdir(exist_ok=True)
    (folder / rel).write_bytes(png)
    theme = _read(folder)
    theme[f"logo_{kind}"] = rel
    theme["updated"] = _now()
    _write(folder, theme)
    return rel


def preview_theme(ref: str, edits: dict) -> dict:
    """Runtime theme with unsaved review edits applied (for live previews)."""
    folder = folder_for(ref)
    data = _read(folder)
    try:
        data.update(_clean_edit(edits or {}, folder))
    except ValueError:
        pass  # keep showing the last valid state while someone types
    return _runtime(folder, data)


def _now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


# ------------------------------------------------------------------ import

def import_upload(filename: str, raw: bytes, api_key: str = "") -> str:
    """Read a Claude Design export (.zip, or tokens.json alone) into an import
    draft. Returns the draft id; the draft is saved only after review."""
    if len(raw) > MAX_UPLOAD:
        raise ValueError("That file is over 60 MB. Export the design system without photos or maps.")
    suffix = Path(filename or "").suffix.lower()
    draft_id = uuid.uuid4().hex[:10]
    folder = DRAFTS / draft_id
    src = folder / "src"
    src.mkdir(parents=True)
    try:
        if suffix == ".zip":
            _unzip(raw, src)
        elif suffix == ".json":
            (src / "tokens.json").write_bytes(raw)
        else:
            raise ValueError("Upload the design system export from Claude Design: a .zip, or its tokens.json.")
        tokens_path = _find_tokens(src)
        base = tokens_path.parent
        try:
            tokens = json.loads(tokens_path.read_text(encoding="utf-8"))
        except ValueError:
            raise ValueError("tokens.json in that file isn't valid JSON.")
        if not isinstance(tokens, dict) or "color" not in tokens:
            raise ValueError("That tokens.json has no colours. Is it a Claude Design design-system export?")
        readme = ""
        for name in ("README.md", "readme.md"):
            if (base / name).exists():
                readme = (base / name).read_text(encoding="utf-8", errors="ignore")
                break

        logos = _collect_logos(base, folder)
        symbols = _collect_symbols(base, folder)
        palette = heuristic_colors(tokens)
        fonts = heuristic_fonts(tokens)
        theme = {
            "id": draft_id,
            "name": str(tokens.get("name") or Path(filename).stem)[:60],
            "colors": palette,
            "fonts": fonts,
            "font_files": {},
            "logo_light": _pick_logo(logos, dark=False),
            "logo_dark": _pick_logo(logos, dark=True),
            "footer": heuristic_footer(readme),
            "brand_rules": readme.split("\n---\n")[0].strip()[:1500],
            "source": f"Claude Design export ({Path(filename).name})",
            "created": _now(), "updated": _now(),
        }
        if api_key:
            try:
                theme.update(_claude_mapping(api_key, tokens, readme, logos, theme))
            except Exception as exc:  # the heuristic proposal is still usable
                theme["import_note"] = f"Colour mapping by Claude failed ({exc}); check the roles."
        theme["font_files"] = _collect_fonts(base, folder, tokens, theme["fonts"])
        _write(folder, theme)
        shutil.rmtree(src, ignore_errors=True)
        return draft_id
    except Exception:
        shutil.rmtree(folder, ignore_errors=True)
        raise


def clear_stale_drafts(max_age_hours: int = 24):
    if not DRAFTS.exists():
        return
    cutoff = datetime.datetime.now().timestamp() - max_age_hours * 3600
    for p in DRAFTS.iterdir():
        if p.is_dir() and p.stat().st_mtime < cutoff:
            shutil.rmtree(p, ignore_errors=True)


def _unzip(raw: bytes, dest: Path):
    try:
        zf = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        raise ValueError("That .zip file can't be opened.")
    total = 0
    members = [m for m in zf.infolist() if not m.is_dir()]
    if len(members) > 3000:
        raise ValueError("That .zip has too many files.")
    for m in members:
        name = m.filename.replace("\\", "/")
        parts = [p for p in name.split("/") if p not in ("", ".")]
        if not parts or ".." in parts or parts[0].startswith("__MACOSX") or parts[-1].startswith("."):
            continue
        total += m.file_size
        if total > 300 * 1024 * 1024:
            raise ValueError("That .zip unpacks to more than 300 MB.")
        target = dest.joinpath(*parts)
        target.parent.mkdir(parents=True, exist_ok=True)
        with zf.open(m) as fin, target.open("wb") as fout:
            shutil.copyfileobj(fin, fout)


def _find_tokens(src: Path) -> Path:
    found = sorted(src.rglob("tokens.json"), key=lambda p: len(p.parts))
    if not found:
        raise ValueError("No tokens.json found. Export the design system from Claude Design and upload that .zip.")
    return found[0]


# ---- colours

def _hex(v) -> Optional[str]:
    if not isinstance(v, str):
        return None
    v = v.strip()
    m = re.fullmatch(r"#([0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})", v)
    if m:
        h = m.group(1)
        if len(h) in (3, 4):
            h = "".join(c * 2 for c in h[:3])
        return "#" + h[:6].upper()
    m = re.fullmatch(r"rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+).*\)", v)
    if m:
        return "#" + "".join(f"{min(255, int(x)):02X}" for x in m.groups())
    return None


def token_colors(tokens: dict) -> dict:
    """{name: {theme_id: '#RRGGBB'}} with aliases resolved."""
    color = tokens.get("color") or {}
    themes = [t.get("id") for t in color.get("themes") or [] if isinstance(t, dict)] or ["light"]
    raw = {t.get("name"): t.get("value") for t in color.get("tokens") or []
           if isinstance(t, dict) and t.get("name")}

    def resolve(v, theme, depth=0):
        if depth > 16:
            return None
        if isinstance(v, dict):
            v = v.get(theme, v.get(themes[0]))
        if isinstance(v, str) and v.startswith("{") and v.endswith("}"):
            return resolve(raw.get(v[1:-1]), theme, depth + 1)
        return _hex(v)

    out = {}
    for name, v in raw.items():
        per = {th: c for th in themes if (c := resolve(v, th))}
        if per:
            out[name.lower()] = per
    return out


def _lum(h: str) -> float:
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (1, 3, 5))
    lin = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in (r, g, b)]
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def _sat(h: str) -> float:
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (1, 3, 5))
    return max(r, g, b) - min(r, g, b)


def heuristic_colors(tokens: dict) -> dict:
    """Map design-system colour tokens onto NDS's nine roles by token name,
    falling back to lightness and saturation."""
    cols = token_colors(tokens)
    color = tokens.get("color") or {}
    themes = [t.get("id") for t in color.get("themes") or [] if isinstance(t, dict)] or ["light"]
    first = themes[0]
    dark_theme = next((t for t in themes if "dark" in (t or "")), None)

    def get(name, theme=None):
        per = cols.get(name)
        if not per:
            return None
        return per.get(theme or first) or per.get(first)

    def first_of(names, theme=None, test=None):
        for n in names:
            for key in cols:
                if key == n or (n.endswith("*") and key.startswith(n[:-1])):
                    v = get(key, theme)
                    if v and (test is None or test(v)):
                        return v
        return None

    darkish = lambda v: _lum(v) < 0.12
    mid = lambda v: 0.05 < _lum(v) < 0.6 and _sat(v) > 0.25
    lightish = lambda v: _lum(v) > 0.55

    base = dict(BUILTIN["neutral"]["colors"])
    all_vals = sorted({v for per in cols.values() for v in per.values()})
    pick = {
        "white": first_of(["surface", "background", "bg", "white", "surface-light"], test=lightish),
        "deep": first_of(["heading", "deep-blue", "primary-dark", "navy", "brand-dark", "deep*",
                          "primary", "brand"], test=darkish),
        "bright": first_of(["accent-fill", "accent", "action", "bright-blue", "bright*", "primary",
                            "brand", "link"], test=mid),
        "light": (first_of(["accent", "link"], theme=dark_theme, test=lambda v: _lum(v) > 0.3)
                  if dark_theme else None)
        or first_of(["light-blue", "cyan", "highlight", "secondary"], test=lambda v: _lum(v) > 0.3),
        "orange": first_of(["signal", "orange", "accent-2", "secondary", "warning", "vault-gold",
                            "gold", "yellow"]),
        "tint": first_of(["blue-tint", "tint", "brand-tint", "primary-tint", "surface-selected",
                          "light-gray"], test=lightish),
        "panel": first_of(["surface-alt", "surface-sunken", "grey-panel", "gray-panel", "panel",
                           "light-gray", "light-grey", "muted"], test=lightish),
        "ink": first_of(["ink", "text", "body", "text-grey", "text-gray", "charcoal", "foreground"],
                        test=lambda v: _lum(v) < 0.25),
        "hairline": first_of(["rule", "border", "hairline", "divider", "line"]),
    }
    if not pick["deep"]:
        dark = [v for v in all_vals if darkish(v)]
        pick["deep"] = min(dark, key=_lum) if dark else None
    if not pick["bright"]:
        bright = [v for v in all_vals if mid(v)]
        pick["bright"] = max(bright, key=_sat) if bright else None
    return {role: (pick[role] or base[role]).upper() for role in ROLES}


# ---- fonts

def _first_family(stack: str) -> Optional[str]:
    if not isinstance(stack, str) or not stack.strip():
        return None
    name = stack.split(",")[0].strip().strip("'\"")
    generic = {"sans-serif", "serif", "system-ui", "monospace", "-apple-system", "ui-sans-serif"}
    return None if name.lower() in generic else name


def heuristic_fonts(tokens: dict) -> dict:
    families = ((tokens.get("type") or {}).get("families") or {})
    sans = None
    for key in ("sans", "body", "base", "text", "heading", "display"):
        if (sans := _first_family(families.get(key))):
            break
    if not sans:
        sans = next((f for v in families.values() if (f := _first_family(v))), None)
    serif = None
    for key in ("serif", "editorial", "display", "heading"):
        if (serif := _first_family(families.get(key))) and serif != sans:
            break
        serif = None
    return {"sans": sans or "Arial", "serif": serif or sans or "Georgia"}


def _font_meta(path: Path):
    """(family, bold, italic) read from a font file, or None."""
    try:
        from fontTools.ttLib import TTFont
        f = TTFont(str(path), lazy=True, fontNumber=0)
        name = f["name"]
        fam = (name.getDebugName(16) or name.getDebugName(1) or "").strip()
        sub = (name.getDebugName(17) or name.getDebugName(2) or "").lower()
        weight = f["OS/2"].usWeightClass if "OS/2" in f else 400
        italic = "italic" in sub or "oblique" in sub
        variable = "fvar" in f
        return fam, (weight >= 600 or "bold" in sub), italic, variable
    except Exception:
        return None


def _to_ttf(path: Path, out_dir: Path) -> Optional[Path]:
    suf = path.suffix.lower()
    out_dir.mkdir(parents=True, exist_ok=True)
    if suf in (".ttf", ".otf"):
        dest = out_dir / path.name
        shutil.copyfile(path, dest)
        return dest
    if suf in (".woff", ".woff2"):
        try:
            from fontTools.ttLib import TTFont
            f = TTFont(str(path))
            f.flavor = None
            dest = out_dir / (path.stem + ".ttf")
            f.save(str(dest))
            return dest
        except Exception:
            return None
    return None


def _collect_fonts(base: Path, folder: Path, tokens: dict, fonts: dict) -> dict:
    """Font files for previews: from the export's fonts/ folder, else Google Fonts."""
    out_dir = folder / "fonts"
    files = [p for p in base.rglob("*") if p.suffix.lower() in (".ttf", ".otf", ".woff", ".woff2")]
    converted = [t for p in files[:40] if (t := _to_ttf(p, out_dir))]
    for fam in {fonts["sans"], fonts["serif"]}:
        if fam.lower() in SAFE_FONTS:
            continue
        have = any((m := _font_meta(p)) and m[0].lower() == fam.lower() for p in converted)
        if not have:
            converted += _google_font_files(fam, out_dir)
    chosen = {}
    metas = [(p, m) for p in converted if (m := _font_meta(p))]
    for key, fam, bold, italic in (("sans", fonts["sans"], False, False),
                                   ("sans_bold", fonts["sans"], True, False),
                                   ("serif", fonts["serif"], False, False),
                                   ("serif_italic", fonts["serif"], False, True)):
        cands = [(p, m) for p, m in metas if m[0].lower() == fam.lower()]
        best = None
        for p, (f, b, i, var) in cands:
            score = (b == bold or var) * 2 + (i == italic) * 3
            if best is None or score > best[0]:
                best = (score, p)
        if best:
            chosen[key] = best[1].relative_to(folder).as_posix()
    return chosen


def _google_font_files(family: str, out_dir: Path) -> list:
    """Fetch TTFs of a Google Fonts family (regular, bold, italic). Best effort."""
    import urllib.parse
    import urllib.request
    fam = urllib.parse.quote_plus(family)
    css = ""
    for axes in (":ital,wght@0,400;0,700;1,400;1,700", ":wght@400;700", ""):
        url = f"https://fonts.googleapis.com/css2?family={fam}{axes}"
        try:
            # a plain client gets TrueType URLs rather than woff2
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/4.0"})
            with urllib.request.urlopen(req, timeout=10) as r:
                css = r.read().decode("utf-8", "ignore")
            break
        except Exception:
            continue
    out = []
    for i, block in enumerate(re.findall(r"@font-face\s*{([^}]*)}", css)):
        m = re.search(r"url\((https://fonts\.gstatic\.com/[^)]+)\)", block)
        if not m:
            continue
        try:
            with urllib.request.urlopen(m.group(1), timeout=15) as r:
                data = r.read()
        except Exception:
            continue
        out_dir.mkdir(parents=True, exist_ok=True)
        style = "italic" if "italic" in block else "normal"
        weight = (re.search(r"font-weight:\s*(\d+)", block) or [None, "400"])[1]
        suffix = ".otf" if data[:4] == b"OTTO" else ".ttf"
        dest = out_dir / f"{family.replace(' ', '')}-{weight}-{style}-{i}{suffix}"
        dest.write_bytes(data)
        out.append(dest)
    return out


# ---- logos and symbols

def _to_png(raw: bytes, suffix: str, height: int = 240) -> bytes:
    from PIL import Image
    if suffix == ".svg" or raw.lstrip()[:5] in (b"<?xml", b"<svg ") or b"<svg" in raw[:400]:
        import resvg_py
        text = raw.decode("utf-8", "ignore")
        png = bytes(resvg_py.svg_to_bytes(svg_string=text, height=height))
        return png
    try:
        img = Image.open(io.BytesIO(raw))
        img.load()
    except Exception:
        raise ValueError("That image can't be read. Use PNG, JPG, WebP or SVG.")
    img = img.convert("RGBA")
    if img.height > height * 2:
        img = img.resize((max(1, int(img.width * height * 2 / img.height)), height * 2))
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def _collect_logos(base: Path, folder: Path) -> list:
    out_dir = folder / "logos"
    cands = []
    for p in sorted(base.rglob("*")):
        rel = p.relative_to(base).as_posix().lower()
        if p.suffix.lower() in IMAGE_SUFFIXES and ("logos/" in rel or "logo" in p.name.lower()):
            cands.append(p)
    out = []
    for p in cands[:24]:
        try:
            png = _to_png(p.read_bytes(), p.suffix.lower())
        except Exception:
            continue
        out_dir.mkdir(parents=True, exist_ok=True)
        dest = out_dir / (p.stem + ".png")
        dest.write_bytes(png)
        out.append(dest.relative_to(folder).as_posix())
    return out


def _pick_logo(logos: list, dark: bool) -> str:
    """Guess by file name: 'white'/'light'/'reverse' marks for dark grounds."""
    on_dark = [l for l in logos if re.search(r"white|light|reverse|negative|on-dark", l.lower())]
    co_brand = re.compile(r"gfk|powered|lockup|partner", re.I)
    if dark:
        pool = [l for l in on_dark if not co_brand.search(l)] or on_dark
    else:
        rest = [l for l in logos if l not in on_dark]
        pool = [l for l in rest if not co_brand.search(l)] or rest
        pool.sort(key=lambda l: (0 if re.search(r"bright|primary|color|colour|full", l.lower()) else 1, l))
    return pool[0] if pool else ""


def _collect_symbols(base: Path, folder: Path) -> list:
    src = next((p for p in base.rglob("*") if p.is_dir() and p.name.lower() in ("symbols", "icons")), None)
    if not src:
        return []
    files = [p for p in sorted(src.rglob("*")) if p.suffix.lower() in (".png", ".svg")]
    by_name = {}
    for p in files:
        stem = p.stem.lower()
        name = re.sub(r"-(bright-blue|deep-blue|primary|light|dark|white|color|colour)$", "", stem)
        name = re.sub(r"[^a-z0-9-]+", "-", name).strip("-")
        prefer = stem.endswith(("bright-blue", "primary", "color", "colour"))
        if name and (name not in by_name or prefer):
            by_name[name] = p
    out_dir = folder / "symbols"
    out = []
    for name, p in list(by_name.items())[:MAX_SYMBOLS]:
        try:
            png = _to_png(p.read_bytes(), p.suffix.lower(), height=256)
        except Exception:
            continue
        out_dir.mkdir(parents=True, exist_ok=True)
        (out_dir / f"{name}.png").write_bytes(png)
        out.append(name)
    return out


def heuristic_footer(readme: str) -> str:
    m = re.search(r"©[^`\n|]{3,120}", readme or "")
    if not m:
        return ""
    line = m.group(0).strip().rstrip(".").strip() + "."
    return re.sub(r"©\s*(19|20)\d\d", "© {year}", line)


# ---- Claude mapping

def _claude_mapping(api_key: str, tokens: dict, readme: str, logos: list, proposal: dict) -> dict:
    """Ask Claude to check the colour roles, fonts, logos and footer, and to
    distil the brand book into slide rules. Returns fields for theme.json."""
    from anthropic import Anthropic
    from pydantic import BaseModel, Field

    from .drafter import MODEL

    class Colors(BaseModel):
        white: str; deep: str; bright: str; light: str; orange: str
        tint: str; panel: str; ink: str; hairline: str

    class Mapping(BaseModel):
        name: str = Field(description="Short template name, e.g. the brand name")
        colors: Colors
        sans_font: str
        serif_font: str
        logo_light: str = Field(description="Logo file for light backgrounds, from the list, or empty")
        logo_dark: str = Field(description="Logo file for dark or brand-colour backgrounds, from the list, or empty")
        footer: str = Field(description="Legal footer line; write the year as {year}; empty if none")
        brand_rules: str = Field(description="6-12 short bullet lines of brand rules that matter for slides")

    color_lines = []
    for t in (tokens.get("color") or {}).get("tokens") or []:
        if isinstance(t, dict):
            color_lines.append(f"- {t.get('name')}: {json.dumps(t.get('value'))} — {str(t.get('usage', ''))[:160]}")
    families = json.dumps((tokens.get("type") or {}).get("families") or {})
    roles = "\n".join(f"- {k}: {v}" for k, v in ROLE_LABELS.items())
    prompt = (
        "Map this brand's design system onto a slide renderer's colour roles.\n\n"
        f"Roles (every value must be #RRGGBB):\n{roles}\n"
        "Rules: white text sits on 'bright' and 'deep' fills, so both must be dark enough for "
        "that; 'deep' is also title text on 'white'; 'ink' is body text on 'white' and 'panel'; "
        "'light' is used as text on 'deep'. Prefer the brand's semantic tokens.\n\n"
        f"Current guess:\n{json.dumps(proposal['colors'])}\n"
        f"Fonts guess: {json.dumps(proposal['fonts'])}\n"
        f"Font family tokens: {families}\n"
        f"Logo files: {json.dumps(logos)}\n\n"
        "Colour tokens:\n" + "\n".join(color_lines[:200]) + "\n\n"
        "Brand book (README):\n<readme>\n" + (readme or "(none)")[:12000] + "\n</readme>"
    )
    client = Anthropic(api_key=api_key)
    resp = client.messages.parse(
        model=MODEL, max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
        output_format=Mapping,
    )
    m = resp.parsed_output
    if m is None:
        raise RuntimeError("no mapping returned")
    colors = {}
    for role in ROLES:
        v = _hex(getattr(m.colors, role))
        colors[role] = v or proposal["colors"][role]
    out = {"colors": colors,
           "fonts": {"sans": (m.sans_font or proposal["fonts"]["sans"]).strip().strip("'\"")[:60],
                     "serif": (m.serif_font or proposal["fonts"]["serif"]).strip().strip("'\"")[:60]}}
    if m.name.strip():
        out["name"] = m.name.strip()[:60]
    for key in ("logo_light", "logo_dark"):
        v = getattr(m, key).strip()
        if v in logos:
            out[key] = v
        elif v == "":
            out[key] = ""
    if m.footer.strip() or not proposal.get("footer"):
        out["footer"] = m.footer.strip()[:160]
    if m.brand_rules.strip():
        out["brand_rules"] = m.brand_rules.strip()[:4000]
    return out
