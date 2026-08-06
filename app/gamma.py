"""NDS — Gamma design engine.

Sends the source text to Gamma's Generations API, which designs a full
presentation (theme, layouts, imagery) and exports it as .pptx. NDS then
narrates that deck exactly like a user-uploaded one.

Requires a Gamma API key (Settings). Gamma's API is in beta and available on
paid Gamma plans; key comes from gamma.app account settings.
"""
import json
import time
import urllib.request
from pathlib import Path
from typing import Callable, Optional

API = "https://public-api.gamma.app/v1.0"

LANGUAGE_CODES = {
    "english": "en", "french": "fr", "german": "de", "spanish": "es", "italian": "it",
}

# NIQ PowerPoint compliance brief for Gamma (kept tight — field is length-limited).
NIQ_BRAND_BRIEF = (
    "Create this deck as if using the latest official NIQ PowerPoint template: keep the visual "
    "system intact — do not invent a new template or alter a slide master. NIQ logo: official "
    "bright-blue mark only, with clear space; never crop to a letter or add text to the logo. "
    "Dominant palette: NIQ primary blues (#060A45, #2C6DF6) and white; secondary colors "
    "(#31D1FF, #EF5F17, #59AD00, #EF5890, #FFB500) sparingly and never as full backgrounds "
    "(grey tints OK). Fonts: Arial for headings/body, Georgia italic for emphasis only — "
    "restrained mixing. Sentence case everywhere (no all-caps titles). Strong contrast, "
    "generous white space, icons sparingly. One insight per slide; takeaway-style titles. "
    "If figures/charts are the point, let them dominate the slide. Recolor charts with NIQ "
    "Custom Colors / official data-viz palette — never PowerPoint default rainbows. NIQ is a "
    "neutral measurer: do not paint third-party data as NIQ brand blue. No isometric/bespoke "
    "illustrations. Internal decks: standard NIQ branding unless source requires approved "
    "GfK or Tech & Durables co-branding."
)


class GammaError(RuntimeError):
    pass


# Gamma's API is behind Cloudflare, which rejects Python's default urllib
# signature (error 1010) — send a normal browser-style User-Agent.
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 NDS/1.0")


def _request(url: str, api_key: str, payload: Optional[dict] = None) -> dict:
    headers = {"X-API-KEY": api_key, "Accept": "application/json", "User-Agent": UA}
    data = None
    if payload is not None:
        data = json.dumps(payload).encode()
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        detail = e.read().decode()[:300]
        raise GammaError(f"Gamma API error {e.code}: {detail}")


def _find_pptx_url(obj) -> Optional[str]:
    """Gamma's response schema is beta — find the pptx link wherever it lives."""
    if isinstance(obj, str):
        return obj if ".pptx" in obj.split("?")[0].lower() else None
    if isinstance(obj, dict):
        for k, v in obj.items():
            if "pptx" in k.lower() and isinstance(v, str) and v.startswith("http"):
                return v
            found = _find_pptx_url(v)
            if found:
                return found
    if isinstance(obj, list):
        for v in obj:
            found = _find_pptx_url(v)
            if found:
                return found
    return None


def generate_deck(
    api_key: str,
    source_text: str,
    out_path: Path,
    *,
    language: str = "English",
    guidance: str = "",
    num_cards: Optional[int] = None,
    theme_name: str = "",
    brand: bool = True,
    progress: Callable[[int, str], None] = lambda pct, msg: None,
    timeout_s: int = 600,
) -> Path:
    """Generate a Gamma-designed deck from source text; download the .pptx.

    When ``brand`` is set, the NIQ 2026 design brief is prepended to Gamma's
    additionalInstructions so the deck comes out on brand. ``theme_name`` (from
    Settings) applies a saved Gamma theme — the most reliable way to lock brand
    colours/fonts if the user has a NIQ theme in their Gamma workspace.
    """
    payload = {
        "inputText": source_text[:100_000],
        "textMode": "generate",
        "format": "presentation",
        "exportAs": "pptx",
        "textOptions": {"language": LANGUAGE_CODES.get(language.lower(), "en")},
    }
    if num_cards:
        payload["numCards"] = num_cards
    if theme_name.strip():
        payload["themeName"] = theme_name.strip()
    # Combine the brand brief with the user's own guidance, newest-first, and keep
    # within Gamma's instruction length budget.
    parts = []
    if brand:
        parts.append(NIQ_BRAND_BRIEF)
    if guidance.strip():
        parts.append(guidance.strip())
    if parts:
        payload["additionalInstructions"] = "\n\n".join(parts)[:2000]

    progress(5, "Sending your content to Gamma…")
    data = _request(f"{API}/generations", api_key, payload)
    generation_id = data.get("generationId") or data.get("id")
    if not generation_id:
        raise GammaError(f"Unexpected Gamma response: {json.dumps(data)[:200]}")

    start = time.time()
    while True:
        if time.time() - start > timeout_s:
            raise GammaError("Gamma generation timed out after 10 minutes.")
        time.sleep(6)
        status_data = _request(f"{API}/generations/{generation_id}", api_key)
        status = (status_data.get("status") or "").lower()
        elapsed = int(time.time() - start)
        if status in ("completed", "succeeded", "done"):
            pptx_url = status_data.get("exportUrl") or _find_pptx_url(status_data)
            if not pptx_url:
                raise GammaError(
                    "Gamma finished but returned no .pptx link — "
                    f"response: {json.dumps(status_data)[:300]}"
                )
            progress(90, "Downloading the designed deck from Gamma…")
            dl = urllib.request.Request(pptx_url, headers={"User-Agent": UA})
            with urllib.request.urlopen(dl, timeout=300) as resp:
                out_path.write_bytes(resp.read())
            return out_path
        if status in ("failed", "error"):
            raise GammaError(f"Gamma generation failed: {json.dumps(status_data)[:300]}")
        progress(min(85, 10 + elapsed // 2), f"Gamma is designing your deck… ({elapsed}s)")
