"""NDS — turn raw source text into a deck plan with narration, via the Claude API."""
from anthropic import Anthropic

from . import themes
from .models import DeckPlan, Slide

MODEL = "claude-opus-4-8"

SYSTEM_TEMPLATE = """You are the drafting engine of NDS (Narrated Deck Studio), a tool that turns a \
source document into a narrated PowerPoint presentation.

From the source material, produce a complete deck plan:

Structure rules:
- The first slide must use layout "title" (deck title + subtitle; bullets empty).
- The last slide must use layout "closing" (short thank-you / key takeaway; bullets may hold 1-3 \
takeaways).
- Use layout "section" sparingly to mark major chapters (title only, bullets empty).
- Default deck length is 8-14 slides unless the user guidance says otherwise.

Layout vocabulary — VARY the layouts; a deck of nothing but bullet slides is a failure:
- "content": 3-6 short bullets (max ~12 words each). Fine for genuinely list-like material, but \
never use it for more than two slides in a row.
- "cards": 2-4 feature cards, each with a short "title" (2-5 words), a one-sentence "desc" and an \
"icon": {icon_rule} \
Perfect for pillars, principles, workstreams, options, personas.
- "stats": 2-5 headline numbers, each with a "value" (e.g. "160PB+", "90%", "$7.4T") and a short \
"label". Put the most important number FIRST: it becomes the slide's big callout. Use whenever the source contains strong figures — numbers deserve their own slide.
- "compare": two side-by-side panels (compare_left / compare_right), each with a "heading" and \
3-6 short items. Perfect for do/don't, permitted/prohibited, before/after, pros/cons, us/them.
- On every content-style slide ("content", "cards", "stats", "compare"), also set "subtitle" to a \
one-line kicker that frames the slide's message (max ~14 words) — e.g. "One unified strategy. \
Trusted data at the core."
- Populate ONLY the fields of the chosen layout (e.g. a "cards" slide has cards, empty bullets).

{design_rules}Narration rules (the most important part):
- Every slide gets a "narration" field: the exact words an artificial voice will speak while the \
slide is shown.
- Write for the ear, not the eye: complete sentences, spoken register, natural transitions \
("Let's move on to...", "So what does this mean?").
- Do NOT read the bullets verbatim — the narration explains and connects them.
- Aim for 40-90 words of narration per content slide (roughly 20-40 seconds of speech); the title \
slide gets a short welcome of 2-3 sentences.
- Never include stage directions, markdown, emoji, or text in brackets — only speakable words.
- Write narration in the requested language. Slide text follows the same language.

Stay faithful to the source material; do not invent facts that are not in it."""

NIQ_RULES = """NIQ PowerPoint compliance (this deck is for NielsenIQ — mandatory):
Template & system:
- Design within the official NIQ visual system that NDS renders (NIQ blues, Arial/Georgia, \
approved layouts). Do NOT invent a new template, theme, footer system, or slide master.
- Do not invent logos, lettermarks, or cropped "N" monograms. Assume the NIQ logo appears \
only as the official bright-blue mark with clear space — never place text on or through it, \
and keep breathing room around where a logo would sit (especially title/closing slides).
- Default to standard NIQ branding for internal content. Mention co-branding (GfK / Tech & \
Durables) ONLY if the source or user guidance explicitly requires those approved rules.

Palette & typography (content decisions that drive the rendered deck):
- Majority of the deck must read as NIQ primary blues and white. Secondary colors (orange, \
green, pink, amber, cyan accents) are sparingly for distinction — never let them overpower \
brand blues, and never treat secondary colors as full-slide backgrounds (grey tints/shades \
are the only common secondary exception).
- Fonts for Microsoft deliverables: Arial for headings and body; Georgia italic only for \
emphasis kickers. Keep font mixing restrained.
- Sentence case for titles, kickers, card titles, headings, and labels — NEVER all caps \
(except unavoidable acronyms already in the source).
- Strong legibility and contrast: short lines, high-contrast wording, no low-contrast or \
decorative text tricks.

Layout & density:
- Keep breathing room / white space on every slide. Sparse text; the narration carries detail.
- Deliver an INSIGHT, not a data dump. One clear takeaway per slide. Titles state the \
takeaway ("Innovation drives 9% of category sales"), not bare topics ("Innovation").
- Use icons sparingly in descriptions — prefer layout structure (cards/stats/compare) over \
icon-heavy concepts. Do not call for isometric or bespoke illustrations.
- If a chart, graph, or set of figures is the main point, give it the "stats" (or compare) \
layout and let numbers dominate the slide — do not bury key figures in a bullet list.
- Charts/figures must be described for NIQ's official data-viz approach: brand Custom Colors \
/ NIQ data palette — never PowerPoint's default rainbow chart palette. NIQ is a NEUTRAL \
measurer: third-party market data is not "painted" as NIQ's own with brand blue fills; keep \
data framing neutral and factual.
- Favor rounded-card layouts ("cards") for pillars/principles/workstreams (signature NIQ \
motif). Use "compare" for do/don't, before/after, permitted/prohibited.
- Spell out acronyms on first use. If featuring a quotation, keep it short and attributed.
- Prefer fewer, sharper slides over dense ones. Stay an insight vehicle, not a data appendix.

"""

TEMPLATE_RULES = """Design rules (a custom template renders the deck):
- The template fixes colours, fonts, logo and footer; choose content and layouts only. Do not \
invent logos, themes or footers.
- Keep breathing room on every slide. Sparse text; the narration carries detail.
- Deliver an INSIGHT, not a data dump. One clear takeaway per slide. Titles state the \
takeaway, not a bare topic.
- If figures are the main point, use the "stats" layout and let numbers dominate the slide.
- Use "cards" for pillars, principles and workstreams; "compare" for do/don't, before/after.
- Spell out acronyms on first use. Keep quotations short and attributed.
{brand_rules}
"""


def system_prompt(theme=None) -> str:
    """The drafting system prompt for a template (id or theme dict). NIQ
    compliance rules apply only to the built-in NIQ template."""
    t = themes.load(theme or "niq")
    if t["symbols"]:
        icon_rule = ("the brand symbol that best fits the card, chosen ONLY from: "
                     + ", ".join(t["symbols"])
                     + '. Use each symbol at most once per slide; leave "icon" empty if none fits.')
    else:
        icon_rule = 'leave "icon" empty (this template has no symbols).'
    if t["niq"]:
        rules = NIQ_RULES
    else:
        brand = t["brand_rules"].strip()
        rules = TEMPLATE_RULES.replace(
            "{brand_rules}",
            ("\nBrand rules from the template (follow them where they apply to slide content):\n"
             + brand + "\n") if brand else "")
    return (SYSTEM_TEMPLATE.replace("{icon_rule}", icon_rule)
            .replace("{design_rules}", rules))


SYSTEM = system_prompt("niq")


CONVERSATION_RULES = """
Narration style for THIS deck: a two-voice conversation. Write every slide's narration as a
short natural dialogue between two hosts, Alex and Sam. Format STRICTLY as one speaker turn
per line, each line starting with the name and a colon:
Alex: <sentence or two>
Sam: <sentence or two>
Alex opens the deck; they alternate naturally (a slide may have 2-5 turns). Keep the warm,
curious tone of a good podcast: Sam asks sharp questions or reacts, Alex explains. No stage
directions, only speakable words after each name."""


def draft_deck(
    source_text: str,
    *,
    api_key: str,
    language: str = "English",
    guidance: str = "",
    narration_style: str = "single",
    theme=None,
) -> DeckPlan:
    client = Anthropic(api_key=api_key)

    user_prompt = (
        f"Language for slides and narration: {language}\n"
        + (f"User guidance: {guidance}\n" if guidance.strip() else "")
        + (CONVERSATION_RULES if narration_style == "conversation" else "")
        + "\nSource material:\n<source>\n"
        + source_text
        + "\n</source>\n\nProduce the deck plan now."
    )

    response = client.messages.parse(
        model=MODEL,
        max_tokens=16000,
        thinking={"type": "adaptive"},
        system=system_prompt(theme),
        messages=[{"role": "user", "content": user_prompt}],
        output_format=DeckPlan,
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("The drafting model declined this request. Try rephrasing the source or guidance.")
    if response.parsed_output is None:
        raise RuntimeError("The drafting model returned no deck plan. Please try again.")
    return response.parsed_output


def draft_narration_for_existing(
    slides: list,
    *,
    api_key: str,
    language: str = "English",
    guidance: str = "",
    narration_style: str = "single",
) -> list:
    """Write one narration per slide of an EXISTING deck (design untouched).

    Where a slide already has speaker notes, they are used as the basis (polished
    for speech); otherwise narration is written from the slide's visible text.
    Returns a list of narration strings, one per slide, same order.
    """
    from typing import List

    from pydantic import BaseModel

    class NarrationPlan(BaseModel):
        narrations: List[str]

    client = Anthropic(api_key=api_key)

    slide_blocks = []
    for s in slides:
        block = f"--- Slide {s['index']} ---\n" + "\n".join(s["texts"])
        if s["notes"]:
            block += f"\n[Existing speaker notes]: {s['notes']}"
        slide_blocks.append(block)

    user_prompt = (
        f"Language for the narration: {language}\n"
        + (f"User guidance: {guidance}\n" if guidance.strip() else "")
        + (CONVERSATION_RULES if narration_style == "conversation" else "")
        + "\nThis is an EXISTING presentation. Do NOT redesign it — your only job is to "
        "write one narration SCRIPT per slide, in slide order. These scripts are the basis "
        "for any later voice narration (TTS or original audio).\n"
        "- If a slide has existing speaker notes, use them as the narration, lightly "
        "polished for natural speech (fix fragments, keep the author's meaning).\n"
        "- If a slide has NO speaker notes, invent nothing beyond the page: write 40-90 "
        "words of ear-friendly narration from that slide's visible titles, bullets, and "
        "labels only, with natural transitions between slides.\n"
        f"- Return exactly {len(slides)} narrations.\n\n"
        + "\n\n".join(slide_blocks)
    )

    response = client.messages.parse(
        model=MODEL,
        max_tokens=16000,
        thinking={"type": "adaptive"},
        system=SYSTEM,
        messages=[{"role": "user", "content": user_prompt}],
        output_format=NarrationPlan,
    )
    if response.stop_reason == "refusal" or response.parsed_output is None:
        raise RuntimeError("The model could not draft the narration. Please try again.")
    narrations = response.parsed_output.narrations
    if len(narrations) < len(slides):  # tolerate a short answer rather than crash
        narrations += [""] * (len(slides) - len(narrations))
    return narrations[: len(slides)]


def regenerate_slide(
    source_text: str,
    plan: DeckPlan,
    index: int,
    instruction: str,
    *,
    api_key: str,
    language: str = "English",
    narration_only: bool = False,
    narration_style: str = "single",
    theme=None,
) -> Slide:
    """Redraft a single slide, keeping the rest of the deck as context."""
    client = Anthropic(api_key=api_key)
    outline = "\n".join(
        f"{i + 1}. [{s.layout}] {s.title}" for i, s in enumerate(plan.slides)
    )
    current = plan.slides[index]
    task = (
        f"User instruction for the redraft: {instruction.strip()}"
        if instruction.strip()
        else "Produce an improved alternative take on this slide."
    )
    lock = (
        "This is an EXISTING presentation — the slide design is untouchable. Redraft ONLY "
        "the 'narration' field; return every other field exactly as given.\n"
        if narration_only
        else "Keep the same layout unless the instruction says otherwise. "
    )
    user_prompt = (
        f"Language for slides and narration: {language}\n"
        + (CONVERSATION_RULES if narration_style == "conversation" else "")
        + f"\nDeck outline (for context):\n{outline}\n\n"
        f"You are redrafting slide {index + 1} ONLY. Its current version:\n"
        f"{current.model_dump_json(indent=2)}\n\n"
        f"{task}\n"
        + lock
        + "Keep the slide coherent with the slides around it (no repeated content, "
        "natural narration hand-off).\n\n"
        f"Source material:\n<source>\n{source_text}\n</source>\n\n"
        "Return the redrafted slide now."
    )

    response = client.messages.parse(
        model=MODEL,
        max_tokens=8000,
        thinking={"type": "adaptive"},
        system=system_prompt(theme),
        messages=[{"role": "user", "content": user_prompt}],
        output_format=Slide,
    )
    if response.stop_reason == "refusal" or response.parsed_output is None:
        raise RuntimeError("The model could not redraft this slide. Try a different instruction.")
    return response.parsed_output
