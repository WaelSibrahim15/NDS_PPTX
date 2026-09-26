"""NDS — deck plan data model shared by the drafter, the builder, and the renderer."""
from typing import List, Literal, Optional

from pydantic import BaseModel


class Card(BaseModel):
    title: str
    desc: str
    icon: Optional[str] = None   # NIQ brand symbol name (see design.SYMBOLS)


class Stat(BaseModel):
    value: str   # e.g. "160PB+", "28,000", "90%"
    label: str   # e.g. "managed data"


class CompareSide(BaseModel):
    heading: str          # e.g. "Permitted", "Before"
    items: List[str]


class ChartPoint(BaseModel):
    label: str    # e.g. "EMEA", "2024"
    value: float  # plotted value, in chart_unit


LAYOUTS = ("title", "section", "content", "cards", "stats", "compare", "closing",
           "quote", "statement", "agenda", "timeline", "chart")


class Slide(BaseModel):
    layout: Literal["title", "section", "content", "cards", "stats", "compare", "closing",
                    "quote", "statement", "agenda", "timeline", "chart"]
    title: str
    subtitle: Optional[str] = None    # kicker line under the title on content-style slides
    bullets: List[str] = []
    cards: Optional[List[Card]] = None            # for layout "cards" (2-4 items)
    stats: Optional[List[Stat]] = None            # for layout "stats" (3-5 items)
    compare_left: Optional[CompareSide] = None    # for layout "compare"
    compare_right: Optional[CompareSide] = None
    steps: Optional[List[Card]] = None            # for layout "timeline" (3-6 steps, in order)
    chart: Optional[List[ChartPoint]] = None      # for layout "chart" (2-8 points)
    chart_unit: Optional[str] = None              # e.g. "%", "$M", "stores"
    variant: Optional[int] = None                 # visual version of the layout; None = automatic
    narration: str


class DeckPlan(BaseModel):
    deck_title: str
    subtitle: str
    slides: List[Slide]
