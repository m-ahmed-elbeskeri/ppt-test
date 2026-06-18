"""Typed, serialisable data contracts shared across the suite.

These pydantic models are the *only* coupling between stages: ground-truth ->
adapters -> metrics -> review -> report all exchange these shapes via JSON on
disk, so any stage can be run independently from the CLI.
"""
from __future__ import annotations

from typing import Any, Optional

from pydantic import BaseModel, Field


# --- ground truth ------------------------------------------------------------
class SlideGT(BaseModel):
    """Canonical, deterministic extraction for a single slide (python-pptx)."""

    slide_number: int
    title: Optional[str] = None
    text_runs: list[str] = Field(default_factory=list)
    # tables[t][row][col]
    tables: list[list[list[str]]] = Field(default_factory=list)
    notes: str = ""
    image_rels: list[str] = Field(default_factory=list)  # media filenames referenced
    n_charts: int = 0
    shape_count: int = 0

    @property
    def text(self) -> str:
        return "\n".join(r for r in self.text_runs if r and r.strip())

    @property
    def table_cells(self) -> list[str]:
        cells: list[str] = []
        for table in self.tables:
            for row in table:
                cells.extend(c for c in row if c and c.strip())
        return cells


class GroundTruth(BaseModel):
    """Deck-level ground truth plus deck totals (the metric denominators)."""

    deck_slug: str
    source_path: str
    n_slides: int = 0
    n_notes: int = 0
    n_tables: int = 0
    n_images: int = 0
    n_charts: int = 0
    gt_tokens: int = 0
    slides: list[SlideGT] = Field(default_factory=list)

    def all_text(self) -> str:
        parts: list[str] = []
        for s in self.slides:
            if s.title:
                parts.append(s.title)
            parts.append(s.text)
            parts.extend(s.table_cells)
        return "\n".join(p for p in parts if p and p.strip())

    def notes_text(self) -> str:
        return "\n".join(s.notes for s in self.slides if s.notes and s.notes.strip())

    def all_table_cells(self) -> list[str]:
        cells: list[str] = []
        for s in self.slides:
            cells.extend(s.table_cells)
        return cells


# --- converter output --------------------------------------------------------
class ConverterOutput(BaseModel):
    """Result of running one adapter on one deck.

    ``markdown`` holds the full text in memory for the metrics stage but is
    excluded from the slim on-disk record (the text lives in output.md).
    """

    converter: str
    status: str = "ok"  # ok | error | skipped
    elapsed_s: float = 0.0
    markdown: str = Field(default="", exclude=True)
    markdown_path: Optional[str] = None
    json_path: Optional[str] = None
    image_paths: list[str] = Field(default_factory=list)
    error: Optional[str] = None
    meta: dict[str, Any] = Field(default_factory=dict)


# --- metrics -----------------------------------------------------------------
class MetricRow(BaseModel):
    converter: str
    status: str = "ok"
    elapsed_s: float = 0.0
    # values may be float, int, or None (N/A, e.g. notes_recall when a deck has no notes)
    metrics: dict[str, Any] = Field(default_factory=dict)


# --- human ratings -----------------------------------------------------------
class Rating(BaseModel):
    converter: str
    mode: str  # deck | sample | allslides
    dimension: str  # one of RATING_DIMENSIONS, or "rank"
    score: float
    slide_number: Optional[int] = None
    note: str = ""
    ts: str = ""

    def key(self) -> tuple:
        return (self.converter, self.mode, self.slide_number, self.dimension)


class RatingsStore(BaseModel):
    deck_slug: str
    ratings: list[Rating] = Field(default_factory=list)

    def upsert(self, r: Rating) -> None:
        k = r.key()
        self.ratings = [x for x in self.ratings if x.key() != k]
        self.ratings.append(r)

    def has(self, converter: str, mode: str, dimension: str, slide_number=None) -> bool:
        return any(
            x.key() == (converter, mode, slide_number, dimension) for x in self.ratings
        )
