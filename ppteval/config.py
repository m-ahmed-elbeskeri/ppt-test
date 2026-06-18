"""Central configuration: paths, tokenizer, rating dimensions, and the tunable
weights used to fold per-converter metrics into a single ``auto_score``.

Everything an operator might reasonably tune lives here so the scoring policy is
auditable in one place (an "enterprise" requirement: no magic numbers buried in
logic).
"""
from __future__ import annotations

import re
from pathlib import Path

# --- paths -------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parents[1]
OUT_ROOT = REPO_ROOT / "out"
DECKS_DIR = REPO_ROOT / "tests" / "decks"

# --- tokenizer ---------------------------------------------------------------
# tiktoken encoding used for the `out_tokens` metric. This is a *consistent
# cross-converter proxy* for how many tokens a deck costs an LLM; it is NOT the
# exact tokenizer for any one model (e.g. Claude). Stated plainly in reports.
TOKENIZER = "o200k_base"

# --- human review ------------------------------------------------------------
# Dimensions the user scores 1-5 in `ppteval review`.
RATING_DIMENSIONS = ["faithfulness", "structure", "tables", "notes", "images", "overall"]
RATING_MIN, RATING_MAX = 1, 5

# --- auto-score policy -------------------------------------------------------
# Weights applied to NORMALIZED (0..1) metric values to form `auto_score`.
# Keys are normalized-metric names produced in report.py. Weights sum to ~1.0.
# Rationale: content fidelity dominates; speed is deliberately light so a fast
# converter that drops content cannot win on speed alone.
AUTO_WEIGHTS = {
    "text_recall": 0.28,        # how much of the deck's text survived
    "notes_recall": 0.14,       # speaker notes are a known differentiator
    "table_cell_recall": 0.12,  # tabular data fidelity
    "image_recall": 0.08,       # images referenced/extracted vs ground truth
    "structure_score": 0.12,    # headings/lists/tables/notes/slide-boundaries
    "compression_fit": 0.10,    # density: penalise both heavy loss and bloat
    "slides_per_sec": 0.06,     # throughput (min-max normalized across converters)
    "status_ok": 0.10,          # ran to completion without error
}

# Default blend when producing the *final* rank in the leaderboard.
# final = HUMAN_BLEND * human_score_norm + (1 - HUMAN_BLEND) * auto_score,
# but only for converters that actually received human ratings; otherwise the
# auto_score is used and the row is flagged "(auto only)".
HUMAN_BLEND = 0.6

# Ideal token-density band for `compression_fit`: out_tokens / gt_tokens.
# Below LO => likely dropping content; above HI => bloat/noise. Inside the band
# scores 1.0, decaying linearly to 0 at the FAR bounds.
COMPRESSION_BAND = (0.9, 1.6)
COMPRESSION_FAR = (0.4, 3.0)


def deck_slug(path: str | Path) -> str:
    """Filesystem-safe slug for a deck, used as its output folder name."""
    stem = Path(path).stem
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("_")
    return s or "deck"


def out_dir_for(path: str | Path) -> Path:
    return OUT_ROOT / deck_slug(path)
