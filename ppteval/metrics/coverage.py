"""Coverage metrics: how much ground-truth content survived the conversion.

All are RELATIVE to python-pptx-extractable content (see groundtruth.py caveat).
Returns None where a deck has nothing to measure (e.g. no notes / no tables), so
the report can skip that dimension instead of scoring a vacuous 1.0 or 0.0.
"""
from __future__ import annotations

import re

from ..schema import ConverterOutput, GroundTruth

_WORD = re.compile(r"[a-z0-9]+")


def _words(text: str) -> set[str]:
    return set(_WORD.findall((text or "").lower()))


def _round(x):
    return round(x, 4) if isinstance(x, float) else x


def coverage_metrics(out: ConverterOutput, gt: GroundTruth) -> dict:
    text = out.markdown or ""
    out_words = _words(text)

    gt_words = _words(gt.all_text())
    text_recall = (len(gt_words & out_words) / len(gt_words)) if gt_words else None

    notes_words = _words(gt.notes_text())
    notes_recall = (
        (len(notes_words & out_words) / len(notes_words)) if notes_words else None
    )

    # Table cells: a cell "survived" if ALL its words appear in the output. Word-set
    # (not contiguous substring) so that correctly splitting a cell into pipe-
    # separated columns isn't penalised vs a plain-text dump. Restrict to cells with
    # >=3 chars and a letter so trivial cells ("1", "%") don't inflate the score.
    cells = [
        c
        for c in gt.all_table_cells()
        if c and len(c.strip()) >= 3 and re.search(r"[a-zA-Z]", c)
    ]
    if cells:
        hits = 0
        for c in cells:
            cw = _words(c)
            if cw and cw <= out_words:
                hits += 1
        table_cell_recall = hits / len(cells)
    else:
        table_cell_recall = None

    return {
        "text_recall": _round(text_recall),
        "notes_recall": _round(notes_recall),
        "table_cell_recall": _round(table_cell_recall),
    }
