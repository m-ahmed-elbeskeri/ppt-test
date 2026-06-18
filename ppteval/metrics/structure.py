"""Structure metrics: does the output preserve LLM-useful structure?

Headings, list items, Markdown tables, image references, and slide boundaries
(the last is what makes slide-level RAG chunking possible). ``structure_score``
is a 0..1 composite of formatting fidelity, kept separate from content recall so
the two aren't double-counted.
"""
from __future__ import annotations

import re

from ..schema import ConverterOutput, GroundTruth

_HEADING = re.compile(r"^\s{0,3}#{1,6}\s+\S", re.M)
_LIST = re.compile(r"^\s*([-*+]|\d+[.)])\s+\S", re.M)
# image references in any common form: Markdown ![](), HTML <img>, or the
# placeholder comment docling/others emit (<!-- image -->). All three count as
# "the converter preserved an anchor for this visual" (useful for a downstream
# VLM captioning pass). We deliberately do NOT count bare filename strings in a
# plain-text dump — that gives false credit (e.g. Tika) for no real handling.
_IMG_MD = re.compile(r"!\[[^\]]*\]\([^)]+\)")
_IMG_HTML = re.compile(r"<img[\s>]", re.I)
_IMG_COMMENT = re.compile(r"<!--\s*image", re.I)
_HTMLTABLE = re.compile(r"<table[\s>]", re.I)
# textual slide markers: "## Slide 3", "Slide number: 3", "Slide 3 -", etc.
_SLIDEMARK = re.compile(r"slide\s*(?:number)?\s*[:#\-]*\s*\d+", re.I)
# thematic breaks (---, ***, ___) used by some converters as slide separators
_HR = re.compile(r"^\s*(-{3,}|\*{3,}|_{3,})\s*$", re.M)
_TABLE_SEP_CHARS = set("|:- ")


def _count_table_separators(text: str) -> int:
    """A Markdown table separator row: only |,:,-,space; >=2 pipes; has a dash.

    Counts one per table and tolerates alignment colons (``| :-: | :-: |``),
    which a naive ``-{3,}`` pattern misses.
    """
    n = 0
    for line in text.splitlines():
        s = line.strip()
        if s.count("|") >= 2 and "-" in s and set(s) <= _TABLE_SEP_CHARS:
            n += 1
    return n


def _round(x):
    return round(x, 4) if isinstance(x, float) else x


def structure_metrics(out: ConverterOutput, gt: GroundTruth) -> dict:
    text = out.markdown or ""

    n_headings = len(_HEADING.findall(text))
    n_list = len(_LIST.findall(text))
    n_md_tables = _count_table_separators(text)
    n_html_tables = len(_HTMLTABLE.findall(text))
    # Both Markdown pipe tables and HTML tables are LLM-parseable structure.
    n_tables = n_md_tables + n_html_tables
    n_imgs = (
        len(_IMG_MD.findall(text))
        + len(_IMG_HTML.findall(text))
        + len(_IMG_COMMENT.findall(text))
    )
    images_extracted = len(out.image_paths or [])

    # slide boundaries: textual "Slide N" markers, thematic breaks, or form feeds
    n_bounds = len(_SLIDEMARK.findall(text)) + len(_HR.findall(text)) + text.count("\f")
    boundary_recall = min(1.0, n_bounds / gt.n_slides) if gt.n_slides else None

    # image_recall = does the output ANCHOR each visual (any ref syntax) so a
    # downstream VLM pass knows it exists? Whether the image *bytes* were written
    # is reported separately as images_extracted (a different capability).
    image_recall = min(1.0, n_imgs / gt.n_images) if gt.n_images else None

    comps = [
        min(1.0, n_headings / gt.n_slides) if gt.n_slides else None,
        min(1.0, n_list / gt.n_slides) if gt.n_slides else None,
        boundary_recall,
        (min(1.0, n_tables / gt.n_tables) if gt.n_tables else None),
    ]
    comps = [c for c in comps if c is not None]
    structure_score = (sum(comps) / len(comps)) if comps else None

    return {
        "n_headings": n_headings,
        "n_list_items": n_list,
        "n_md_tables": n_md_tables,
        "n_html_tables": n_html_tables,
        "n_image_refs": n_imgs,
        "images_extracted": images_extracted,
        "slide_boundary_recall": _round(boundary_recall),
        "image_recall": _round(image_recall),
        "structure_score": _round(structure_score),
    }
