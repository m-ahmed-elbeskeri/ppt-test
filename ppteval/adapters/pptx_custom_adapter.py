"""Custom python-pptx extractor (MIT) — the "maximum control" option.

Emits the recommended LLM-ready schema (per-slide Markdown with metadata,
tables, speaker notes, image references) plus a lossless per-slide JSON. Because
it is built on the same extraction as the ground truth, it is the *reference /
upper-bound* converter: its text recall is ~1.0 by construction, so it competes
on structure, tokens, speed and human ratings — flagged as such in the report.
"""
from __future__ import annotations

from pathlib import Path

from ..schema import ConverterOutput, GroundTruth, SlideGT
from .base import Adapter, extract_media


def _esc(cell: str) -> str:
    return (cell or "").replace("|", "\\|").replace("\n", " ").strip()


def _md_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    width = max((len(r) for r in rows), default=0)
    rows = [list(r) + [""] * (width - len(r)) for r in rows]
    header = rows[0]
    lines = [
        "| " + " | ".join(_esc(c) for c in header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    for r in rows[1:]:
        lines.append("| " + " | ".join(_esc(c) for c in r) + " |")
    return "\n".join(lines)


def render_markdown(gt: GroundTruth, image_files: set[str]) -> str:
    out: list[str] = []
    out.append(f"# Deck: {gt.deck_slug}")
    out.append(
        f"\n_{gt.n_slides} slides · {gt.n_tables} tables · {gt.n_images} images · "
        f"{gt.n_charts} charts · {gt.n_notes} with speaker notes_\n"
    )
    for s in gt.slides:
        title = s.title or "(untitled)"
        out.append(f"\n## Slide {s.slide_number} — {title}")
        out.append(
            f"<!-- slide_number: {s.slide_number}; images: {len(s.image_rels)}; "
            f"tables: {len(s.tables)}; charts: {s.n_charts} -->"
        )
        body = [r for r in s.text_runs if r.strip()]
        if body:
            out.append("\n### Text")
            out.extend(f"- {line}" for line in body)
        for i, tbl in enumerate(s.tables, 1):
            out.append(f"\n### Table {i}")
            out.append(_md_table(tbl))
        if s.notes.strip():
            out.append("\n### Speaker notes")
            out.append(s.notes.strip())
        if s.image_rels:
            out.append("\n### Visuals")
            for fname in s.image_rels:
                rel = f"images/{fname}" if fname in image_files else fname
                out.append(f"![image: {fname}]({rel})")
    return "\n".join(out) + "\n"


class PptxCustomAdapter(Adapter):
    name = "pptx_custom"
    license = "MIT"
    note = "Reference extractor (defines ground truth); per-slide MD + JSON + images."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        import importlib.util

        if importlib.util.find_spec("pptx") is None:
            return False, "not installed (pip install python-pptx)"
        return True, "python-pptx"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        from ..groundtruth import extract_ground_truth

        gt = extract_ground_truth(pptx)

        # Extract referenced media so the schema's image links resolve.
        img_dir = outdir / "images"
        written = extract_media(pptx, img_dir)
        image_files = {Path(p).name for p in written}

        md = render_markdown(gt, image_files)
        jpath = outdir / "output.json"
        jpath.write_text(gt.model_dump_json(indent=2), encoding="utf-8")

        return ConverterOutput(
            converter=self.name,
            status="ok",
            markdown=md,
            json_path=str(jpath),
            image_paths=written,
            meta={"slides": gt.n_slides, "reference": True},
        )
