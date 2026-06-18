"""PyMuPDF4LLM (AGPL) — render PPTX->PDF (PowerPoint COM) then parse to Markdown.

Represents the "visual fidelity" path: the deck is rendered to PDF (preserving
layout/reading order) and PyMuPDF4LLM produces page-chunked Markdown. Requires
PowerPoint COM for the PDF step (no LibreOffice here).
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

from ..schema import ConverterOutput
from .base import Adapter


class Pymupdf4llmAdapter(Adapter):
    name = "pymupdf4llm"
    license = "AGPL-3.0"
    note = "PPTX->PDF (PowerPoint) -> layout-aware Markdown; needs COM for PDF step."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        if importlib.util.find_spec("pymupdf4llm") is None:
            return False, "not installed (pip install pymupdf4llm)"
        from ..render import available as render_available

        ok, info = render_available()
        if not ok:
            return False, f"PDF step needs PowerPoint COM: {info}"
        return True, "pymupdf4llm + PowerPoint COM"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        import pymupdf4llm

        from ..render import pptx_to_pdf

        pdf = outdir / "deck.pdf"
        ok, err = pptx_to_pdf(pptx, pdf)
        if not ok:
            raise RuntimeError(f"PowerPoint PDF export failed: {err}")

        # Extract embedded images to disk + reference them (default writes none).
        img_dir = outdir / "images"
        img_dir.mkdir(parents=True, exist_ok=True)
        try:
            md = pymupdf4llm.to_markdown(
                str(pdf), write_images=True, image_path=str(img_dir), image_format="png"
            )
        except TypeError:
            # older signature without image kwargs
            md = pymupdf4llm.to_markdown(str(pdf))
        if isinstance(md, list):  # page_chunks form
            md = "\n\n".join(
                (c.get("text", "") if isinstance(c, dict) else str(c)) for c in md
            )
        imgs = [str(p) for p in img_dir.glob("*") if p.is_file()]
        return ConverterOutput(
            converter=self.name,
            status="ok",
            markdown=md,
            image_paths=imgs,
            meta={"via": "powerpoint-pdf+pymupdf4llm", "pdf": str(pdf)},
        )
