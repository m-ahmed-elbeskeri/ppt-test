"""Datalab Marker (GPL) — high-quality parser via PPTX->PDF. Opt-in (PyTorch).

Marker is PDF-first and downloads models on first run. We render the deck to PDF
via PowerPoint COM, then run Marker. GPL — see the README license table.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

from ..schema import ConverterOutput
from .base import Adapter


class MarkerAdapter(Adapter):
    name = "marker"
    license = "GPL-3.0"
    heavy = True
    note = "PPTX->PDF -> Marker (PyTorch models); high accuracy, heavy first run."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        if importlib.util.find_spec("marker") is None:
            return False, "not installed (pip install -r requirements-heavy.txt)"
        from ..render import available as render_available

        ok, info = render_available()
        if not ok:
            return False, f"PDF step needs PowerPoint COM: {info}"
        return True, "marker + PowerPoint COM"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        from marker.converters.pdf import PdfConverter
        from marker.models import create_model_dict
        from marker.output import text_from_rendered

        from ..render import pptx_to_pdf

        pdf = outdir / "deck.pdf"
        ok, err = pptx_to_pdf(pptx, pdf)
        if not ok:
            raise RuntimeError(f"PowerPoint PDF export failed: {err}")

        converter = PdfConverter(artifact_dict=create_model_dict())
        rendered = converter(str(pdf))
        text, _, images = text_from_rendered(rendered)

        img_paths: list[str] = []
        if isinstance(images, dict):
            img_dir = outdir / "images"
            img_dir.mkdir(parents=True, exist_ok=True)
            for name, img in images.items():
                try:
                    dest = img_dir / Path(name).name
                    img.save(dest)
                    img_paths.append(str(dest))
                except Exception:
                    continue
        return ConverterOutput(
            converter=self.name,
            status="ok",
            markdown=text or "",
            image_paths=img_paths,
            meta={"via": "powerpoint-pdf+marker"},
        )
