"""pptx2md (MIT) — PPTX-specific Markdown with extracted images."""
from __future__ import annotations

import importlib.util
from pathlib import Path

from ..schema import ConverterOutput
from .base import Adapter


class Pptx2mdAdapter(Adapter):
    name = "pptx2md"
    license = "MIT"
    note = "Titles, nested lists, formatting, merged-cell tables, extracted images."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        if importlib.util.find_spec("pptx2md") is None:
            return False, "not installed (pip install pptx2md)"
        return True, "pptx2md"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        from pptx2md import ConversionConfig, convert

        out_md = outdir / "output.md"
        img_dir = outdir / "images"

        # Only pass kwargs this pptx2md version actually supports.
        try:
            fields = set(ConversionConfig.model_fields)  # pydantic model
        except Exception:
            fields = set()
        kwargs = {"pptx_path": pptx, "output_path": out_md, "image_dir": img_dir}
        for key, val in (("disable_notes", False), ("enable_slides", True)):
            if key in fields:
                kwargs[key] = val

        convert(ConversionConfig(**kwargs))

        text = out_md.read_text(encoding="utf-8") if out_md.exists() else ""
        imgs = [str(p) for p in img_dir.glob("*")] if img_dir.exists() else []
        return ConverterOutput(
            converter=self.name,
            status="ok",
            markdown=text,
            image_paths=imgs,
            meta={"enable_slides": "enable_slides" in fields},
        )
