"""Microsoft MarkItDown — PPTX -> Markdown (MIT)."""
from __future__ import annotations

import importlib.util
from pathlib import Path

from ..schema import ConverterOutput
from .base import Adapter


class MarkitdownAdapter(Adapter):
    name = "markitdown"
    license = "MIT"
    note = "LLM-oriented Markdown; token-efficient; OCR via optional plugin."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        if importlib.util.find_spec("markitdown") is None:
            return False, "not installed (pip install markitdown)"
        return True, "markitdown"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        from markitdown import MarkItDown

        res = MarkItDown().convert(str(pptx))
        text = getattr(res, "markdown", None) or getattr(res, "text_content", "") or ""
        return ConverterOutput(
            converter=self.name, status="ok", markdown=text, meta={"lib": "markitdown"}
        )
