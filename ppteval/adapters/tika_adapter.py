"""Apache Tika (Apache-2.0) — mature text+metadata extraction (needs a JRE).

Baseline: strong, cheap plain-text extraction across many formats, but weak on
slide hierarchy/tables/images. The tika-python package auto-downloads a
tika-server jar on first use (requires Java on PATH and network once).
"""
from __future__ import annotations

import importlib.util
import shutil
from pathlib import Path

from ..schema import ConverterOutput
from .base import Adapter


class TikaAdapter(Adapter):
    name = "tika"
    license = "Apache-2.0"
    note = "Plain text + metadata; mature ingestion; weak on slide structure."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        if importlib.util.find_spec("tika") is None:
            return False, "not installed (pip install tika)"
        if shutil.which("java") is None:
            return False, "Java runtime not found on PATH (Tika needs a JRE)"
        return True, "tika + Java"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        from tika import parser

        parsed = parser.from_file(str(pptx))
        text = (parsed.get("content") or "").strip()
        meta = parsed.get("metadata") or {}
        # Keep meta small/serialisable.
        keep = {
            k: meta[k]
            for k in ("Content-Type", "slide-count", "xmpTPg:NPages")
            if k in meta
        }
        return ConverterOutput(
            converter=self.name, status="ok", markdown=text, meta={"tika_meta": keep}
        )
