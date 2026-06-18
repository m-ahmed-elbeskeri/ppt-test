"""Unstructured (Apache-2.0 core) — typed elements -> Markdown. Opt-in."""
from __future__ import annotations

import importlib.util
from pathlib import Path

from ..schema import ConverterOutput
from .base import Adapter


class UnstructuredAdapter(Adapter):
    name = "unstructured"
    license = "Apache-2.0"
    heavy = True
    note = "Typed elements (Title/NarrativeText/Table/...); RAG-friendly partitioning."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        if importlib.util.find_spec("unstructured") is None:
            return False, "not installed (pip install -r requirements-heavy.txt)"
        return True, "unstructured"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        from unstructured.partition.pptx import partition_pptx

        els = partition_pptx(filename=str(pptx), include_page_breaks=True)
        lines: list[str] = []
        cur_page = None
        for el in els:
            page = getattr(getattr(el, "metadata", None), "page_number", None)
            if page is not None and page != cur_page:
                cur_page = page
                lines.append(f"\n## Slide {page}\n")
            cat = getattr(el, "category", type(el).__name__)
            text = (getattr(el, "text", "") or "").strip()
            if cat == "Title":
                if text:
                    lines.append(f"### {text}")
            elif cat == "Table":
                html = getattr(getattr(el, "metadata", None), "text_as_html", None)
                lines.append(html or text)
            elif text:
                lines.append(text)
        md = "\n".join(lines)
        return ConverterOutput(
            converter=self.name, status="ok", markdown=md, meta={"n_elements": len(els)}
        )
