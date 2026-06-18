"""Docling (MIT) — unified document model -> Markdown. Opt-in (larger deps)."""
from __future__ import annotations

import importlib.util
from pathlib import Path

from ..schema import ConverterOutput
from .base import Adapter


class DoclingAdapter(Adapter):
    name = "docling"
    license = "MIT"
    heavy = True
    note = "GenAI document pipeline; Markdown/JSON/DocTags; runs locally."

    @classmethod
    def available(cls) -> tuple[bool, str]:
        if importlib.util.find_spec("docling") is None:
            return False, "not installed (pip install -r requirements-heavy.txt)"
        return True, "docling"

    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        from docling.document_converter import DocumentConverter

        res = DocumentConverter().convert(str(pptx))
        doc = res.document

        # Docling's PPTX backend loads picture bitmaps into the document model, but
        # save_as_markdown(REFERENCED) doesn't write them for PPTX (dangling refs).
        # So extract the bitmaps straight from the model — reliable, no torch needed.
        img_dir = outdir / "images"
        img_dir.mkdir(parents=True, exist_ok=True)
        imgs: list[str] = []
        for i, pic in enumerate(getattr(doc, "pictures", []) or []):
            try:
                pil = pic.get_image(doc) if hasattr(pic, "get_image") else None
                if pil is None and getattr(pic, "image", None) is not None:
                    pil = getattr(pic.image, "pil_image", None)
                if pil is not None:
                    p = img_dir / f"picture_{i + 1:03d}.png"
                    pil.save(p)
                    imgs.append(str(p))
            except Exception:
                continue

        md = doc.export_to_markdown()  # placeholder <!-- image --> anchors (counted)
        jpath = outdir / "output.json"
        try:
            jpath.write_text(doc.export_to_json(), encoding="utf-8")
            json_path = str(jpath)
        except Exception:
            json_path = None
        return ConverterOutput(
            converter=self.name,
            status="ok",
            markdown=md,
            json_path=json_path,
            image_paths=imgs,
            meta={"pictures": len(getattr(doc, "pictures", []) or [])},
        )
