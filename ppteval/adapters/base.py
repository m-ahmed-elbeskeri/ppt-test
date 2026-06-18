"""Adapter base class + small shared helpers.

Each converter is wrapped so the runner can treat them uniformly:
- ``available()`` import-guards optional libs (missing => SKIPPED, never fatal),
- ``convert()`` times the run with ``perf_counter`` (adapters cannot under-report),
  catches exceptions into a structured ``error`` result, and persists output.md.
"""
from __future__ import annotations

import time
import traceback
import zipfile
from abc import ABC, abstractmethod
from pathlib import Path

from ..schema import ConverterOutput


class Adapter(ABC):
    name: str = "adapter"
    license: str = ""
    heavy: bool = False  # True => part of the opt-in heavy set (large deps)
    note: str = ""

    @classmethod
    def available(cls) -> tuple[bool, str]:
        """(installed_and_usable, human-readable reason/version)."""
        return True, ""

    @abstractmethod
    def _convert(self, pptx: Path, outdir: Path) -> ConverterOutput:
        ...

    def convert(self, pptx: str | Path, outdir: str | Path) -> ConverterOutput:
        pptx = Path(pptx)
        outdir = Path(outdir)
        outdir.mkdir(parents=True, exist_ok=True)

        ok, info = self.available()
        if not ok:
            return ConverterOutput(converter=self.name, status="skipped", error=info)

        t0 = time.perf_counter()
        try:
            out = self._convert(pptx, outdir)
        except Exception as e:
            return ConverterOutput(
                converter=self.name,
                status="error",
                elapsed_s=time.perf_counter() - t0,
                error=f"{type(e).__name__}: {e}",
                meta={"traceback": traceback.format_exc()[-2000:]},
            )

        out.elapsed_s = time.perf_counter() - t0
        out.converter = self.name
        if out.status == "ok" and out.markdown:
            mdp = outdir / "output.md"
            mdp.write_text(out.markdown, encoding="utf-8")
            out.markdown_path = str(mdp)
        return out


# --- shared helpers ----------------------------------------------------------
def extract_media(pptx: Path, out_images: Path) -> list[str]:
    """Extract ppt/media/* from the pptx zip into ``out_images``; return paths."""
    out_images.mkdir(parents=True, exist_ok=True)
    written: list[str] = []
    with zipfile.ZipFile(pptx) as z:
        for n in z.namelist():
            if n.startswith("ppt/media/"):
                fname = n.split("/")[-1]
                dest = out_images / fname
                try:
                    dest.write_bytes(z.read(n))
                    written.append(str(dest))
                except Exception:
                    continue
    return written
