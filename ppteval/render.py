"""Render slides to PNG via PowerPoint COM automation (Windows only).

Used by the review CLI so a human can rate each converter's text against the
*actual* slide. Best-effort by design: if PowerPoint/pywin32 is unavailable or a
slide fails, we skip gracefully rather than abort the eval.

Safety: we attach to PowerPoint (single-instance on Windows) and only Quit it if
no other presentations remain open, so we never close a deck the user had open.
"""
from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class RenderResult:
    ok: bool
    png_paths: list[str] = field(default_factory=list)
    error: str | None = None
    skipped: bool = False


def available() -> tuple[bool, str]:
    if sys.platform != "win32":
        return False, "not Windows (PowerPoint COM unavailable)"
    try:
        import win32com.client  # noqa: F401
    except Exception as e:  # pragma: no cover
        return False, f"pywin32 missing: {e}"
    return True, "pywin32 + PowerPoint COM"


def render_deck(pptx_path: str | Path, out_dir: str | Path, width: int = 1280) -> RenderResult:
    ok, info = available()
    if not ok:
        return RenderResult(ok=False, skipped=True, error=info)

    import pythoncom
    import win32com.client

    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    powerpoint = None
    pres = None
    we_opened_app = False
    pre_existing = 0
    png_paths: list[str] = []
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        try:
            pre_existing = powerpoint.Presentations.Count
        except Exception:
            pre_existing = 0
        # PowerPoint refuses to run fully invisible on many builds; make visible.
        try:
            powerpoint.Visible = 1
        except Exception:
            pass
        we_opened_app = True

        try:
            pres = powerpoint.Presentations.Open(
                str(pptx_path), ReadOnly=True, WithWindow=False
            )
        except Exception:
            # Some installs require a window.
            pres = powerpoint.Presentations.Open(str(pptx_path), ReadOnly=True)

        # Compute target height from the slide aspect ratio.
        try:
            sw = float(pres.PageSetup.SlideWidth)
            sh = float(pres.PageSetup.SlideHeight)
            height = max(1, int(round(width * sh / sw)))
        except Exception:
            height = int(width * 9 / 16)

        n = pres.Slides.Count
        for i in range(1, n + 1):
            target = out_dir / f"slide_{i:03d}.png"
            try:
                pres.Slides(i).Export(str(target), "PNG", width, height)
                png_paths.append(str(target))
            except Exception:
                # Skip the individual slide; keep going.
                continue
        return RenderResult(ok=True, png_paths=png_paths)
    except Exception as e:
        return RenderResult(ok=False, error=f"{type(e).__name__}: {e}", png_paths=png_paths)
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        # Only Quit if we created the app AND no presentation the user had open remains.
        try:
            if powerpoint is not None and we_opened_app and pre_existing == 0:
                if powerpoint.Presentations.Count == 0:
                    powerpoint.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def pptx_to_pdf(pptx_path: str | Path, pdf_path: str | Path) -> tuple[bool, str | None]:
    """Export a deck to PDF via PowerPoint COM (for PDF-based adapters).

    Returns (ok, error). Best-effort with the same single-instance safety as
    ``render_deck`` (never Quits an app the user already had running).
    """
    ok, info = available()
    if not ok:
        return False, info

    import pythoncom
    import win32com.client

    pptx_path = Path(pptx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    powerpoint = None
    pres = None
    pre_existing = 0
    PP_SAVE_AS_PDF = 32
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        try:
            pre_existing = powerpoint.Presentations.Count
        except Exception:
            pre_existing = 0
        try:
            powerpoint.Visible = 1
        except Exception:
            pass
        try:
            pres = powerpoint.Presentations.Open(str(pptx_path), ReadOnly=True, WithWindow=False)
        except Exception:
            pres = powerpoint.Presentations.Open(str(pptx_path), ReadOnly=True)
        pres.SaveAs(str(pdf_path), PP_SAVE_AS_PDF)
        return (pdf_path.exists(), None if pdf_path.exists() else "PDF not written")
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if powerpoint is not None and pre_existing == 0 and powerpoint.Presentations.Count == 0:
                powerpoint.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
