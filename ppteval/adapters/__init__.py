"""Converter adapters. Each wraps one OSS PPTX->text/markdown tool behind a common
interface so the runner can execute and rank any selected subset."""
from __future__ import annotations

from .base import Adapter
from .docling_adapter import DoclingAdapter
from .marker_adapter import MarkerAdapter
from .markitdown_adapter import MarkitdownAdapter
from .pptx2md_adapter import Pptx2mdAdapter
from .pptx_custom_adapter import PptxCustomAdapter
from .pymupdf4llm_adapter import Pymupdf4llmAdapter
from .tika_adapter import TikaAdapter
from .unstructured_adapter import UnstructuredAdapter

# Registration order = default run/report order (light set first, heavy last).
_REGISTRY: dict[str, Adapter] = {}
for _cls in (
    MarkitdownAdapter,
    PptxCustomAdapter,
    Pptx2mdAdapter,
    Pymupdf4llmAdapter,
    TikaAdapter,
    DoclingAdapter,
    UnstructuredAdapter,
    MarkerAdapter,
):
    _REGISTRY[_cls.name] = _cls()


def all_names() -> list[str]:
    return list(_REGISTRY.keys())


def get_adapter(name: str) -> Adapter:
    if name not in _REGISTRY:
        raise KeyError(f"unknown converter '{name}'. Known: {', '.join(_REGISTRY)}")
    return _REGISTRY[name]


def select(names: list[str] | None = None, include_heavy: bool = True) -> list[Adapter]:
    """Resolve a list of converter names to adapter instances.

    names=None  -> every registered adapter (optionally excluding heavy ones).
    """
    if names:
        return [get_adapter(n.strip()) for n in names if n.strip()]
    return [a for a in _REGISTRY.values() if include_heavy or not a.heavy]


def availability() -> list[dict]:
    """Status table for `ppteval list-converters`."""
    rows = []
    for a in _REGISTRY.values():
        ok, info = a.available()
        rows.append(
            {
                "name": a.name,
                "license": a.license,
                "heavy": a.heavy,
                "available": ok,
                "info": info,
                "note": a.note,
            }
        )
    return rows
