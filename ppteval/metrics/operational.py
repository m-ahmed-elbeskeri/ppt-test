"""Operational metrics: speed and output size/cost (user-requested)."""
from __future__ import annotations

from ..schema import ConverterOutput, GroundTruth
from ..tokens import count_tokens


def operational_metrics(out: ConverterOutput, gt: GroundTruth) -> dict:
    text = out.markdown or ""
    out_tokens = count_tokens(text)
    n = max(1, gt.n_slides)
    sps = (gt.n_slides / out.elapsed_s) if out.elapsed_s > 0 else None
    comp = (out_tokens / gt.gt_tokens) if gt.gt_tokens > 0 else None
    return {
        "parse_time_s": round(out.elapsed_s, 3),
        "slides_per_sec": round(sps, 3) if sps is not None else None,
        "out_chars": len(text),
        "out_tokens": out_tokens,
        "tokens_per_slide": round(out_tokens / n, 1),
        "compression": round(comp, 3) if comp is not None else None,
        "status_ok": 1.0 if out.status == "ok" else 0.0,
    }
