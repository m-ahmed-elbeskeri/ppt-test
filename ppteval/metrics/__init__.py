"""Deterministic (non-LLM) metrics: operational, coverage-vs-ground-truth, structure.

``compute_metrics`` merges all three families into one flat dict per converter.
None values mean "not applicable to this deck" and are skipped by the scorer.
"""
from __future__ import annotations

from ..schema import ConverterOutput, GroundTruth
from .coverage import coverage_metrics
from .operational import operational_metrics
from .structure import structure_metrics

__all__ = ["compute_metrics", "operational_metrics", "coverage_metrics", "structure_metrics"]


def compute_metrics(out: ConverterOutput, gt: GroundTruth) -> dict:
    if out.status != "ok":
        # Failed/skipped run: only operational + status flag are meaningful.
        return {
            "status_ok": 0.0,
            "parse_time_s": round(out.elapsed_s, 3),
            "out_tokens": 0,
        }
    m: dict = {}
    m.update(operational_metrics(out, gt))
    m.update(coverage_metrics(out, gt))
    m.update(structure_metrics(out, gt))
    return m
