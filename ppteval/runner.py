"""Orchestration: ground-truth -> render -> convert(selected) -> metrics.

Each stage persists its artifacts under ``out/<deck_slug>/`` so any stage can be
run on its own from the CLI (e.g. re-run metrics without re-converting).
"""
from __future__ import annotations

import json
from pathlib import Path

from . import adapters as adapters_mod
from .config import out_dir_for
from .groundtruth import extract_ground_truth
from .metrics import compute_metrics
from .render import RenderResult, render_deck
from .schema import ConverterOutput, GroundTruth


def _gt_path(od: Path) -> Path:
    return od / "ground_truth.json"


# --- ground truth ------------------------------------------------------------
def run_groundtruth(deck: str | Path, od: Path | None = None) -> GroundTruth:
    od = od or out_dir_for(deck)
    od.mkdir(parents=True, exist_ok=True)
    gt = extract_ground_truth(deck)
    _gt_path(od).write_text(gt.model_dump_json(indent=2), encoding="utf-8")
    return gt


def load_groundtruth(od: Path) -> GroundTruth:
    return GroundTruth.model_validate_json(_gt_path(od).read_text(encoding="utf-8"))


# --- render ------------------------------------------------------------------
def run_render(deck: str | Path, od: Path | None = None, width: int = 1280) -> RenderResult:
    od = od or out_dir_for(deck)
    return render_deck(deck, od / "renders", width=width)


# --- convert -----------------------------------------------------------------
def run_convert(deck: str | Path, adapter_list, od: Path | None = None) -> list[ConverterOutput]:
    od = od or out_dir_for(deck)
    outputs: list[ConverterOutput] = []
    for a in adapter_list:
        cdir = od / "converters" / a.name
        outputs.append(a.convert(deck, cdir))

    # MERGE (upsert by name) so converting subsets at different times accumulates
    # into one leaderboard, ordered by registry order for stable display.
    path = od / "converters.json"
    merged: dict[str, dict] = {}
    if path.exists():
        for d in json.loads(path.read_text(encoding="utf-8")):
            merged[d.get("converter")] = d
    for o in outputs:
        merged[o.converter] = o.model_dump()

    order = {name: i for i, name in enumerate(adapters_mod.all_names())}
    ordered = sorted(merged.values(), key=lambda d: order.get(d.get("converter"), 999))
    path.write_text(json.dumps(ordered, indent=2), encoding="utf-8")
    return outputs


def load_outputs(od: Path) -> list[ConverterOutput]:
    data = json.loads((od / "converters.json").read_text(encoding="utf-8"))
    outs: list[ConverterOutput] = []
    for d in data:
        o = ConverterOutput.model_validate(d)
        if o.markdown_path and Path(o.markdown_path).exists():
            o.markdown = Path(o.markdown_path).read_text(encoding="utf-8")
        outs.append(o)
    return outs


# --- metrics -----------------------------------------------------------------
def run_metrics(gt: GroundTruth, outputs: list[ConverterOutput], od: Path) -> list[dict]:
    rows: list[dict] = []
    for o in outputs:
        rows.append(
            {
                "converter": o.converter,
                "status": o.status,
                "elapsed_s": round(o.elapsed_s, 3),
                "error": o.error,
                "metrics": compute_metrics(o, gt),
            }
        )
    (od / "metrics.json").write_text(json.dumps(rows, indent=2), encoding="utf-8")
    return rows


# --- end to end --------------------------------------------------------------
def run_all(
    deck: str | Path,
    adapter_list,
    do_render: bool = True,
    width: int = 1280,
) -> tuple[Path, GroundTruth, list[ConverterOutput], RenderResult | None]:
    od = out_dir_for(deck)
    od.mkdir(parents=True, exist_ok=True)
    gt = run_groundtruth(deck, od)
    render = run_render(deck, od, width=width) if do_render else None
    outputs = run_convert(deck, adapter_list, od)
    run_metrics(gt, outputs, od)
    return od, gt, outputs, render
