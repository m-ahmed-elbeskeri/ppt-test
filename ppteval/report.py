"""Leaderboard generation: merge automated metrics + human ratings -> ranking.

Outputs (under out/<deck_slug>/):
  - results.json    full structured result (raw + normalized + scores)
  - scorecard.csv   flat table for spreadsheets
  - leaderboard.md  human-readable ranked report with caveats

Scoring policy lives in config.py (AUTO_WEIGHTS, HUMAN_BLEND, COMPRESSION_*).
"""
from __future__ import annotations

import json
from pathlib import Path

from .config import (
    AUTO_WEIGHTS,
    COMPRESSION_BAND,
    COMPRESSION_FAR,
    HUMAN_BLEND,
    RATING_DIMENSIONS,
)
from .schema import RatingsStore


# --- scoring helpers ---------------------------------------------------------
def compression_fit(c: float | None) -> float | None:
    """1.0 inside the ideal density band; decays to 0 at the far bounds."""
    if c is None:
        return None
    lo, hi = COMPRESSION_BAND
    flo, fhi = COMPRESSION_FAR
    if lo <= c <= hi:
        return 1.0
    if c < lo:
        return max(0.0, (c - flo) / (lo - flo)) if lo > flo else 0.0
    return max(0.0, (fhi - c) / (fhi - hi)) if fhi > hi else 0.0


def _minmax(values: dict[str, float]) -> dict[str, float]:
    vals = [v for v in values.values() if v is not None]
    if not vals:
        return {k: None for k in values}
    mn, mx = min(vals), max(vals)
    if mx <= mn:
        return {k: (1.0 if v is not None else None) for k, v in values.items()}
    return {k: ((v - mn) / (mx - mn) if v is not None else None) for k, v in values.items()}


def _auto_score(norm: dict[str, float | None]) -> float:
    num = 0.0
    wsum = 0.0
    for key, w in AUTO_WEIGHTS.items():
        v = norm.get(key)
        if v is None:
            continue
        num += w * v
        wsum += w
    return round(num / wsum, 4) if wsum > 0 else 0.0


# --- human ratings -----------------------------------------------------------
def _human_scores(ratings_path: Path, converters: list[str]) -> dict[str, dict]:
    out = {c: {"avg": None, "n": 0, "human_norm": None} for c in converters}
    if not ratings_path.exists():
        return out
    store = RatingsStore.model_validate_json(ratings_path.read_text(encoding="utf-8"))

    # per-converter dimension ratings (1-5), excluding the ordinal "rank"
    by_conv: dict[str, list[float]] = {c: [] for c in converters}
    rank_by_conv: dict[str, list[float]] = {c: [] for c in converters}
    for r in store.ratings:
        if r.converter not in by_conv:
            continue
        if r.dimension == "rank":
            rank_by_conv[r.converter].append(r.score)
        elif r.dimension in RATING_DIMENSIONS:
            by_conv[r.converter].append(r.score)

    n_ranked = sum(1 for c in converters if rank_by_conv[c])
    for c in converters:
        dims = by_conv[c]
        parts = []
        avg = None
        if dims:
            avg = sum(dims) / len(dims)
            parts.append((avg - 1) / 4.0)  # normalise 1..5 -> 0..1
        if rank_by_conv[c] and n_ranked > 1:
            pos = sum(rank_by_conv[c]) / len(rank_by_conv[c])
            parts.append(max(0.0, (n_ranked - pos) / (n_ranked - 1)))
        out[c] = {
            "avg": round(avg, 2) if avg is not None else None,
            "n": len(dims),
            "human_norm": round(sum(parts) / len(parts), 4) if parts else None,
        }
    return out


# --- main entry --------------------------------------------------------------
def build_report(od: Path) -> Path:
    metrics_rows = json.loads((od / "metrics.json").read_text(encoding="utf-8"))
    gt = json.loads((od / "ground_truth.json").read_text(encoding="utf-8"))
    converters = [r["converter"] for r in metrics_rows]
    human = _human_scores(od / "ratings.json", converters)

    # cross-converter normalisation for speed (ok rows only)
    sps_raw = {
        r["converter"]: (r["metrics"].get("slides_per_sec") if r["status"] == "ok" else None)
        for r in metrics_rows
    }
    sps_norm = _minmax(sps_raw)

    results = []
    for r in metrics_rows:
        m = r["metrics"]
        ok = r["status"] == "ok"
        norm = {
            "text_recall": m.get("text_recall") if ok else None,
            "notes_recall": m.get("notes_recall") if ok else None,
            "table_cell_recall": m.get("table_cell_recall") if ok else None,
            "image_recall": m.get("image_recall") if ok else None,
            "structure_score": m.get("structure_score") if ok else None,
            "compression_fit": compression_fit(m.get("compression")) if ok else None,
            "slides_per_sec": sps_norm.get(r["converter"]) if ok else None,
            "status_ok": m.get("status_ok", 0.0),
        }
        auto = _auto_score(norm) if ok else 0.0
        h = human.get(r["converter"], {})
        hn = h.get("human_norm")
        if hn is not None:
            final = round(HUMAN_BLEND * hn + (1 - HUMAN_BLEND) * auto, 4)
            basis = "human+auto"
        else:
            final = auto
            basis = "auto only"
        results.append(
            {
                "converter": r["converter"],
                "status": r["status"],
                "error": r.get("error"),
                "final_score": final,
                "auto_score": auto,
                "human_avg": h.get("avg"),
                "human_n": h.get("n", 0),
                "basis": basis,
                "metrics": m,
                "normalized": norm,
            }
        )

    results.sort(key=lambda x: (x["status"] == "ok", x["final_score"]), reverse=True)
    for i, row in enumerate(results, 1):
        row["rank"] = i

    payload = {"deck": gt.get("deck_slug"), "ground_truth_totals": _gt_totals(gt), "results": results}
    (od / "results.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    _write_csv(od, results)
    _write_markdown(od, gt, results)
    return od / "leaderboard.md"


def _gt_totals(gt: dict) -> dict:
    return {
        k: gt.get(k)
        for k in ("n_slides", "n_notes", "n_tables", "n_images", "n_charts", "gt_tokens")
    }


def _write_csv(od: Path, results: list[dict]) -> bool:
    """Write scorecard.csv. Non-fatal: a locked file (e.g. open in Excel on
    Windows) must not abort the rest of the report. Returns True on success."""
    import pandas as pd

    flat = []
    for r in results:
        row = {
            "rank": r["rank"],
            "converter": r["converter"],
            "status": r["status"],
            "final_score": r["final_score"],
            "auto_score": r["auto_score"],
            "human_avg": r["human_avg"],
            "human_n": r["human_n"],
        }
        row.update(r["metrics"])
        flat.append(row)
    try:
        pd.DataFrame(flat).to_csv(od / "scorecard.csv", index=False)
        return True
    except (PermissionError, OSError):
        # leave the stale CSV; leaderboard.md + results.json are the source of truth
        return False


def _fmt(v, pct=False):
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v*100:.0f}%" if pct else f"{v:.3f}".rstrip("0").rstrip(".")
    return str(v)


def _write_markdown(od: Path, gt: dict, results: list[dict]) -> None:
    t = _gt_totals(gt)
    lines: list[str] = []
    lines.append(f"# PPTX -> LLM converter leaderboard — `{gt.get('deck_slug')}`\n")
    lines.append(
        f"Deck ground truth (python-pptx): **{t['n_slides']} slides**, "
        f"{t['n_tables']} tables, {t['n_images']} images, {t['n_charts']} charts, "
        f"{t['n_notes']} slides with notes, ~{t['gt_tokens']} content tokens.\n"
    )

    # ranking table
    lines.append("## Ranking\n")
    head = [
        "Rank", "Converter", "Final", "Auto", "Human", "Text", "Notes", "Tables",
        "Images", "Struct", "Parse s", "Tokens", "Tok/slide", "Status",
    ]
    lines.append("| " + " | ".join(head) + " |")
    lines.append("| " + " | ".join("---" for _ in head) + " |")
    for r in results:
        m = r["metrics"]
        human = f"{r['human_avg']}/5" if r["human_avg"] is not None else "—"
        row = [
            str(r["rank"]),
            f"`{r['converter']}`",
            _fmt(r["final_score"]),
            _fmt(r["auto_score"]),
            human,
            _fmt(m.get("text_recall"), pct=True),
            _fmt(m.get("notes_recall"), pct=True),
            _fmt(m.get("table_cell_recall"), pct=True),
            _fmt(m.get("image_recall"), pct=True),
            _fmt(m.get("structure_score"), pct=True),
            _fmt(m.get("parse_time_s")),
            _fmt(m.get("out_tokens")),
            _fmt(m.get("tokens_per_slide")),
            r["status"],
        ]
        lines.append("| " + " | ".join(row) + " |")

    # per-converter notes
    lines.append("\n## Per-converter notes\n")
    for r in results:
        lines.append(f"### {r['rank']}. `{r['converter']}` — {r['basis']}")
        if r["status"] != "ok":
            lines.append(f"- **{r['status'].upper()}**: {r.get('error')}")
            lines.append("")
            continue
        lines.extend(f"- {s}" for s in _highlights(r))
        lines.append("")

    # caveats
    lines.append("## How to read this / caveats\n")
    lines.append(
        "- **Recall is relative to python-pptx-extractable content**, not absolute "
        "truth. python-pptx cannot see text inside SmartArt without a text frame, "
        "text rasterised into images (PNG/EMF/WDP), or some embedded objects. Treat "
        "recall as a consistent cross-converter yardstick; your **human ratings** + "
        "the rendered slides are the corrective for what ground truth misses."
    )
    lines.append(
        "- `pptx_custom` is the **reference extractor** that defines ground truth, so "
        "its text/table/image recall is ~100% by construction — it competes on "
        "structure, tokens, speed and your ratings, not recall."
    )
    lines.append(
        "- **Tokens** are counted with tiktoken `o200k_base` — a consistent proxy for "
        "LLM cost, not the exact tokenizer of any one model (e.g. Claude)."
    )
    lines.append(
        f"- **Final score** = {int(HUMAN_BLEND*100)}% human + {int((1-HUMAN_BLEND)*100)}%"
        " auto where you rated a converter; auto-only otherwise (flagged above). "
        "Weights are tunable in `ppteval/config.py`."
    )
    (od / "leaderboard.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def _highlights(r: dict) -> list[str]:
    """Plain-language strengths/weaknesses from metric outliers."""
    m = r["metrics"]
    out: list[str] = []
    tr = m.get("text_recall")
    if tr is not None:
        out.append(f"Text recall {tr*100:.0f}% · structure {(_z(m.get('structure_score')))*100:.0f}%")
    out.append(
        f"{m.get('out_tokens','—')} tokens ({m.get('tokens_per_slide','—')}/slide), "
        f"parsed in {m.get('parse_time_s','—')}s "
        f"({_z(m.get('slides_per_sec')):.1f} slides/s)"
    )
    ier = m.get("image_recall")
    if ier is not None:
        out.append(
            f"Images: {ier*100:.0f}% referenced, "
            f"{m.get('images_extracted', 0)} extracted to disk"
        )
    nr = m.get("notes_recall")
    if nr is not None and nr < 0.5:
        out.append("⚠ dropped most speaker notes")
    if (
        m.get("n_md_tables", 0) == 0
        and m.get("n_html_tables", 0) == 0
        and m.get("table_cell_recall") not in (None, 0)
    ):
        out.append("⚠ table data present but not formatted as Markdown/HTML tables")
    comp = m.get("compression")
    if comp is not None and comp > 2.0:
        out.append(f"⚠ token bloat (×{comp:.1f} vs content)")
    if comp is not None and comp < 0.7:
        out.append(f"⚠ likely content loss (only ×{comp:.1f} of content tokens)")
    if m.get("slide_boundary_recall") not in (None,) and m.get("slide_boundary_recall", 0) < 0.5:
        out.append("⚠ weak slide boundaries (harder to chunk per-slide)")
    return out


def _z(v):
    return v if isinstance(v, (int, float)) else 0.0
