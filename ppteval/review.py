"""Interactive human rating CLI (no LLM judging — the user is the judge).

Modes:
  deck       rate each converter once over the whole deck, then optional ranking.
  sample     rate converters on N auto-selected representative slides.
  allslides  rate converters on every slide.

Every rating is written to ratings.json immediately (crash/resume safe); already
rated (converter, mode, slide, dimension) tuples are skipped. Set env
PPTEVAL_NO_OPEN=1 to suppress opening files/images (used by tests).
"""
from __future__ import annotations

import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from .config import RATING_DIMENSIONS, RATING_MAX, RATING_MIN, out_dir_for
from .runner import load_groundtruth, load_outputs
from .schema import GroundTruth, Rating, RatingsStore

_MARK = re.compile(r"slide\s*(?:number)?\s*[:#\-]*\s*(\d+)", re.I)


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _store_path(od: Path) -> Path:
    return od / "ratings.json"


def _load_store(od: Path, slug: str) -> RatingsStore:
    p = _store_path(od)
    if p.exists():
        return RatingsStore.model_validate_json(p.read_text(encoding="utf-8"))
    return RatingsStore(deck_slug=slug)


def _save_store(od: Path, store: RatingsStore) -> None:
    _store_path(od).write_text(store.model_dump_json(indent=2), encoding="utf-8")


def _open(path: str | Path) -> bool:
    if os.environ.get("PPTEVAL_NO_OPEN"):
        return False
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
            return True
        import subprocess

        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.Popen([opener, str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


def _ask_score(prompt: str):
    """Return float score, None to skip, or 'QUIT'."""
    while True:
        try:
            raw = input(prompt).strip().lower()
        except EOFError:
            return "QUIT"
        if raw in ("", "s"):
            return None
        if raw in ("q", "quit"):
            return "QUIT"
        if raw.isdigit() and RATING_MIN <= int(raw) <= RATING_MAX:
            return float(int(raw))
        print(f"    -> enter {RATING_MIN}-{RATING_MAX}, 's' to skip, 'q' to quit")


def _clip(text: str, n: int = 1500) -> str:
    text = text or ""
    return text if len(text) <= n else text[:n] + f"\n…[+{len(text)-n} chars]"


def _metric_line(o, metrics_by_conv: dict) -> str:
    m = metrics_by_conv.get(o.converter, {})
    return (
        f"  parse {m.get('parse_time_s','?')}s · {m.get('out_tokens','?')} tokens · "
        f"text {_pct(m.get('text_recall'))} · struct {_pct(m.get('structure_score'))}"
    )


def _pct(v):
    return f"{v*100:.0f}%" if isinstance(v, float) else "—"


def _representative_slides(gt: GroundTruth, n: int) -> list[int]:
    scored = []
    for s in gt.slides:
        richness = (
            2.0 * len(s.tables)
            + 2.0 * s.n_charts
            + (1.0 if s.notes.strip() else 0.0)
            + 0.6 * len(s.image_rels)
            + 0.002 * len(" ".join(s.text_runs))
        )
        scored.append((richness, s.slide_number))
    scored.sort(reverse=True)
    return sorted({num for _, num in scored[:n]})


def slice_slide(text: str, n: int) -> str | None:
    matches = list(_MARK.finditer(text or ""))
    if not matches:
        return None
    for i, mt in enumerate(matches):
        num = int(mt.group(1))
        if num == n:
            start = mt.start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            return text[start:end].strip()
    return None


def _metrics_map(od: Path) -> dict:
    import json

    p = od / "metrics.json"
    if not p.exists():
        return {}
    rows = json.loads(p.read_text(encoding="utf-8"))
    return {r["converter"]: r.get("metrics", {}) for r in rows}


# --- deck mode ---------------------------------------------------------------
def _review_deck(od, gt, outputs, store, metrics_by_conv) -> None:
    print("\n=== DECK review: rate each converter 1-5 (Enter/s=skip, q=quit) ===")
    for o in outputs:
        print(f"\n######## {o.converter}  ########")
        print(_metric_line(o, metrics_by_conv))
        if o.markdown_path:
            try:
                ans = input("  open full output.md? [y/N] ").strip().lower()
            except EOFError:
                ans = "n"
            if ans == "y":
                _open(o.markdown_path)
            else:
                print("  --- preview ---")
                print(_clip(o.markdown, 1200))
                print("  --- end preview ---")
        for dim in RATING_DIMENSIONS:
            if store.has(o.converter, "deck", dim):
                print(f"  {dim}: already rated, skipping")
                continue
            sc = _ask_score(f"  rate {dim} (1-{RATING_MAX}): ")
            if sc == "QUIT":
                return
            if sc is None:
                continue
            store.upsert(Rating(converter=o.converter, mode="deck", dimension=dim, score=sc, ts=_now()))
            _save_store(od, store)

    # optional head-to-head ranking
    try:
        ans = input("\nRank converters head-to-head now? [y/N] ").strip().lower()
    except EOFError:
        ans = "n"
    if ans == "y":
        names = [o.converter for o in outputs]
        for i, nm in enumerate(names, 1):
            print(f"  {i}. {nm}")
        try:
            order = input("  enter numbers best->worst (comma-sep): ").strip()
        except EOFError:
            order = ""
        picks = [p.strip() for p in order.split(",") if p.strip()]
        for pos, p in enumerate(picks, 1):
            nm = names[int(p) - 1] if p.isdigit() and 1 <= int(p) <= len(names) else None
            if nm:
                store.upsert(Rating(converter=nm, mode="deck", dimension="rank", score=float(pos), ts=_now()))
        _save_store(od, store)


# --- slide modes -------------------------------------------------------------
def _review_slides(od, gt, outputs, store, mode, slide_numbers, renders_dir) -> None:
    print(f"\n=== {mode.upper()} review: {len(slide_numbers)} slides × {len(outputs)} converters ===")
    for sn in slide_numbers:
        png = renders_dir / f"slide_{sn:03d}.png"
        print(f"\n========== Slide {sn} ==========")
        if png.exists():
            opened = _open(png)
            print(f"(slide image: {png.name}{' — opened' if opened else ''})")
        else:
            print("(no rendered image; rate from text + ground truth)")
        gts = gt.slides[sn - 1] if 1 <= sn <= len(gt.slides) else None
        if gts:
            print(f"GROUND TRUTH — title: {gts.title or '(none)'}")
            print(_clip("\n".join(gts.text_runs), 600))
        for o in outputs:
            if store.has(o.converter, mode, "overall", sn):
                print(f"  {o.converter}: already rated, skipping")
                continue
            seg = slice_slide(o.markdown, sn)
            print(f"\n--- {o.converter} (slide {sn}) ---")
            print(_clip(seg, 800) if seg else "(no per-slide markers — rate from image + deck output)")
            sc = _ask_score(f"  rate {o.converter} slide {sn} overall (1-{RATING_MAX}): ")
            if sc == "QUIT":
                return
            if sc is None:
                continue
            store.upsert(
                Rating(converter=o.converter, mode=mode, slide_number=sn, dimension="overall", score=sc, ts=_now())
            )
            _save_store(od, store)


# --- entry -------------------------------------------------------------------
def run_review(deck, mode: str = "deck", converter_names=None, n: int = 10) -> Path:
    od = out_dir_for(deck)
    if not (od / "ground_truth.json").exists():
        raise SystemExit(f"no ground truth at {od}. Run `ppteval convert` (or `run`) first.")
    gt = load_groundtruth(od)
    outputs = [o for o in load_outputs(od) if o.status == "ok" and o.markdown]
    if converter_names:
        sel = {c.strip() for c in converter_names}
        outputs = [o for o in outputs if o.converter in sel]
    if not outputs:
        raise SystemExit("no converter outputs to review. Run `ppteval convert` first.")

    store = _load_store(od, gt.deck_slug)
    metrics_by_conv = _metrics_map(od)
    renders_dir = od / "renders"

    if mode == "deck":
        _review_deck(od, gt, outputs, store, metrics_by_conv)
    elif mode == "sample":
        _review_slides(od, gt, outputs, store, mode, _representative_slides(gt, n), renders_dir)
    elif mode == "allslides":
        _review_slides(od, gt, outputs, store, mode, [s.slide_number for s in gt.slides], renders_dir)
    else:
        raise SystemExit(f"unknown review mode '{mode}' (deck|sample|allslides)")

    _save_store(od, store)
    # refresh leaderboard so ratings show up immediately
    from .report import build_report

    build_report(od)
    print(f"\nRatings saved -> {_store_path(od)}\nLeaderboard refreshed -> {od/'leaderboard.md'}")
    return _store_path(od)
