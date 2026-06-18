"""ppteval command-line interface.

Subcommands:
  list-converters             show every adapter, its license and availability
  ground-truth DECK           extract + persist ground_truth.json
  render DECK                 render slides to PNG via PowerPoint COM
  convert DECK [--converters] run selected converters, save outputs
  metrics DECK                (re)compute automated metrics over saved outputs
  review DECK [--mode ...]     interactive human rating
  report DECK                 (re)build leaderboard.md / scorecard.csv / results.json
  run DECK [...]              end-to-end: gt -> render -> convert -> metrics -> report
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import adapters as adapters_mod
from . import runner
from .config import DECKS_DIR, out_dir_for
from .report import build_report


def _utf8_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except Exception:
            pass


def _resolve_deck(s: str) -> Path:
    p = Path(s)
    if p.exists():
        return p
    cand = DECKS_DIR / s
    if cand.exists():
        return cand
    matches = sorted(DECKS_DIR.glob(f"*{s}*")) if DECKS_DIR.exists() else []
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise SystemExit(f"ambiguous deck '{s}': {[m.name for m in matches]}")
    raise SystemExit(f"deck not found: {s} (looked in . and {DECKS_DIR})")


def _select(args) -> list:
    if getattr(args, "converters", None):
        names = [x for x in args.converters.split(",") if x.strip()]
        return adapters_mod.select(names)
    if getattr(args, "all", False):
        return adapters_mod.select(None, include_heavy=True)
    return adapters_mod.select(None, include_heavy=False)


# --- command handlers --------------------------------------------------------
def cmd_list(args) -> None:
    rows = adapters_mod.availability()
    width = max(len(r["name"]) for r in rows)
    print(f"{'converter'.ljust(width)}  avail  set    license      detail")
    print("-" * (width + 45))
    for r in rows:
        flag = " ok  " if r["available"] else "  -  "
        kind = "heavy" if r["heavy"] else "light"
        print(f"{r['name'].ljust(width)}  {flag}  {kind}  {r['license']:<11}  {r['info']}")
    print("\nlight = installed by default · heavy = pip install -r requirements-heavy.txt")


def cmd_groundtruth(args) -> None:
    deck = _resolve_deck(args.deck)
    gt = runner.run_groundtruth(deck)
    print(
        f"{gt.deck_slug}: {gt.n_slides} slides, {gt.n_tables} tables, {gt.n_images} images, "
        f"{gt.n_charts} charts, {gt.n_notes} with notes, ~{gt.gt_tokens} tokens"
    )
    print(f"-> {out_dir_for(deck) / 'ground_truth.json'}")


def cmd_render(args) -> None:
    deck = _resolve_deck(args.deck)
    print("rendering slides via PowerPoint COM (this may briefly open PowerPoint)…")
    res = runner.run_render(deck, width=args.width)
    if res.skipped:
        print(f"SKIPPED: {res.error}")
    elif res.ok:
        print(f"rendered {len(res.png_paths)} slides -> {out_dir_for(deck) / 'renders'}")
    else:
        print(f"render FAILED: {res.error} (rendered {len(res.png_paths)} before failing)")


def cmd_convert(args) -> None:
    deck = _resolve_deck(args.deck)
    adapters = _select(args)
    print(f"converting with: {', '.join(a.name for a in adapters)}")
    outputs = runner.run_convert(deck, adapters)
    for o in outputs:
        extra = f" ({o.error})" if o.error else ""
        print(f"  {o.converter:<13} {o.status:<8} {o.elapsed_s:6.2f}s{extra}")
    print(f"-> {out_dir_for(deck) / 'converters'}")


def cmd_metrics(args) -> None:
    deck = _resolve_deck(args.deck)
    od = out_dir_for(deck)
    gt = runner.load_groundtruth(od)
    outputs = runner.load_outputs(od)
    runner.run_metrics(gt, outputs, od)
    print(f"metrics -> {od / 'metrics.json'}")


def cmd_review(args) -> None:
    from .review import run_review

    deck = _resolve_deck(args.deck)
    names = [x for x in args.converters.split(",") if x.strip()] if args.converters else None
    run_review(deck, mode=args.mode, converter_names=names, n=args.n)


def cmd_report(args) -> None:
    deck = _resolve_deck(args.deck)
    path = build_report(out_dir_for(deck))
    print(f"leaderboard -> {path}")
    _print_leaderboard_preview(out_dir_for(deck))


def cmd_tui(args) -> None:
    from .tui import run as run_tui

    run_tui()


def cmd_run(args) -> None:
    deck = _resolve_deck(args.deck)
    adapters = _select(args)
    print(f"deck: {deck.name}")
    print(f"converters: {', '.join(a.name for a in adapters)}")
    od, gt, outputs, render = runner.run_all(deck, adapters, do_render=not args.no_render, width=args.width)
    print(f"ground truth: {gt.n_slides} slides, ~{gt.gt_tokens} tokens")
    if render is not None:
        if render.ok:
            print(f"rendered {len(render.png_paths)} slides")
        else:
            print(f"render skipped/failed: {render.error}")
    for o in outputs:
        extra = f" ({o.error})" if o.error else ""
        print(f"  {o.converter:<13} {o.status:<8} {o.elapsed_s:6.2f}s{extra}")

    if args.review:
        from .review import run_review

        run_review(deck, mode=args.review, n=args.n)
    else:
        build_report(od)
    _print_leaderboard_preview(od)
    print(f"\nartifacts in: {od}")


def _print_leaderboard_preview(od: Path) -> None:
    import json

    p = od / "results.json"
    if not p.exists():
        return
    data = json.loads(p.read_text(encoding="utf-8"))
    print("\nLEADERBOARD (final score):")
    for r in data["results"]:
        human = f"{r['human_avg']}/5" if r.get("human_avg") is not None else "auto"
        print(f"  {r['rank']}. {r['converter']:<13} {r['final_score']:.3f}  [{human}] {r['status']}")
    print(f"\nfull report: {od / 'leaderboard.md'}")


# --- parser ------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="ppteval", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    def add_deck(sp):
        sp.add_argument("deck", help="path to .pptx, or a name/substring under tests/decks/")

    def add_select(sp):
        sp.add_argument("--converters", help="comma-separated subset (default: light set)")
        sp.add_argument("--all", action="store_true", help="include heavy converters (docling/unstructured/marker)")

    sp = sub.add_parser("tui", help="launch the interactive TUI")
    sp.set_defaults(func=cmd_tui)

    sp = sub.add_parser("list-converters", help="show adapters + availability")
    sp.set_defaults(func=cmd_list)

    sp = sub.add_parser("ground-truth", help="extract ground truth")
    add_deck(sp)
    sp.set_defaults(func=cmd_groundtruth)

    sp = sub.add_parser("render", help="render slides to PNG (PowerPoint COM)")
    add_deck(sp)
    sp.add_argument("--width", type=int, default=1280)
    sp.set_defaults(func=cmd_render)

    sp = sub.add_parser("convert", help="run converters")
    add_deck(sp)
    add_select(sp)
    sp.set_defaults(func=cmd_convert)

    sp = sub.add_parser("metrics", help="(re)compute metrics")
    add_deck(sp)
    sp.set_defaults(func=cmd_metrics)

    sp = sub.add_parser("review", help="interactive human rating")
    add_deck(sp)
    sp.add_argument("--mode", choices=["deck", "sample", "allslides"], default="deck")
    sp.add_argument("--converters", help="comma-separated subset to review")
    sp.add_argument("--n", type=int, default=10, help="sample size for --mode sample")
    sp.set_defaults(func=cmd_review)

    sp = sub.add_parser("report", help="(re)build leaderboard")
    add_deck(sp)
    sp.set_defaults(func=cmd_report)

    sp = sub.add_parser("run", help="end-to-end pipeline")
    add_deck(sp)
    add_select(sp)
    sp.add_argument("--no-render", action="store_true", help="skip slide rendering")
    sp.add_argument("--review", choices=["deck", "sample", "allslides"], help="rate interactively after the run")
    sp.add_argument("--n", type=int, default=10)
    sp.add_argument("--width", type=int, default=1280)
    sp.set_defaults(func=cmd_run)

    return p


def main(argv=None) -> int:
    _utf8_stdout()
    args = build_parser().parse_args(argv)
    args.func(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
