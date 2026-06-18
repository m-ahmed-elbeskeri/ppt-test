"""Interactive Textual TUI for ppteval.

Pick or add a deck, toggle converters on/off, run the pipeline (with live log),
watch the leaderboard fill in, and rate outputs 1-5 — all without leaving the
terminal.

Launch:  ppteval tui     (or the console script: ppteval-tui)
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

from textual import on, work
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual.screen import ModalScreen
from textual.widgets import (
    Button,
    DataTable,
    Footer,
    Header,
    Input,
    Label,
    RichLog,
    Rule,
    Select,
    SelectionList,
    Static,
    TextArea,
)
from textual.widgets.selection_list import Selection

from . import adapters as adapters_mod
from . import runner
from .config import DECKS_DIR, RATING_DIMENSIONS, RATING_MAX, out_dir_for
from .report import build_report
from .schema import Rating, RatingsStore

# Full metric set shown in the leaderboard table: (header, result-key).
FULL_COLS = [
    ("#", "rank"), ("Converter", "converter"), ("Final", "final_score"),
    ("Auto", "auto_score"), ("Human", "human_avg"),
    ("Text%", "text_recall"), ("Notes%", "notes_recall"), ("Tbl%", "table_cell_recall"),
    ("Img%", "image_recall"), ("Refs", "n_image_refs"), ("Extract", "images_extracted"),
    ("Struct%", "structure_score"), ("Head", "n_headings"), ("List", "n_list_items"),
    ("MdTbl", "n_md_tables"), ("HtmlTbl", "n_html_tables"), ("Bound%", "slide_boundary_recall"),
    ("Parse s", "parse_time_s"), ("Sl/s", "slides_per_sec"), ("Tokens", "out_tokens"),
    ("Tok/sl", "tokens_per_slide"), ("Chars", "out_chars"), ("Compr", "compression"),
    ("Status", "status"),
]
_PCT_KEYS = {
    "text_recall", "notes_recall", "table_cell_recall", "image_recall",
    "structure_score", "slide_boundary_recall",
}


def _cell(r: dict, key: str) -> str:
    if key in ("rank", "converter", "status"):
        return str(r.get(key, ""))
    if key in ("final_score", "auto_score"):
        return f"{r.get(key, 0):.3f}"
    if key == "human_avg":
        v = r.get("human_avg")
        return f"{v}/5" if v is not None else "—"
    v = r.get("metrics", {}).get(key)
    if key in _PCT_KEYS:
        return f"{v*100:.0f}%" if isinstance(v, float) else "—"
    return _num(v)


# What each column / rating dimension means (used by tooltips + the Help screen).
METRIC_HELP = {
    "rank": "Final ranking position (by Final score).",
    "converter": "Converter name.",
    "final_score": "Final = 60% your human rating + 40% auto score (auto-only if unrated).",
    "auto_score": "Weighted blend of the normalized automated metrics (weights in config.py).",
    "human_avg": "Average of your 1-5 ratings for this converter.",
    "text_recall": "% of ground-truth text (python-pptx) found in the output.",
    "notes_recall": "% of speaker-notes text captured.",
    "table_cell_recall": "% of table cells whose text survived (word-set match).",
    "image_recall": "% of slide visuals ANCHORED in output (![](), <img>, or <!-- image -->).",
    "n_image_refs": "Count of image references/anchors in the output.",
    "images_extracted": "Count of image FILES written to disk (anchoring is NOT extracting).",
    "structure_score": "0-1 composite: headings, lists, md+html tables, slide boundaries.",
    "n_headings": "Markdown headings emitted.",
    "n_list_items": "Markdown list items (bullets / numbered).",
    "n_md_tables": "Markdown pipe tables.",
    "n_html_tables": "HTML <table> blocks (also LLM-parseable).",
    "slide_boundary_recall": "% of slides delimited — enables per-slide RAG chunking.",
    "parse_time_s": "Wall-clock seconds to convert the deck.",
    "slides_per_sec": "Throughput = slides / parse time.",
    "out_tokens": "Output size in tiktoken o200k tokens (LLM-cost proxy).",
    "tokens_per_slide": "out_tokens / number of slides.",
    "out_chars": "Output size in characters.",
    "compression": "out_tokens / ground-truth tokens (~1 ideal; high = bloat, low = loss).",
    "status": "ok / error / skipped.",
}

DIMENSION_HELP = {
    "faithfulness": "Accuracy & completeness vs the real slide — no hallucination or dropped content.",
    "structure": "Titles, bullets, tables and reading order preserved in a clean, LLM-usable shape.",
    "tables": "Tables captured and well-formed (rows / columns intact).",
    "notes": "Speaker notes captured (or correctly absent).",
    "images": "Visuals handled usefully — referenced and/or extracted.",
    "overall": "Your gut call: how usable is this output for an LLM / RAG pipeline?",
}


def _column_glossary() -> str:
    return "Leaderboard columns (press F1 for full help):\n" + "\n".join(
        f"  {h} — {METRIC_HELP.get(k, '')}" for h, k in FULL_COLS
    )


def _list_decks() -> list[Path]:
    DECKS_DIR.mkdir(parents=True, exist_ok=True)
    return sorted(DECKS_DIR.glob("*.pptx"))


def _converter_selections() -> list[Selection]:
    """One SelectionList item (✓ checkbox) per converter; light+available pre-ticked."""
    sels: list[Selection] = []
    for r in adapters_mod.availability():
        heavy = " ·heavy" if r["heavy"] else ""
        tag = "" if r["available"] else " (unavailable)"
        prompt = f"{r['name']}  [{r['license']}]{heavy}{tag}"
        sels.append(Selection(prompt, r["name"], r["available"] and not r["heavy"]))
    return sels


def _open_path(path: Path) -> str | None:
    """Open a file/folder in the OS. Returns an error string or None."""
    if not Path(path).exists():
        return f"not found: {path}"
    if os.environ.get("PPTEVAL_NO_OPEN"):
        return None  # tests / headless: don't actually launch an app
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            import subprocess

            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return None
    except Exception as e:  # noqa: BLE001
        return f"{type(e).__name__}: {e}"


def _pct(v) -> str:
    return f"{v*100:.0f}%" if isinstance(v, float) else "—"


def _num(v) -> str:
    return f"{v:g}" if isinstance(v, (int, float)) else "—"


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _safe_report(od: Path) -> str | None:
    """Rebuild the report; return an error string instead of raising (e.g. when
    leaderboard.md / scorecard.csv is open elsewhere and locked on Windows)."""
    try:
        build_report(od)
        return None
    except Exception as e:  # noqa: BLE001 - surfaced to the log, never crashes UI
        return f"{type(e).__name__}: {e}"


# --- add-deck modal ----------------------------------------------------------
class AddDeckScreen(ModalScreen[str | None]):
    CSS = """
    AddDeckScreen { align: center middle; }
    #box { width: 84; height: auto; padding: 1 2; background: $panel; border: thick $accent; }
    #row { height: auto; margin-top: 1; }
    #row Button { width: 1fr; margin-right: 1; }
    """

    def compose(self) -> ComposeResult:
        with Vertical(id="box"):
            yield Label("Add a deck — full path to a .pptx (it's copied into tests/decks/):")
            yield Input(placeholder=r"C:\path\to\deck.pptx", id="path")
            with Horizontal(id="row"):
                yield Button("Add", variant="success", id="ok")
                yield Button("Cancel", id="cancel")

    @on(Button.Pressed, "#ok")
    @on(Input.Submitted, "#path")
    def _ok(self) -> None:
        self.dismiss(self.query_one("#path", Input).value.strip())

    @on(Button.Pressed, "#cancel")
    def _cancel(self) -> None:
        self.dismiss(None)


# --- help / glossary modal ---------------------------------------------------
class HelpScreen(ModalScreen[None]):
    CSS = """
    HelpScreen { align: center middle; }
    #help { width: 92%; height: 90%; padding: 1 2; background: $panel; border: thick $accent; }
    #helpbody { height: 1fr; border: round $primary; padding: 0 1; }
    .title { text-style: bold; color: $accent; }
    #help Button { width: auto; margin-top: 1; }
    """
    BINDINGS = [
        Binding("escape", "dismiss_help", "Close"),
        Binding("f1", "dismiss_help", "Close"),
        Binding("question_mark", "dismiss_help", "Close"),
    ]

    def compose(self) -> ComposeResult:
        with Vertical(id="help"):
            yield Label("ppteval — metric & rating glossary", classes="title")
            with VerticalScroll(id="helpbody"):
                yield Static(self._text())
            yield Button("Close", variant="primary", id="help-close")

    def _text(self) -> str:
        lines = ["[b]Leaderboard columns[/b]\n"]
        for h, k in FULL_COLS:
            lines.append(f"  [b]{h}[/b] — {METRIC_HELP.get(k, '')}")
        lines.append("\n[b]Rating dimensions (1-5)[/b]\n")
        for d in RATING_DIMENSIONS:
            lines.append(f"  [b]{d}[/b] — {DIMENSION_HELP.get(d, '')}")
        lines.append(
            "\n[dim]Recall is measured vs python-pptx-extractable content — a consistent "
            "cross-converter yardstick, not absolute truth. Tokens use tiktoken o200k "
            "(a proxy for LLM cost). Final blends your ratings 60/40 with the auto score; "
            "weights live in ppteval/config.py.[/dim]"
        )
        return "\n".join(lines)

    @on(Button.Pressed, "#help-close")
    def _close(self) -> None:
        self.dismiss(None)

    def action_dismiss_help(self) -> None:
        self.dismiss(None)


# --- review modal ------------------------------------------------------------
class ReviewScreen(ModalScreen[None]):
    CSS = """
    ReviewScreen { align: center middle; }
    #panel { width: 96%; height: 94%; padding: 1 2; background: $panel; border: thick $accent; }
    #substat { color: $text-muted; }
    #openrow { height: auto; margin: 1 0; }
    #openrow Button { width: auto; margin-right: 1; }
    #preview { height: 1fr; border: round $primary; }
    #dims { height: auto; max-height: 14; border: round $primary; padding: 0 1; }
    .dimrow { height: 3; }
    .dimrow Label { width: 16; padding-top: 1; }
    #btns { height: auto; margin-top: 1; }
    #btns Button { width: 1fr; margin-right: 1; }
    .title { text-style: bold; color: $accent; }
    """
    BINDINGS = [
        Binding("escape", "close", "Close"),
        Binding("ctrl+o", "open_folder", "Output folder"),
        Binding("ctrl+p", "open_pptx", "Open PowerPoint"),
        Binding("f1", "help", "Help"),
    ]

    def action_help(self) -> None:
        self.app.push_screen(HelpScreen())

    def __init__(self, deck: Path) -> None:
        super().__init__()
        self.deck = Path(deck)
        self.od = out_dir_for(deck)
        self.outputs = [
            o for o in runner.load_outputs(self.od) if o.status == "ok" and o.markdown
        ]
        self.idx = 0
        self.store = self._load_store()
        self.metrics: dict[str, dict] = {}
        mp = self.od / "metrics.json"
        if mp.exists():
            try:
                for row in json.loads(mp.read_text(encoding="utf-8")):
                    self.metrics[row["converter"]] = row.get("metrics", {})
            except Exception:
                pass

    def _load_store(self) -> RatingsStore:
        p = self.od / "ratings.json"
        if p.exists():
            return RatingsStore.model_validate_json(p.read_text(encoding="utf-8"))
        return RatingsStore(deck_slug=self.od.name)

    def compose(self) -> ComposeResult:
        with Vertical(id="panel"):
            yield Label(id="title", classes="title")
            yield Label(id="substat")
            with Horizontal(id="openrow"):
                yield Button("📁 Output folder", id="rev-folder")
                yield Button("📊 Open PowerPoint", id="rev-pptx")
            yield TextArea("", id="preview", read_only=True, soft_wrap=True)
            with VerticalScroll(id="dims"):
                for dim in RATING_DIMENSIONS:
                    with Horizontal(classes="dimrow"):
                        dim_label = Label(dim)
                        dim_label.tooltip = DIMENSION_HELP.get(dim, "")
                        yield dim_label
                        yield Select(
                            [(str(i), i) for i in range(1, RATING_MAX + 1)],
                            id=f"d-{dim}",
                            allow_blank=True,
                            prompt="– skip –",
                        )
            with Horizontal(id="btns"):
                yield Button("Save & Next ▶", variant="success", id="save")
                yield Button("Skip", id="skip")
                yield Button("Close & rebuild", variant="primary", id="close")

    def on_mount(self) -> None:
        if not self.outputs:
            self.notify("No converter outputs — run the pipeline first.", severity="warning")
            self.dismiss(None)
            return
        self.query_one("#rev-folder", Button).tooltip = (
            "Open this converter's output folder (output.md + any extracted images)."
        )
        self.query_one("#rev-pptx", Button).tooltip = "Open the source .pptx in PowerPoint."
        self._show()

    def _substat(self, m: dict) -> str:
        def pct(k: str) -> str:
            v = m.get(k)
            return f"{v*100:.0f}%" if isinstance(v, float) else "—"

        return (
            f"parse {m.get('parse_time_s', '?')}s · {m.get('out_tokens', '?')} tokens · "
            f"text {pct('text_recall')} · struct {pct('structure_score')} · "
            f"img refs {m.get('n_image_refs', '?')} / extracted {m.get('images_extracted', '?')}"
        )

    def _show(self) -> None:
        o = self.outputs[self.idx]
        self.query_one("#title", Label).update(
            f"[{self.idx + 1}/{len(self.outputs)}]  {o.converter}   —   "
            f"scroll the full output, then rate 1-{RATING_MAX} (Ctrl+O folder · Ctrl+P PowerPoint)"
        )
        self.query_one("#substat", Label).update(self._substat(self.metrics.get(o.converter, {})))
        ta = self.query_one("#preview", TextArea)
        ta.text = o.markdown or ""
        ta.scroll_home(animate=False)
        for dim in RATING_DIMENSIONS:
            sel = self.query_one(f"#d-{dim}", Select)
            existing = next(
                (
                    r.score
                    for r in self.store.ratings
                    if r.converter == o.converter and r.mode == "deck" and r.dimension == dim
                ),
                None,
            )
            if existing is not None:
                sel.value = int(existing)
            else:
                sel.clear()  # canonical way to blank a Select (allow_blank=True)

    @on(Button.Pressed, "#rev-folder")
    def action_open_folder(self) -> None:
        o = self.outputs[self.idx]
        err = _open_path(self.od / "converters" / o.converter)
        self.notify(err or f"opened {o.converter} output folder", severity="warning" if err else "information")

    @on(Button.Pressed, "#rev-pptx")
    def action_open_pptx(self) -> None:
        err = _open_path(self.deck)
        self.notify(err or f"opening {self.deck.name}", severity="warning" if err else "information")

    @on(Button.Pressed, "#save")
    def _save(self) -> None:
        o = self.outputs[self.idx]
        ts = _now()
        for dim in RATING_DIMENSIONS:
            v = self.query_one(f"#d-{dim}", Select).value
            if isinstance(v, (int, float)):  # skip the NoSelection blank sentinel
                self.store.upsert(
                    Rating(converter=o.converter, mode="deck", dimension=dim, score=float(v), ts=ts)
                )
        (self.od / "ratings.json").write_text(self.store.model_dump_json(indent=2), encoding="utf-8")
        self.notify(f"Saved ratings for {o.converter}")
        self._advance()

    @on(Button.Pressed, "#skip")
    def _skip(self) -> None:
        self._advance()

    def _advance(self) -> None:
        if self.idx + 1 >= len(self.outputs):
            _safe_report(self.od)
            self.dismiss(None)
        else:
            self.idx += 1
            self._show()

    @on(Button.Pressed, "#close")
    def _close_btn(self) -> None:
        self.action_close()

    def action_close(self) -> None:
        _safe_report(self.od)
        self.dismiss(None)


# --- main app ----------------------------------------------------------------
class PptEvalApp(App):
    TITLE = "ppteval"
    SUB_TITLE = "PPTX → LLM converter eval"
    CSS = """
    #main { height: 1fr; }
    #sidebar { width: 42; border-right: solid $primary; padding: 0 1; }
    #content { padding: 0 1; }
    #converters { height: 1fr; border: round $primary; }
    #opts { height: auto; max-height: 5; border: round $primary; }
    #board { height: 1fr; }
    #hint { height: 1; color: $accent; }
    #openbar { height: auto; }
    #openbar Button { width: auto; margin-right: 1; }
    #log { height: 7; border: round $primary; }
    .title { text-style: bold; color: $accent; }
    #sidebar Button { width: 1fr; margin-top: 1; }
    """
    BINDINGS = [
        Binding("r", "run", "Run"),
        Binding("e", "review", "Review"),
        Binding("a", "add_deck", "Add deck"),
        Binding("b", "report", "Rebuild report"),
        Binding("o", "open_deck", "Open folder"),
        Binding("c", "open_csv", "Open CSV"),
        Binding("f", "open_converter", "Conv folder"),
        Binding("f1", "help", "Help"),
        Binding("question_mark", "help", "Help"),
        Binding("q", "quit", "Quit"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.decks = _list_decks()
        self.current: Path | None = self.decks[0] if self.decks else None
        self._row_converters: list[str] = []  # board row index -> converter name

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        with Horizontal(id="main"):
            with Vertical(id="sidebar"):
                yield Label("Deck", classes="title")
                yield Select(
                    [(d.name, str(d)) for d in self.decks],
                    id="deck",
                    allow_blank=True,
                    value=str(self.current) if self.current else Select.BLANK,
                )
                yield Button("➕ Add deck", id="add")
                yield Rule()
                yield Label("Converters — space/click toggles ✓", classes="title")
                yield SelectionList(*_converter_selections(), id="converters")
                yield Rule()
                yield SelectionList(
                    Selection("Render slides (PowerPoint COM)", "render", True),
                    id="opts",
                )
                yield Button("▶ Run selected", variant="success", id="run")
                yield Button("★ Review / rate", variant="primary", id="review")
                yield Button("↻ Rebuild report", id="report")
                yield Button("❔ Help (F1)", id="help")
            with Vertical(id="content"):
                yield Label(id="deckinfo", classes="title")
                yield DataTable(id="board", zebra_stripes=True)
                yield Label(id="hint")
                with Horizontal(id="openbar"):
                    yield Button("📂 Deck folder", id="open-deck")
                    yield Button("📄 CSV", id="open-csv")
                    yield Button("📝 leaderboard.md", id="open-md")
                    yield Button("📁 Converter folder (selected row)", id="open-conv")
                yield RichLog(id="log", highlight=True, markup=True)
        yield Footer()

    def on_mount(self) -> None:
        dt = self.query_one("#board", DataTable)
        dt.add_columns(*[h for h, _ in FULL_COLS])
        dt.cursor_type = "cell"  # so moving across columns reveals each one's meaning
        dt.fixed_columns = 2  # keep #/Converter pinned while scrolling metrics
        self.query_one("#converters", SelectionList).tooltip = (
            "Tick the converters to run. Heavy ones (docling/unstructured/marker) pull large deps."
        )
        self.query_one("#opts", SelectionList).tooltip = (
            "Tick to render each slide to PNG (PowerPoint) for side-by-side review."
        )
        self._update_deckinfo()
        self._load_board()
        self.query_one("#hint", Label).update(
            "[dim]Click a cell / arrow ← → across columns to see each metric's meaning here · F1 = full glossary[/dim]"
        )
        self._log(
            "[b]ppteval[/] ready. Toggle converters (space/click), press [b]Run[/] (r). "
            "Move across table columns (or press [b]F1[/]) for metric meanings; pick a row + [b]f[/] to open its folder."
        )

    # --- helpers ---
    def _log(self, msg: str) -> None:
        self.query_one("#log", RichLog).write(msg)

    def _update_deckinfo(self) -> None:
        info = self.query_one("#deckinfo", Label)
        if not self.current:
            info.update("No decks yet — press [b]Add deck[/] (a)")
            return
        gtp = out_dir_for(self.current) / "ground_truth.json"
        if gtp.exists():
            gt = json.loads(gtp.read_text(encoding="utf-8"))
            info.update(
                f"{self.current.name}  —  {gt['n_slides']} slides · {gt['n_tables']} tables · "
                f"{gt['n_images']} images · {gt['n_charts']} charts · ~{gt['gt_tokens']} tokens"
            )
        else:
            info.update(f"{self.current.name}  —  not analyzed yet (press Run)")

    def _load_board(self) -> None:
        dt = self.query_one("#board", DataTable)
        dt.clear()
        self._row_converters = []
        if not self.current:
            return
        rp = out_dir_for(self.current) / "results.json"
        if not rp.exists():
            return
        data = json.loads(rp.read_text(encoding="utf-8"))
        for r in data["results"]:
            dt.add_row(*[_cell(r, key) for _, key in FULL_COLS])
            self._row_converters.append(r["converter"])

    def _selected(self) -> list[str]:
        return list(self.query_one("#converters", SelectionList).selected)

    def _open(self, path) -> None:
        err = _open_path(Path(path))
        self._log(f"[red]{err}[/]" if err else f"opened {Path(path).name or path}")

    # --- events ---
    @on(Select.Changed, "#deck")
    def _deck_changed(self, e: Select.Changed) -> None:
        if e.value and e.value != Select.BLANK:
            self.current = Path(str(e.value))
            self._update_deckinfo()
            self._load_board()

    @on(DataTable.CellHighlighted, "#board")
    def _cell_hint(self, e: DataTable.CellHighlighted) -> None:
        col = e.coordinate.column
        if 0 <= col < len(FULL_COLS):
            header, key = FULL_COLS[col]
            self.query_one("#hint", Label).update(f"[b]{header}[/b] — {METRIC_HELP.get(key, '')}")

    @on(Button.Pressed, "#run")
    def _run_btn(self) -> None:
        self.action_run()

    @on(Button.Pressed, "#review")
    def _review_btn(self) -> None:
        self.action_review()

    @on(Button.Pressed, "#report")
    def _report_btn(self) -> None:
        self.action_report()

    @on(Button.Pressed, "#add")
    def _add_btn(self) -> None:
        self.action_add_deck()

    @on(Button.Pressed, "#open-deck")
    def _open_deck_btn(self) -> None:
        self.action_open_deck()

    @on(Button.Pressed, "#open-csv")
    def _open_csv_btn(self) -> None:
        self.action_open_csv()

    @on(Button.Pressed, "#open-md")
    def _open_md_btn(self) -> None:
        if self.current:
            self._open(out_dir_for(self.current) / "leaderboard.md")

    @on(Button.Pressed, "#open-conv")
    def _open_conv_btn(self) -> None:
        self.action_open_converter()

    @on(Button.Pressed, "#help")
    def _help_btn(self) -> None:
        self.action_help()

    def action_help(self) -> None:
        self.push_screen(HelpScreen())

    # --- actions ---
    def action_run(self) -> None:
        if not self.current:
            self._log("[red]No deck selected.[/]")
            return
        names = self._selected()
        if not names:
            self._log("[red]No converters selected.[/]")
            return
        do_render = "render" in self.query_one("#opts", SelectionList).selected
        self._log(f"[b]Running[/] {self.current.name} → {', '.join(names)} (render={do_render})…")
        self._run_worker(self.current, names, do_render)

    @work(thread=True, exclusive=True)
    def _run_worker(self, deck: Path, names: list[str], do_render: bool) -> None:
        try:
            adapters = adapters_mod.select(names)
            od, gt, outputs, render = runner.run_all(deck, adapters, do_render=do_render)
            rep_err = _safe_report(od)
            if rep_err:
                self.call_from_thread(self._log, f"[yellow]report: {rep_err}[/]")
            for o in outputs:
                tail = f" [red]{o.error}[/]" if o.error else ""
                self.call_from_thread(self._log, f"  {o.converter}: {o.status} ({o.elapsed_s:.1f}s){tail}")
            if render is not None:
                rmsg = f"ok {len(render.png_paths)} slides" if render.ok else f"skip/fail ({render.error})"
                self.call_from_thread(self._log, f"  render: {rmsg}")
            self.call_from_thread(self._finish_run)
        except Exception as e:  # surface, never crash the UI
            self.call_from_thread(self._log, f"[red]Run failed: {type(e).__name__}: {e}[/]")

    def _finish_run(self) -> None:
        self._update_deckinfo()
        self._load_board()
        self._log("[green]Done.[/] Press [b]Review[/] (e) to rate, or toggle converters and Run again.")

    def action_review(self) -> None:
        if not self.current:
            return
        if not (out_dir_for(self.current) / "converters.json").exists():
            self._log("[yellow]Run the pipeline first (no outputs to review).[/]")
            return
        self.push_screen(ReviewScreen(self.current), lambda _: self._after_review())

    def _after_review(self) -> None:
        self._load_board()
        self._log("[green]Ratings saved; leaderboard updated.[/]")

    def action_report(self) -> None:
        if self.current and (out_dir_for(self.current) / "metrics.json").exists():
            err = _safe_report(out_dir_for(self.current))
            self._load_board()
            self._log(f"[red]report: {err}[/]" if err else "Report rebuilt.")

    def action_open_deck(self) -> None:
        if self.current:
            self._open(out_dir_for(self.current))

    def action_open_csv(self) -> None:
        if self.current:
            self._open(out_dir_for(self.current) / "scorecard.csv")

    def action_open_converter(self) -> None:
        if not self.current:
            return
        dt = self.query_one("#board", DataTable)
        idx = dt.cursor_row
        if idx is None or not (0 <= idx < len(self._row_converters)):
            self._log("[yellow]Select a row in the table first.[/]")
            return
        name = self._row_converters[idx]
        self._open(out_dir_for(self.current) / "converters" / name)

    def action_add_deck(self) -> None:
        self.push_screen(AddDeckScreen(), self._added)

    def _added(self, path: str | None) -> None:
        if not path:
            return
        src = Path(path)
        if not src.exists() or src.suffix.lower() != ".pptx":
            self._log(f"[red]Not a .pptx file: {path}[/]")
            return
        dest = DECKS_DIR / src.name
        try:
            shutil.copy(src, dest)
        except Exception as e:
            self._log(f"[red]Copy failed: {e}[/]")
            return
        self.decks = _list_decks()
        self.current = dest
        sel = self.query_one("#deck", Select)
        sel.set_options([(d.name, str(d)) for d in self.decks])
        sel.value = str(dest)
        self._update_deckinfo()
        self._load_board()
        self._log(f"[green]Added {src.name}[/] — press Run to analyze it.")


def run() -> None:
    PptEvalApp().run()


if __name__ == "__main__":
    run()
