# ppteval — rank PowerPoint → LLM-format converters

A free / open-source evaluation suite for R&D on the **"PPTX → LLM-ready format"**
problem. It runs a *selectable* set of OSS converters over your decks, scores them
with **deterministic metrics** (parse time, output tokens, content recall,
structure), lets **you rate the outputs yourself** in the CLI (no LLM judge), and
produces a combined **leaderboard**.

No paid services, no cloud APIs, no LLM-as-judge — quality = objective metrics +
your ratings.

---

## Why it exists

Slides are multimodal and layout-heavy; no single extractor wins on every deck.
This harness lets you measure converters on *your* decks instead of trusting
marketing claims — and re-run it as new tools appear, or as you tune what
"LLM-ready" means for your pipeline (weights live in one file).

## What's measured

**Operational**
- `parse_time_s`, `slides_per_sec` — speed / throughput
- `out_tokens` (tiktoken `o200k_base`), `tokens_per_slide`, `out_chars` — LLM cost
- `compression` = out_tokens / ground-truth-tokens — flags content loss vs bloat

**Content recall** (vs a python-pptx ground truth)
- `text_recall`, `notes_recall`, `table_cell_recall`
- `image_recall` — does the output *anchor* each visual (Markdown `![]()`, HTML
  `<img>`, or a `<!-- image -->` placeholder) so a downstream VLM pass knows it
  exists. Bare filename strings in a plain-text dump do **not** count.
- `images_extracted` — how many image *files* the converter actually wrote to
  disk (a different capability from anchoring; informational, not scored by
  default — add a weight in `config.py` if your pipeline needs the bytes).

**Structure**
- `n_headings`, `n_list_items`, `n_md_tables` + `n_html_tables`, `slide_boundary_recall`
- `structure_score` — 0..1 composite of formatting fidelity (what makes per-slide
  RAG chunking possible). Credits both Markdown and HTML tables.

> **Honest caveat.** Recall is measured against *python-pptx-extractable* content.
> python-pptx can't see text inside SmartArt without a text frame, text rasterised
> into images (PNG/EMF/WDP), or some embedded objects. So recall is a **consistent
> cross-converter yardstick, not absolute truth** — your human ratings + the
> rendered slides are the corrective for what ground truth misses. tiktoken counts
> are a proxy for LLM cost, not the exact tokenizer of any one model.

## Converters (all free / OSS)

| name | license | set | notes |
|---|---|---|---|
| `markitdown` | MIT | light | MS MarkItDown — LLM-oriented Markdown |
| `pptx_custom` | MIT | light | our python-pptx extractor; **reference** (defines ground truth) |
| `pptx2md` | MIT | light | PPTX-specific Markdown + images |
| `pymupdf4llm` | AGPL-3.0 | light | renders PPTX→PDF (PowerPoint COM) then parses |
| `tika` | Apache-2.0 | light | Apache Tika; plain text (needs a JRE) |
| `docling` | MIT | heavy | unified document model → Markdown/JSON |
| `unstructured` | Apache-2.0 | heavy | typed elements for RAG |
| `marker` | GPL-3.0 | heavy | Datalab Marker via PDF (PyTorch) |

`pptx_custom` is the reference extractor, so its recall is ~100% by construction —
it competes on structure, tokens, speed and your ratings.

## Install

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -e .          # light set (default)
# optional, large deps (incl. PyTorch):
.venv/Scripts/python -m pip install -r requirements-heavy.txt
```

- **Slide rendering** (for review) uses **PowerPoint via COM** on Windows
  (`pywin32`) — no LibreOffice needed. The PDF-based converters (`pymupdf4llm`,
  `marker`) use the same path.
- **Tika** downloads a `tika-server` jar on first use (needs Java on PATH + network
  once).
- Any converter whose deps are missing is **SKIPPED**, never fatal.

## Interactive TUI (easiest)

```bash
.venv/Scripts/ppteval tui          # or the console script: ppteval-tui
```

A full-screen terminal app (Textual) that drives everything without memorising
commands:

- **pick a deck** from `tests/decks/`, or **➕ Add deck** to copy a new `.pptx` in;
- **toggle converters** with ✓ checkboxes (space/click) — pre-ticked for the
  available light set; flip on the heavy ones when you want them;
- **▶ Run** the selected converters (optional slide render) — live log, then the
  **full-metric leaderboard** fills in: Final/Auto/Human, text·notes·table·image
  recall, image **Refs vs Extracted**, structure (headings/lists/md+html tables/
  boundaries), and timing/tokens. Scroll ← → for every column; `#`/Converter stay pinned;
- **open buttons** (and keys): 📂 deck folder (`o`), 📄 `scorecard.csv` (`c`),
  📝 `leaderboard.md`, 📁 the **selected row's converter folder** (`f`);
- **★ Review** opens a rating screen: **scroll the converter's full output**, score
  1–5 per dimension (saved instantly, folded into the board), and jump to the
  **📁 output folder** or **📊 open the source PowerPoint** for what you're rating.
- **Column meanings** — move the cell cursor across the table (click a cell or
  arrow ← →) and a **live hint line** under it explains that column; **❔ Help / F1**
  shows the full glossary of every column and rating dimension; rating dimensions
  also have hover tooltips.

Keys: `r` run · `e` review · `a` add deck · `b` rebuild report · `o` open folder ·
`c` open CSV · `f` converter folder · `F1`/`?` help · `q` quit.
In review: `Ctrl+O` output folder · `Ctrl+P` open PowerPoint · `Esc` close.

## Quickstart (CLI)

```bash
# what's installed?
.venv/Scripts/ppteval list-converters

# end-to-end on the bundled deck (light converters), then rate it yourself:
.venv/Scripts/ppteval run "ED20473 Vycarb WP3 Draft v0.3.pptx" --review deck

# pick specific converters / include heavy ones:
.venv/Scripts/ppteval run mydeck.pptx --converters markitdown,docling,tika
.venv/Scripts/ppteval run mydeck.pptx --all          # every available converter
.venv/Scripts/ppteval run mydeck.pptx --no-render     # skip slide PNGs (faster)
```

`run` does: ground-truth → render → convert(selected) → metrics → report. The
deck arg accepts a path or a name/substring under `tests/decks/`.

### Rate outputs yourself

```bash
ppteval review mydeck.pptx --mode deck        # one rating per converter + ranking
ppteval review mydeck.pptx --mode sample --n 10   # N auto-picked rich slides
ppteval review mydeck.pptx --mode allslides   # every slide
```

Review opens the **rendered slide** next to each converter's text so you rate
against the real slide. Ratings (1–5 across faithfulness / structure / tables /
notes / images / overall, plus optional head-to-head ranking) are saved after
every entry — interrupt and resume anytime. `final_score` blends your ratings with
the auto score (default 60/40, tunable).

### Run stages individually

```bash
ppteval ground-truth mydeck.pptx
ppteval render mydeck.pptx
ppteval convert mydeck.pptx --converters markitdown,pptx2md   # accumulates
ppteval metrics mydeck.pptx
ppteval report mydeck.pptx
```

Converting subsets at different times **accumulates** into one leaderboard.

## Outputs

```
out/<deck_slug>/
  ground_truth.json
  renders/slide_001.png …                 # PowerPoint-rendered slides
  converters/<name>/output.md (+ output.json, images/)
  converters.json                          # slim per-converter records
  metrics.json                             # raw automated metrics
  ratings.json                             # your human ratings
  leaderboard.md  scorecard.csv  results.json
```

## Tuning the scoring

All scoring policy is in [`ppteval/config.py`](ppteval/config.py): metric weights
(`AUTO_WEIGHTS`), the human/auto blend (`HUMAN_BLEND`), the ideal token-density
band (`COMPRESSION_*`), tokenizer, and rating dimensions. Raw metrics are always
shown in the report, so re-weighting never hides data.

## Add a converter

Drop a module in `ppteval/adapters/` subclassing `Adapter` (implement
`available()` and `_convert()`), register it in `ppteval/adapters/__init__.py`, and
it's picked up by `list-converters`, `--converters`, metrics and the leaderboard.

## Tests

```bash
.venv/Scripts/python -m pytest tests/ -q
```

Smoke tests build a tiny synthetic deck (no large fixture, no COM/network) and
exercise availability → ground truth → convert → metrics → report.
