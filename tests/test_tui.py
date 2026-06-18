"""Headless smoke test for the Textual TUI (composes + key widgets + actions).

Uses Textual's run_test harness via asyncio.run, so no pytest-asyncio dependency.
PPTEVAL_NO_OPEN stops the open-folder/CSV actions from actually launching apps.
"""
from __future__ import annotations

import asyncio
import os

from textual.widgets import DataTable, SelectionList

from ppteval import adapters
from ppteval.tui import FULL_COLS, HelpScreen, PptEvalApp


def test_tui_composes_and_actions():
    os.environ["PPTEVAL_NO_OPEN"] = "1"

    async def go():
        app = PptEvalApp()
        async with app.run_test(size=(160, 40)) as pilot:
            await pilot.pause()
            for wid in ("#board", "#deck", "#log", "#deckinfo", "#opts", "#converters"):
                app.query_one(wid)
            # one ✓-toggle per converter; light+available ones pre-selected
            sl = app.query_one("#converters", SelectionList)
            assert sl.option_count == len(adapters.availability())
            assert "markitdown" in app._selected()
            # render is a ticked option, on by default
            assert "render" in app.query_one("#opts", SelectionList).selected
            # full metric set present as columns; per-column hint line exists
            dt = app.query_one("#board", DataTable)
            assert len(dt.columns) == len(FULL_COLS)
            app.query_one("#hint")
            # moving the cell cursor (which drives the hint) must not raise
            if dt.row_count:
                dt.move_cursor(row=0, column=5)
                await pilot.pause()
            # open + report actions must not raise
            app.action_report()
            app.action_open_csv()
            app.action_open_deck()
            await pilot.pause()
            # help glossary opens
            app.action_help()
            await pilot.pause()
            assert isinstance(app.screen, HelpScreen)
            app.pop_screen()
            await pilot.pause()

    asyncio.run(go())
