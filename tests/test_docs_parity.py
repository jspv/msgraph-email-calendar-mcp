"""The documented tool surface must match the registered one, in both directions.

`AGENTS.md` requires docs to be updated alongside any change to the tool surface.
That rule was ignored once already and the docs drifted badly: `README.md`
advertised a `list_attachments` tool that does not exist, and `AGENTS.md`
described 14 tools against a 29-tool server. A rule with no enforcement is a
comment, so this test is the enforcement.

Both directions matter. A tool missing from the docs leaves agents unaware of a
capability; a documented tool that does not exist sends them to call something
that will fail.
"""
from __future__ import annotations

import asyncio
import importlib
import re
from dataclasses import replace
from pathlib import Path

import pytest

from msgraph_mcp import config, tools

REPO_ROOT = Path(__file__).resolve().parent.parent

# Every scope the server knows about, so the full surface registers.
ALL_SCOPES = (
    "User.Read",
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.Read",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "People.Read",
)

# README lists tools in a three-column table keyed by area:
#     | Mail | `send_message` | Send a new email (dry-run by default) |
_README_ROW = re.compile(r"^\|\s*(?:Auth|Mail|Calendar|Contacts)\s*\|\s*`([a-z][a-z0-9_]*)`\s*\|")

# AGENTS.md splits tools across per-scope tables, tool name in the first column:
#     | `send_message` | `to`, `subject`, `body`, `dry_run=True` | ... |
_AGENTS_ROW = re.compile(r"^\|\s*`([a-z][a-z0-9_]*)`\s*\|")


@pytest.fixture(autouse=True)
def _restore_tools_module():
    yield
    importlib.reload(tools)


def _registered_tool_names() -> set[str]:
    original = config.settings
    config.settings = replace(original, scopes=ALL_SCOPES)
    try:
        importlib.reload(tools)
        return {t.name for t in asyncio.run(tools.mcp.list_tools())}
    finally:
        config.settings = original
        importlib.reload(tools)


def _documented(filename: str, pattern: re.Pattern[str]) -> set[str]:
    text = (REPO_ROOT / filename).read_text(encoding="utf-8")
    return {m.group(1) for line in text.splitlines() if (m := pattern.match(line))}


class TestDocsParity:
    def test_readme_documents_every_registered_tool(self):
        missing = _registered_tool_names() - _documented("README.md", _README_ROW)
        assert not missing, (
            f"README.md is missing rows for registered tools: {sorted(missing)}. "
            "Add them to the 'Tool reference' table."
        )

    def test_readme_documents_no_phantom_tools(self):
        phantom = _documented("README.md", _README_ROW) - _registered_tool_names()
        assert not phantom, (
            f"README.md documents tools that do not exist: {sorted(phantom)}. "
            "Remove the rows or register the tools."
        )

    def test_agents_documents_every_registered_tool(self):
        missing = _registered_tool_names() - _documented("AGENTS.md", _AGENTS_ROW)
        assert not missing, (
            f"AGENTS.md is missing rows for registered tools: {sorted(missing)}. "
            "Add them to the per-scope tool tables."
        )

    def test_agents_documents_no_phantom_tools(self):
        phantom = _documented("AGENTS.md", _AGENTS_ROW) - _registered_tool_names()
        assert not phantom, (
            f"AGENTS.md documents tools that do not exist: {sorted(phantom)}. "
            "Remove the rows or register the tools."
        )

    def test_parsers_actually_find_the_tables(self):
        # A regex that silently matches nothing would make every assertion above
        # vacuously true, so pin the parsers to a known row.
        assert "send_message" in _documented("README.md", _README_ROW)
        assert "send_message" in _documented("AGENTS.md", _AGENTS_ROW)
