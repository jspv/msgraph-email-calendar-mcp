"""The dry-run default has to reach the MCP tool surface, not just the module.

`calendar.py` is internal; agents only ever call `tools.py`. A safe default that
the wrapper silently drops would be worse than no default at all, because the
docs would claim a protection that is not there.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp import tools


class TestToolLayerPassesDryRunThrough:
    @patch("msgraph_mcp.calendar.create_event")
    def test_create_event_defaults_to_dry_run(self, create):
        tools.create_event(
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
        )
        assert create.call_args[1]["dry_run"] is True

    @patch("msgraph_mcp.calendar.create_event")
    def test_create_event_forwards_explicit_false(self, create):
        tools.create_event(
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
            dry_run=False,
        )
        assert create.call_args[1]["dry_run"] is False

    @patch("msgraph_mcp.calendar.update_event")
    def test_update_event_defaults_to_dry_run(self, update):
        tools.update_event(event_id="event-1", subject="New")
        assert update.call_args[1]["dry_run"] is True

    @patch("msgraph_mcp.calendar.update_event")
    def test_update_event_forwards_explicit_false(self, update):
        tools.update_event(event_id="event-1", subject="New", dry_run=False)
        assert update.call_args[1]["dry_run"] is False

    @patch("msgraph_mcp.calendar.delete_event")
    def test_delete_event_defaults_to_dry_run(self, delete):
        tools.delete_event(event_id="event-1")
        assert delete.call_args[1]["dry_run"] is True

    @patch("msgraph_mcp.calendar.delete_event")
    def test_delete_event_forwards_explicit_false(self, delete):
        tools.delete_event(event_id="event-1", dry_run=False)
        assert delete.call_args[1]["dry_run"] is False

    @patch("msgraph_mcp.calendar.respond_to_event")
    def test_respond_to_event_has_no_dry_run(self, respond):
        # Responding is reversible by responding again; gating it is friction
        # with no safety payoff.
        with pytest.raises(TypeError):
            tools.respond_to_event(event_id="event-1", response="accept", dry_run=True)
