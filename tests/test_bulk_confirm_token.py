"""Destructive bulk actions require a token echoed from a dry-run.

The token is derived from the *matched message ids*, not from the filter
arguments, so it cannot authorise a set the caller never saw. The live run
recomputes it from its own scan: no server state, no salt, no clock, which keeps
it correct across a Lambda cold start between the two calls.

Scope is deliberate. ``delete`` and ``move`` are gated; ``mark_read`` and
``mark_unread`` are trivially reversible, so gating them would be friction with
no safety payoff.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp import tools
from msgraph_mcp.mail import _confirm_token, bulk_manage_messages_multi_pass

from tests.test_bulk_manage import _fake_request, _msg


def _pages(n: int) -> list[list[dict]]:
    """One short page of *n* messages, so the scan reports the folder exhausted."""
    return [[_msg(i, f"2026-01-{i:02d}T00:00:00Z") for i in range(1, n + 1)]]


def _run(MockClient, *, pages=None, **kwargs):
    request, _calls = _fake_request(pages or _pages(3), total=3)
    MockClient.return_value.request.side_effect = request
    return bulk_manage_messages_multi_pass(folder="inbox", **kwargs)


class TestTokenDerivation:
    def test_token_is_stable_for_the_same_matched_set(self):
        assert _confirm_token("delete", None, ["b", "a"]) == _confirm_token(
            "delete", None, ["a", "b"]
        )

    def test_token_changes_when_the_matched_set_changes(self):
        assert _confirm_token("delete", None, ["a", "b"]) != _confirm_token(
            "delete", None, ["a", "b", "c"]
        )

    def test_token_changes_when_the_action_changes(self):
        # A token authorising a preview of a move must not authorise a delete.
        assert _confirm_token("delete", None, ["a"]) != _confirm_token("move", None, ["a"])

    def test_token_changes_when_the_destination_changes(self):
        assert _confirm_token("move", "archive", ["a"]) != _confirm_token(
            "move", "junk", ["a"]
        )

    def test_token_is_prefixed_with_the_match_count(self):
        # The server keeps no state, so without this prefix a mismatch could
        # only say "changed" -- never "42 became 43".
        assert _confirm_token("delete", None, ["a", "b", "c"]).startswith("3-")


class TestDryRunIssuesToken:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_delete_dry_run_returns_a_token(self, MockClient):
        result = _run(MockClient, action="delete", dry_run=True)
        assert result["confirm_token"] == _confirm_token("delete", None, ["m1", "m2", "m3"])

    @patch("msgraph_mcp.mail.GraphClient")
    def test_move_dry_run_returns_a_token_bound_to_the_destination(self, MockClient):
        result = _run(MockClient, action="move", destination="archive", dry_run=True)
        assert result["confirm_token"] == _confirm_token(
            "move", "archive", ["m1", "m2", "m3"]
        )

    @patch("msgraph_mcp.mail.GraphClient")
    def test_reversible_actions_get_no_token(self, MockClient):
        # Presence of the field is the signal that confirmation is required.
        result = _run(MockClient, action="mark_read", dry_run=True)
        assert "confirm_token" not in result


class TestLiveRunRequiresToken:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_delete_without_a_token_is_refused(self, MockClient):
        with pytest.raises(ValueError, match="dry_run=True"):
            _run(MockClient, action="delete", dry_run=False)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_move_without_a_token_is_refused(self, MockClient):
        with pytest.raises(ValueError, match="confirm_token"):
            _run(MockClient, action="move", destination="archive", dry_run=False)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_nothing_is_deleted_when_the_token_is_missing(self, MockClient):
        request, calls = _fake_request(_pages(3), total=3)
        MockClient.return_value.request.side_effect = request
        with pytest.raises(ValueError):
            bulk_manage_messages_multi_pass(folder="inbox", action="delete", dry_run=False)
        assert not [c for c in calls if c[0] == "POST"]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_correct_token_lets_the_run_proceed(self, MockClient):
        token = _confirm_token("delete", None, ["m1", "m2", "m3"])
        result = _run(MockClient, action="delete", dry_run=False, confirm_token=token)
        assert result["acted"] == 3

    @patch("msgraph_mcp.mail.GraphClient")
    def test_stale_token_is_refused_when_the_match_set_grew(self, MockClient):
        stale = _confirm_token("delete", None, ["m1", "m2"])
        with pytest.raises(ValueError) as exc:
            _run(MockClient, action="delete", dry_run=False, confirm_token=stale)
        assert "2" in str(exc.value) and "3" in str(exc.value)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_mismatch_error_supplies_the_fresh_token(self, MockClient):
        # Recovery must be one call, not a guessing game.
        stale = _confirm_token("delete", None, ["m1", "m2"])
        with pytest.raises(ValueError) as exc:
            _run(MockClient, action="delete", dry_run=False, confirm_token=stale)
        assert _confirm_token("delete", None, ["m1", "m2", "m3"]) in str(exc.value)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_token_for_a_different_action_is_refused(self, MockClient):
        move_token = _confirm_token("move", "archive", ["m1", "m2", "m3"])
        with pytest.raises(ValueError):
            _run(MockClient, action="delete", dry_run=False, confirm_token=move_token)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_reversible_actions_need_no_token(self, MockClient):
        result = _run(MockClient, action="mark_read", dry_run=False)
        assert result["acted"] == 3


class TestToolLayerExposesConfirmToken:
    """The gate has to reach the MCP surface; `mail.py` is internal."""

    @patch("msgraph_mcp.mail.bulk_manage_messages_multi_pass")
    def test_tool_forwards_the_token(self, bulk):
        tools.bulk_manage_messages(
            action="delete", dry_run=False, confirm_token="3-abc123def456"
        )
        assert bulk.call_args[1]["confirm_token"] == "3-abc123def456"

    @patch("msgraph_mcp.mail.bulk_manage_messages_multi_pass")
    def test_tool_defaults_the_token_to_none(self, bulk):
        tools.bulk_manage_messages(action="delete", dry_run=True)
        assert bulk.call_args[1]["confirm_token"] is None
