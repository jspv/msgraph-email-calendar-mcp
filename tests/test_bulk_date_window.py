"""Date bounds are pushed into Graph's ``$filter`` (issue #4).

Reaching old mail was never impossible -- a default scan walks the whole folder --
but it was expensive, because both bounds were applied client-side *after* the
fetch. On a 50k-message mailbox, "what did this sender send me last March" paged
all 50k rows to match a few dozen.

Both bounds now travel to Graph, which turns a date-scoped query from O(folder)
into O(window). ``received_before`` also falls out of the existing pagination for
free: ``_collect_matches`` already anchors its cursor on ``receivedDateTime le``,
so the upper bound is simply the initial cursor -- the scan *starts* inside the
window rather than at the newest message.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp.mail import bulk_manage_messages_multi_pass

from tests.test_bulk_manage import _msg


def _capture(pages=None, total=3):
    """Record the params of every scan GET.

    Non-final pages carry ``@odata.nextLink`` so the scan keeps paging: a short
    page with no nextLink is (correctly) treated as the end of the window, which
    would otherwise make a multi-page test silently one page long.
    """
    seen: list[dict] = []
    remaining = list(pages if pages is not None else [[_msg(1, "2026-03-15T00:00:00Z")]])

    def _request(method, path, *, params=None, json_body=None):
        if method == "GET" and path.endswith("/messages"):
            seen.append(params or {})
            page = remaining.pop(0) if remaining else []
            payload: dict = {"value": page}
            if remaining:
                payload["@odata.nextLink"] = "https://graph.microsoft.com/v1.0/next"
            return payload
        if method == "GET":
            return {"totalItemCount": total}
        return {"id": "ok"}

    return _request, seen


def _filters(seen: list[dict]) -> list[str]:
    return [str(p.get("$filter", "")) for p in seen]


class TestServerSideDateBounds:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_received_after_becomes_a_ge_clause(self, MockClient):
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(
            action="mark_read", dry_run=True, received_after="2026-03-01T00:00:00Z"
        )
        assert "receivedDateTime ge 2026-03-01T00:00:00Z" in _filters(seen)[0]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_received_before_becomes_an_le_clause(self, MockClient):
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(
            action="mark_read", dry_run=True, received_before="2026-04-01T00:00:00Z"
        )
        assert "receivedDateTime le 2026-04-01T00:00:00Z" in _filters(seen)[0]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_both_bounds_combine(self, MockClient):
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(
            action="mark_read",
            dry_run=True,
            received_after="2026-03-01T00:00:00Z",
            received_before="2026-04-01T00:00:00Z",
        )
        first = _filters(seen)[0]
        assert "ge 2026-03-01T00:00:00Z" in first
        assert "le 2026-04-01T00:00:00Z" in first
        assert " and " in first

    @patch("msgraph_mcp.mail.GraphClient")
    def test_no_filter_when_no_bounds_are_given(self, MockClient):
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(action="mark_read", dry_run=True)
        assert seen[0].get("$filter") is None

    @patch("msgraph_mcp.mail.GraphClient")
    def test_the_lower_bound_survives_into_later_pages(self, MockClient):
        # The cursor rewrites the `le` half each page; dropping `ge` would make
        # page two silently scan past the window.
        pages = [
            [_msg(i, f"2026-03-{i:02d}T00:00:00Z") for i in range(20, 0, -1)],
            [_msg(99, "2026-02-01T00:00:00Z")],
        ]
        request, seen = _capture(pages=pages)
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(
            action="mark_read", dry_run=True, received_after="2026-03-01T00:00:00Z"
        )
        assert len(seen) >= 2
        assert all("ge 2026-03-01T00:00:00Z" in f for f in _filters(seen))

    @patch("msgraph_mcp.mail.GraphClient")
    def test_upper_bound_anchors_the_first_page_not_just_later_ones(self, MockClient):
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(
            action="mark_read", dry_run=True, received_before="2026-04-01T00:00:00Z"
        )
        # Exactly one `le` clause: the bound *is* the initial cursor, not an extra.
        assert _filters(seen)[0].count("le ") == 1


class TestValidation:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_malformed_received_before_fails_before_any_request(self, MockClient):
        with pytest.raises(ValueError, match="received_before"):
            bulk_manage_messages_multi_pass(
                action="mark_read", dry_run=True, received_before="last tuesday"
            )
        MockClient.return_value.request.assert_not_called()

    @patch("msgraph_mcp.mail.GraphClient")
    def test_inverted_window_is_rejected(self, MockClient):
        # Silently returning zero matches would read as "nothing there".
        with pytest.raises(ValueError, match="received_after"):
            bulk_manage_messages_multi_pass(
                action="mark_read",
                dry_run=True,
                received_after="2026-04-01T00:00:00Z",
                received_before="2026-03-01T00:00:00Z",
            )
        MockClient.return_value.request.assert_not_called()


class TestReporting:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_a_windowed_scan_reports_window_exhausted(self, MockClient):
        # "folder_exhausted" would overclaim: the folder was not scanned, the
        # window was -- which is still complete coverage of what was asked.
        request, _ = _capture()
        MockClient.return_value.request.side_effect = request
        result = bulk_manage_messages_multi_pass(
            action="mark_read", dry_run=True, received_after="2026-03-01T00:00:00Z"
        )
        assert result["stop_reason"] == "window_exhausted"
        assert result["truncated"] is False

    @patch("msgraph_mcp.mail.GraphClient")
    def test_an_unwindowed_scan_still_reports_folder_exhausted(self, MockClient):
        request, _ = _capture()
        MockClient.return_value.request.side_effect = request
        result = bulk_manage_messages_multi_pass(action="mark_read", dry_run=True)
        assert result["stop_reason"] == "folder_exhausted"


class TestToolSurface:
    @patch("msgraph_mcp.mail.bulk_manage_messages_multi_pass")
    def test_bulk_tool_forwards_received_before(self, bulk):
        from msgraph_mcp import tools

        tools.bulk_manage_messages(received_before="2026-04-01T00:00:00Z")
        assert bulk.call_args[1]["received_before"] == "2026-04-01T00:00:00Z"

    @patch("msgraph_mcp.mail.list_messages")
    def test_list_messages_tool_forwards_until(self, listed):
        from msgraph_mcp import tools

        listed.return_value = []
        tools.list_messages(until="2026-04-01T00:00:00Z")
        assert listed.call_args[1]["until"] == "2026-04-01T00:00:00Z"
