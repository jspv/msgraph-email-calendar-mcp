"""Tests for bulk_manage_messages_multi_pass: collect-then-act and reporting.

These cover the two behaviours the rewrite guarantees:

* Collection and action are separate phases -- the scan never mutates while
  paginating (issue #1), so the live run acts on exactly the collected ids.
* The response distinguishes a truncated window from an exhausted folder and
  tolerates messages that vanish between collection and action (issue #2).
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp.errors import GraphRequestError
from msgraph_mcp.mail import bulk_manage_messages_multi_pass


def _msg(i: int, dt: str) -> dict:
    return {
        "id": f"m{i}",
        "subject": f"subject {i}",
        "from": {"emailAddress": {"name": f"Sender {i}", "address": f"a{i}@x.com"}},
        "receivedDateTime": dt,
        "isRead": False,
        "hasAttachments": False,
        "bodyPreview": "",
    }


def _fake_request(pages: list[list[dict]], *, total: int = 100, fail_move_ids=()):
    """Build a GraphClient.request side_effect and a call log.

    ``pages`` are returned in order for scan GETs; a folder-metadata GET returns
    ``total``; action calls (POST .../move) succeed unless the target id is in
    ``fail_move_ids``, in which case they raise a 404 GraphRequestError.
    """
    page_iter = iter(pages)
    calls: list[tuple[str, str]] = []

    def _request(method, path, *, params=None, json_body=None):
        calls.append((method, path))
        if method == "GET" and path.endswith("/messages"):
            return {"value": next(page_iter, [])}
        if method == "GET":  # folder metadata (/me/mailFolders/{id})
            return {"totalItemCount": total}
        # action: delete_message issues POST /me/messages/{id}/move
        for bad in fail_move_ids:
            if f"/me/messages/{bad}/" in path:
                raise GraphRequestError("Not found", status_code=404)
        return {"id": "moved"}

    return _request, calls


class TestReporting:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_dry_run_exhausted_folder(self, MockClient):
        # One short page (< default $top) => folder fully scanned.
        request, calls = _fake_request(
            [[_msg(1, "2026-01-02T00:00:00Z"), _msg(2, "2026-01-01T00:00:00Z")]],
            total=2,
        )
        MockClient.return_value.request.side_effect = request

        result = bulk_manage_messages_multi_pass(folder="inbox", action="delete", dry_run=True)

        assert result["dry_run"] is True
        assert result["scanned"] == 2
        assert result["matched"] == 2
        assert result["match_count"] == 2  # backwards-compatible alias
        assert result["truncated"] is False
        assert result["stop_reason"] == "folder_exhausted"
        assert result["total_in_folder"] == 2
        assert result["results"] is None
        assert len(result["matches"]) == 2
        # dry run performs no mutations: only GETs, no POST/DELETE.
        assert all(m == "GET" for m, _ in calls)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_truncated_when_max_passes_reached(self, MockClient):
        # Full pages with descending timestamps so the cursor keeps advancing.
        request, _ = _fake_request(
            [
                [_msg(1, "2026-01-04T00:00:00Z"), _msg(2, "2026-01-03T00:00:00Z")],
                [_msg(3, "2026-01-02T00:00:00Z"), _msg(4, "2026-01-01T00:00:00Z")],
            ],
            total=100,
        )
        MockClient.return_value.request.side_effect = request

        result = bulk_manage_messages_multi_pass(
            folder="inbox", action="delete", limit_per_pass=2, max_passes=1, dry_run=True
        )

        assert result["passes"] == 1
        assert result["scanned"] == 2
        assert result["truncated"] is True
        assert result["stop_reason"] == "max_passes_reached"


class TestCollectThenAct:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_live_run_acts_on_collected_ids_after_full_scan(self, MockClient):
        request, calls = _fake_request(
            [[_msg(1, "2026-01-02T00:00:00Z"), _msg(2, "2026-01-01T00:00:00Z")]],
            total=2,
        )
        MockClient.return_value.request.side_effect = request

        result = bulk_manage_messages_multi_pass(folder="inbox", action="delete", dry_run=False)

        assert result["acted"] == 2
        assert result["already_gone"] == 0
        assert result["failed"] == 0
        assert result["match_count"] == 2  # live-run alias counts actions

        # Issue #1 guarantee: every scan GET precedes the first mutation.
        first_action = next(i for i, (m, _) in enumerate(calls) if m != "GET")
        scan_gets = [
            i for i, (m, p) in enumerate(calls)
            if m == "GET" and p.endswith("/messages")
        ]
        assert max(scan_gets) < first_action

    @patch("msgraph_mcp.mail.GraphClient")
    def test_already_gone_is_not_an_error(self, MockClient):
        request, _ = _fake_request(
            [[_msg(1, "2026-01-02T00:00:00Z"), _msg(2, "2026-01-01T00:00:00Z")]],
            total=2,
            fail_move_ids=("m1",),  # rule moved/deleted m1 between scan and action
        )
        MockClient.return_value.request.side_effect = request

        result = bulk_manage_messages_multi_pass(folder="inbox", action="delete", dry_run=False)

        assert result["acted"] == 1
        assert result["already_gone"] == 1
        assert result["failed"] == 0
        gone = [r for r in result["results"] if r.get("status") == "already_gone"]
        assert len(gone) == 1
        assert gone[0]["message_id"] == "m1"


class TestValidation:
    def test_rejects_unknown_action(self):
        with pytest.raises(ValueError, match="action must be one of"):
            bulk_manage_messages_multi_pass(folder="inbox", action="archive", dry_run=True)

    def test_move_requires_destination(self):
        with pytest.raises(ValueError, match="destination is required"):
            bulk_manage_messages_multi_pass(folder="inbox", action="move", dry_run=True)
