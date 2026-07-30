"""Tests for bulk_manage_messages_multi_pass: collect-then-act and reporting.

These cover the two behaviours the rewrite guarantees:

* Collection and action are separate phases -- the scan never mutates while
  paginating (issue #1), so the live run acts on exactly the collected ids.
* The response distinguishes a truncated window from an exhausted folder and
  tolerates messages that vanish between collection and action (issue #2).
"""
from __future__ import annotations

from datetime import datetime

from unittest.mock import patch

import pytest

from msgraph_mcp.errors import GraphRequestError
from msgraph_mcp.mail import _matches_filters, bulk_manage_messages_multi_pass
from msgraph_mcp.models import MailMessageSummary


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
    def test_truncated_when_scan_limit_reached(self, MockClient):
        # A folder with more than scan_limit messages: stop at the cap and
        # report truncated, since more may exist deeper.
        request, _ = _fake_request(
            [
                [_msg(1, "2026-01-04T00:00:00Z"), _msg(2, "2026-01-03T00:00:00Z")],
                [_msg(3, "2026-01-02T00:00:00Z"), _msg(4, "2026-01-01T00:00:00Z")],
            ],
            total=100,
        )
        MockClient.return_value.request.side_effect = request

        result = bulk_manage_messages_multi_pass(
            folder="inbox", action="delete", scan_limit=2, dry_run=True
        )

        assert result["scanned"] == 2
        assert result["truncated"] is True
        assert result["stop_reason"] == "scan_limit_reached"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_default_scans_whole_folder(self, MockClient):
        # No scan_limit: walk to the end across multiple pages. The last page is
        # short, signalling the folder end -> not truncated, true total.
        request, _ = _fake_request(
            [
                [_msg(i, f"2026-03-01T{i // 60:02d}:{i % 60:02d}:00Z") for i in range(1000)],
                [_msg(1000 + i, f"2026-01-01T00:00:0{i}Z") for i in range(3)],
            ],
            total=1003,
        )
        MockClient.return_value.request.side_effect = request

        result = bulk_manage_messages_multi_pass(folder="inbox", action="delete", dry_run=True)

        assert result["scanned"] == 1003
        assert result["matched"] == 1003
        assert result["truncated"] is False
        assert result["stop_reason"] == "folder_exhausted"
        assert result["passes"] == 2


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


class TestShortPageWithNextLink:
    """A short page is only the folder end when Graph says there is no more.

    Graph may return fewer items than ``$top`` while still emitting
    ``@odata.nextLink``. Treating that as the folder end would report
    ``truncated=False`` on an incomplete scan -- the one field callers are
    told they can trust as a true folder total.
    """

    def _paged_request(self, pages: list[dict]):
        page_iter = iter(pages)

        def _request(method, path, *, params=None, json_body=None):
            if method == "GET" and path.endswith("/messages"):
                return next(page_iter, {"value": []})
            if method == "GET":
                return {"totalItemCount": 3}
            return {"id": "moved"}

        return _request

    @patch("msgraph_mcp.mail.GraphClient")
    def test_scan_continues_past_a_short_page_carrying_next_link(self, MockClient):
        MockClient.return_value.request.side_effect = self._paged_request([
            {
                "value": [_msg(1, "2026-01-03T00:00:00Z")],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
            },
            {"value": [_msg(2, "2026-01-02T00:00:00Z"), _msg(3, "2026-01-01T00:00:00Z")]},
        ])

        result = bulk_manage_messages_multi_pass(folder="inbox", action="delete", dry_run=True)

        assert result["scanned"] == 3
        assert result["passes"] == 2
        assert result["truncated"] is False
        assert result["stop_reason"] == "folder_exhausted"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_empty_page_with_next_link_stalls_rather_than_claiming_exhausted(self, MockClient):
        # Nothing to advance the cursor with, but Graph says more exists:
        # report the stall honestly instead of a false folder total.
        MockClient.return_value.request.side_effect = self._paged_request([
            {
                "value": [],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
            },
        ])

        result = bulk_manage_messages_multi_pass(folder="inbox", action="delete", dry_run=True)

        assert result["truncated"] is True
        assert result["stop_reason"] == "cursor_stalled"


class TestReceivedAfterFilter:
    """`received_after` accepts the date forms a caller actually types.

    Graph always returns tz-aware ``receivedDateTime``; a bare date parses
    naive, so the comparison has to normalise both sides rather than trusting
    the caller to supply an offset.
    """

    def _summary(self, received: str = "2026-07-29T12:00:00Z") -> MailMessageSummary:
        return MailMessageSummary(id="m1", subject="hi", received_datetime=received)

    def test_bare_date_is_treated_as_utc_midnight(self):
        assert _matches_filters(self._summary(), received_after="2026-01-01") is True

    def test_bare_date_excludes_older_messages(self):
        older = self._summary("2025-12-31T23:00:00Z")
        assert _matches_filters(older, received_after="2026-01-01") is False

    def test_offset_aware_cutoff_still_works(self):
        # 2026-07-29T06:00-07:00 == 13:00Z, so a 12:00Z message is older.
        message = self._summary("2026-07-29T12:00:00Z")
        assert _matches_filters(message, received_after="2026-07-29T06:00:00-07:00") is False

    @patch("msgraph_mcp.mail.GraphClient")
    def test_bulk_scan_accepts_a_bare_date(self, MockClient):
        request, _ = _fake_request(
            [[_msg(1, "2026-01-02T00:00:00Z"), _msg(2, "2025-06-01T00:00:00Z")]],
            total=2,
        )
        MockClient.return_value.request.side_effect = request

        result = bulk_manage_messages_multi_pass(
            folder="inbox", action="delete", received_after="2026-01-01", dry_run=True
        )

        assert result["scanned"] == 2
        assert result["matched"] == 1

    def test_naive_datetime_cutoff_is_normalised(self):
        # A caller passing a pre-parsed datetime can hand over a naive one;
        # it must not reach the comparison unnormalised.
        cutoff = datetime(2026, 1, 1)
        assert _matches_filters(self._summary(), received_after=cutoff) is True

    def test_rejects_malformed_cutoff_up_front(self):
        with pytest.raises(ValueError, match="received_after must be an ISO-8601 datetime"):
            bulk_manage_messages_multi_pass(
                folder="inbox", action="delete", received_after="last tuesday", dry_run=True
            )


class TestValidation:
    def test_rejects_unknown_action(self):
        with pytest.raises(ValueError, match="action must be one of"):
            bulk_manage_messages_multi_pass(folder="inbox", action="archive", dry_run=True)

    def test_move_requires_destination(self):
        with pytest.raises(ValueError, match="destination is required"):
            bulk_manage_messages_multi_pass(folder="inbox", action="move", dry_run=True)

    def test_rejects_non_positive_scan_limit(self):
        with pytest.raises(ValueError, match="scan_limit must be a positive integer"):
            bulk_manage_messages_multi_pass(folder="inbox", action="delete", scan_limit=0, dry_run=True)
