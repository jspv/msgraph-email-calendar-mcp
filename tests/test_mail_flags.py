"""Follow-up flags and categories are readable, not just writable.

``flag_message`` and ``categorize_message`` could set both, but no read path
returned either, so an agent could flag a message and then had no way to tell
which messages were flagged -- including one it had flagged itself moments
earlier. "Show me my flagged mail" was unanswerable.

The ``fields`` override was not a workaround: it widens ``$select``, and Graph
did return ``flag``, but the summary model dropped the value on the way back --
the same failure mode recipients had before #11.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp.mail import _SUMMARY_SELECT, _matches_filters, _message_summary
from msgraph_mcp.models import MailMessageSummary


def _item(flag_status: str | None = "flagged", categories=None, **overrides) -> dict:
    item = {
        "id": "m1",
        "subject": "Renew the lease",
        "from": {"emailAddress": {"name": "Landlord", "address": "landlord@example.com"}},
        "toRecipients": [],
        "ccRecipients": [],
        "receivedDateTime": "2026-01-02T00:00:00Z",
        "isRead": True,
        "hasAttachments": False,
        "bodyPreview": "",
        "categories": [] if categories is None else categories,
    }
    if flag_status is not None:
        item["flag"] = {"flagStatus": flag_status}
    item.update(overrides)
    return item


class TestSummarySelect:
    def test_flag_is_requested(self):
        assert "flag" in _SUMMARY_SELECT

    def test_categories_are_requested(self):
        assert "categories" in _SUMMARY_SELECT


class TestSummaryModel:
    def test_flag_status_is_surfaced(self):
        assert _message_summary(_item("flagged")).flag_status == "flagged"

    def test_completed_flag_is_distinguishable_from_flagged(self):
        # Three states, so a bool would lose "complete".
        assert _message_summary(_item("complete")).flag_status == "complete"

    def test_unflagged_is_reported_as_such(self):
        assert _message_summary(_item("notFlagged")).flag_status == "notFlagged"

    def test_missing_flag_object_is_none_not_a_crash(self):
        # A `fields` override that omits `flag` must not blow up the model.
        assert _message_summary(_item(flag_status=None)).flag_status is None

    def test_categories_are_surfaced(self):
        assert _message_summary(_item(categories=["Rent", "Urgent"])).categories == [
            "Rent",
            "Urgent",
        ]

    def test_absent_categories_are_empty_not_none(self):
        item = _item()
        del item["categories"]
        assert _message_summary(item).categories == []

    def test_flagged_shows_in_the_summary_string(self):
        assert "flagged" in _message_summary(_item("flagged")).summary

    def test_unflagged_does_not_clutter_the_summary(self):
        # Most mail is unflagged; saying so on every row is noise.
        assert "flag" not in _message_summary(_item("notFlagged")).summary.lower()

    def test_categories_show_in_the_summary_string(self):
        assert "Rent" in _message_summary(_item(categories=["Rent"])).summary


class TestFlagStatusFilter:
    def _summary(self, flag_status=None, categories=()):
        return MailMessageSummary(
            id="m1", flag_status=flag_status, categories=list(categories)
        )

    def test_matches_flagged(self):
        assert _matches_filters(self._summary("flagged"), flag_status="flagged")

    def test_rejects_a_different_status(self):
        assert not _matches_filters(self._summary("notFlagged"), flag_status="flagged")

    def test_can_select_completed_follow_ups(self):
        assert _matches_filters(self._summary("complete"), flag_status="complete")

    def test_rejects_when_status_is_unknown(self):
        assert not _matches_filters(self._summary(None), flag_status="flagged")

    def test_absent_filter_matches_everything(self):
        assert _matches_filters(self._summary("notFlagged"))


class TestCategoryFilter:
    def _summary(self, categories=()):
        return MailMessageSummary(id="m1", categories=list(categories))

    def test_matches_a_category(self):
        assert _matches_filters(self._summary(["Rent"]), category="Rent")

    def test_is_case_insensitive(self):
        assert _matches_filters(self._summary(["Rent"]), category="rent")

    def test_rejects_a_non_match(self):
        assert not _matches_filters(self._summary(["Rent"]), category="Travel")

    def test_rejects_when_there_are_no_categories(self):
        assert not _matches_filters(self._summary(), category="Rent")


class TestBulkValidation:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_invalid_flag_status_is_rejected_before_any_request(self, MockClient):
        from msgraph_mcp.mail import bulk_manage_messages_multi_pass

        with pytest.raises(ValueError, match="flag_status"):
            bulk_manage_messages_multi_pass(
                action="mark_read", dry_run=True, flag_status="starred"
            )
        MockClient.return_value.request.assert_not_called()


class TestServerSideFlagFilterOnList:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_list_messages_filters_flagged_server_side(self, MockClient):
        from msgraph_mcp.mail import list_messages

        MockClient.return_value.paginate.return_value = []
        list_messages(flag_status="flagged")
        params = MockClient.return_value.paginate.call_args[1]["params"]
        assert "flag/flagStatus eq 'flagged'" in params["$filter"]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_flag_filter_combines_with_a_date_bound(self, MockClient):
        from msgraph_mcp.mail import list_messages

        MockClient.return_value.paginate.return_value = []
        list_messages(since="2026-01-01T00:00:00Z", flag_status="flagged")
        f = MockClient.return_value.paginate.call_args[1]["params"]["$filter"]
        assert "receivedDateTime ge" in f and "flag/flagStatus" in f and " and " in f

    @patch("msgraph_mcp.mail.GraphClient")
    def test_invalid_flag_status_rejected(self, MockClient):
        from msgraph_mcp.mail import list_messages

        with pytest.raises(ValueError, match="flag_status"):
            list_messages(flag_status="starred")
        MockClient.return_value.paginate.assert_not_called()


class TestDetailPath:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_get_message_surfaces_flag_and_categories(self, MockClient):
        from msgraph_mcp.mail import get_message

        MockClient.return_value.request.return_value = {
            "id": "m1",
            "subject": "Renew the lease",
            "from": {"emailAddress": {"address": "landlord@example.com"}},
            "toRecipients": [],
            "ccRecipients": [],
            "receivedDateTime": "2026-01-02T00:00:00Z",
            "isRead": True,
            "hasAttachments": False,
            "flag": {"flagStatus": "flagged"},
            "categories": ["Rent"],
            "body": {"contentType": "text", "content": "hi"},
        }
        detail = get_message(None, "m1")
        assert detail.flag_status == "flagged"
        assert detail.categories == ["Rent"]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_get_message_requests_the_columns(self, MockClient):
        from msgraph_mcp.mail import get_message

        MockClient.return_value.request.return_value = {"id": "m1", "body": {}}
        get_message(None, "m1")
        select = MockClient.return_value.request.call_args[1]["params"]["$select"]
        assert "flag" in select and "categories" in select


class TestToolSurface:
    @patch("msgraph_mcp.mail.bulk_manage_messages_multi_pass")
    def test_bulk_tool_forwards_flag_status_and_category(self, bulk):
        from msgraph_mcp import tools

        tools.bulk_manage_messages(flag_status="flagged", category="Rent")
        assert bulk.call_args[1]["flag_status"] == "flagged"
        assert bulk.call_args[1]["category"] == "Rent"

    @patch("msgraph_mcp.mail.list_messages")
    def test_list_tool_forwards_flag_status(self, listed):
        from msgraph_mcp import tools

        listed.return_value = []
        tools.list_messages(flag_status="flagged")
        assert listed.call_args[1]["flag_status"] == "flagged"


class TestFlagFilterRespectsGraphsRestriction:
    """Graph rejects `$filter` on flag/flagStatus alongside `$orderby`.

    Verified against the live API, not inferred: every shape with an `$orderby`
    returns "The restriction or sort order is too complex for this operation",
    while the same filter without one succeeds -- including combined with a
    receivedDateTime clause. Mocked tests cannot catch this, so the constraint
    is pinned here explicitly.
    """

    @patch("msgraph_mcp.mail.GraphClient")
    def test_orderby_is_dropped_when_filtering_on_flag(self, MockClient):
        from msgraph_mcp.mail import list_messages

        MockClient.return_value.paginate.return_value = []
        list_messages(flag_status="flagged")
        params = MockClient.return_value.paginate.call_args[1]["params"]
        assert "$orderby" not in params

    @patch("msgraph_mcp.mail.GraphClient")
    def test_orderby_is_kept_when_not_filtering_on_flag(self, MockClient):
        from msgraph_mcp.mail import list_messages

        MockClient.return_value.paginate.return_value = []
        list_messages(since="2026-01-01T00:00:00Z")
        params = MockClient.return_value.paginate.call_args[1]["params"]
        assert params["$orderby"] == "receivedDateTime desc"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_results_are_still_newest_first(self, MockClient):
        # Graph will not sort these, so the ordering has to be restored here --
        # otherwise a flagged-mail listing comes back in arbitrary order.
        from msgraph_mcp.mail import list_messages

        def row(mid, dt):
            return {"id": mid, "subject": mid, "from": {"emailAddress": {}},
                    "receivedDateTime": dt, "isRead": True, "hasAttachments": False,
                    "flag": {"flagStatus": "flagged"}, "categories": [], "bodyPreview": ""}

        MockClient.return_value.paginate.return_value = [
            row("old", "2026-01-01T00:00:00Z"),
            row("new", "2026-06-01T00:00:00Z"),
            row("mid", "2026-03-01T00:00:00Z"),
        ]
        got = [m.id for m in list_messages(flag_status="flagged")]
        assert got == ["new", "mid", "old"]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_a_date_bound_still_travels_with_the_flag_filter(self, MockClient):
        # Confirmed live: flag + receivedDateTime in one $filter is accepted so
        # long as there is no $orderby, so the window still narrows server-side.
        from msgraph_mcp.mail import list_messages

        MockClient.return_value.paginate.return_value = []
        list_messages(since="2026-01-01T00:00:00Z", flag_status="flagged")
        f = MockClient.return_value.paginate.call_args[1]["params"]["$filter"]
        assert "receivedDateTime ge" in f and "flag/flagStatus" in f
