"""Category filtering is server-side (issue: flags/categories were write-only).

Probed against the live API before implementing, because a mocked test proves
only that a ``$filter`` string was built, not that Graph accepts it -- the gap
that let a broken flag filter ship.

What the probe established, with a real message tagged and a negative control:

* ``categories/any(c:c eq 'X')`` matches correctly, and a non-matching category
  returns nothing.
* Unlike ``flag/flagStatus``, it is accepted **alongside** ``$orderby``. So
  category filtering keeps newest-first ordering *and* can be pushed into the
  bulk scan, whose cursor pagination depends on that sort.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp.mail import _odata_string, bulk_manage_messages_multi_pass, list_messages

from tests.test_bulk_date_window import _capture, _filters


class TestODataStringEscaping:
    """Category names are free text going straight into a ``$filter``."""

    def test_plain_value_is_quoted(self):
        assert _odata_string("Rent") == "'Rent'"

    def test_single_quote_is_doubled(self):
        # OData escapes ' by doubling it. Without this, "Bob's" terminates the
        # literal early and the rest of the name is parsed as operators.
        assert _odata_string("Bob's stuff") == "'Bob''s stuff'"

    def test_quote_only_value(self):
        assert _odata_string("'") == "''''"

    def test_attempted_injection_is_neutralised(self):
        got = _odata_string("x' or startswith(subject,'a")
        assert got.startswith("'") and got.endswith("'")
        assert got.count("'") % 2 == 0


class TestListMessagesCategoryFilter:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_category_filters_server_side(self, MockClient):
        MockClient.return_value.paginate.return_value = []
        list_messages(category="Rent")
        params = MockClient.return_value.paginate.call_args[1]["params"]
        assert "categories/any(c:c eq 'Rent')" in params["$filter"]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_orderby_is_preserved_unlike_the_flag_filter(self, MockClient):
        # Confirmed live: Exchange accepts a category restriction with a sort.
        MockClient.return_value.paginate.return_value = []
        list_messages(category="Rent")
        params = MockClient.return_value.paginate.call_args[1]["params"]
        assert params["$orderby"] == "receivedDateTime desc"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_combines_with_a_date_bound(self, MockClient):
        MockClient.return_value.paginate.return_value = []
        list_messages(since="2026-01-01T00:00:00Z", category="Rent")
        f = MockClient.return_value.paginate.call_args[1]["params"]["$filter"]
        assert "receivedDateTime ge" in f and "categories/any" in f

    @patch("msgraph_mcp.mail.GraphClient")
    def test_a_flag_filter_still_forces_the_sort_off(self, MockClient):
        # Category tolerates the sort, flag does not; combined, flag wins.
        MockClient.return_value.paginate.return_value = []
        list_messages(category="Rent", flag_status="flagged")
        params = MockClient.return_value.paginate.call_args[1]["params"]
        assert "$orderby" not in params

    @patch("msgraph_mcp.mail.GraphClient")
    def test_quotes_in_the_name_are_escaped(self, MockClient):
        MockClient.return_value.paginate.return_value = []
        list_messages(category="Bob's stuff")
        f = MockClient.return_value.paginate.call_args[1]["params"]["$filter"]
        assert "'Bob''s stuff'" in f


class TestBulkCategoryFilter:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_category_is_pushed_into_the_scan_filter(self, MockClient):
        # The whole point: "everything categorised X" should not page a folder.
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(action="mark_read", dry_run=True, category="Rent")
        assert "categories/any(c:c eq 'Rent')" in _filters(seen)[0]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_it_survives_into_later_pages(self, MockClient):
        from tests.test_bulk_manage import _msg

        pages = [[_msg(i, f"2026-03-{i:02d}T00:00:00Z") for i in range(20, 0, -1)],
                 [_msg(99, "2026-02-01T00:00:00Z")]]
        request, seen = _capture(pages=pages)
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(action="mark_read", dry_run=True, category="Rent")
        assert len(seen) >= 2
        assert all("categories/any" in f for f in _filters(seen))

    @patch("msgraph_mcp.mail.GraphClient")
    def test_it_combines_with_a_date_window(self, MockClient):
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(
            action="mark_read", dry_run=True, category="Rent",
            received_after="2026-03-01T00:00:00Z",
        )
        f = _filters(seen)[0]
        assert "categories/any" in f and "receivedDateTime ge" in f

    @patch("msgraph_mcp.mail.GraphClient")
    def test_the_sort_is_kept_so_cursor_paging_still_works(self, MockClient):
        # Bulk paging anchors on receivedDateTime desc; losing it would break
        # the cursor entirely, which is why flags stay client-side here.
        request, seen = _capture()
        MockClient.return_value.request.side_effect = request
        bulk_manage_messages_multi_pass(action="mark_read", dry_run=True, category="Rent")
        assert seen[0]["$orderby"] == "receivedDateTime desc"
