"""Immutable ids (#14) and delta sync (#13).

Both probed live before implementing:

* ``Prefer: IdType="ImmutableId"`` is honoured -- Graph answers
  ``Preference-Applied: IdType=ImmutableId`` and returns a visibly different id.
* ``/messages/delta`` returns ``@odata.deltaLink`` on the terminating page,
  which ``paginate`` discards today.
"""
from __future__ import annotations

from dataclasses import replace
from unittest.mock import Mock, patch

import pytest

from msgraph_mcp import graph
from msgraph_mcp.errors import GraphRequestError
from msgraph_mcp.graph import GraphClient


def _resp(payload=None, status=200):
    r = Mock()
    r.status_code = status
    r.content = b"{}"
    r.json.return_value = payload or {}
    r.raise_for_status.return_value = None
    return r


# ------------------------------------------------------------- #14 headers


class TestPerRequestHeaders:
    @patch.object(GraphClient, "_headers", return_value={"Authorization": "Bearer x"})
    def test_extra_headers_are_merged_not_replaced(self, _h):
        with patch.object(graph, "_http_client") as pool:
            pool.return_value.request.return_value = _resp()
            GraphClient(None).request("GET", "/me/messages", headers={"X-Test": "1"})
        sent = pool.return_value.request.call_args[1]["headers"]
        assert sent["Authorization"] == "Bearer x"
        assert sent["X-Test"] == "1"

    @patch.object(GraphClient, "_headers", return_value={})
    def test_prefer_is_combined_rather_than_clobbered(self, _h):
        # Prefer is a comma-separated list. Assigning over it would silently drop
        # outlook.body-content-type on exactly the calls that fetch a body.
        with patch.object(graph, "_http_client") as pool:
            pool.return_value.request.return_value = _resp()
            GraphClient(None).request(
                "GET", "/me/messages",
                params={"$select": "body"},
                headers={"Prefer": 'IdType="ImmutableId"'},
            )
        prefer = pool.return_value.request.call_args[1]["headers"]["Prefer"]
        assert 'outlook.body-content-type="text"' in prefer
        assert 'IdType="ImmutableId"' in prefer
        assert prefer.count(",") == 1


class TestImmutableIdSetting:
    def test_off_by_default(self):
        from msgraph_mcp.config import Settings

        assert Settings(client_id="x").immutable_ids is False

    def _settings(self, on):
        # Settings is a frozen dataclass, so swap the whole object.
        return patch.object(graph, "settings", replace(graph.settings, immutable_ids=on))

    def test_header_absent_when_off(self):
        with self._settings(False):
            with patch("msgraph_mcp.graph.get_access_token", return_value="t"):
                assert "Prefer" not in GraphClient(None)._headers()

    def test_header_present_when_on(self):
        with self._settings(True):
            with patch("msgraph_mcp.graph.get_access_token", return_value="t"):
                assert GraphClient(None)._headers()["Prefer"] == 'IdType="ImmutableId"'

    def test_it_rides_on_continuation_pages_too(self):
        # paginate re-enters request with a full URL and no params, so anything
        # keyed off params would apply to page 1 only -- and page 2 would come
        # back with the other id type. Living in _headers is what avoids that.
        with self._settings(True):
            with patch("msgraph_mcp.graph.get_access_token", return_value="t"):
                h = GraphClient(None)._headers()
        assert 'IdType="ImmutableId"' in h["Prefer"]


# --------------------------------------------------------------- #13 delta


class TestPaginateDelta:
    @patch.object(GraphClient, "_headers", return_value={})
    def test_returns_rows_and_the_delta_token(self, _h):
        with patch.object(graph, "_http_client") as pool:
            pool.return_value.request.return_value = _resp(
                {"value": [{"id": "m1"}], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/x?$deltatoken=abc"}
            )
            rows, token = GraphClient(None).paginate_delta("/me/mailFolders/inbox/messages/delta")
        assert [r["id"] for r in rows] == ["m1"]
        assert token.endswith("$deltatoken=abc")

    @patch.object(GraphClient, "_headers", return_value={})
    def test_follows_nextlink_until_the_delta_link(self, _h):
        pages = [
            _resp({"value": [{"id": "m1"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/p2"}),
            _resp({"value": [{"id": "m2"}], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/d"}),
        ]
        with patch.object(graph, "_http_client") as pool:
            pool.return_value.request.side_effect = pages
            rows, token = GraphClient(None).paginate_delta("/me/mailFolders/inbox/messages/delta")
        assert [r["id"] for r in rows] == ["m1", "m2"]
        assert token is not None

    @patch.object(GraphClient, "_headers", return_value={})
    def test_limit_stops_early_without_a_token(self, _h):
        # Stopping short means the token would not cover the unread rows, so
        # returning one would silently skip them on the next sync.
        with patch.object(graph, "_http_client") as pool:
            pool.return_value.request.return_value = _resp(
                {"value": [{"id": "m1"}, {"id": "m2"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/p2"}
            )
            rows, token = GraphClient(None).paginate_delta("/me/x/delta", limit=1)
        assert len(rows) == 1
        assert token is None


class TestSyncMessages:
    def _payload(self, rows, delta="https://graph.microsoft.com/v1.0/d?$deltatoken=t2"):
        return {"value": rows, "@odata.deltaLink": delta}

    @patch("msgraph_mcp.mail.GraphClient")
    def test_returns_changes_and_a_token(self, MockClient):
        from msgraph_mcp.mail import sync_messages

        MockClient.return_value.paginate_delta.return_value = (
            [{"id": "m1", "subject": "hi", "from": {"emailAddress": {}},
              "receivedDateTime": "2026-01-01T00:00:00Z", "isRead": True,
              "hasAttachments": False, "categories": [], "bodyPreview": ""}],
            "tok",
        )
        out = sync_messages(folder="inbox")
        assert out["delta_token"] == "tok"
        assert out["changes"][0]["id"] == "m1"
        assert out["changes"][0]["removed"] is False

    @patch("msgraph_mcp.mail.GraphClient")
    def test_removals_are_represented_not_inferred(self, MockClient):
        from msgraph_mcp.mail import sync_messages

        MockClient.return_value.paginate_delta.return_value = (
            [{"id": "gone-1", "@removed": {"reason": "deleted"}}], "tok",
        )
        change = sync_messages(folder="inbox")["changes"][0]
        assert change["removed"] is True
        assert change["id"] == "gone-1"
        assert change["reason"] == "deleted"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_an_existing_token_is_resumed_rather_than_re_listing(self, MockClient):
        from msgraph_mcp.mail import sync_messages

        MockClient.return_value.paginate_delta.return_value = ([], "tok2")
        sync_messages(folder="inbox", delta_token="https://graph.microsoft.com/v1.0/d?$deltatoken=t1")
        path = MockClient.return_value.paginate_delta.call_args[0][0]
        assert "deltatoken=t1" in path

    @patch("msgraph_mcp.mail.GraphClient")
    def test_expired_token_surfaces_as_resync_not_a_generic_error(self, MockClient):
        from msgraph_mcp.mail import sync_messages

        MockClient.return_value.paginate_delta.side_effect = GraphRequestError(
            "gone", status_code=410
        )
        with pytest.raises(GraphRequestError, match="re-sync"):
            sync_messages(folder="inbox", delta_token="stale")

    @patch("msgraph_mcp.mail.GraphClient")
    def test_other_errors_are_not_reworded(self, MockClient):
        from msgraph_mcp.mail import sync_messages

        MockClient.return_value.paginate_delta.side_effect = GraphRequestError(
            "boom", status_code=500
        )
        with pytest.raises(GraphRequestError, match="boom"):
            sync_messages(folder="inbox")
