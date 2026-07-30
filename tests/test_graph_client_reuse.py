"""The HTTP connection pool is shared across every ``GraphClient``.

A new ``httpx.Client`` per request means a fresh TCP + TLS handshake per request.
The workload that makes this expensive is ``bulk_manage_messages_multi_pass``:
its action loop calls ``delete_message(account_id, item.id)`` per message, and
each of those constructs its *own* ``GraphClient``. Pooling per instance would
therefore buy nothing for the one case that needs it -- the pool has to outlive
the instance, which is why it lives at module scope.
"""
from __future__ import annotations

from unittest.mock import Mock, patch

import httpx
import pytest

from msgraph_mcp import graph
from msgraph_mcp.errors import GraphRequestError
from msgraph_mcp.graph import GraphClient


def _ok_response(payload: dict | None = None) -> Mock:
    response = Mock()
    response.status_code = 200
    response.content = b"{}"
    response.json.return_value = payload if payload is not None else {}
    response.raise_for_status.return_value = None
    return response


@pytest.fixture(autouse=True)
def _reset_pool():
    graph.close_http_client()
    yield
    graph.close_http_client()


class TestPooledClient:
    def test_repeated_calls_return_the_same_client(self):
        assert graph._http_client() is graph._http_client()

    def test_client_is_configured_with_the_settings_timeout(self):
        from msgraph_mcp.config import settings

        assert graph._http_client().timeout.read == settings.timeout_seconds

    def test_close_disposes_the_client_and_the_next_call_rebuilds(self):
        first = graph._http_client()
        graph.close_http_client()
        assert first.is_closed
        assert graph._http_client() is not first

    def test_rebuilds_if_the_client_was_closed_underneath_us(self):
        first = graph._http_client()
        first.close()
        second = graph._http_client()
        assert second is not first
        assert not second.is_closed


class TestGraphClientUsesThePool:
    @patch.object(GraphClient, "_headers", return_value={})
    def test_separate_graph_clients_share_one_connection_pool(self, _headers):
        pooled = graph._http_client()
        with patch.object(pooled, "request", return_value=_ok_response()) as request:
            GraphClient(account_id=None).request("GET", "/me/messages")
            GraphClient(account_id=None).request("GET", "/me/events")
        assert request.call_count == 2

    @patch.object(GraphClient, "_headers", return_value={})
    def test_pool_stays_open_after_a_request_completes(self, _headers):
        pooled = graph._http_client()
        with patch.object(pooled, "request", return_value=_ok_response()):
            GraphClient(account_id=None).request("GET", "/me/messages")
        assert not pooled.is_closed
        assert graph._http_client() is pooled

    @patch("msgraph_mcp.graph.time.sleep")  # don't pay the retry backoff
    @patch.object(GraphClient, "_headers", return_value={})
    def test_pool_stays_open_after_a_failed_request(self, _headers, _sleep):
        # A transport error must not poison the shared pool for later callers.
        pooled = graph._http_client()
        with patch.object(pooled, "request", side_effect=httpx.ConnectError("boom")):
            with pytest.raises(GraphRequestError):
                GraphClient(account_id=None).request("GET", "/me/messages")
        assert not pooled.is_closed
