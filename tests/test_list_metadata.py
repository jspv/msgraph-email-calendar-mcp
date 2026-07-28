"""Tests for PR2: conversation_id, since filter, fields override, job_title.

All additions are backward-compatible; these lock in the new behavior and the
guardrails on the caller-supplied `since` and `fields` inputs.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp import config, contacts, mail


def _raw(**over) -> dict:
    base = {
        "id": "m1",
        "subject": "s",
        "from": {"emailAddress": {"name": "N", "address": "a@x.com"}},
        "receivedDateTime": "2026-07-01T00:00:00Z",
        "isRead": False,
        "hasAttachments": False,
        "conversationId": "conv-1",
        "bodyPreview": "hi",
    }
    base.update(over)
    return base


class TestConversationId:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_list_messages_populates_conversation_id(self, MockClient):
        MockClient.return_value.paginate.return_value = [_raw()]
        out = mail.list_messages(None, "inbox", 5)
        assert out[0].conversation_id == "conv-1"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_default_select_includes_conversation_id(self, MockClient):
        client = MockClient.return_value
        client.paginate.return_value = []
        mail.list_messages(None, "inbox", 5)
        params = client.paginate.call_args.kwargs["params"]
        assert "conversationId" in params["$select"]


class TestSinceFilter:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_since_builds_server_side_filter(self, MockClient):
        client = MockClient.return_value
        client.paginate.return_value = []
        mail.list_messages(None, "inbox", 5, since="2026-07-01T00:00:00Z")
        params = client.paginate.call_args.kwargs["params"]
        assert params["$filter"] == "receivedDateTime ge 2026-07-01T00:00:00Z"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_no_filter_without_since(self, MockClient):
        client = MockClient.return_value
        client.paginate.return_value = []
        mail.list_messages(None, "inbox", 5)
        assert "$filter" not in client.paginate.call_args.kwargs["params"]

    def test_since_rejects_non_iso(self):
        with pytest.raises(ValueError, match="ISO-8601"):
            mail.list_messages(None, "inbox", 5, since="yesterday")


class TestFieldsOverride:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_fields_override_forces_id_first(self, MockClient):
        client = MockClient.return_value
        client.paginate.return_value = []
        mail.list_messages(None, "inbox", 5, fields=["subject", "receivedDateTime"])
        select = client.paginate.call_args.kwargs["params"]["$select"]
        assert select.split(",")[0] == "id"
        assert "subject" in select
        assert "bodyPreview" not in select  # not requested

    @patch("msgraph_mcp.mail.GraphClient")
    def test_fields_omitting_from_yields_null_sender(self, MockClient):
        MockClient.return_value.paginate.return_value = [
            {"id": "m1", "subject": "s", "receivedDateTime": "2026-07-01T00:00:00Z"}
        ]
        out = mail.list_messages(None, "inbox", 5, fields=["id", "subject", "receivedDateTime"])
        assert out[0].id == "m1"
        assert out[0].sender_label is None

    def test_fields_rejects_injection(self):
        with pytest.raises(ValueError, match="invalid field name"):
            mail.list_messages(None, "inbox", 5, fields=["id", "subject) or 1 eq 1"])


class TestJobTitle:
    @patch("msgraph_mcp.contacts.GraphClient")
    def test_search_people_returns_job_title(self, MockClient):
        MockClient.return_value.request.return_value = {
            "value": [
                {
                    "displayName": "Jane Doe",
                    "jobTitle": "CFO",
                    "scoredEmailAddresses": [{"address": "jane@x.com"}],
                }
            ]
        }
        out = contacts.search_people(None, query="jane")
        assert out[0].job_title == "CFO"
        assert out[0].email == "jane@x.com"

    @patch("msgraph_mcp.contacts.GraphClient")
    def test_search_people_selects_job_title(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"value": []}
        contacts.search_people(None, query="x")
        assert "jobTitle" in client.request.call_args.kwargs["params"]["$select"]


class TestMaxListLimit:
    def test_default_is_1000(self, monkeypatch):
        monkeypatch.delenv("MAX_LIST_LIMIT", raising=False)
        assert config.load_settings().max_list_limit == 1000

    def test_env_override(self, monkeypatch):
        monkeypatch.setenv("MAX_LIST_LIMIT", "250")
        assert config.load_settings().max_list_limit == 250
