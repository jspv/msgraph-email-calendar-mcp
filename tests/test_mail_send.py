from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

from msgraph_mcp.mail import (
    _build_recipients,
    _build_from,
    _draft_preview,
    send_message,
)


class TestBuildRecipients:
    def test_single_email(self):
        result = _build_recipients(["alice@example.com"])
        assert result == [{"emailAddress": {"address": "alice@example.com"}}]

    def test_multiple_emails(self):
        result = _build_recipients(["a@x.com", "b@x.com"])
        assert len(result) == 2
        assert result[1]["emailAddress"]["address"] == "b@x.com"

    def test_empty_list(self):
        assert _build_recipients([]) == []

    def test_none(self):
        assert _build_recipients(None) == []


class TestBuildFrom:
    def test_none_returns_none(self):
        assert _build_from(None) is None

    def test_email_string(self):
        result = _build_from("alias@example.com")
        assert result == {"emailAddress": {"address": "alias@example.com"}}


class TestDraftPreview:
    def test_extracts_fields(self):
        draft_payload = {
            "id": "draft-123",
            "subject": "Hello",
            "from": {"emailAddress": {"address": "me@example.com"}},
            "toRecipients": [{"emailAddress": {"address": "you@example.com"}}],
            "ccRecipients": [],
            "bccRecipients": [],
            "bodyPreview": "Hi there",
        }
        preview = _draft_preview(draft_payload)
        assert preview.id == "draft-123"
        assert preview.subject == "Hello"
        assert preview.from_address == "me@example.com"
        assert preview.to_recipients == ["you@example.com"]
        assert preview.body_preview == "Hi there"


class TestSendMessage:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_dry_run_creates_and_deletes_draft(self, MockClient):
        client = MockClient.return_value
        client.request.side_effect = [
            # First call: POST create draft
            {
                "id": "draft-1",
                "subject": "Test",
                "from": {"emailAddress": {"address": "me@example.com"}},
                "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
                "ccRecipients": [],
                "bccRecipients": [],
                "bodyPreview": "Body text",
            },
            # Second call: DELETE draft
            None,
        ]
        result = send_message(
            account_id=None,
            to=["to@example.com"],
            subject="Test",
            body="Body text",
            dry_run=True,
        )
        assert result["dry_run"] is True
        assert result["preview"]["subject"] == "Test"
        # Verify draft was created then deleted
        calls = client.request.call_args_list
        assert calls[0][0][0] == "POST"  # create draft
        assert calls[1][0][0] == "DELETE"  # cleanup

    @patch("msgraph_mcp.mail.GraphClient")
    def test_live_send(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None  # sendMail returns empty
        result = send_message(
            account_id=None,
            to=["to@example.com"],
            subject="Test",
            body="Body text",
            dry_run=False,
        )
        assert result["ok"] is True
        assert result["dry_run"] is False
        call_args = client.request.call_args
        assert call_args[0][0] == "POST"
        assert call_args[0][1] == "/me/sendMail"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_send_with_send_as(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = send_message(
            account_id=None,
            to=["to@example.com"],
            subject="Test",
            body="Body",
            send_as="alias@example.com",
            dry_run=False,
        )
        call_args = client.request.call_args
        body = call_args[1]["json_body"]
        assert body["message"]["from"]["emailAddress"]["address"] == "alias@example.com"
