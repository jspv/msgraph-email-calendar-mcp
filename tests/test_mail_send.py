from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

from msgraph_mcp.errors import GraphRequestError
from msgraph_mcp.mail import (
    _build_recipients,
    _build_from,
    _draft_preview,
    send_message,
    reply_to_message,
    forward_message,
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


class TestReplyToMessage:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_dry_run_reply(self, MockClient):
        client = MockClient.return_value
        client.request.side_effect = [
            # createReply returns draft
            {
                "id": "reply-draft-1",
                "subject": "Re: Original",
                "from": {"emailAddress": {"address": "me@example.com"}},
                "toRecipients": [{"emailAddress": {"address": "sender@example.com"}}],
                "ccRecipients": [],
                "bccRecipients": [],
                "bodyPreview": "",
            },
            # PATCH to set body
            None,
            # GET refreshed draft
            {
                "id": "reply-draft-1",
                "subject": "Re: Original",
                "from": {"emailAddress": {"address": "me@example.com"}},
                "toRecipients": [{"emailAddress": {"address": "sender@example.com"}}],
                "ccRecipients": [],
                "bccRecipients": [],
                "bodyPreview": "My reply",
            },
            # DELETE draft
            None,
        ]
        result = reply_to_message(
            account_id=None,
            message_id="msg-1",
            body="My reply",
            dry_run=True,
        )
        assert result["dry_run"] is True
        assert result["preview"]["subject"] == "Re: Original"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_live_reply(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = reply_to_message(
            account_id=None,
            message_id="msg-1",
            body="My reply",
            dry_run=False,
        )
        assert result["ok"] is True
        call_args = client.request.call_args
        assert "/reply" in call_args[0][1]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_live_reply_all(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = reply_to_message(
            account_id=None,
            message_id="msg-1",
            body="My reply",
            reply_all=True,
            dry_run=False,
        )
        call_args = client.request.call_args
        assert "/replyAll" in call_args[0][1]


class TestForwardMessage:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_dry_run_forward(self, MockClient):
        client = MockClient.return_value
        client.request.side_effect = [
            # createForward returns draft
            {
                "id": "fwd-draft-1",
                "subject": "Fw: Original",
                "from": {"emailAddress": {"address": "me@example.com"}},
                "toRecipients": [],
                "ccRecipients": [],
                "bccRecipients": [],
                "bodyPreview": "",
            },
            # PATCH to set recipients + body
            None,
            # GET refreshed draft
            {
                "id": "fwd-draft-1",
                "subject": "Fw: Original",
                "from": {"emailAddress": {"address": "me@example.com"}},
                "toRecipients": [{"emailAddress": {"address": "someone@example.com"}}],
                "ccRecipients": [],
                "bccRecipients": [],
                "bodyPreview": "FYI",
            },
            # DELETE draft
            None,
        ]
        result = forward_message(
            account_id=None,
            message_id="msg-1",
            to=["someone@example.com"],
            body="FYI",
            dry_run=True,
        )
        assert result["dry_run"] is True
        assert result["preview"]["to_recipients"] == ["someone@example.com"]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_live_forward(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = forward_message(
            account_id=None,
            message_id="msg-1",
            to=["someone@example.com"],
            dry_run=False,
        )
        assert result["ok"] is True
        call_args = client.request.call_args
        assert "/forward" in call_args[0][1]
        body = call_args[1]["json_body"]
        assert body["toRecipients"][0]["emailAddress"]["address"] == "someone@example.com"


from msgraph_mcp.mail import create_draft, update_draft, send_draft


class TestCreateDraft:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_creates_draft(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "id": "draft-1",
            "subject": "Draft Subject",
            "from": {"emailAddress": {"address": "me@example.com"}},
            "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
            "ccRecipients": [],
            "bccRecipients": [],
            "bodyPreview": "Draft body",
        }
        result = create_draft(
            account_id=None,
            to=["to@example.com"],
            subject="Draft Subject",
            body="Draft body",
        )
        assert result["ok"] is True
        assert result["draft"]["id"] == "draft-1"
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/messages")


class TestUpdateDraft:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_updates_draft(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "id": "draft-1",
            "subject": "Updated Subject",
            "from": {"emailAddress": {"address": "me@example.com"}},
            "toRecipients": [{"emailAddress": {"address": "new@example.com"}}],
            "ccRecipients": [],
            "bccRecipients": [],
            "bodyPreview": "Updated body",
        }
        result = update_draft(
            account_id=None,
            message_id="draft-1",
            subject="Updated Subject",
            body="Updated body",
        )
        assert result["ok"] is True
        call_args = client.request.call_args
        assert call_args[0] == ("PATCH", "/me/messages/draft-1")

    @patch("msgraph_mcp.mail.GraphClient")
    def test_only_sends_provided_fields(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "id": "draft-1", "subject": "New Subject",
            "from": None, "toRecipients": [], "ccRecipients": [], "bccRecipients": [],
            "bodyPreview": "",
        }
        update_draft(account_id=None, message_id="draft-1", subject="New Subject")
        body = client.request.call_args[1]["json_body"]
        assert "subject" in body
        assert "body" not in body
        assert "toRecipients" not in body


class TestSendDraft:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_sends_draft(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = send_draft(account_id=None, message_id="draft-1")
        assert result["ok"] is True
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/messages/draft-1/send")


class TestDryRunDraftLifecycle:
    """The dry-run preview builds a real draft, so it must always clean it up.

    Both failure modes matter: a Graph response with no usable id should raise
    a typed error rather than KeyError, and a failure after the draft exists
    must not leave it sitting in the user's Drafts folder.
    """

    @patch("msgraph_mcp.mail.GraphClient")
    def test_send_raises_typed_error_when_draft_has_no_id(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {}  # Graph returned an empty body

        with pytest.raises(GraphRequestError, match="did not return a draft id"):
            send_message(
                account_id=None, to=["a@x.com"], subject="s", body="b", dry_run=True
            )

    @patch("msgraph_mcp.mail.GraphClient")
    def test_send_deletes_the_draft_when_preview_fails(self, MockClient):
        client = MockClient.return_value

        def _request(method, path, *, params=None, json_body=None):
            if method == "POST" and path == "/me/messages":
                return {"id": "draft-1"}
            if method == "DELETE":
                return None
            raise GraphRequestError("boom", status_code=500)

        client.request.side_effect = _request

        # Force a failure between create and delete by making the GET blow up.
        with patch("msgraph_mcp.mail._draft_preview", side_effect=RuntimeError("boom")):
            with pytest.raises(RuntimeError):
                send_message(
                    account_id=None, to=["a@x.com"], subject="s", body="b", dry_run=True
                )

        deletes = [
            c for c in client.request.call_args_list
            if c[0][0] == "DELETE" and c[0][1] == "/me/messages/draft-1"
        ]
        assert len(deletes) == 1, "dry-run draft was left behind"

    @patch("msgraph_mcp.mail.GraphClient")
    def test_reply_raises_typed_error_when_draft_has_no_id(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {}

        with pytest.raises(GraphRequestError, match="did not return a draft id"):
            reply_to_message(account_id=None, message_id="m1", body="b", dry_run=True)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_forward_raises_typed_error_when_draft_has_no_id(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {}

        with pytest.raises(GraphRequestError, match="did not return a draft id"):
            forward_message(account_id=None, message_id="m1", to=["a@x.com"], dry_run=True)

    @patch("msgraph_mcp.mail.GraphClient")
    def test_forward_deletes_the_draft_when_patch_fails(self, MockClient):
        client = MockClient.return_value

        def _request(method, path, *, params=None, json_body=None):
            if method == "POST" and path.endswith("/createForward"):
                return {"id": "draft-2"}
            if method == "PATCH":
                raise GraphRequestError("patch failed", status_code=500)
            return None

        client.request.side_effect = _request

        with pytest.raises(GraphRequestError, match="patch failed"):
            forward_message(account_id=None, message_id="m1", to=["a@x.com"], dry_run=True)

        deletes = [
            c for c in client.request.call_args_list
            if c[0][0] == "DELETE" and c[0][1] == "/me/messages/draft-2"
        ]
        assert len(deletes) == 1, "dry-run draft was left behind"
