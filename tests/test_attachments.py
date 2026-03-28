from __future__ import annotations

from unittest.mock import patch

from msgraph_mcp.mail import list_attachments, get_attachment, add_attachment_to_draft


class TestListAttachments:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_returns_summaries(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "value": [
                {
                    "id": "att1",
                    "name": "doc.pdf",
                    "size": 1024,
                    "contentType": "application/pdf",
                    "isInline": False,
                },
                {
                    "id": "att2",
                    "name": "image.png",
                    "size": 2048,
                    "contentType": "image/png",
                    "isInline": True,
                },
            ]
        }
        result = list_attachments(account_id=None, message_id="msg-1")
        assert len(result) == 2
        assert result[0].name == "doc.pdf"
        assert result[0].size == 1024
        assert result[1].is_inline is True

    @patch("msgraph_mcp.mail.GraphClient")
    def test_empty_attachments(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"value": []}
        result = list_attachments(account_id=None, message_id="msg-1")
        assert result == []


class TestGetAttachment:
    @patch("msgraph_mcp.mail.settings")
    @patch("msgraph_mcp.mail.GraphClient")
    def test_inline_small_attachment(self, MockClient, mock_settings):
        mock_settings.max_attachment_inline_size = 1_572_864
        client = MockClient.return_value
        client.request.return_value = {
            "id": "att1",
            "name": "doc.pdf",
            "size": 1024,
            "contentType": "application/pdf",
            "isInline": False,
            "contentBytes": "dGVzdCBjb250ZW50",
        }
        result = get_attachment(account_id=None, message_id="msg-1", attachment_id="att1")
        assert result.content_base64 == "dGVzdCBjb250ZW50"
        assert result.content_omitted is False

    @patch("msgraph_mcp.mail.settings")
    @patch("msgraph_mcp.mail.GraphClient")
    def test_omits_large_attachment(self, MockClient, mock_settings):
        mock_settings.max_attachment_inline_size = 100  # Very small limit
        client = MockClient.return_value
        client.request.return_value = {
            "id": "att1",
            "name": "huge.zip",
            "size": 10_000_000,
            "contentType": "application/zip",
            "isInline": False,
            "contentBytes": "dGVzdA==",
        }
        result = get_attachment(account_id=None, message_id="msg-1", attachment_id="att1")
        assert result.content_omitted is True
        assert result.content_base64 is None
        assert "exceeds" in result.omit_reason


class TestAddAttachmentToDraft:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_adds_attachment(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "id": "att-new",
            "name": "report.pdf",
            "size": 512,
            "contentType": "application/pdf",
        }
        result = add_attachment_to_draft(
            account_id=None,
            message_id="draft-1",
            name="report.pdf",
            content_base64="dGVzdA==",
            content_type="application/pdf",
        )
        assert result["ok"] is True
        assert result["attachment_id"] == "att-new"
        call_args = client.request.call_args
        body = call_args[1]["json_body"]
        assert body["@odata.type"] == "#microsoft.graph.fileAttachment"
        assert body["name"] == "report.pdf"
        assert body["contentBytes"] == "dGVzdA=="
