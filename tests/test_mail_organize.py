from __future__ import annotations

from unittest.mock import patch

from msgraph_mcp.mail import (
    create_folder,
    list_child_folders,
    flag_message,
    categorize_message,
    list_aliases,
)


class TestCreateFolder:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_create_root_folder(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "id": "folder-new",
            "displayName": "Reports",
            "totalItemCount": 0,
            "unreadItemCount": 0,
            "childFolderCount": 0,
        }
        result = create_folder(account_id=None, name="Reports")
        assert result["ok"] is True
        assert result["folder"]["display_name"] == "Reports"
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/mailFolders")

    @patch("msgraph_mcp.mail.GraphClient")
    def test_create_child_folder(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "id": "folder-child",
            "displayName": "Sub",
            "totalItemCount": 0,
            "unreadItemCount": 0,
            "childFolderCount": 0,
        }
        result = create_folder(account_id=None, name="Sub", parent_folder_id="parent-id")
        call_args = client.request.call_args
        assert "parent-id/childFolders" in call_args[0][1]


class TestListChildFolders:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_returns_child_folders(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "value": [
                {
                    "id": "child-1",
                    "displayName": "Sub A",
                    "totalItemCount": 5,
                    "unreadItemCount": 2,
                    "childFolderCount": 0,
                },
            ]
        }
        result = list_child_folders(account_id=None, folder_id="parent-1")
        assert len(result) == 1
        assert result[0].display_name == "Sub A"


class TestFlagMessage:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_flag(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = flag_message(account_id=None, message_id="msg-1", flag_status="flagged")
        assert result["ok"] is True
        body = client.request.call_args[1]["json_body"]
        assert body == {"flag": {"flagStatus": "flagged"}}

    def test_rejects_invalid_status(self):
        import pytest
        with pytest.raises(ValueError, match="flag_status"):
            flag_message(account_id=None, message_id="msg-1", flag_status="invalid")


class TestCategorizeMessage:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_set_categories(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = categorize_message(
            account_id=None,
            message_id="msg-1",
            categories=["Red category", "Blue category"],
        )
        assert result["ok"] is True
        body = client.request.call_args[1]["json_body"]
        assert body == {"categories": ["Red category", "Blue category"]}


class TestListAliases:
    @patch("msgraph_mcp.mail.GraphClient")
    def test_parses_proxy_addresses(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "mail": "primary@example.com",
            "userPrincipalName": "user@example.com",
            "proxyAddresses": [
                "SMTP:primary@example.com",
                "smtp:alias1@example.com",
                "smtp:alias2@example.com",
            ],
        }
        result = list_aliases(account_id=None)
        assert result["primary"] == "primary@example.com"
        assert "alias1@example.com" in result["aliases"]
        assert "alias2@example.com" in result["aliases"]
        assert "primary@example.com" in result["all_addresses"]

    @patch("msgraph_mcp.mail.GraphClient")
    def test_empty_proxy_addresses(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "mail": "user@example.com",
            "userPrincipalName": "user@example.com",
            "proxyAddresses": [],
        }
        result = list_aliases(account_id=None)
        assert result["primary"] == "user@example.com"
        assert result["aliases"] == []
