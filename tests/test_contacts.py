from __future__ import annotations

from unittest.mock import patch

from msgraph_mcp.contacts import search_people


class TestSearchPeople:
    @patch("msgraph_mcp.contacts.GraphClient")
    def test_returns_results(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "value": [
                {
                    "displayName": "Alice Smith",
                    "scoredEmailAddresses": [
                        {"address": "alice@example.com", "relevanceScore": 10.0},
                    ],
                },
                {
                    "displayName": "Alice Jones",
                    "scoredEmailAddresses": [
                        {"address": "alice.j@example.com", "relevanceScore": 8.0},
                    ],
                },
            ]
        }
        result = search_people(account_id=None, query="Alice", limit=10)
        assert len(result) == 2
        assert result[0].name == "Alice Smith"
        assert result[0].email == "alice@example.com"
        assert result[1].name == "Alice Jones"

    @patch("msgraph_mcp.contacts.GraphClient")
    def test_empty_results(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"value": []}
        result = search_people(account_id=None, query="Nobody")
        assert result == []

    @patch("msgraph_mcp.contacts.GraphClient")
    def test_no_email(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "value": [
                {
                    "displayName": "No Email Person",
                    "scoredEmailAddresses": [],
                },
            ]
        }
        result = search_people(account_id=None, query="No Email")
        assert len(result) == 1
        assert result[0].email is None

    @patch("msgraph_mcp.contacts.GraphClient")
    def test_limit_respected(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"value": []}
        search_people(account_id=None, query="test", limit=5)
        call_args = client.request.call_args
        assert call_args[1]["params"]["$top"] == 5
