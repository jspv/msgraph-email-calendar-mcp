from __future__ import annotations

import pytest

from msgraph_mcp.graph import GraphClient, validate_path_segment


class TestValidatePathSegment:
    def test_valid_base64_id(self):
        assert validate_path_segment("AAMkAGE1M2IyNGZi") == "AAMkAGE1M2IyNGZi"

    def test_valid_with_special_chars(self):
        assert validate_path_segment("abc-def_ghi=jkl+mno.pqr") == "abc-def_ghi=jkl+mno.pqr"

    def test_rejects_empty(self):
        with pytest.raises(ValueError, match="Invalid"):
            validate_path_segment("")

    def test_rejects_slash(self):
        with pytest.raises(ValueError, match="Invalid"):
            validate_path_segment("../../admin")

    def test_rejects_question_mark(self):
        with pytest.raises(ValueError, match="Invalid"):
            validate_path_segment("id?query=1")

    def test_rejects_hash(self):
        with pytest.raises(ValueError, match="Invalid"):
            validate_path_segment("id#fragment")

    def test_rejects_spaces(self):
        with pytest.raises(ValueError, match="Invalid"):
            validate_path_segment("id with spaces")

    def test_custom_label_in_error(self):
        with pytest.raises(ValueError, match="Invalid message_id"):
            validate_path_segment("bad/id", "message_id")


class TestNormalizePath:
    """Tests for GraphClient._normalize_path, including the pagination double-/v1.0 fix."""

    def _client(self) -> GraphClient:
        return GraphClient(account_id=None)

    def test_relative_path_passthrough(self):
        assert self._client()._normalize_path("/me/messages") == "/me/messages"

    def test_rejects_path_without_leading_slash(self):
        with pytest.raises(ValueError, match="must start with '/'"):
            self._client()._normalize_path("me/messages")

    def test_full_url_strips_base_path(self):
        """The critical pagination fix: nextLink URLs must not double /v1.0."""
        result = self._client()._normalize_path(
            "https://graph.microsoft.com/v1.0/me/messages?$skip=10"
        )
        assert result == "/me/messages?$skip=10"

    def test_full_url_without_query(self):
        result = self._client()._normalize_path(
            "https://graph.microsoft.com/v1.0/me/calendars"
        )
        assert result == "/me/calendars"

    def test_full_url_preserves_complex_query(self):
        result = self._client()._normalize_path(
            "https://graph.microsoft.com/v1.0/me/messages?$top=50&$skiptoken=abc123"
        )
        assert result == "/me/messages?$top=50&$skiptoken=abc123"

    def test_rejects_http_url(self):
        with pytest.raises(ValueError, match="Refusing to follow"):
            self._client()._normalize_path("http://graph.microsoft.com/v1.0/me/messages")

    def test_rejects_wrong_host(self):
        with pytest.raises(ValueError, match="Refusing to follow"):
            self._client()._normalize_path("https://evil.example.com/v1.0/me/messages")
