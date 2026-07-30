"""Tests for scope-based tool registration.

Tools are exposed to MCP clients only when a Graph scope that satisfies them is
present in ``settings.scopes``. These tests reload the ``tools`` module under
different scope sets and assert which tools register. The real settings are
restored afterwards so other test modules are unaffected.
"""
from __future__ import annotations

import asyncio
import importlib
from dataclasses import replace

from unittest.mock import patch

import pytest

from msgraph_mcp import config, tools

FULL_SCOPES = (
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "People.Read",
)
READ_ONLY_SCOPES = ("User.Read", "Mail.Read", "Calendars.Read")

ALWAYS_ON = {"auth_status", "start_auth", "finish_auth"}


def _registered_tool_names(scopes) -> set[str]:
    """Reload the tools module with *scopes* configured and return tool names."""
    original = config.settings
    config.settings = replace(original, scopes=tuple(scopes))
    try:
        importlib.reload(tools)
        return {t.name for t in asyncio.run(tools.mcp.list_tools())}
    finally:
        config.settings = original
        importlib.reload(tools)


@pytest.fixture(autouse=True)
def _restore_tools_module():
    # Guarantee the module is back to real settings even if a test asserts fail.
    yield
    importlib.reload(tools)


class TestScopeFiltering:
    def test_full_scopes_expose_everything(self):
        names = _registered_tool_names(FULL_SCOPES)
        # The full complement of tools (29 at time of writing) registers.
        assert len(names) >= 29
        for expected in ("send_message", "delete_message", "create_event", "search_people"):
            assert expected in names

    def test_read_only_hides_mutations(self):
        names = _registered_tool_names(READ_ONLY_SCOPES)
        # Read + auth tools present.
        for expected in ("list_messages", "search_messages", "get_message", "list_events"):
            assert expected in names
        # Write / send / people tools absent.
        for hidden in (
            "delete_message",
            "move_message",
            "update_message",
            "bulk_manage_messages",
            "send_message",
            "reply_to_message",
            "create_event",
            "delete_event",
            "search_people",
        ):
            assert hidden not in names

    def test_auth_tools_always_registered(self):
        # Even with no scopes at all, authentication must remain possible.
        names = _registered_tool_names(())
        assert ALWAYS_ON.issubset(names)

    def test_readwrite_satisfies_read_requirement(self):
        # Mail.ReadWrite alone should expose read tools (which accept it).
        names = _registered_tool_names(("Mail.ReadWrite",))
        assert "list_messages" in names
        assert "delete_message" in names
        # No calendar/people scope -> those stay hidden.
        assert "list_events" not in names
        assert "search_people" not in names

    def test_send_scope_exposes_send_only(self):
        names = _registered_tool_names(("Mail.Send",))
        assert "send_message" in names
        assert "reply_to_message" in names
        # Sending does not grant read or delete.
        assert "list_messages" not in names
        assert "delete_message" not in names

    def test_people_scope_gates_search_people(self):
        assert "search_people" in _registered_tool_names(("People.Read",))
        assert "search_people" not in _registered_tool_names(("Mail.ReadWrite",))


class TestManageDraftScopeGuard:
    """`manage_draft` spans two scopes: updating needs write, sending needs send.

    It stays registered under either scope so a send-only deployment can still
    dispatch a draft, but the update path has to refuse rather than issue a
    request the token cannot satisfy.
    """

    def _call_manage_draft(self, scopes, **kwargs):
        original = config.settings
        config.settings = replace(original, scopes=tuple(scopes))
        try:
            importlib.reload(tools)
            return tools.manage_draft(**kwargs)
        finally:
            config.settings = original
            importlib.reload(tools)

    def test_registered_under_send_only(self):
        assert "manage_draft" in _registered_tool_names(("Mail.Send",))

    def test_registered_under_write_only(self):
        assert "manage_draft" in _registered_tool_names(("Mail.ReadWrite",))

    def test_update_refused_without_write_scope(self):
        # Patched so a regressed guard fails the assertion rather than firing a
        # real Graph request.
        with patch("msgraph_mcp.mail.update_draft") as update:
            with pytest.raises(ValueError, match="Mail.ReadWrite"):
                self._call_manage_draft(
                    ("Mail.Send",), message_id="draft-1", subject="new subject"
                )
        update.assert_not_called()

    def test_update_allowed_with_write_scope(self):
        with patch("msgraph_mcp.mail.update_draft", return_value={"ok": True}) as update:
            self._call_manage_draft(
                ("Mail.ReadWrite",), message_id="draft-1", subject="new subject"
            )
        update.assert_called_once()

    def test_send_only_scope_can_still_send(self):
        with patch("msgraph_mcp.mail.send_draft", return_value={"ok": True}) as send:
            result = self._call_manage_draft(
                ("Mail.Send",), message_id="draft-1", send=True
            )
        assert result["sent"] is True
        send.assert_called_once()
