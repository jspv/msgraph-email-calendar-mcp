"""FastMCP tool definitions exposed to MCP clients."""

from __future__ import annotations

from fastmcp import FastMCP

from . import auth, calendar, mail
from .config import settings

mcp = FastMCP("msgraph-mcp")


@mcp.tool()
def auth_status() -> dict:
    """Show current Microsoft auth configuration and cached accounts."""
    return auth.auth_status()


@mcp.tool()
def start_auth() -> dict:
    """Start Microsoft device-code authentication."""
    flow = auth.begin_device_flow()
    return {
        "verification_uri": flow.verification_uri,
        "user_code": flow.user_code,
        "expires_in": flow.expires_in,
        "message": flow.message,
    }


@mcp.tool()
def finish_auth() -> dict:
    """Complete Microsoft device-code authentication after user approval."""
    result = auth.complete_device_flow()
    claims = result.get("id_token_claims", {}) or {}
    return {
        "account": {
            "preferred_username": claims.get("preferred_username"),
            "name": claims.get("name"),
            "tenant_id": claims.get("tid"),
        },
        "scopes": result.get("scope"),
        "token_type": result.get("token_type"),
    }


@mcp.tool()
def list_accounts() -> list[dict]:
    """List authenticated Microsoft accounts available in local token cache."""
    return [account.model_dump() for account in auth.list_accounts()]


@mcp.tool()
def list_folders(account_id: str | None = None, include_hidden: bool = False) -> list[dict]:
    """List available mail folders for the authenticated account."""
    return [item.model_dump() for item in mail.list_folders(account_id, include_hidden)]


@mcp.tool()
def list_messages(account_id: str | None = None, folder: str = "inbox", limit: int = 10) -> list[dict]:
    """List recent mail messages from a folder."""
    bounded_limit = max(1, min(limit, settings.max_list_limit))
    return [item.model_dump() for item in mail.list_messages(account_id, folder, bounded_limit)]


@mcp.tool()
def get_message(message_id: str, account_id: str | None = None) -> dict:
    """Get a full read-only view of a specific mail message."""
    return mail.get_message(account_id, message_id).model_dump()


@mcp.tool()
def mark_message_read(message_id: str, account_id: str | None = None, is_read: bool = True) -> dict:
    """Mark a message as read or unread. Requires Mail.ReadWrite permission."""
    return mail.mark_message_read(account_id, message_id, is_read)


@mcp.tool()
def move_message(message_id: str, destination: str, account_id: str | None = None) -> dict:
    """Move a message to another folder. Requires Mail.ReadWrite permission."""
    return mail.move_message(account_id, message_id, destination)


@mcp.tool()
def delete_message(message_id: str, account_id: str | None = None, permanent: bool = False) -> dict:
    """Delete a message or move it to Deleted Items. Requires Mail.ReadWrite permission."""
    return mail.delete_message(account_id, message_id, permanent)


@mcp.tool()
def search_messages(query: str, account_id: str | None = None, limit: int = 10) -> list[dict]:
    """Search messages using Microsoft Graph message search."""
    bounded_limit = max(1, min(limit, settings.max_list_limit))
    safe_query = query.strip()
    if not safe_query:
        raise ValueError("query must not be empty")
    return [item.model_dump() for item in mail.search_messages(account_id, safe_query, bounded_limit)]


@mcp.tool()
def bulk_manage_messages(
    account_id: str | None = None,
    folder: str = "inbox",
    sender_contains: str | None = None,
    subject_contains: str | None = None,
    received_after: str | None = None,
    unread_only: bool = False,
    action: str = "delete",
    destination: str | None = None,
    limit: int = 50,
    max_passes: int = 5,
    dry_run: bool = True,
) -> dict:
    """Bulk-manage messages with filtering. Supports dry-run, delete, move, and mark-read actions."""
    bounded_limit = max(1, min(limit, settings.max_event_limit))
    bounded_passes = max(1, min(max_passes, 10))
    return mail.bulk_manage_messages_multi_pass(
        account_id=account_id,
        folder=folder,
        sender_contains=sender_contains,
        subject_contains=subject_contains,
        received_after=received_after,
        unread_only=unread_only,
        action=action,
        destination=destination,
        limit_per_pass=bounded_limit,
        max_passes=bounded_passes,
        dry_run=dry_run,
    )


@mcp.tool()
def list_calendars(account_id: str | None = None) -> list[dict]:
    """List readable calendars for the authenticated account."""
    return [item.model_dump() for item in calendar.list_calendars(account_id)]


@mcp.tool()
def list_events(
    account_id: str | None = None,
    start_iso: str | None = None,
    end_iso: str | None = None,
    calendar_id: str | None = None,
    limit: int = 25,
) -> list[dict]:
    """List events from the default or specified calendar in a time range."""
    bounded_limit = max(1, min(limit, settings.max_event_limit))
    return [
        item.model_dump()
        for item in calendar.list_events(account_id, start_iso, end_iso, calendar_id, bounded_limit)
    ]


@mcp.tool()
def get_event(event_id: str, account_id: str | None = None) -> dict:
    """Get full details for a specific calendar event."""
    return calendar.get_event(account_id, event_id).model_dump()
