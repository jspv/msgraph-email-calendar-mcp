"""FastMCP tool definitions exposed to MCP clients."""

from __future__ import annotations

import os

from fastmcp import FastMCP

from . import auth, calendar, contacts, mail
from .config import settings

mcp = FastMCP("msgraph-mcp")


# ── Scope-based tool registration ──────────────────────────────────────
#
# Each capability maps to the Graph scopes that satisfy it. A tool is exposed
# to MCP clients only when the configured scope set (from MICROSOFT_SCOPES,
# via ``settings.scopes``) contains at least one acceptable scope. With the
# default full scope set every tool registers, so behavior is unchanged unless
# scopes are deliberately narrowed -- a read-only deployment simply never
# advertises actions its token could not perform. Auth tools are exempt and
# always register, so authentication remains possible before any scope exists.

_MAIL_READ = ("Mail.Read", "Mail.ReadWrite")
_MAIL_WRITE = ("Mail.ReadWrite",)
_MAIL_SEND = ("Mail.Send",)
_CALENDAR_READ = (
    "Calendars.Read",
    "Calendars.Read.Shared",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
)
_CALENDAR_WRITE = ("Calendars.ReadWrite", "Calendars.ReadWrite.Shared")
_USER = ("User.Read",)
_PEOPLE = ("People.Read",)


def _requires_scope(*alternatives: str):
    """Register the decorated function as a tool only if a required scope is set.

    Applies ``mcp.tool()`` when at least one of *alternatives* is present in
    ``settings.scopes``; otherwise returns the function unregistered so it is
    never surfaced to the model.
    """
    def decorator(func):
        if set(settings.scopes).intersection(alternatives):
            return mcp.tool()(func)
        return func
    return decorator


@mcp.tool()
def auth_status() -> dict:
    """Show current Microsoft auth configuration and cached accounts.

    When running inside the MCP Lambda wrapper framework, reports
    framework-managed authentication status instead of local MSAL state.
    """
    # Framework-managed auth (Lambda mode)
    if os.environ.get("OAUTH_AUTHENTICATED") == "true":
        user_id = os.environ.get("OAUTH_USER_ID", "unknown")
        return {
            "configured": True,
            "authenticated": True,
            "mode": "framework-managed",
            "accounts": [{"username": "framework-managed", "account_id": user_id}],
            "message": "Authenticated via MCP Lambda wrapper framework.",
        }
    auth_url = os.environ.get("OAUTH_AUTH_URL")
    if auth_url:
        return {
            "configured": True,
            "authenticated": False,
            "mode": "framework-managed",
            "auth_url_available": True,
            "message": "Not authenticated. Use start_auth to get the authentication URL.",
        }
    # Local MSAL mode (standalone / development)
    return auth.auth_status()


@mcp.tool()
def start_auth() -> dict:
    """Start Microsoft authentication.

    In Lambda mode, returns an OAuth authorization URL.
    In local mode, starts the device-code flow.
    """
    # Framework-managed auth (Lambda mode)
    auth_url = os.environ.get("OAUTH_AUTH_URL")
    if auth_url:
        return {
            "mode": "oauth",
            "auth_url": auth_url,
            "message": (
                "Open this URL in your browser to authenticate with Microsoft. "
                "Once complete, call auth_status to verify."
            ),
        }
    # Local MSAL device-code flow
    flow = auth.begin_device_flow()
    return {
        "mode": "device_code",
        "verification_uri": flow.verification_uri,
        "user_code": flow.user_code,
        "expires_in": flow.expires_in,
        "message": flow.message,
    }


@mcp.tool()
def finish_auth() -> dict:
    """Complete Microsoft authentication after user approval.

    In Lambda mode, this is not needed — call auth_status instead.
    In local mode, completes the device-code flow.
    """
    if os.environ.get("OAUTH_AUTHENTICATED") == "true":
        return {"message": "Already authenticated via framework."}
    if os.environ.get("OAUTH_AUTH_URL"):
        return {
            "message": (
                "Authentication is managed by the framework. After opening "
                "the auth URL and completing login, call auth_status to check."
            ),
        }
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



@_requires_scope(*_MAIL_READ)
def list_folders(
    account_id: str | None = None,
    include_hidden: bool = False,
    parent_folder_id: str | None = None,
) -> list[dict]:
    """List mail folders. Pass parent_folder_id to list subfolders (check child_folder_count to know if subfolders exist)."""
    if parent_folder_id:
        return [item.model_dump() for item in mail.list_child_folders(account_id, folder_id=parent_folder_id)]
    return [item.model_dump() for item in mail.list_folders(account_id, include_hidden)]


@_requires_scope(*_MAIL_READ)
def list_messages(
    account_id: str | None = None,
    folder: str = "inbox",
    limit: int = 10,
    since: str | None = None,
    until: str | None = None,
    flag_status: str | None = None,
    category: str | None = None,
    fields: list[str] | None = None,
) -> list[dict]:
    """List recent mail messages from a folder, newest first.

    Pass `since` and/or `until` (ISO-8601) to bound the window server-side;
    together they read one date slice instead of the newest N messages, which is
    how you reach older mail without paging everything above it. Pass `fields`
    to override the selected columns (`id` is always included) for a leaner or
    extended payload. Each summary includes `conversation_id` for threading, and
    `to_recipient_labels` / `cc_recipient_labels` so Sent Items rows are
    distinguishable and Inbox rows show which alias was addressed.

    Pass `flag_status` (`flagged` | `complete` | `notFlagged`) or `category` to
    filter server-side — this is how you answer "show me my flagged mail" or
    "everything I filed under Rent" in one request rather than a folder scan. Every summary also reports `flag_status` and
    `categories`, which were previously writable but not readable.
    """
    bounded_limit = max(1, min(limit, settings.max_list_limit))
    return [
        item.model_dump()
        for item in mail.list_messages(
            account_id, folder, bounded_limit, since=since, until=until,
            flag_status=flag_status, category=category, fields=fields
        )
    ]


@_requires_scope(*_MAIL_READ)
def get_message(message_id: str, account_id: str | None = None) -> dict:
    """Get a full read-only view of a specific mail message."""
    return mail.get_message(account_id, message_id).model_dump()


@_requires_scope(*_MAIL_WRITE)
def update_message(
    message_id: str,
    is_read: bool | None = None,
    flag_status: str | None = None,
    categories: list[str] | None = None,
    account_id: str | None = None,
) -> dict:
    """Update message properties: read/unread, follow-up flag (flagged/complete/notFlagged), or color categories."""
    results: dict = {"ok": True, "message_id": message_id, "updated": []}
    if is_read is not None:
        mail.mark_message_read(account_id, message_id, is_read)
        results["updated"].append("is_read")
        results["is_read"] = is_read
    if flag_status is not None:
        mail.flag_message(account_id, message_id=message_id, flag_status=flag_status)
        results["updated"].append("flag_status")
        results["flag_status"] = flag_status
    if categories is not None:
        mail.categorize_message(account_id, message_id=message_id, categories=categories)
        results["updated"].append("categories")
        results["categories"] = categories
    return results


@_requires_scope(*_MAIL_WRITE)
def move_message(message_id: str, destination: str, account_id: str | None = None) -> dict:
    """Move a message to another folder. Requires Mail.ReadWrite permission."""
    return mail.move_message(account_id, message_id, destination)


@_requires_scope(*_MAIL_WRITE)
def delete_message(message_id: str, account_id: str | None = None, permanent: bool = False) -> dict:
    """Delete a message or move it to Deleted Items. Requires Mail.ReadWrite permission."""
    return mail.delete_message(account_id, message_id, permanent)


@_requires_scope(*_MAIL_READ)
def search_messages(query: str, account_id: str | None = None, limit: int = 10) -> list[dict]:
    """Search messages using Microsoft Graph message search."""
    bounded_limit = max(1, min(limit, settings.max_list_limit))
    safe_query = query.strip()
    if not safe_query:
        raise ValueError("query must not be empty")
    return [item.model_dump() for item in mail.search_messages(account_id, safe_query, bounded_limit)]


@_requires_scope(*_MAIL_WRITE)
def bulk_manage_messages(
    account_id: str | None = None,
    folder: str = "inbox",
    sender_contains: str | None = None,
    subject_contains: str | None = None,
    recipient_contains: str | None = None,
    flag_status: str | None = None,
    category: str | None = None,
    received_after: str | None = None,
    received_before: str | None = None,
    unread_only: bool = False,
    action: str = "delete",
    destination: str | None = None,
    limit: int | None = None,
    dry_run: bool = True,
    confirm_token: str | None = None,
) -> dict:
    """Bulk-manage messages with filtering. Supports dry-run, delete, move, and mark-read actions.

    By default (``limit=None``) this scans the **entire** folder newest-first, so
    "find/act on all messages matching X" sees the whole folder, not a window.
    Pass ``limit`` only to cap how many messages are scanned (e.g. ``limit=200``
    scans at most the newest 200); the value is honored exactly, never silently
    clamped. Filtering is client-side.

    The response reports coverage honestly: ``truncated`` is False only when the
    scan reached the end of the folder (``stop_reason="folder_exhausted"``), so a
    match count is a true total; if ``limit`` cut the scan short it is True. The
    mailbox is live, so counts are a point-in-time snapshot. Collection and action
    are separate phases; messages moved or deleted between them are reported as
    ``already_gone``, not errors.

    ``delete`` and ``move`` are two-step: run with ``dry_run=True``, review the
    listed matches, then repeat with ``dry_run=False`` and the ``confirm_token``
    from that preview. The token is derived from the matched message ids, so if
    the mailbox changed in between the live run refuses and hands back a fresh
    token rather than acting on a set you never saw. ``mark_read`` /
    ``mark_unread`` are reversible and need no token.

    ``received_after`` / ``received_before`` are applied by Graph rather than
    after the fetch, so a date-scoped query reads only the window. Use them to
    reach old mail cheaply: without a bound, "what did this sender send me last
    March" pages the entire folder to match a handful of rows. When either is
    given, ``stop_reason`` is ``window_exhausted`` rather than
    ``folder_exhausted`` -- complete coverage of what was asked, not of the
    folder.

    ``recipient_contains`` matches across both To and Cc. It is what makes
    "everything sent to my <vendor> alias" a single call.

    ``category`` is applied by Graph, like the date bounds, so filtering a large
    folder by category does not page the whole thing. ``flag_status``
    (``flagged`` | ``complete`` | ``notFlagged``) is matched client-side instead:
    Exchange rejects a flag restriction combined with the ``receivedDateTime``
    sort this scan's cursor depends on.
    """
    if limit is not None and limit < 1:
        raise ValueError("limit must be a positive integer, or None to scan the whole folder")
    return mail.bulk_manage_messages_multi_pass(
        account_id=account_id,
        folder=folder,
        sender_contains=sender_contains,
        subject_contains=subject_contains,
        recipient_contains=recipient_contains,
        flag_status=flag_status,
        category=category,
        received_after=received_after,
        received_before=received_before,
        unread_only=unread_only,
        action=action,
        destination=destination,
        scan_limit=limit,
        dry_run=dry_run,
        confirm_token=confirm_token,
    )


@_requires_scope(*_CALENDAR_READ)
def list_calendars(account_id: str | None = None, user_id: str | None = None) -> list[dict]:
    """List readable calendars. Pass user_id for shared calendars (e.g. another user's email or object ID)."""
    return [item.model_dump() for item in calendar.list_calendars(account_id, user_id=user_id)]


@_requires_scope(*_CALENDAR_READ)
def list_events(
    account_id: str | None = None,
    start_iso: str | None = None,
    end_iso: str | None = None,
    calendar_id: str | None = None,
    limit: int = 25,
    user_id: str | None = None,
) -> list[dict]:
    """List events from a calendar in a time range. Pass user_id for shared calendars."""
    bounded_limit = max(1, min(limit, settings.max_event_limit))
    return [
        item.model_dump()
        for item in calendar.list_events(account_id, start_iso, end_iso, calendar_id, bounded_limit, user_id=user_id)
    ]


@_requires_scope(*_CALENDAR_READ)
def get_event(event_id: str, account_id: str | None = None, user_id: str | None = None) -> dict:
    """Get full details for a specific calendar event. Pass user_id for shared calendars."""
    return calendar.get_event(account_id, event_id, user_id=user_id).model_dump()


# ── Mail: Send ──────────────────────────────────────────────────────────


@_requires_scope(*_MAIL_SEND)
def send_message(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    send_as: str | None = None,
    dry_run: bool = True,
    account_id: str | None = None,
) -> dict:
    """Compose and send an email (no attachments). Defaults to dry-run (preview only). Set dry_run=False to actually send. To send with attachments, use create_draft + add_attachment_to_draft + manage_draft instead."""
    return mail.send_message(
        account_id, to=to, subject=subject, body=body,
        cc=cc, bcc=bcc, send_as=send_as, dry_run=dry_run,
    )


@_requires_scope(*_MAIL_SEND)
def reply_to_message(
    message_id: str,
    body: str,
    reply_all: bool = False,
    send_as: str | None = None,
    dry_run: bool = True,
    account_id: str | None = None,
) -> dict:
    """Reply to a message. Defaults to dry-run (preview only). Set dry_run=False to actually send."""
    return mail.reply_to_message(
        account_id, message_id=message_id, body=body,
        reply_all=reply_all, send_as=send_as, dry_run=dry_run,
    )


@_requires_scope(*_MAIL_SEND)
def forward_message(
    message_id: str,
    to: list[str],
    body: str | None = None,
    send_as: str | None = None,
    dry_run: bool = True,
    account_id: str | None = None,
) -> dict:
    """Forward a message. Defaults to dry-run (preview only). Set dry_run=False to actually send."""
    return mail.forward_message(
        account_id, message_id=message_id, to=to,
        body=body, send_as=send_as, dry_run=dry_run,
    )


@_requires_scope(*_MAIL_WRITE)
def create_draft(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    send_as: str | None = None,
    account_id: str | None = None,
) -> dict:
    """Create a draft email (saved to Drafts folder). Use this when you need to add attachments before sending, or want to build a message over multiple steps. Then use manage_draft to send."""
    return mail.create_draft(
        account_id, to=to, subject=subject, body=body,
        cc=cc, bcc=bcc, send_as=send_as,
    )


# ── Mail: Drafts ───────────────────────────────────────────────────────


@_requires_scope(*_MAIL_WRITE, *_MAIL_SEND)
def manage_draft(
    message_id: str,
    to: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    send_as: str | None = None,
    send: bool = False,
    account_id: str | None = None,
) -> dict:
    """Update and/or send a previously created draft. Provide fields to update, set send=True to send. Use after create_draft and optionally add_attachment_to_draft."""
    if any(v is not None for v in [to, subject, body, cc, bcc, send_as]):
        # This tool registers under either _MAIL_WRITE or _MAIL_SEND so that a
        # send-only deployment can still dispatch a draft. Editing one needs
        # write, though, so refuse here rather than issue a request the token
        # cannot satisfy.
        if not set(settings.scopes).intersection(_MAIL_WRITE):
            raise ValueError(
                "Updating a draft requires the Mail.ReadWrite scope; this server "
                "is configured for sending only. Call manage_draft with send=True "
                "and no update fields, or re-authenticate with Mail.ReadWrite."
            )
        result = mail.update_draft(
            account_id, message_id=message_id, to=to, subject=subject,
            body=body, cc=cc, bcc=bcc, send_as=send_as,
        )
    else:
        result = {"ok": True}
    if send:
        mail.send_draft(account_id, message_id=message_id)
        result["sent"] = True
        result["action"] = "sent"
    return result


# ── Mail: Attachments ──────────────────────────────────────────────────


@_requires_scope(*_MAIL_READ)
def get_attachments(
    message_id: str,
    attachment_id: str | None = None,
    account_id: str | None = None,
) -> dict | list[dict]:
    """Get attachments for a message. Without attachment_id: list all metadata. With attachment_id: download that attachment (base64 if under 1.5 MB)."""
    if attachment_id:
        return mail.get_attachment(account_id, message_id=message_id, attachment_id=attachment_id).model_dump()
    return [a.model_dump() for a in mail.list_attachments(account_id, message_id=message_id)]


@_requires_scope(*_MAIL_WRITE)
def add_attachment_to_draft(
    message_id: str,
    name: str,
    content_base64: str,
    content_type: str = "application/octet-stream",
    account_id: str | None = None,
) -> dict:
    """Attach a file (base64-encoded) to a draft message."""
    return mail.add_attachment_to_draft(
        account_id, message_id=message_id, name=name,
        content_base64=content_base64, content_type=content_type,
    )


# ── Mail: Organization ─────────────────────────────────────────────────


@_requires_scope(*_MAIL_WRITE)
def create_folder(
    name: str,
    parent_folder_id: str | None = None,
    account_id: str | None = None,
) -> dict:
    """Create a new mail folder, optionally under a parent folder."""
    return mail.create_folder(account_id, name=name, parent_folder_id=parent_folder_id)



# ── Mail: Aliases ──────────────────────────────────────────────────────


@_requires_scope(*_USER)
def list_aliases(account_id: str | None = None) -> dict:
    """List email aliases (send-from addresses) for the authenticated user."""
    return mail.list_aliases(account_id)


# ── Calendar: Write ────────────────────────────────────────────────────


@_requires_scope(*_CALENDAR_WRITE)
def create_event(
    subject: str,
    start_iso: str,
    end_iso: str,
    attendees: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_all_day: bool = False,
    calendar_id: str | None = None,
    account_id: str | None = None,
    user_id: str | None = None,
    timezone: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Create a calendar event. Dry-run by default. Pass user_id for shared calendars.

    Graph emails invitations the instant an event with attendees is created and
    there is no undo, so this previews by default. The preview costs no Graph
    call and echoes the exact body that would be sent, including times converted
    to UTC -- check those before setting ``dry_run=False``.

    Times without a UTC offset need a zone: pass ``timezone`` (IANA, e.g.
    ``"America/New_York"``) or set ``MSGRAPH_DEFAULT_TIMEZONE`` on the server.
    Otherwise the call is refused rather than guessing UTC, which would book the
    event at the wrong hour. A time that already carries an offset is used as-is.
    """
    return calendar.create_event(
        account_id, subject=subject, start_iso=start_iso, end_iso=end_iso,
        attendees=attendees, body=body, location=location,
        is_all_day=is_all_day, calendar_id=calendar_id, user_id=user_id,
        timezone=timezone, dry_run=dry_run,
    )


@_requires_scope(*_CALENDAR_WRITE)
def update_event(
    event_id: str,
    subject: str | None = None,
    start_iso: str | None = None,
    end_iso: str | None = None,
    attendees: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_all_day: bool | None = None,
    account_id: str | None = None,
    user_id: str | None = None,
    timezone: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Update a calendar event. Dry-run by default. Pass user_id for shared calendars.

    Edits notify attendees. The preview spends one read to show the event's
    current state alongside the proposed changes.

    Times without a UTC offset need a zone: pass ``timezone`` (IANA, e.g.
    ``"America/New_York"``) or set ``MSGRAPH_DEFAULT_TIMEZONE`` on the server.
    Otherwise the call is refused rather than guessing UTC, which would book the
    event at the wrong hour. A time that already carries an offset is used as-is.
    """
    return calendar.update_event(
        account_id, event_id=event_id, subject=subject,
        start_iso=start_iso, end_iso=end_iso, attendees=attendees,
        body=body, location=location, is_all_day=is_all_day, user_id=user_id,
        timezone=timezone, dry_run=dry_run,
    )


@_requires_scope(*_CALENDAR_WRITE)
def delete_event(
    event_id: str,
    cancel_message: str | None = None,
    account_id: str | None = None,
    user_id: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Delete or cancel an event. Dry-run by default. Pass user_id for shared calendars.

    The preview names the event -- subject, time, organizer, attendee count --
    rather than echoing an opaque id, and states which of the two very different
    outcomes applies: with ``cancel_message`` the event is cancelled and
    attendees are **notified**; without it the event is hard-deleted and
    **nobody is told**.
    """
    return calendar.delete_event(
        account_id, event_id=event_id, cancel_message=cancel_message,
        user_id=user_id, dry_run=dry_run,
    )


@_requires_scope(*_CALENDAR_WRITE)
def respond_to_event(
    event_id: str,
    response: str,
    message: str | None = None,
    account_id: str | None = None,
    user_id: str | None = None,
) -> dict:
    """Respond to a meeting invite. Pass user_id for shared calendars."""
    return calendar.respond_to_event(
        account_id, event_id=event_id, response=response, message=message, user_id=user_id,
    )


# ── Calendar: Scheduling ──────────────────────────────────────────────


@_requires_scope(*_CALENDAR_READ)
def check_availability(
    emails: list[str],
    start_iso: str,
    end_iso: str,
    mode: str = "free_busy",
    duration_minutes: int = 60,
    account_id: str | None = None,
    timezone: str | None = None,
) -> list[dict]:
    """Check calendar availability. mode='free_busy': get free/busy schedule. mode='suggest': suggest meeting times.

    Times without a UTC offset need a zone: pass ``timezone`` (IANA, e.g.
    ``"America/New_York"``) or set ``MSGRAPH_DEFAULT_TIMEZONE`` on the server.
    """
    if mode == "suggest":
        return [
            s.model_dump()
            for s in calendar.find_meeting_times(
                account_id, attendees=emails, duration_minutes=duration_minutes,
                start_iso=start_iso, end_iso=end_iso, timezone=timezone,
            )
        ]
    return [
        s.model_dump()
        for s in calendar.get_schedule(
            account_id, emails=emails, start_iso=start_iso, end_iso=end_iso,
            timezone=timezone,
        )
    ]


# ── People ─────────────────────────────────────────────────────────────


@_requires_scope(*_PEOPLE)
def search_people(query: str, limit: int = 10, account_id: str | None = None) -> list[dict]:
    """Search for people by name to find their email addresses."""
    bounded_limit = max(1, min(limit, settings.max_list_limit))
    return [p.model_dump() for p in contacts.search_people(account_id, query=query, limit=bounded_limit)]
