"""Pydantic models and formatting helpers for Graph API responses.

Models carry both raw Graph fields and pre-formatted ``*_label`` /
``summary`` fields so downstream MCP consumers do less formatting work.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


def _safe_parse_datetime(value: str | None) -> datetime | None:
    """Parse an ISO-8601 datetime string, returning ``None`` on failure."""
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def _format_datetime_label(value: str | None) -> str | None:
    """Format an ISO-8601 string as ``YYYY-MM-DD HH:MM TZ``."""
    parsed = _safe_parse_datetime(value)
    if not parsed:
        return value
    return parsed.strftime("%Y-%m-%d %H:%M %Z")


def _clean_text_snippet(value: str | None, max_len: int = 240) -> str | None:
    """Collapse whitespace and truncate to *max_len* characters."""
    if value is None:
        return None
    cleaned = " ".join(value.split())
    if not cleaned:
        return None
    if len(cleaned) <= max_len:
        return cleaned
    return cleaned[: max_len - 1].rstrip() + "…"


def _address_label(person: dict[str, Any] | None) -> str | None:
    """Format a Graph ``emailAddress`` object as ``Name <address>``."""
    email = (person or {}).get("emailAddress") or {}
    name = email.get("name")
    address = email.get("address")
    if name and address:
        return f"{name} <{address}>"
    return name or address


def _recipient_labels(items: list[dict[str, Any]]) -> list[str]:
    """Format a list of Graph recipient objects into display labels."""
    labels: list[str] = []
    for item in items:
        label = _address_label(item)
        if label:
            labels.append(label)
    return labels


def _location_label(location: dict[str, Any] | None) -> str | None:
    """Extract the display name from a Graph location object."""
    if not location:
        return None
    return location.get("displayName") or None


def _event_time_label(start: dict[str, Any] | None, end: dict[str, Any] | None, is_all_day: bool) -> str | None:
    """Build a human-readable time range label for a calendar event."""
    start_value = (start or {}).get("dateTime")
    end_value = (end or {}).get("dateTime")
    start_label = _format_datetime_label(start_value)
    end_label = _format_datetime_label(end_value)
    if is_all_day:
        if start_label and end_label:
            return f"All day ({start_label} → {end_label})"
        return "All day"
    if start_label and end_label:
        return f"{start_label} → {end_label}"
    return start_label or end_label


class AccountInfo(BaseModel):
    """Authenticated Microsoft account from the local token cache."""
    username: str
    account_id: str


class MailFolderSummary(BaseModel):
    """Lightweight view of a mail folder."""
    id: str
    display_name: str
    total_item_count: int | None = None
    unread_item_count: int | None = None
    child_folder_count: int | None = None
    summary: str | None = None


class MailMessageSummary(BaseModel):
    """List-level view of a mail message (no full body)."""
    id: str
    subject: str | None = None
    sender_name: str | None = None
    sender_email: str | None = None
    received_datetime: str | None = None
    received_label: str | None = None
    sender_label: str | None = None
    is_read: bool = False
    has_attachments: bool = False
    body_preview: str | None = None
    summary: str | None = None


class MailMessageDetail(BaseModel):
    """Full view of a single mail message including body content."""
    id: str
    subject: str | None = None
    sender: dict[str, Any] | None = None
    sender_label: str | None = None
    to_recipients: list[dict[str, Any]] = Field(default_factory=list)
    to_recipient_labels: list[str] = Field(default_factory=list)
    cc_recipients: list[dict[str, Any]] = Field(default_factory=list)
    cc_recipient_labels: list[str] = Field(default_factory=list)
    received_datetime: str | None = None
    received_label: str | None = None
    is_read: bool = False
    has_attachments: bool = False
    importance: str | None = None
    body_preview: str | None = None
    body_content_type: str | None = None
    body_content: str | None = None
    body_preview_clean: str | None = None
    summary: str | None = None


class CalendarSummary(BaseModel):
    """Lightweight view of a calendar."""
    id: str
    name: str | None = None
    color: str | None = None
    is_default: bool = False
    can_edit: bool | None = None
    summary: str | None = None


class CalendarEventSummary(BaseModel):
    """List-level view of a calendar event."""
    id: str
    subject: str | None = None
    start: dict[str, Any] | None = None
    end: dict[str, Any] | None = None
    location: str | None = None
    location_label: str | None = None
    is_all_day: bool = False
    time_label: str | None = None
    web_link: str | None = None
    summary: str | None = None


class CalendarEventDetail(BaseModel):
    """Full view of a single calendar event including body and attendees."""
    id: str
    subject: str | None = None
    start: dict[str, Any] | None = None
    end: dict[str, Any] | None = None
    is_all_day: bool = False
    time_label: str | None = None
    location: dict[str, Any] | None = None
    location_label: str | None = None
    body: dict[str, Any] | None = None
    body_content_type: str | None = None
    body_content: str | None = None
    body_preview_clean: str | None = None
    attendees: list[dict[str, Any]] = Field(default_factory=list)
    attendee_labels: list[str] = Field(default_factory=list)
    organizer: dict[str, Any] | None = None
    organizer_label: str | None = None
    web_link: str | None = None
    is_cancelled: bool = False
    is_online_meeting: bool = False
    summary: str | None = None
