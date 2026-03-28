"""Outlook calendar operations: list calendars, list events, and get event details."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

from .graph import GraphClient, validate_path_segment
from .models import (
    CalendarEventDetail,
    CalendarEventSummary,
    CalendarSummary,
    _address_label,
    _clean_text_snippet,
    _event_time_label,
    _location_label,
    _recipient_labels,
)



def _resolve_time_window(start_iso: str | None, end_iso: str | None) -> tuple[str, str]:
    """Fill in missing start/end with sensible defaults (14-day window)."""
    now = datetime.now(timezone.utc)
    if not start_iso and not end_iso:
        start = now - timedelta(days=1)
        end = now + timedelta(days=14)
        return start.isoformat(), end.isoformat()
    if not start_iso and end_iso:
        end = datetime.fromisoformat(end_iso.replace("Z", "+00:00"))
        start = end - timedelta(days=14)
        return start.isoformat(), end.isoformat()
    if start_iso and not end_iso:
        start = datetime.fromisoformat(start_iso.replace("Z", "+00:00"))
        end = start + timedelta(days=14)
        return start.isoformat(), end.isoformat()
    return start_iso, end_iso



def list_calendars(account_id: str | None = None) -> list[CalendarSummary]:
    """Return all readable calendars for the authenticated account."""
    client = GraphClient(account_id)
    payload = client.request(
        "GET",
        "/me/calendars",
        params={
            "$select": "id,name,color,isDefaultCalendar,canEdit",
            "$top": 50,
        },
    ) or {"value": []}
    return [
        CalendarSummary(
            id=item["id"],
            name=item.get("name"),
            color=item.get("color"),
            is_default=bool(item.get("isDefaultCalendar", False)),
            can_edit=item.get("canEdit"),
            summary=" • ".join(
                bit
                for bit in [
                    item.get("name") or "(unnamed calendar)",
                    "default" if bool(item.get("isDefaultCalendar", False)) else None,
                    "editable" if item.get("canEdit") else "read-only" if item.get("canEdit") is not None else None,
                    item.get("color"),
                ]
                if bit
            ),
        )
        for item in payload.get("value", [])
    ]



def list_events(
    account_id: str | None = None,
    start_iso: str | None = None,
    end_iso: str | None = None,
    calendar_id: str | None = None,
    limit: int = 25,
) -> list[CalendarEventSummary]:
    """List events in a time range from the default or specified calendar."""
    client = GraphClient(account_id)
    start_iso, end_iso = _resolve_time_window(start_iso, end_iso)

    if calendar_id:
        validate_path_segment(calendar_id, "calendar_id")
        path = f"/me/calendars/{calendar_id}/calendarView"
    else:
        path = "/me/calendar/calendarView"

    items = client.paginate(
        path,
        params={
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "$top": min(limit, 100),
            "$orderby": "start/dateTime",
            "$select": "id,subject,start,end,location,isAllDay,webLink",
        },
        limit=limit,
    )
    output: list[CalendarEventSummary] = []
    for item in items:
        is_all_day = bool(item.get("isAllDay", False))
        location_label = (item.get("location") or {}).get("displayName")
        time_label = _event_time_label(item.get("start"), item.get("end"), is_all_day)
        subject = item.get("subject") or "(no subject)"
        summary = " — ".join(
            part
            for part in [subject, time_label, location_label]
            if part
        )
        output.append(
            CalendarEventSummary(
                id=item["id"],
                subject=item.get("subject"),
                start=item.get("start"),
                end=item.get("end"),
                location=location_label,
                location_label=location_label,
                is_all_day=is_all_day,
                time_label=time_label,
                web_link=item.get("webLink"),
                summary=summary,
            )
        )
    return output



def create_event(
    account_id: str | None = None,
    *,
    subject: str,
    start_iso: str,
    end_iso: str,
    attendees: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_all_day: bool = False,
    calendar_id: str | None = None,
) -> dict:
    """Create a new calendar event."""
    client = GraphClient(account_id)
    event_body: dict = {
        "subject": subject,
        "start": {"dateTime": start_iso, "timeZone": "UTC"},
        "end": {"dateTime": end_iso, "timeZone": "UTC"},
        "isAllDay": is_all_day,
    }
    if attendees:
        event_body["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ]
    if body:
        event_body["body"] = {"contentType": "text", "content": body}
    if location:
        event_body["location"] = {"displayName": location}

    if calendar_id:
        validate_path_segment(calendar_id, "calendar_id")
        path = f"/me/calendars/{calendar_id}/events"
    else:
        path = "/me/calendar/events"

    result = client.request("POST", path, json_body=event_body) or {}
    return {
        "ok": True,
        "event": {
            "id": result.get("id"),
            "subject": result.get("subject"),
            "start": result.get("start"),
            "end": result.get("end"),
            "web_link": result.get("webLink"),
        },
    }


def update_event(
    account_id: str | None = None,
    *,
    event_id: str,
    subject: str | None = None,
    start_iso: str | None = None,
    end_iso: str | None = None,
    attendees: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_all_day: bool | None = None,
) -> dict:
    """Update an existing calendar event.  Only provided fields are changed."""
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    update: dict = {}
    if subject is not None:
        update["subject"] = subject
    if start_iso is not None:
        update["start"] = {"dateTime": start_iso, "timeZone": "UTC"}
    if end_iso is not None:
        update["end"] = {"dateTime": end_iso, "timeZone": "UTC"}
    if attendees is not None:
        update["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ]
    if body is not None:
        update["body"] = {"contentType": "text", "content": body}
    if location is not None:
        update["location"] = {"displayName": location}
    if is_all_day is not None:
        update["isAllDay"] = is_all_day

    result = client.request("PATCH", f"/me/events/{event_id}", json_body=update) or {}
    return {
        "ok": True,
        "event": {
            "id": result.get("id"),
            "subject": result.get("subject"),
            "start": result.get("start"),
            "end": result.get("end"),
        },
    }


_VALID_RESPONSES = {"accept", "decline", "tentativelyAccept"}


def delete_event(
    account_id: str | None = None,
    *,
    event_id: str,
    cancel_message: str | None = None,
) -> dict:
    """Delete or cancel a calendar event.

    If *cancel_message* is provided, sends a cancellation notification
    to attendees instead of silently deleting.
    """
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    if cancel_message:
        client.request(
            "POST",
            f"/me/events/{event_id}/cancel",
            json_body={"comment": cancel_message},
        )
        return {"ok": True, "event_id": event_id, "action": "cancelled"}
    client.request("DELETE", f"/me/events/{event_id}")
    return {"ok": True, "event_id": event_id, "action": "deleted"}


def respond_to_event(
    account_id: str | None = None,
    *,
    event_id: str,
    response: str,
    message: str | None = None,
) -> dict:
    """Accept, decline, or tentatively accept a meeting invite."""
    if response not in _VALID_RESPONSES:
        raise ValueError(f"response must be one of {_VALID_RESPONSES}")
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    json_body: dict = {}
    if message:
        json_body["comment"] = message
    json_body["sendResponse"] = True
    client.request("POST", f"/me/events/{event_id}/{response}", json_body=json_body)
    return {"ok": True, "event_id": event_id, "response": response}


def get_event(account_id: str | None, event_id: str) -> CalendarEventDetail:
    """Fetch full details for a single calendar event."""
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    item = client.request(
        "GET",
        f"/me/events/{event_id}",
        params={
            "$select": "id,subject,start,end,isAllDay,location,body,attendees,organizer,webLink,isCancelled,isOnlineMeeting",
        },
    ) or {}
    is_all_day = bool(item.get("isAllDay", False))
    location = item.get("location")
    body = item.get("body") or {}
    attendees = item.get("attendees") or []
    organizer = item.get("organizer")
    time_label = _event_time_label(item.get("start"), item.get("end"), is_all_day)
    location_label = _location_label(location)
    body_content = body.get("content")
    body_content_type = body.get("contentType")
    body_preview_clean = _clean_text_snippet(body_content)
    attendee_labels = _recipient_labels(attendees)
    organizer_label = _address_label(organizer)
    summary = " — ".join(
        part
        for part in [item.get("subject") or "(no subject)", time_label, location_label, body_preview_clean]
        if part
    )
    return CalendarEventDetail(
        id=item["id"],
        subject=item.get("subject"),
        start=item.get("start"),
        end=item.get("end"),
        is_all_day=is_all_day,
        time_label=time_label,
        location=location,
        location_label=location_label,
        body=body,
        body_content_type=body_content_type,
        body_content=body_content,
        body_preview_clean=body_preview_clean,
        attendees=attendees,
        attendee_labels=attendee_labels,
        organizer=organizer,
        organizer_label=organizer_label,
        web_link=item.get("webLink"),
        is_cancelled=bool(item.get("isCancelled", False)),
        is_online_meeting=bool(item.get("isOnlineMeeting", False)),
        summary=summary,
    )
