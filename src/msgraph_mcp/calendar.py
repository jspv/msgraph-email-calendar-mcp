"""Outlook calendar operations: list calendars, list events, and get event details.

Supports shared calendars via the *user_id* parameter.  When provided,
requests target ``/users/{user_id}/…`` instead of ``/me/…``.  Requires
the ``Calendars.ReadWrite.Shared`` scope.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

from .graph import GraphClient, validate_path_segment
from .models import (
    CalendarEventDetail,
    CalendarEventSummary,
    CalendarSummary,
    MeetingTimeSuggestion,
    ScheduleEntry,
    _address_label,
    _clean_text_snippet,
    _event_time_label,
    _location_label,
    _parse_utc,
    _recipient_labels,
)


def _graph_datetime(value: str, label: str) -> dict[str, str]:
    """Build a Graph ``dateTimeTimeZone`` object in real UTC.

    ``dateTime`` carries no offset of its own, so an offset-bearing input must
    be *converted* rather than relabelled -- pairing ``14:00:00-07:00`` with
    ``timeZone: "UTC"`` would otherwise book the event seven hours early. A
    value with no offset is taken as UTC.
    """
    return {
        "dateTime": _parse_utc(value, label).strftime("%Y-%m-%dT%H:%M:%S"),
        "timeZone": "UTC",
    }



def _attendees_phrase(count: int) -> str:
    """``"1 attendee"`` / ``"3 attendees"`` -- these strings are read by humans."""
    return f"{count} attendee" if count == 1 else f"{count} attendees"


def _base_path(user_id: str | None) -> str:
    """Return ``/me`` or ``/users/{user_id}`` as the request base."""
    if user_id:
        validate_path_segment(user_id, "user_id")
        return f"/users/{user_id}"
    return "/me"


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



def list_calendars(account_id: str | None = None, user_id: str | None = None) -> list[CalendarSummary]:
    """Return all readable calendars.  Pass *user_id* for shared calendars."""
    client = GraphClient(account_id)
    base = _base_path(user_id)
    payload = client.request(
        "GET",
        f"{base}/calendars",
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
    user_id: str | None = None,
) -> list[CalendarEventSummary]:
    """List events in a time range.  Pass *user_id* for shared calendars."""
    client = GraphClient(account_id)
    base = _base_path(user_id)
    start_iso, end_iso = _resolve_time_window(start_iso, end_iso)

    if calendar_id:
        validate_path_segment(calendar_id, "calendar_id")
        path = f"{base}/calendars/{calendar_id}/calendarView"
    else:
        path = f"{base}/calendar/calendarView"

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
    user_id: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Create a new calendar event.  Pass *user_id* for shared calendars.

    Dry-run by default: Graph mails invitations the moment an event with
    attendees is created, so there is no undo. Unlike the mail preview -- which
    has to create a real draft to see Graph's rendering -- the event body is
    fully known here, so the preview costs no Graph call at all.
    """
    client = GraphClient(account_id)
    base = _base_path(user_id)
    event_body: dict = {
        "subject": subject,
        "start": _graph_datetime(start_iso, "start_iso"),
        "end": _graph_datetime(end_iso, "end_iso"),
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
        path = f"{base}/calendars/{calendar_id}/events"
    else:
        path = f"{base}/calendar/events"

    if dry_run:
        return {
            "ok": True,
            "dry_run": True,
            "action": "create",
            "preview": {"path": path, "event": event_body},
            "message": (
                "Dry-run: event NOT created. Set dry_run=False to create it"
                + (
                    f" and invite {_attendees_phrase(len(attendees))}."
                    if attendees
                    else "."
                )
            ),
        }

    result = client.request("POST", path, json_body=event_body) or {}
    return {
        "ok": True,
        "dry_run": False,
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
    user_id: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Update an existing calendar event.  Pass *user_id* for shared calendars.

    Dry-run by default: edits notify attendees. The preview spends one GET so it
    can show current-versus-proposed rather than a bare patch dict -- knowing a
    field is about to become "Q4 Planning" is only useful next to what it is now.
    """
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    base = _base_path(user_id)
    update: dict = {}
    if subject is not None:
        update["subject"] = subject
    if start_iso is not None:
        update["start"] = _graph_datetime(start_iso, "start_iso")
    if end_iso is not None:
        update["end"] = _graph_datetime(end_iso, "end_iso")
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

    if dry_run:
        # If this GET fails the error propagates rather than degrading to a
        # partial preview: a caller who cannot read the event cannot patch it
        # either, so there is no half-state worth inventing.
        current = get_event(account_id, event_id, user_id)
        return {
            "ok": True,
            "dry_run": True,
            "action": "update",
            "event_id": event_id,
            "preview": {
                "current": {
                    "subject": current.subject,
                    "time_label": current.time_label,
                    "location": current.location_label,
                    "attendee_count": len(current.attendees),
                },
                "changes": update,
            },
            "message": (
                "Dry-run: event NOT updated. Set dry_run=False to apply"
                + (
                    f" and notify {_attendees_phrase(len(current.attendees))}."
                    if current.attendees
                    else "."
                )
            ),
        }

    result = client.request("PATCH", f"{base}/events/{event_id}", json_body=update) or {}
    return {
        "ok": True,
        "dry_run": False,
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
    user_id: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Delete or cancel a calendar event.  Pass *user_id* for shared calendars.

    Dry-run by default, and the preview spends one GET to name the event: an
    opaque id tells a reviewing human nothing about the meeting they are about
    to destroy.

    The two paths differ in a way the API shape understates -- with
    *cancel_message* Graph cancels and **notifies attendees**; without it the
    event is hard-deleted and **nobody is told**. The preview says which.
    """
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    base = _base_path(user_id)

    if dry_run:
        current = get_event(account_id, event_id, user_id)
        attendee_count = len(current.attendees)
        if cancel_message:
            consequence = (
                f"cancel it and notify {_attendees_phrase(attendee_count)}."
                if attendee_count
                else "cancel it."
            )
        else:
            consequence = (
                f"delete it without notifying its {_attendees_phrase(attendee_count)}."
                if attendee_count
                else "delete it without notifying anyone."
            )
        return {
            "ok": True,
            "dry_run": True,
            "action": "cancel" if cancel_message else "delete",
            "event_id": event_id,
            "preview": {
                "subject": current.subject,
                "time_label": current.time_label,
                "location": current.location_label,
                "attendee_count": attendee_count,
                "organizer": current.organizer_label,
                "is_cancelled": current.is_cancelled,
            },
            "message": f"Dry-run: event NOT removed. Set dry_run=False to {consequence}",
        }

    if cancel_message:
        client.request(
            "POST",
            f"{base}/events/{event_id}/cancel",
            json_body={"comment": cancel_message},
        )
        return {"ok": True, "dry_run": False, "event_id": event_id, "action": "cancelled"}
    client.request("DELETE", f"{base}/events/{event_id}")
    return {"ok": True, "dry_run": False, "event_id": event_id, "action": "deleted"}


def respond_to_event(
    account_id: str | None = None,
    *,
    event_id: str,
    response: str,
    message: str | None = None,
    user_id: str | None = None,
) -> dict:
    """Accept, decline, or tentatively accept a meeting invite.  Pass *user_id* for shared calendars."""
    if response not in _VALID_RESPONSES:
        raise ValueError(f"response must be one of {_VALID_RESPONSES}")
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    base = _base_path(user_id)
    json_body: dict = {}
    if message:
        json_body["comment"] = message
    json_body["sendResponse"] = True
    client.request("POST", f"{base}/events/{event_id}/{response}", json_body=json_body)
    return {"ok": True, "event_id": event_id, "response": response}


def get_event(account_id: str | None, event_id: str, user_id: str | None = None) -> CalendarEventDetail:
    """Fetch full details for a single calendar event.  Pass *user_id* for shared calendars."""
    validate_path_segment(event_id, "event_id")
    client = GraphClient(account_id)
    base = _base_path(user_id)
    item = client.request(
        "GET",
        f"{base}/events/{event_id}",
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


def find_meeting_times(
    account_id: str | None = None,
    *,
    attendees: list[str],
    duration_minutes: int = 60,
    start_iso: str | None = None,
    end_iso: str | None = None,
) -> list[MeetingTimeSuggestion]:
    """Suggest available meeting times for a set of attendees."""
    client = GraphClient(account_id)
    body: dict = {
        "attendees": [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ],
        "meetingDuration": f"PT{duration_minutes}M",
    }
    if start_iso and end_iso:
        body["timeConstraint"] = {
            "timeslots": [
                {
                    "start": _graph_datetime(start_iso, "start_iso"),
                    "end": _graph_datetime(end_iso, "end_iso"),
                }
            ]
        }
    result = client.request("POST", "/me/findMeetingTimes", json_body=body) or {}
    suggestions = result.get("meetingTimeSuggestions") or []
    return [
        MeetingTimeSuggestion(
            start=(s.get("meetingTimeSlot") or {}).get("start", {}).get("dateTime"),
            end=(s.get("meetingTimeSlot") or {}).get("end", {}).get("dateTime"),
            confidence=s.get("confidence"),
            organizer_availability=s.get("organizerAvailability"),
            attendee_availability=s.get("attendeeAvailability") or [],
        )
        for s in suggestions
    ]


def get_schedule(
    account_id: str | None = None,
    *,
    emails: list[str],
    start_iso: str,
    end_iso: str,
) -> list[ScheduleEntry]:
    """Get free/busy information for one or more users."""
    client = GraphClient(account_id)
    body = {
        "schedules": emails,
        "startTime": _graph_datetime(start_iso, "start_iso"),
        "endTime": _graph_datetime(end_iso, "end_iso"),
    }
    result = client.request("POST", "/me/calendar/getSchedule", json_body=body) or {}
    items = result.get("value") or []
    return [
        ScheduleEntry(
            email=item.get("scheduleId", ""),
            availability_view=item.get("availabilityView"),
            schedule_items=item.get("scheduleItems") or [],
        )
        for item in items
    ]
