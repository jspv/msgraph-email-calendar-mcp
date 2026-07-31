"""Outlook calendar operations: list calendars, list events, and get event details.

Supports shared calendars via the *user_id* parameter.  When provided,
requests target ``/users/{user_id}/…`` instead of ``/me/…``.  Requires
the ``Calendars.ReadWrite.Shared`` scope.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from . import config
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


_GRAPH_DT_FORMAT = "%Y-%m-%dT%H:%M:%S"


def _validate_timezone(name: str, label: str) -> str:
    """Reject a mistyped IANA zone locally; let Windows zone ids through.

    Graph accepts both IANA (``America/New_York``) and Windows (``Eastern
    Standard Time``) names, and only the former can be checked with
    :mod:`zoneinfo`. A name containing ``/`` is unambiguously IANA-style, so a
    typo there fails here rather than as a Graph 400 a round-trip later.
    """
    if "/" in name:
        try:
            ZoneInfo(name)
        except (ZoneInfoNotFoundError, ValueError) as exc:
            raise ValueError(
                f"Unknown timezone {name!r} for {label}: expected an IANA name "
                f"such as 'America/New_York'."
            ) from exc
    return name


def _parse_naive_or_aware(value: str, label: str) -> datetime:
    """Parse an ISO-8601 string *without* forcing a timezone onto it.

    ``models._parse_utc`` deliberately reads a naive value as UTC, which is the
    right default for a mail ``received_after`` cutoff. Calendar writes must be
    able to tell the two cases apart, so they parse here instead.
    """
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except (ValueError, AttributeError) as exc:
        raise ValueError(
            f"{label} must be an ISO-8601 datetime, got {value!r}"
        ) from exc


def _graph_datetime(
    value: str,
    label: str,
    *,
    timezone_name: str | None = None,
    all_day: bool = False,
) -> dict[str, str]:
    """Build a Graph ``dateTimeTimeZone`` object without relocating the caller.

    Graph pairs a *naive* wall-clock string with a separate zone name, so the
    pairing has to be built deliberately. Three inputs, three answers:

    * **Offset-bearing** -- the instant is unambiguous, so convert it to real
      UTC. Relabelling ``14:00:00-07:00`` as ``timeZone: "UTC"`` would book the
      event seven hours early.
    * **Offsetless with a known zone** -- hand the wall-clock time to Graph
      alongside that zone rather than converting. Graph resolves it, and a
      recurring event stays at 14:00 local across a DST boundary, which a
      one-shot UTC conversion cannot do.
    * **Offsetless with no zone** -- refuse. Reading it as UTC is how a 2pm
      Eastern meeting silently becomes 10:00 EDT, and on this server the caller
      is usually a model emitting a bare local time as its normal output.

    *all_day* is a different contract: Graph wants midnight in the stated zone,
    so the calendar **date** must survive and the instant must not. The date is
    taken as the caller wrote it -- ``2026-04-01T23:00:00-04:00`` means April 1
    even though it is April 2 in UTC -- and no zone is required, because a
    calendar date is unambiguous without one.
    """
    parsed = _parse_naive_or_aware(value, label)

    if all_day:
        zone = timezone_name or config.settings.default_timezone or "UTC"
        return {
            "dateTime": parsed.date().strftime("%Y-%m-%d") + "T00:00:00",
            "timeZone": _validate_timezone(zone, label),
        }

    if parsed.tzinfo is not None:
        return {
            "dateTime": parsed.astimezone(timezone.utc).strftime(_GRAPH_DT_FORMAT),
            "timeZone": "UTC",
        }

    zone = timezone_name or config.settings.default_timezone
    if not zone:
        raise ValueError(
            f"{label} has no UTC offset and no timezone was given, so the intended "
            f"time is ambiguous. Pass timezone=\"America/New_York\", include an "
            f"offset (e.g. {value}-04:00), or set MSGRAPH_DEFAULT_TIMEZONE on the "
            f"server."
        )
    return {
        "dateTime": parsed.strftime(_GRAPH_DT_FORMAT),
        "timeZone": _validate_timezone(zone, label),
    }



def _graph_instant(value: str, label: str, timezone_name: str | None = None) -> str:
    """Resolve a caller time to an unambiguous UTC instant for a ``$filter``.

    The write paths hand Graph a wall-clock string *plus* a zone name, which is
    what preserves intent across DST. A ``$filter`` has no such pairing -- it
    compares instants -- so an offsetless value must be resolved through its
    zone here rather than passed along.

    Same refusal as ``_graph_datetime``: with no offset and no zone the intended
    time is unknown, and guessing UTC silently shifts the whole query window by
    the caller's offset. Reads were left out when #9 fixed the writes, so the
    same string was refused by ``create_event`` and quietly accepted here.
    """
    parsed = _parse_naive_or_aware(value, label)
    if parsed.tzinfo is not None:
        return parsed.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    zone = timezone_name or config.settings.default_timezone
    if not zone:
        raise ValueError(
            f"{label} has no UTC offset and no timezone was given, so the "
            f"intended time is ambiguous. Pass timezone=\"America/New_York\", "
            f"include an offset (e.g. {value}-04:00), or set "
            f"MSGRAPH_DEFAULT_TIMEZONE on the server."
        )
    _validate_timezone(zone, label)
    return (
        parsed.replace(tzinfo=ZoneInfo(zone))
        .astimezone(timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ")
    )


_WEEKDAYS = {
    "monday": "monday", "mon": "monday",
    "tuesday": "tuesday", "tue": "tuesday", "tues": "tuesday",
    "wednesday": "wednesday", "wed": "wednesday",
    "thursday": "thursday", "thu": "thursday", "thur": "thursday", "thurs": "thursday",
    "friday": "friday", "fri": "friday",
    "saturday": "saturday", "sat": "saturday",
    "sunday": "sunday", "sun": "sunday",
}

_REPEAT_TYPES = {
    "daily": "daily",
    "weekly": "weekly",
    "monthly": "absoluteMonthly",
    "yearly": "absoluteYearly",
}


def _recurrence(
    repeat: str | None,
    interval: int,
    days: list[str] | None,
    count: int | None,
    until: str | None,
    start_iso: str,
) -> dict | None:
    """Build Graph's ``recurrence`` object from a compact caller vocabulary.

    Graph's own shape is a pattern/range pair with six pattern types and three
    range types, most of which a caller saying "every Tuesday" does not want to
    think about. This exposes ``daily``/``weekly``/``monthly``/``yearly`` and
    derives the rest -- monthly and yearly take their day from *start_iso*,
    which is what the caller already told us.

    Returns ``None`` when no repeat was asked for, so the key is omitted rather
    than sent empty.
    """
    if repeat is None:
        return None
    if repeat not in _REPEAT_TYPES:
        raise ValueError(
            f"repeat must be one of {sorted(_REPEAT_TYPES)}, got {repeat!r}"
        )
    if count is not None and until is not None:
        raise ValueError("repeat_count and repeat_until are mutually exclusive")
    if interval < 1:
        raise ValueError("repeat_interval must be a positive integer")

    pattern: dict = {"type": _REPEAT_TYPES[repeat], "interval": interval}

    if repeat == "weekly":
        if not days:
            raise ValueError("repeat_days is required for a weekly repeat")
        normalised = []
        for day in days:
            key = day.strip().lower()
            if key not in _WEEKDAYS:
                raise ValueError(f"unknown day {day!r} in repeat_days")
            if _WEEKDAYS[key] not in normalised:
                normalised.append(_WEEKDAYS[key])
        pattern["daysOfWeek"] = normalised

    start_date = _parse_naive_or_aware(start_iso, "start_iso").date()
    if repeat == "monthly":
        pattern["dayOfMonth"] = start_date.day
    elif repeat == "yearly":
        pattern["dayOfMonth"] = start_date.day
        pattern["month"] = start_date.month

    if count is not None:
        if count < 1:
            raise ValueError("repeat_count must be a positive integer")
        rng = {
            "type": "numbered",
            "startDate": start_date.isoformat(),
            "numberOfOccurrences": count,
        }
    elif until is not None:
        rng = {
            "type": "endDate",
            "startDate": start_date.isoformat(),
            "endDate": _parse_naive_or_aware(until, "repeat_until").date().isoformat(),
        }
    else:
        rng = {"type": "noEnd", "startDate": start_date.isoformat()}

    return {"pattern": pattern, "range": rng}


def _repeat_phrase(repeat: str, interval: int, days, count, until) -> str:
    """Describe a repeat in words, for a dry-run a human has to approve."""
    every = repeat if interval == 1 else f"every {interval} {repeat.rstrip('ly')}s"
    base = {"daily": "daily", "weekly": "weekly",
            "monthly": "monthly", "yearly": "yearly"}[repeat] if interval == 1 else every
    text = f"repeats {base}"
    if days:
        text += " on " + ", ".join(days)
    if count:
        text += f", {count} times"
    elif until:
        text += f", until {until}"
    else:
        text += ", with no end date"
    return text


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
    timezone: str | None = None,
) -> list[CalendarEventSummary]:
    """List events in a time range.  Pass *user_id* for shared calendars.

    *start_iso* / *end_iso* follow the same rule as the write paths: an
    offsetless value needs *timezone* or ``MSGRAPH_DEFAULT_TIMEZONE``, or the
    call is refused rather than silently shifting the window.
    """
    client = GraphClient(account_id)
    base = _base_path(user_id)
    start_iso, end_iso = _resolve_time_window(start_iso, end_iso)
    start_iso = _graph_instant(start_iso, "start_iso", timezone)
    end_iso = _graph_instant(end_iso, "end_iso", timezone)

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
            "$select": "id,subject,start,end,location,isAllDay,webLink,type,seriesMasterId",
        },
        limit=limit,
    )
    output: list[CalendarEventSummary] = []
    for item in items:
        is_all_day = bool(item.get("isAllDay", False))
        location_label = (item.get("location") or {}).get("displayName")
        time_label = _event_time_label(item.get("start"), item.get("end"), is_all_day)
        subject = item.get("subject") or "(no subject)"
        event_type = item.get("type")
        # calendarView expands a series into occurrences, so without this a
        # repeat is indistinguishable from a one-off in a list result.
        repeat_note = "repeat" if event_type in {"occurrence", "exception"} else None
        summary = " — ".join(
            part
            for part in [subject, time_label, location_label, repeat_note]
            if part
        )
        output.append(
            CalendarEventSummary(
                id=item["id"],
                subject=item.get("subject"),
                type=event_type,
                series_master_id=item.get("seriesMasterId"),
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
    timezone: str | None = None,
    repeat: str | None = None,
    repeat_interval: int = 1,
    repeat_days: list[str] | None = None,
    repeat_count: int | None = None,
    repeat_until: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Create a new calendar event.  Pass *user_id* for shared calendars.

    *repeat* (``daily``/``weekly``/``monthly``/``yearly``) creates a recurring
    series. A weekly repeat needs *repeat_days*; monthly and yearly take their
    day from *start_iso*. Bound the series with *repeat_count* or
    *repeat_until* -- with neither, it never ends.

    Dry-run by default: Graph mails invitations the moment an event with
    attendees is created, so there is no undo. Unlike the mail preview -- which
    has to create a real draft to see Graph's rendering -- the event body is
    fully known here, so the preview costs no Graph call at all.
    """
    client = GraphClient(account_id)
    base = _base_path(user_id)
    event_body: dict = {
        "subject": subject,
        "start": _graph_datetime(
            start_iso, "start_iso", timezone_name=timezone, all_day=is_all_day
        ),
        "end": _graph_datetime(
            end_iso, "end_iso", timezone_name=timezone, all_day=is_all_day
        ),
        "isAllDay": is_all_day,
    }
    recurrence = _recurrence(
        repeat, repeat_interval, repeat_days, repeat_count, repeat_until, start_iso
    )
    if recurrence:
        event_body["recurrence"] = recurrence
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
                    f" and invite {_attendees_phrase(len(attendees))}"
                    if attendees
                    else ""
                )
                + (
                    "; "
                    + _repeat_phrase(
                        repeat, repeat_interval, repeat_days, repeat_count, repeat_until
                    )
                    + "."
                    if repeat
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
    timezone: str | None = None,
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
        update["start"] = _graph_datetime(
            start_iso, "start_iso", timezone_name=timezone, all_day=bool(is_all_day)
        )
    if end_iso is not None:
        update["end"] = _graph_datetime(
            end_iso, "end_iso", timezone_name=timezone, all_day=bool(is_all_day)
        )
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
                "type": current.type,
                "series_master_id": current.series_master_id,
            },
            "message": (
                f"Dry-run: event NOT removed. Set dry_run=False to {consequence}"
                # Deleting a master takes the whole series with it, which the
                # subject and time alone give no hint of.
                + (
                    " This is a recurring series: removing it deletes every"
                    " occurrence, not just this one."
                    if current.type == "seriesMaster"
                    else ""
                )
            ),
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
            "$select": "id,subject,start,end,isAllDay,location,body,attendees,organizer,webLink,isCancelled,isOnlineMeeting,type,seriesMasterId,recurrence",
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
        type=item.get("type"),
        series_master_id=item.get("seriesMasterId"),
        recurrence=item.get("recurrence"),
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
    timezone: str | None = None,
) -> list[MeetingTimeSuggestion]:
    """Suggest available meeting times for a set of attendees.

    *start_iso* / *end_iso* follow the same rule as the write paths: an
    offsetless time needs either *timezone* or ``MSGRAPH_DEFAULT_TIMEZONE``.
    """
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
                    "start": _graph_datetime(
                        start_iso, "start_iso", timezone_name=timezone
                    ),
                    "end": _graph_datetime(end_iso, "end_iso", timezone_name=timezone),
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
    timezone: str | None = None,
) -> list[ScheduleEntry]:
    """Get free/busy information for one or more users.

    *start_iso* / *end_iso* follow the same rule as the write paths: an
    offsetless time needs either *timezone* or ``MSGRAPH_DEFAULT_TIMEZONE``.
    """
    client = GraphClient(account_id)
    body = {
        "schedules": emails,
        "startTime": _graph_datetime(start_iso, "start_iso", timezone_name=timezone),
        "endTime": _graph_datetime(end_iso, "end_iso", timezone_name=timezone),
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
