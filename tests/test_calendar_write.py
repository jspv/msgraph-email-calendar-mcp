from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp.calendar import (
    create_event,
    update_event,
    delete_event,
    respond_to_event,
)


class TestCreateEvent:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_basic_event(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "id": "event-1",
            "subject": "Team Standup",
            "start": {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-04-01T09:30:00", "timeZone": "UTC"},
            "isAllDay": False,
            "location": {"displayName": ""},
            "webLink": "https://outlook.com/event-1",
        }
        result = create_event(
            account_id=None,
            subject="Team Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
            dry_run=False,
        )
        assert result["ok"] is True
        assert result["event"]["id"] == "event-1"
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/calendar/events")

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_event_with_attendees(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "event-2", "subject": "Meeting",
            "start": {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-04-01T10:00:00", "timeZone": "UTC"},
            "isAllDay": False, "location": {}, "webLink": ""}
        create_event(
            account_id=None,
            subject="Meeting",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T10:00:00",
            attendees=["alice@example.com", "bob@example.com"],
            dry_run=False,
        )
        body = client.request.call_args[1]["json_body"]
        assert len(body["attendees"]) == 2
        assert body["attendees"][0]["emailAddress"]["address"] == "alice@example.com"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_event_on_specific_calendar(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "event-3", "subject": "Cal Event",
            "start": {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-04-01T10:00:00", "timeZone": "UTC"},
            "isAllDay": False, "location": {}, "webLink": ""}
        create_event(
            account_id=None,
            subject="Cal Event",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T10:00:00",
            calendar_id="cal-123",
            dry_run=False,
        )
        call_args = client.request.call_args
        assert "calendars/cal-123/events" in call_args[0][1]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_all_day_event(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "event-4", "subject": "Holiday",
            "start": {"dateTime": "2026-04-01", "timeZone": "UTC"},
            "end": {"dateTime": "2026-04-02", "timeZone": "UTC"},
            "isAllDay": True, "location": {}, "webLink": ""}
        create_event(
            account_id=None,
            subject="Holiday",
            start_iso="2026-04-01",
            end_iso="2026-04-02",
            is_all_day=True,
            dry_run=False,
        )
        body = client.request.call_args[1]["json_body"]
        assert body["isAllDay"] is True


class TestUpdateEvent:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_update_subject(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "event-1", "subject": "Updated",
            "start": {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-04-01T10:00:00", "timeZone": "UTC"},
            "isAllDay": False, "location": {}, "webLink": ""}
        result = update_event(
            account_id=None, event_id="event-1", subject="Updated", dry_run=False
        )
        assert result["ok"] is True
        call_args = client.request.call_args
        assert call_args[0] == ("PATCH", "/me/events/event-1")
        body = call_args[1]["json_body"]
        assert body["subject"] == "Updated"
        assert "start" not in body  # Only provided fields

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_update_time(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "event-1", "subject": "Mtg",
            "start": {"dateTime": "2026-04-01T10:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-04-01T11:00:00", "timeZone": "UTC"},
            "isAllDay": False, "location": {}, "webLink": ""}
        update_event(
            account_id=None,
            event_id="event-1",
            start_iso="2026-04-01T10:00:00",
            end_iso="2026-04-01T11:00:00",
            dry_run=False,
        )
        body = client.request.call_args[1]["json_body"]
        assert "start" in body
        assert "end" in body


class TestDeleteEvent:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_simple_delete(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = delete_event(account_id=None, event_id="event-1", dry_run=False)
        assert result["ok"] is True
        call_args = client.request.call_args
        assert call_args[0] == ("DELETE", "/me/events/event-1")

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_cancel_with_message(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = delete_event(
            account_id=None,
            event_id="event-1",
            cancel_message="Meeting cancelled due to conflict",
            dry_run=False,
        )
        assert result["ok"] is True
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/events/event-1/cancel")
        body = call_args[1]["json_body"]
        assert body["comment"] == "Meeting cancelled due to conflict"


class TestRespondToEvent:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_accept(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = respond_to_event(account_id=None, event_id="event-1", response="accept")
        assert result["ok"] is True
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/events/event-1/accept")

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_decline_with_message(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = respond_to_event(
            account_id=None,
            event_id="event-1",
            response="decline",
            message="Can't make it",
        )
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/events/event-1/decline")
        body = call_args[1]["json_body"]
        assert body["comment"] == "Can't make it"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_tentative(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = respond_to_event(account_id=None, event_id="event-1", response="tentativelyAccept")
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/events/event-1/tentativelyAccept")

    def test_rejects_invalid_response(self):
        import pytest
        with pytest.raises(ValueError, match="response"):
            respond_to_event(account_id=None, event_id="event-1", response="maybe")


from msgraph_mcp.calendar import find_meeting_times, get_schedule


class TestFindMeetingTimes:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_basic_request(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "meetingTimeSuggestions": [
                {
                    "meetingTimeSlot": {
                        "start": {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"},
                        "end": {"dateTime": "2026-04-01T10:00:00", "timeZone": "UTC"},
                    },
                    "confidence": 100.0,
                    "organizerAvailability": "free",
                    "attendeeAvailability": [],
                }
            ],
        }
        result = find_meeting_times(
            account_id=None,
            attendees=["alice@example.com"],
            duration_minutes=60,
        )
        assert len(result) == 1
        assert result[0].confidence == 100.0
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/findMeetingTimes")


class TestGetSchedule:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_basic_request(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {
            "value": [
                {
                    "scheduleId": "alice@example.com",
                    "availabilityView": "0010220",
                    "scheduleItems": [
                        {"status": "busy", "start": {"dateTime": "2026-04-01T10:00:00"}, "end": {"dateTime": "2026-04-01T11:00:00"}},
                    ],
                },
            ],
        }
        result = get_schedule(
            account_id=None,
            emails=["alice@example.com"],
            start_iso="2026-04-01T00:00:00",
            end_iso="2026-04-01T23:59:59",
        )
        assert len(result) == 1
        assert result[0].email == "alice@example.com"
        assert result[0].availability_view == "0010220"
        call_args = client.request.call_args
        assert call_args[0] == ("POST", "/me/calendar/getSchedule")


class TestEventTimeNormalisation:
    """Event times are converted to real UTC, not just labelled UTC.

    ``dateTimeTimeZone.dateTime`` carries no offset of its own, so pairing an
    offset-bearing string with ``timeZone: "UTC"`` would land the event at the
    wrong hour.
    """

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_create_converts_offset_bearing_times_to_utc(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "e1"}

        create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-07-30T14:00:00-07:00",
            end_iso="2026-07-30T14:30:00-07:00",
            dry_run=False,
        )

        body = client.request.call_args[1]["json_body"]
        assert body["start"] == {"dateTime": "2026-07-30T21:00:00", "timeZone": "UTC"}
        assert body["end"] == {"dateTime": "2026-07-30T21:30:00", "timeZone": "UTC"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_create_reads_naive_times_as_utc_unchanged(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "e2"}

        create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
            dry_run=False,
        )

        body = client.request.call_args[1]["json_body"]
        assert body["start"] == {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_create_accepts_trailing_z(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "e3"}

        create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-04-01T09:00:00Z",
            end_iso="2026-04-01T09:30:00Z",
            dry_run=False,
        )

        body = client.request.call_args[1]["json_body"]
        assert body["start"] == {"dateTime": "2026-04-01T09:00:00", "timeZone": "UTC"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_update_converts_offset_bearing_times_to_utc(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "e4"}

        update_event(
            account_id=None, event_id="e4", start_iso="2026-07-30T14:00:00-07:00",
            dry_run=False,
        )

        body = client.request.call_args[1]["json_body"]
        assert body["start"] == {"dateTime": "2026-07-30T21:00:00", "timeZone": "UTC"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_rejects_malformed_start(self, MockClient):
        with pytest.raises(ValueError, match="start_iso must be an ISO-8601 datetime"):
            create_event(
                account_id=None,
                subject="Standup",
                start_iso="next friday",
                end_iso="2026-04-01T09:30:00",
            )


_EVENT_PAYLOAD = {
    "id": "event-1",
    "subject": "Q3 Planning",
    "start": {"dateTime": "2026-08-04T14:00:00", "timeZone": "UTC"},
    "end": {"dateTime": "2026-08-04T15:00:00", "timeZone": "UTC"},
    "isAllDay": False,
    "location": {"displayName": "Room 4"},
    "body": {"contentType": "text", "content": "agenda"},
    "attendees": [
        {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
        {"emailAddress": {"name": "Bob", "address": "bob@example.com"}},
    ],
    "organizer": {"emailAddress": {"name": "Carol", "address": "carol@example.com"}},
    "webLink": "https://outlook.com/event-1",
    "isCancelled": False,
    "isOnlineMeeting": False,
}


class TestCreateEventDryRun:
    """Creating an event mails invitations immediately, so preview is the default."""

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_is_the_default_and_makes_no_graph_call(self, MockClient):
        client = MockClient.return_value
        result = create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
        )
        assert result["dry_run"] is True
        # The body is fully known client-side, so unlike mail's draft-based
        # preview this costs nothing.
        client.request.assert_not_called()

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_preview_shows_the_body_that_would_be_sent(self, MockClient):
        result = create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
            attendees=["alice@example.com"],
        )
        event = result["preview"]["event"]
        assert event["subject"] == "Standup"
        assert event["attendees"][0]["emailAddress"]["address"] == "alice@example.com"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_preview_shows_converted_utc_times(self, MockClient):
        # The preview doubles as a check on the timezone conversion: a caller
        # who types a -07:00 time should see 21:00 UTC before anything is booked.
        result = create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-07-30T14:00:00-07:00",
            end_iso="2026-07-30T14:30:00-07:00",
        )
        event = result["preview"]["event"]
        assert event["start"] == {"dateTime": "2026-07-30T21:00:00", "timeZone": "UTC"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_preview_names_the_target_path(self, MockClient):
        result = create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
            calendar_id="cal-123",
        )
        assert "calendars/cal-123/events" in result["preview"]["path"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_message_says_nothing_was_created(self, MockClient):
        result = create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
        )
        assert "NOT created" in result["message"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_false_actually_posts(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "event-1"}
        result = create_event(
            account_id=None,
            subject="Standup",
            start_iso="2026-04-01T09:00:00",
            end_iso="2026-04-01T09:30:00",
            dry_run=False,
        )
        assert result["dry_run"] is False
        assert client.request.call_args[0] == ("POST", "/me/calendar/events")

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_validation_still_runs_in_dry_run(self, MockClient):
        # A preview that accepts an unparseable time would be worse than useless.
        with pytest.raises(ValueError, match="start_iso must be an ISO-8601 datetime"):
            create_event(
                account_id=None,
                subject="Standup",
                start_iso="next friday",
                end_iso="2026-04-01T09:30:00",
            )


class TestUpdateEventDryRun:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_is_the_default_and_does_not_patch(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = _EVENT_PAYLOAD
        result = update_event(account_id=None, event_id="event-1", subject="Q4 Planning")
        assert result["dry_run"] is True
        methods = {call[0][0] for call in client.request.call_args_list}
        assert methods == {"GET"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_preview_pairs_current_state_with_proposed_changes(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = _EVENT_PAYLOAD
        result = update_event(account_id=None, event_id="event-1", subject="Q4 Planning")
        assert result["preview"]["current"]["subject"] == "Q3 Planning"
        assert result["preview"]["changes"]["subject"] == "Q4 Planning"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_preview_converts_times_in_the_change_set(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = _EVENT_PAYLOAD
        result = update_event(
            account_id=None, event_id="event-1", start_iso="2026-07-30T14:00:00-07:00"
        )
        changes = result["preview"]["changes"]
        assert changes["start"] == {"dateTime": "2026-07-30T21:00:00", "timeZone": "UTC"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_false_actually_patches(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = {"id": "event-1"}
        result = update_event(
            account_id=None, event_id="event-1", subject="Q4 Planning", dry_run=False
        )
        assert result["dry_run"] is False
        assert client.request.call_args[0] == ("PATCH", "/me/events/event-1")


class TestDeleteEventDryRun:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_is_the_default_and_does_not_delete(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = _EVENT_PAYLOAD
        result = delete_event(account_id=None, event_id="event-1")
        assert result["dry_run"] is True
        methods = {call[0][0] for call in client.request.call_args_list}
        assert methods == {"GET"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_preview_names_the_event_rather_than_echoing_an_id(self, MockClient):
        # An opaque id tells a reviewing human nothing about what is being lost.
        client = MockClient.return_value
        client.request.return_value = _EVENT_PAYLOAD
        preview = delete_event(account_id=None, event_id="event-1")["preview"]
        assert preview["subject"] == "Q3 Planning"
        assert preview["attendee_count"] == 2
        assert "2026-08-04" in preview["time_label"]
        assert "carol@example.com" in preview["organizer"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_message_warns_that_cancelling_notifies_attendees(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = _EVENT_PAYLOAD
        result = delete_event(
            account_id=None, event_id="event-1", cancel_message="Conflict"
        )
        assert result["action"] == "cancel"
        assert "notify 2 attendees" in result["message"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_message_notes_that_hard_delete_is_silent(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = _EVENT_PAYLOAD
        result = delete_event(account_id=None, event_id="event-1")
        assert result["action"] == "delete"
        assert "without notifying" in result["message"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_false_actually_deletes(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = delete_event(account_id=None, event_id="event-1", dry_run=False)
        assert result["dry_run"] is False
        assert client.request.call_args[0] == ("DELETE", "/me/events/event-1")

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_false_with_message_cancels(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = delete_event(
            account_id=None, event_id="event-1", cancel_message="Conflict", dry_run=False
        )
        assert client.request.call_args[0] == ("POST", "/me/events/event-1/cancel")
        assert result["action"] == "cancelled"


class TestRespondToEventIsNotGated:
    """Accept/decline is reversible by responding again, so no dry-run gate."""

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_responds_immediately_without_a_dry_run_flag(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = respond_to_event(account_id=None, event_id="event-1", response="accept")
        assert result["ok"] is True
        assert "dry_run" not in result
        assert client.request.call_args[0] == ("POST", "/me/events/event-1/accept")
