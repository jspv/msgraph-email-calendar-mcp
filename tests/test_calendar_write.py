from __future__ import annotations

from unittest.mock import patch

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
        result = update_event(account_id=None, event_id="event-1", subject="Updated")
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
        )
        body = client.request.call_args[1]["json_body"]
        assert "start" in body
        assert "end" in body


class TestDeleteEvent:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_simple_delete(self, MockClient):
        client = MockClient.return_value
        client.request.return_value = None
        result = delete_event(account_id=None, event_id="event-1")
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
