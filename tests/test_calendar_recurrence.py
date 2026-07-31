"""Recurring events: creation, and telling a series apart from a single event.

Probed live first. What that established:

* A ``seriesMaster`` is created by attaching ``recurrence`` (pattern + range) to
  an otherwise ordinary event body.
* ``calendarView`` -- which ``list_events`` already uses -- expands occurrences,
  so recurring events *display* correctly. Plain ``/events`` returns only the
  master, which is why the endpoint choice matters.
* Occurrences carry ``type="occurrence"`` and a ``seriesMasterId`` pointing at
  the master. Without those on the model a caller cannot tell a repeat from a
  one-off, nor find the master in order to edit the whole series.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from msgraph_mcp.calendar import _recurrence, create_event, delete_event


class TestRecurrencePattern:
    def test_none_when_no_repeat_requested(self):
        assert _recurrence(None, 1, None, None, None, "2027-05-03") is None

    def test_daily(self):
        r = _recurrence("daily", 1, None, None, None, "2027-05-03")
        assert r["pattern"]["type"] == "daily"
        assert r["pattern"]["interval"] == 1

    def test_interval_is_carried(self):
        r = _recurrence("weekly", 2, ["monday"], None, None, "2027-05-03")
        assert r["pattern"]["interval"] == 2

    def test_weekly_requires_days(self):
        # Graph needs daysOfWeek for a weekly pattern; failing here beats a 400.
        with pytest.raises(ValueError, match="repeat_days"):
            _recurrence("weekly", 1, None, None, None, "2027-05-03")

    def test_weekly_days_are_normalised(self):
        r = _recurrence("weekly", 1, ["Monday", "WED"], None, None, "2027-05-03")
        assert r["pattern"]["daysOfWeek"] == ["monday", "wednesday"]

    def test_rejects_an_unknown_day(self):
        with pytest.raises(ValueError, match="funday"):
            _recurrence("weekly", 1, ["funday"], None, None, "2027-05-03")

    def test_rejects_an_unknown_frequency(self):
        with pytest.raises(ValueError, match="repeat"):
            _recurrence("fortnightly", 1, None, None, None, "2027-05-03")

    def test_monthly_uses_the_start_day(self):
        r = _recurrence("monthly", 1, None, None, None, "2027-05-03")
        assert r["pattern"]["type"] == "absoluteMonthly"
        assert r["pattern"]["dayOfMonth"] == 3

    def test_yearly(self):
        r = _recurrence("yearly", 1, None, None, None, "2027-05-03")
        assert r["pattern"]["type"] == "absoluteYearly"

    def test_default_range_never_ends(self):
        r = _recurrence("daily", 1, None, None, None, "2027-05-03")
        assert r["range"]["type"] == "noEnd"

    def test_count_gives_a_numbered_range(self):
        r = _recurrence("daily", 1, None, 4, None, "2027-05-03")
        assert r["range"] == {
            "type": "numbered",
            "startDate": "2027-05-03",
            "numberOfOccurrences": 4,
        }

    def test_until_gives_an_end_date_range(self):
        r = _recurrence("daily", 1, None, None, "2027-06-30", "2027-05-03")
        assert r["range"]["type"] == "endDate"
        assert r["range"]["endDate"] == "2027-06-30"

    def test_count_and_until_together_are_rejected(self):
        with pytest.raises(ValueError, match="repeat_count"):
            _recurrence("daily", 1, None, 4, "2027-06-30", "2027-05-03")


class TestCreateRecurringEvent:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_recurrence_reaches_the_request_body(self, MockClient):
        MockClient.return_value.request.return_value = {"id": "e"}
        create_event(
            subject="Standup",
            start_iso="2027-05-03T09:00:00Z",
            end_iso="2027-05-03T09:30:00Z",
            repeat="weekly",
            repeat_days=["monday", "wednesday"],
            repeat_count=4,
            dry_run=False,
        )
        body = MockClient.return_value.request.call_args[1]["json_body"]
        assert body["recurrence"]["pattern"]["daysOfWeek"] == ["monday", "wednesday"]
        assert body["recurrence"]["range"]["numberOfOccurrences"] == 4

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_non_repeating_events_send_no_recurrence_key(self, MockClient):
        MockClient.return_value.request.return_value = {"id": "e"}
        create_event(subject="One off", start_iso="2027-05-03T09:00:00Z",
                     end_iso="2027-05-03T09:30:00Z", dry_run=False)
        assert "recurrence" not in MockClient.return_value.request.call_args[1]["json_body"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_preview_describes_the_repeat_in_words(self, MockClient):
        # The preview is what a human approves; a raw pattern dict is not review.
        result = create_event(
            subject="Standup", start_iso="2027-05-03T09:00:00Z",
            end_iso="2027-05-03T09:30:00Z", repeat="weekly",
            repeat_days=["monday"], repeat_count=4,
        )
        assert "repeats" in result["message"].lower()
        assert "4" in result["message"]


class TestDeletingASeries:
    """Deleting a master destroys every occurrence -- the preview must say so."""

    def _payload(self, **over):
        p = {
            "id": "e1", "subject": "Standup",
            "start": {"dateTime": "2027-05-03T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2027-05-03T09:30:00", "timeZone": "UTC"},
            "isAllDay": False, "location": {}, "body": {},
            "attendees": [], "organizer": None, "webLink": "",
            "isCancelled": False, "isOnlineMeeting": False,
            "type": "singleInstance",
        }
        p.update(over)
        return p

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_series_master_deletion_is_flagged_in_the_preview(self, MockClient):
        MockClient.return_value.request.return_value = self._payload(type="seriesMaster")
        r = delete_event(event_id="e1")
        assert r["preview"]["type"] == "seriesMaster"
        assert "every occurrence" in r["message"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_single_event_preview_says_nothing_about_series(self, MockClient):
        MockClient.return_value.request.return_value = self._payload()
        r = delete_event(event_id="e1")
        assert "every occurrence" not in r["message"]

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_occurrence_preview_reports_its_master(self, MockClient):
        MockClient.return_value.request.return_value = self._payload(
            type="occurrence", seriesMasterId="master-1"
        )
        r = delete_event(event_id="e1")
        assert r["preview"]["series_master_id"] == "master-1"


class TestReadModelsExposeSeriesInfo:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_get_event_surfaces_type_and_master(self, MockClient):
        from msgraph_mcp.calendar import get_event

        MockClient.return_value.request.return_value = {
            "id": "e1", "subject": "Standup", "type": "occurrence",
            "seriesMasterId": "master-1", "body": {},
            "start": {"dateTime": "2027-05-03T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2027-05-03T09:30:00", "timeZone": "UTC"},
        }
        d = get_event(None, "e1")
        assert d.type == "occurrence"
        assert d.series_master_id == "master-1"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_list_events_surfaces_type(self, MockClient):
        from msgraph_mcp.calendar import list_events

        MockClient.return_value.paginate.return_value = [{
            "id": "e1", "subject": "Standup", "type": "occurrence",
            "seriesMasterId": "master-1",
            "start": {"dateTime": "2027-05-03T09:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2027-05-03T09:30:00", "timeZone": "UTC"},
            "isAllDay": False, "location": {}, "webLink": "",
        }]
        ev = list_events(start_iso="2027-05-01T00:00:00Z", end_iso="2027-06-01T00:00:00Z")[0]
        assert ev.type == "occurrence"
        assert ev.series_master_id == "master-1"
        assert "repeat" in (ev.summary or "").lower()
