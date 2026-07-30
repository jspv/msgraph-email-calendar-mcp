"""Calendar times must never be silently relocated (issues #9 and #10).

Graph's ``dateTimeTimeZone`` pairs a *naive* wall-clock string with a separate
zone name. Getting that pairing wrong does not fail loudly -- it books a real
meeting at the wrong hour, with invitations already sent.

Three inputs, three correct answers:

* offset-bearing  -> convert the instant to UTC (fixed in fcc17cc)
* offsetless + a known zone -> hand the wall-clock time to Graph with that zone,
  which also keeps recurring events correct across DST
* offsetless + nothing configured -> refuse, rather than guess UTC (issue #9)

All-day events are a separate contract: Graph wants midnight in the stated zone,
so the *date* survives and the instant does not (issue #10).
"""
from __future__ import annotations

from dataclasses import replace
from unittest.mock import patch

import pytest

from msgraph_mcp import config
from msgraph_mcp.calendar import _graph_datetime, create_event, update_event


def _with_default_tz(name: str | None):
    """Swap ``settings.default_timezone`` for the duration of a with-block."""
    return patch.object(config, "settings", replace(config.settings, default_timezone=name))


class TestOffsetBearingInputStillConvertsToUtc:
    def test_offset_is_converted_not_relabelled(self):
        assert _graph_datetime("2026-08-01T14:00:00-04:00", "start_iso") == {
            "dateTime": "2026-08-01T18:00:00",
            "timeZone": "UTC",
        }

    def test_trailing_z_is_utc(self):
        assert _graph_datetime("2026-08-01T14:00:00Z", "start_iso") == {
            "dateTime": "2026-08-01T14:00:00",
            "timeZone": "UTC",
        }

    def test_explicit_timezone_does_not_override_an_explicit_offset(self):
        # The offset is unambiguous; a zone argument must not silently reinterpret it.
        assert _graph_datetime(
            "2026-08-01T14:00:00-04:00", "start_iso", timezone_name="Asia/Tokyo"
        ) == {"dateTime": "2026-08-01T18:00:00", "timeZone": "UTC"}


class TestOffsetlessInputUsesTheGivenZone:
    def test_wall_clock_time_is_handed_to_graph_with_the_zone(self):
        # Not converted to UTC: preserving the zone is what keeps a recurring
        # event at 14:00 local on both sides of a DST boundary.
        assert _graph_datetime(
            "2026-08-01T14:00:00", "start_iso", timezone_name="America/New_York"
        ) == {"dateTime": "2026-08-01T14:00:00", "timeZone": "America/New_York"}

    def test_seconds_are_normalised(self):
        assert _graph_datetime(
            "2026-08-01T14:00", "start_iso", timezone_name="America/New_York"
        )["dateTime"] == "2026-08-01T14:00:00"

    def test_configured_default_is_used_when_no_argument_is_given(self):
        with _with_default_tz("America/New_York"):
            assert _graph_datetime("2026-08-01T14:00:00", "start_iso") == {
                "dateTime": "2026-08-01T14:00:00",
                "timeZone": "America/New_York",
            }

    def test_explicit_argument_beats_the_configured_default(self):
        with _with_default_tz("America/New_York"):
            assert _graph_datetime(
                "2026-08-01T14:00:00", "start_iso", timezone_name="Asia/Tokyo"
            )["timeZone"] == "Asia/Tokyo"

    def test_windows_zone_names_pass_through(self):
        # Graph accepts Windows zone ids as well as IANA ones.
        assert _graph_datetime(
            "2026-08-01T14:00:00", "start_iso", timezone_name="Eastern Standard Time"
        )["timeZone"] == "Eastern Standard Time"


class TestOffsetlessInputWithoutAZoneIsRefused:
    """Issue #9: reading a bare wall-clock time as UTC misplaces real meetings."""

    def test_refused_when_nothing_is_configured(self):
        with _with_default_tz(None):
            with pytest.raises(ValueError, match="no UTC offset"):
                _graph_datetime("2026-08-01T14:00:00", "start_iso")

    def test_error_names_the_offending_parameter(self):
        with _with_default_tz(None):
            with pytest.raises(ValueError, match="start_iso"):
                _graph_datetime("2026-08-01T14:00:00", "start_iso")

    def test_error_explains_all_three_ways_out(self):
        with _with_default_tz(None):
            with pytest.raises(ValueError) as exc:
                _graph_datetime("2026-08-01T14:00:00", "start_iso")
        message = str(exc.value)
        assert "timezone" in message
        assert "offset" in message
        assert "MSGRAPH_DEFAULT_TIMEZONE" in message

    def test_create_event_refuses_rather_than_booking_at_the_wrong_hour(self):
        with _with_default_tz(None):
            with patch("msgraph_mcp.calendar.GraphClient") as MockClient:
                with pytest.raises(ValueError, match="no UTC offset"):
                    create_event(
                        subject="Sync",
                        start_iso="2026-08-01T14:00:00",
                        end_iso="2026-08-01T15:00:00",
                        dry_run=False,
                    )
                MockClient.return_value.request.assert_not_called()


class TestTimezoneValidation:
    def test_a_mistyped_iana_zone_fails_locally(self):
        # Fast local feedback beats a round-trip and a Graph 400.
        with pytest.raises(ValueError, match="timezone"):
            _graph_datetime(
                "2026-08-01T14:00:00", "start_iso", timezone_name="America/New_Yrok"
            )


class TestAllDayEventsKeepTheirDate:
    """Issue #10: Graph wants midnight in the stated zone for an all-day event."""

    def test_offset_bearing_all_day_stays_at_midnight(self):
        # Converting the instant to UTC would emit 04:00 and draw a 400.
        assert _graph_datetime(
            "2026-04-01T00:00:00-04:00", "start_iso", all_day=True
        ) == {"dateTime": "2026-04-01T00:00:00", "timeZone": "UTC"}

    def test_date_only_input_is_unaffected(self):
        assert _graph_datetime("2026-04-01", "start_iso", all_day=True) == {
            "dateTime": "2026-04-01T00:00:00",
            "timeZone": "UTC",
        }

    def test_all_day_keeps_the_callers_calendar_date_not_the_utc_one(self):
        # 2026-04-01T23:00-04:00 is 2026-04-02T03:00Z. The caller means April 1.
        assert _graph_datetime(
            "2026-04-01T23:00:00-04:00", "start_iso", all_day=True
        )["dateTime"] == "2026-04-01T00:00:00"

    def test_all_day_does_not_require_a_configured_zone(self):
        # A calendar date is unambiguous without one, so #9's refusal must not fire.
        with _with_default_tz(None):
            assert _graph_datetime("2026-04-01", "start_iso", all_day=True)[
                "dateTime"
            ] == "2026-04-01T00:00:00"

    def test_all_day_uses_the_given_zone_when_there_is_one(self):
        assert _graph_datetime(
            "2026-04-01", "start_iso", timezone_name="America/New_York", all_day=True
        ) == {"dateTime": "2026-04-01T00:00:00", "timeZone": "America/New_York"}


class TestCreateAndUpdateWireBodies:
    """The fix has to reach the actual request body, not just the helper."""

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_create_all_day_emits_midnight(self, MockClient):
        MockClient.return_value.request.return_value = {"id": "e"}
        create_event(
            subject="Holiday",
            start_iso="2026-04-01T00:00:00-04:00",
            end_iso="2026-04-02T00:00:00-04:00",
            is_all_day=True,
            dry_run=False,
        )
        body = MockClient.return_value.request.call_args[1]["json_body"]
        assert body["start"] == {"dateTime": "2026-04-01T00:00:00", "timeZone": "UTC"}
        assert body["end"] == {"dateTime": "2026-04-02T00:00:00", "timeZone": "UTC"}

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_create_passes_the_timezone_argument_through(self, MockClient):
        MockClient.return_value.request.return_value = {"id": "e"}
        create_event(
            subject="Sync",
            start_iso="2026-08-01T14:00:00",
            end_iso="2026-08-01T15:00:00",
            timezone="America/New_York",
            dry_run=False,
        )
        body = MockClient.return_value.request.call_args[1]["json_body"]
        assert body["start"]["timeZone"] == "America/New_York"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_update_passes_the_timezone_argument_through(self, MockClient):
        MockClient.return_value.request.return_value = {"id": "e"}
        update_event(
            event_id="e",
            start_iso="2026-08-01T14:00:00",
            timezone="America/New_York",
            dry_run=False,
        )
        body = MockClient.return_value.request.call_args[1]["json_body"]
        assert body["start"]["timeZone"] == "America/New_York"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_update_all_day_emits_midnight(self, MockClient):
        MockClient.return_value.request.return_value = {"id": "e"}
        update_event(
            event_id="e",
            start_iso="2026-04-01T00:00:00-04:00",
            is_all_day=True,
            dry_run=False,
        )
        body = MockClient.return_value.request.call_args[1]["json_body"]
        assert body["start"]["dateTime"] == "2026-04-01T00:00:00"

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_dry_run_preview_shows_the_corrected_all_day_body(self, MockClient):
        # The preview is what a human reviews, so it must not differ from the wire.
        result = create_event(
            subject="Holiday",
            start_iso="2026-04-01T00:00:00-04:00",
            end_iso="2026-04-02T00:00:00-04:00",
            is_all_day=True,
        )
        assert result["preview"]["event"]["start"]["dateTime"] == "2026-04-01T00:00:00"
