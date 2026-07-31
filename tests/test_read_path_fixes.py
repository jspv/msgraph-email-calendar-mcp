"""Three fixes: read-path timezones, datetime labels, and a stable message key.

1. `list_events` never got the rule #9 established for writes, so the same bare
   wall-clock string was refused by `create_event` and silently accepted (and
   shifted by the caller's offset) by `list_events`.
2. `_format_datetime_label` applied `%Z` to naive datetimes, yielding a trailing
   space and -- worse since 0.3.0 -- never naming the zone at all.
3. `internetMessageId` was never selected or surfaced, so callers had no key
   that survives a folder move (issue #12).
"""
from __future__ import annotations

from dataclasses import replace
from unittest.mock import patch

import pytest

from msgraph_mcp import config
from msgraph_mcp.calendar import list_events
from msgraph_mcp.mail import _SUMMARY_SELECT, _message_summary
from msgraph_mcp.models import _event_time_label, _format_datetime_label


def _with_default_tz(name):
    return patch.object(config, "settings", replace(config.settings, default_timezone=name))


# ---------------------------------------------------------------- 1. read tz


class TestListEventsTimezone:
    @patch("msgraph_mcp.calendar.GraphClient")
    def test_offsetless_window_is_refused_like_the_write_paths(self, MockClient):
        with _with_default_tz(None):
            with pytest.raises(ValueError, match="no UTC offset"):
                list_events(start_iso="2026-07-31T00:00:00", end_iso="2026-08-01T00:00:00")
        MockClient.return_value.paginate.assert_not_called()

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_explicit_timezone_is_converted_to_an_instant(self, MockClient):
        # A $filter needs a single instant, so unlike a write the wall-clock
        # time must be resolved through the zone rather than handed over with it.
        MockClient.return_value.paginate.return_value = []
        list_events(
            start_iso="2026-07-31T00:00:00",
            end_iso="2026-08-01T00:00:00",
            timezone="America/New_York",
        )
        params = MockClient.return_value.paginate.call_args[1]["params"]
        assert "2026-07-31T04:00:00" in str(params)

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_offset_bearing_window_is_honoured(self, MockClient):
        MockClient.return_value.paginate.return_value = []
        list_events(start_iso="2026-07-31T00:00:00-04:00", end_iso="2026-08-01T00:00:00-04:00")
        assert "2026-07-31T04:00:00" in str(MockClient.return_value.paginate.call_args[1]["params"])

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_configured_default_is_used(self, MockClient):
        MockClient.return_value.paginate.return_value = []
        with _with_default_tz("America/New_York"):
            list_events(start_iso="2026-07-31T00:00:00", end_iso="2026-08-01T00:00:00")
        assert "2026-07-31T04:00:00" in str(MockClient.return_value.paginate.call_args[1]["params"])

    @patch("msgraph_mcp.calendar.GraphClient")
    def test_defaults_still_work_with_no_window_at_all(self, MockClient):
        # The computed default window is timezone-aware, so it must not trip the
        # new refusal -- otherwise plain list_events() breaks.
        MockClient.return_value.paginate.return_value = []
        with _with_default_tz(None):
            list_events()
        assert MockClient.return_value.paginate.called


# ------------------------------------------------------------------ 2. labels


class TestDatetimeLabels:
    def test_utc_input_keeps_its_zone_name(self):
        assert _format_datetime_label("2026-08-04T14:00:00Z") == "2026-08-04 14:00 UTC"

    def test_naive_input_has_no_trailing_space(self):
        got = _format_datetime_label("2026-08-04T14:00:00")
        assert got == got.strip()
        assert got == "2026-08-04 14:00"

    def test_naive_input_can_be_told_its_zone(self):
        # Graph's dateTimeTimeZone carries the zone beside the naive string, so
        # the label can name it instead of silently dropping it.
        assert (
            _format_datetime_label("2026-08-04T14:00:00", "Eastern Standard Time")
            == "2026-08-04 14:00 Eastern Standard Time"
        )

    def test_event_range_has_no_doubled_spaces(self):
        label = _event_time_label(
            {"dateTime": "2026-08-04T14:00:00", "timeZone": "UTC"},
            {"dateTime": "2026-08-04T15:00:00", "timeZone": "UTC"},
            False,
        )
        assert "  " not in label
        assert label == label.strip()

    def test_event_range_names_the_zone_from_the_payload(self):
        label = _event_time_label(
            {"dateTime": "2026-08-04T14:00:00", "timeZone": "UTC"},
            {"dateTime": "2026-08-04T15:00:00", "timeZone": "UTC"},
            False,
        )
        assert "UTC" in label

    def test_all_day_label_is_unchanged_in_shape(self):
        label = _event_time_label(
            {"dateTime": "2026-08-04T00:00:00", "timeZone": "UTC"},
            {"dateTime": "2026-08-05T00:00:00", "timeZone": "UTC"},
            True,
        )
        assert label.startswith("All day")


# ------------------------------------------------- 3. internetMessageId (#12)


class TestInternetMessageId:
    def _item(self, **over):
        item = {
            "id": "AAMkGRAPH_ID_1",
            "internetMessageId": "<abc123@mail.example.com>",
            "subject": "Renew the lease",
            "from": {"emailAddress": {"address": "a@b.c"}},
            "toRecipients": [],
            "ccRecipients": [],
            "receivedDateTime": "2026-01-02T00:00:00Z",
            "isRead": True,
            "hasAttachments": False,
            "categories": [],
            "bodyPreview": "",
        }
        item.update(over)
        return item

    def test_it_is_selected(self):
        assert "internetMessageId" in _SUMMARY_SELECT

    def test_summary_surfaces_it(self):
        assert (
            _message_summary(self._item()).internet_message_id
            == "<abc123@mail.example.com>"
        )

    def test_absent_value_is_none_not_a_crash(self):
        item = self._item()
        del item["internetMessageId"]
        assert _message_summary(item).internet_message_id is None

    def test_it_is_distinct_from_the_graph_id(self):
        # The whole point: Graph's id is reminted on every folder move.
        s = _message_summary(self._item())
        assert s.id != s.internet_message_id

    @patch("msgraph_mcp.mail.GraphClient")
    def test_get_message_selects_and_surfaces_it(self, MockClient):
        from msgraph_mcp.mail import get_message

        MockClient.return_value.request.return_value = self._item() | {"body": {}}
        detail = get_message(None, "m1")
        select = MockClient.return_value.request.call_args[1]["params"]["$select"]
        assert "internetMessageId" in select
        assert detail.internet_message_id == "<abc123@mail.example.com>"
