from __future__ import annotations

from datetime import datetime, timezone

from msgraph_mcp.calendar import _resolve_time_window


class TestResolveTimeWindow:
    def test_both_none_returns_default_range(self):
        start, end = _resolve_time_window(None, None)
        s = datetime.fromisoformat(start)
        e = datetime.fromisoformat(end)
        assert s.tzinfo is not None
        assert e.tzinfo is not None
        # Default window is -1 day to +14 days
        delta = e - s
        assert delta.days == 15

    def test_only_end_provided(self):
        end_iso = "2026-06-15T00:00:00+00:00"
        start, end = _resolve_time_window(None, end_iso)
        s = datetime.fromisoformat(start)
        e = datetime.fromisoformat(end)
        delta = e - s
        assert delta.days == 14

    def test_only_start_provided(self):
        start_iso = "2026-06-01T00:00:00+00:00"
        start, end = _resolve_time_window(start_iso, None)
        s = datetime.fromisoformat(start)
        e = datetime.fromisoformat(end)
        delta = e - s
        assert delta.days == 14

    def test_both_provided_passthrough(self):
        start_iso = "2026-06-01T00:00:00+00:00"
        end_iso = "2026-06-30T00:00:00+00:00"
        start, end = _resolve_time_window(start_iso, end_iso)
        assert start == start_iso
        assert end == end_iso
