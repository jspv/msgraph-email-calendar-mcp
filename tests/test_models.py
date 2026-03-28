from __future__ import annotations

from msgraph_mcp.models import (
    _address_label,
    _clean_text_snippet,
    _event_time_label,
    _format_datetime_label,
    _location_label,
    _recipient_labels,
    _safe_parse_datetime,
    AttachmentSummary,
    AttachmentDetail,
    DraftPreview,
    PersonResult,
    MeetingTimeSuggestion,
    ScheduleEntry,
)


class TestSafeParseDatetime:
    def test_iso_with_z(self):
        result = _safe_parse_datetime("2026-03-26T12:00:00Z")
        assert result is not None
        assert result.year == 2026

    def test_iso_with_offset(self):
        result = _safe_parse_datetime("2026-03-26T12:00:00+05:00")
        assert result is not None

    def test_none(self):
        assert _safe_parse_datetime(None) is None

    def test_empty_string(self):
        assert _safe_parse_datetime("") is None

    def test_invalid(self):
        assert _safe_parse_datetime("not a date") is None


class TestFormatDatetimeLabel:
    def test_formats_utc(self):
        result = _format_datetime_label("2026-03-26T14:30:00Z")
        assert "2026-03-26" in result
        assert "14:30" in result

    def test_none(self):
        assert _format_datetime_label(None) is None

    def test_invalid_returns_original(self):
        assert _format_datetime_label("not a date") == "not a date"


class TestCleanTextSnippet:
    def test_collapses_whitespace(self):
        assert _clean_text_snippet("hello   world\n\tfoo") == "hello world foo"

    def test_truncates_with_ellipsis(self):
        result = _clean_text_snippet("a" * 300, max_len=50)
        assert len(result) == 50
        assert result.endswith("\u2026")

    def test_short_string_unchanged(self):
        assert _clean_text_snippet("hello") == "hello"

    def test_none(self):
        assert _clean_text_snippet(None) is None

    def test_empty_string(self):
        assert _clean_text_snippet("") is None

    def test_whitespace_only(self):
        assert _clean_text_snippet("   \n\t  ") is None


class TestAddressLabel:
    def test_name_and_address(self):
        person = {"emailAddress": {"name": "Alice", "address": "alice@example.com"}}
        assert _address_label(person) == "Alice <alice@example.com>"

    def test_name_only(self):
        person = {"emailAddress": {"name": "Alice"}}
        assert _address_label(person) == "Alice"

    def test_address_only(self):
        person = {"emailAddress": {"address": "alice@example.com"}}
        assert _address_label(person) == "alice@example.com"

    def test_none(self):
        assert _address_label(None) is None

    def test_empty_dict(self):
        assert _address_label({}) is None


class TestRecipientLabels:
    def test_multiple(self):
        items = [
            {"emailAddress": {"name": "A", "address": "a@x.com"}},
            {"emailAddress": {"name": "B", "address": "b@x.com"}},
        ]
        result = _recipient_labels(items)
        assert len(result) == 2
        assert "A <a@x.com>" in result

    def test_empty_list(self):
        assert _recipient_labels([]) == []


class TestLocationLabel:
    def test_with_display_name(self):
        assert _location_label({"displayName": "Room 42"}) == "Room 42"

    def test_empty_display_name(self):
        assert _location_label({"displayName": ""}) is None

    def test_none(self):
        assert _location_label(None) is None


class TestEventTimeLabel:
    def test_all_day(self):
        result = _event_time_label(
            {"dateTime": "2026-03-26T00:00:00"},
            {"dateTime": "2026-03-27T00:00:00"},
            is_all_day=True,
        )
        assert result.startswith("All day")

    def test_timed_event(self):
        result = _event_time_label(
            {"dateTime": "2026-03-26T09:00:00"},
            {"dateTime": "2026-03-26T10:00:00"},
            is_all_day=False,
        )
        assert "\u2192" in result

    def test_none_values(self):
        assert _event_time_label(None, None, is_all_day=False) is None


class TestAttachmentSummary:
    def test_defaults(self):
        a = AttachmentSummary(id="att1")
        assert a.id == "att1"
        assert a.is_inline is False
        assert a.size is None

    def test_full(self):
        a = AttachmentSummary(id="att1", name="doc.pdf", size=1024, content_type="application/pdf", is_inline=False)
        assert a.name == "doc.pdf"
        assert a.size == 1024


class TestAttachmentDetail:
    def test_with_content(self):
        a = AttachmentDetail(id="att1", name="doc.pdf", size=100, content_base64="dGVzdA==")
        assert a.content_base64 == "dGVzdA=="
        assert a.content_omitted is False

    def test_omitted(self):
        a = AttachmentDetail(id="att1", name="big.zip", size=10_000_000, content_omitted=True, omit_reason="exceeds size limit")
        assert a.content_omitted is True
        assert a.content_base64 is None


class TestDraftPreview:
    def test_defaults(self):
        d = DraftPreview(id="draft1", subject="Hello", to_recipients=["alice@example.com"])
        assert d.subject == "Hello"
        assert d.to_recipients == ["alice@example.com"]
        assert d.cc_recipients == []


class TestPersonResult:
    def test_person(self):
        p = PersonResult(name="Alice Smith", email="alice@example.com")
        assert p.name == "Alice Smith"
        assert p.email == "alice@example.com"


class TestMeetingTimeSuggestion:
    def test_defaults(self):
        m = MeetingTimeSuggestion(start="2026-04-01T09:00:00", end="2026-04-01T10:00:00", confidence=100.0)
        assert m.confidence == 100.0


class TestScheduleEntry:
    def test_defaults(self):
        s = ScheduleEntry(email="alice@example.com", availability_view="0010")
        assert s.email == "alice@example.com"
        assert s.schedule_items == []
