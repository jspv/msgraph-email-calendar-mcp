from __future__ import annotations

from msgraph_mcp.mail import FOLDERS, _matches_filters
from msgraph_mcp.models import MailMessageSummary


def _make_message(
    *,
    sender_label: str | None = None,
    subject: str | None = None,
    body_preview: str | None = None,
    is_read: bool = False,
    received_datetime: str | None = None,
) -> MailMessageSummary:
    return MailMessageSummary(
        id="test-id",
        sender_label=sender_label,
        subject=subject,
        body_preview=body_preview,
        is_read=is_read,
        received_datetime=received_datetime,
    )


class TestMatchesFilters:
    def test_no_filters_matches_all(self):
        msg = _make_message(sender_label="anyone", subject="anything")
        assert _matches_filters(msg) is True

    def test_sender_contains_match(self):
        msg = _make_message(sender_label="Alice <alice@example.com>")
        assert _matches_filters(msg, sender_contains="alice") is True

    def test_sender_contains_miss(self):
        msg = _make_message(sender_label="Bob <bob@example.com>")
        assert _matches_filters(msg, sender_contains="alice") is False

    def test_subject_contains_match(self):
        msg = _make_message(subject="Meeting tomorrow")
        assert _matches_filters(msg, subject_contains="meeting") is True

    def test_subject_contains_miss(self):
        msg = _make_message(subject="Hello world")
        assert _matches_filters(msg, subject_contains="meeting") is False

    def test_unread_only_filters_read(self):
        read = _make_message(is_read=True)
        unread = _make_message(is_read=False)
        assert _matches_filters(read, unread_only=True) is False
        assert _matches_filters(unread, unread_only=True) is True

    def test_received_after_filters_old(self):
        old = _make_message(received_datetime="2026-01-01T00:00:00Z")
        new = _make_message(received_datetime="2026-06-01T00:00:00Z")
        assert _matches_filters(old, received_after="2026-03-01T00:00:00Z") is False
        assert _matches_filters(new, received_after="2026-03-01T00:00:00Z") is True



class TestFoldersMapping:
    def test_known_aliases(self):
        assert FOLDERS["inbox"] == "inbox"
        assert FOLDERS["sent"] == "sentitems"
        assert FOLDERS["sentitems"] == "sentitems"
        assert FOLDERS["deleted"] == "deleteditems"
        assert FOLDERS["junk"] == "junkemail"
        assert FOLDERS["archive"] == "archive"
        assert FOLDERS["drafts"] == "drafts"
