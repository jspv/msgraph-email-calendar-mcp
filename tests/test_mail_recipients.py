"""List-level results carry recipients, not just senders (issue #11).

Sender alone is not enough context at list level:

* In **Sent Items** every row is the mailbox owner, so the only field that tells
  the messages apart -- who they went to -- was missing entirely.
* In the **Inbox** the delivery address identifies which vendor or signup a
  message belongs to when the owner uses a per-service alias, and whether the
  owner is in ``to`` versus ``cc`` versus neither is one of the strongest
  available priority signals.

Neither could be recovered by the ``fields`` override, which only widens
``$select`` -- the model dropped the values on the way back.
"""
from __future__ import annotations

from unittest.mock import patch

from msgraph_mcp.mail import _SUMMARY_SELECT, _matches_filters, _message_summary
from msgraph_mcp.models import MailMessageSummary


def _item(**overrides) -> dict:
    item = {
        "id": "m1",
        "subject": "Invoice",
        "from": {"emailAddress": {"name": "Me", "address": "me@example.com"}},
        "toRecipients": [
            {"emailAddress": {"name": "Vendor Billing", "address": "billing@vendor.com"}}
        ],
        "ccRecipients": [
            {"emailAddress": {"name": "Accounts", "address": "accounts@vendor.com"}}
        ],
        "receivedDateTime": "2026-01-02T00:00:00Z",
        "isRead": True,
        "hasAttachments": False,
        "bodyPreview": "",
    }
    item.update(overrides)
    return item


class TestSummarySelect:
    def test_recipients_are_requested_for_every_folder(self):
        # Deliberately unconditional: an Inbox-skipping optimisation would drop
        # exactly the alias and to-vs-cc signals that motivate this.
        assert "toRecipients" in _SUMMARY_SELECT
        assert "ccRecipients" in _SUMMARY_SELECT


class TestSummaryModel:
    def test_to_recipient_labels_are_populated(self):
        summary = _message_summary(_item())
        assert summary.to_recipient_labels == ["Vendor Billing <billing@vendor.com>"]

    def test_cc_recipient_labels_are_populated(self):
        summary = _message_summary(_item())
        assert summary.cc_recipient_labels == ["Accounts <accounts@vendor.com>"]

    def test_raw_to_recipients_are_kept(self):
        summary = _message_summary(_item())
        assert summary.to_recipients[0]["emailAddress"]["address"] == "billing@vendor.com"

    def test_missing_recipients_are_empty_not_none(self):
        summary = _message_summary(_item(toRecipients=None, ccRecipients=None))
        assert summary.to_recipient_labels == []
        assert summary.cc_recipient_labels == []

    def test_recipients_appear_in_the_generated_summary(self):
        # A model reading the list should see them without inspecting fields.
        assert "billing@vendor.com" in _message_summary(_item()).summary

    def test_cc_is_distinguishable_from_to_in_the_summary(self):
        summary = _message_summary(_item()).summary
        assert "to " in summary
        assert "cc " in summary


class TestRecipientContainsFilter:
    def _summary(self, to=(), cc=()):
        return MailMessageSummary(
            id="m1",
            to_recipient_labels=list(to),
            cc_recipient_labels=list(cc),
        )

    def test_matches_a_to_recipient(self):
        item = self._summary(to=["Vendor <billing@vendor.com>"])
        assert _matches_filters(item, recipient_contains="billing@vendor.com")

    def test_matches_a_cc_recipient(self):
        item = self._summary(cc=["Accounts <accounts@vendor.com>"])
        assert _matches_filters(item, recipient_contains="accounts@vendor.com")

    def test_is_case_insensitive_like_sender_contains(self):
        item = self._summary(to=["Vendor <Billing@Vendor.com>"])
        assert _matches_filters(item, recipient_contains="billing@vendor")

    def test_rejects_a_non_match(self):
        item = self._summary(to=["Vendor <billing@vendor.com>"])
        assert not _matches_filters(item, recipient_contains="landlord@")

    def test_rejects_when_there_are_no_recipients_at_all(self):
        assert not _matches_filters(self._summary(), recipient_contains="anyone@")


class TestToolSurface:
    @patch("msgraph_mcp.mail.bulk_manage_messages_multi_pass")
    def test_bulk_tool_forwards_recipient_contains(self, bulk):
        from msgraph_mcp import tools

        tools.bulk_manage_messages(recipient_contains="billing@vendor.com")
        assert bulk.call_args[1]["recipient_contains"] == "billing@vendor.com"
