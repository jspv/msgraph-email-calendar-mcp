"""Outlook mail operations: list, get, search, move, delete, and bulk-manage messages."""

from __future__ import annotations

import hashlib
import re
from datetime import datetime, timezone

from .config import settings
from .errors import GraphRequestError
from .graph import GraphClient, validate_path_segment
from .models import (
    AttachmentDetail,
    AttachmentSummary,
    DraftPreview,
    MailFolderSummary,
    MailMessageDetail,
    MailMessageSummary,
    _address_label,
    _clean_text_snippet,
    _flag_status,
    _format_datetime_label,
    _parse_utc,
    _recipient_labels,
)


#: Maps friendly folder names to their Graph API well-known folder IDs.
FOLDERS = {
    "inbox": "inbox",
    "drafts": "drafts",
    "sent": "sentitems",
    "sentitems": "sentitems",
    "archive": "archive",
    "deleted": "deleteditems",
    "deleteditems": "deleteditems",
    "junk": "junkemail",
    "junkemail": "junkemail",
}




def list_folders(account_id: str | None = None, include_hidden: bool = False) -> list[MailFolderSummary]:
    """Return all mail folders, optionally including hidden ones."""
    client = GraphClient(account_id)
    params = {
        "$top": 100,
        "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount,isHidden",
    }
    items = client.paginate("/me/mailFolders", params=params, limit=100)
    output: list[MailFolderSummary] = []
    for item in items:
        if not include_hidden and item.get("isHidden"):
            continue
        display_name = item.get("displayName") or ""
        total_item_count = item.get("totalItemCount")
        unread_item_count = item.get("unreadItemCount")
        child_folder_count = item.get("childFolderCount")
        folder_bits = [display_name]
        if total_item_count is not None:
            folder_bits.append(f"{total_item_count} total")
        if unread_item_count is not None:
            folder_bits.append(f"{unread_item_count} unread")
        if child_folder_count:
            folder_bits.append(f"{child_folder_count} child folders")
        output.append(
            MailFolderSummary(
                id=item["id"],
                display_name=display_name,
                total_item_count=total_item_count,
                unread_item_count=unread_item_count,
                child_folder_count=child_folder_count,
                summary=" • ".join(folder_bits),
            )
        )
    return output



def _sender_parts(payload: dict) -> tuple[str | None, str | None]:
    sender = payload.get("from", {}) or {}
    email = sender.get("emailAddress", {}) or {}
    return email.get("name"), email.get("address")



#: Graph's follow-up flag states. Read by the list/detail paths and the bulk
#: filter, written by ``flag_message`` -- the same vocabulary both ways, so a
#: value read off a message feeds straight back into a write.
_VALID_FLAG_STATUSES = {"flagged", "complete", "notFlagged"}


#: Columns selected for list-level message summaries. ``conversationId`` lets a
#: client thread/group messages without a follow-up get_message per item.
#:
#: ``toRecipients``/``ccRecipients`` are requested for **every** folder, not just
#: Sent Items. Restricting them to sent mail would drop the two signals that make
#: them worth fetching on the Inbox: which alias a message was delivered to (when
#: the owner uses a per-vendor address), and whether the owner was addressed
#: directly or merely copied. That costs two extra arrays per row on a
#: whole-folder scan; correct context beats a lighter page. If it ever does hurt,
#: the answer is a leaner projection (addresses only), not dropping them.
_SUMMARY_SELECT = (
    "id",
    "internetMessageId",
    "subject",
    "from",
    "toRecipients",
    "ccRecipients",
    "receivedDateTime",
    "isRead",
    "hasAttachments",
    "conversationId",
    "flag",
    "categories",
    "bodyPreview",
)


def _sanitize_select(fields: list[str]) -> list[str]:
    """Validate caller-supplied ``$select`` field names; always force ``id`` in.

    Guards the OData ``$select`` against injection: each entry must be a simple
    Graph property path (letters/digits/dot/slash). ``id`` is always present
    because summaries and de-duplication depend on it.
    """
    clean: list[str] = []
    for field in fields:
        name = str(field).strip()
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9./]*", name):
            raise ValueError(f"invalid field name: {field!r}")
        if name not in clean:
            clean.append(name)
    if "id" not in clean:
        clean.insert(0, "id")
    return clean


def _odata_string(value: str) -> str:
    """Quote *value* as an OData string literal, escaping embedded quotes.

    Unlike dates and flag states -- which are validated against a parser and an
    enum respectively -- a category name is free text that lands directly in a
    ``$filter``. OData escapes a single quote by doubling it; without that, a
    name like ``Bob's stuff`` closes the literal early and the remainder is
    parsed as operators.
    """
    return "'" + value.replace("'", "''") + "'"


def _validate_iso(value: str, label: str = "since") -> str:
    """Validate that *value* is an ISO-8601 datetime and return it unchanged.

    Parsing also guards the ``$filter`` interpolation: a value that parses as a
    datetime cannot smuggle OData operators. The caller's exact spelling is
    returned so the filter carries the offset they asked for.
    """
    _parse_utc(value, label)
    return value


def _message_summary(item: dict) -> MailMessageSummary:
    sender_name, sender_email = _sender_parts(item)
    sender_label = _address_label(item.get("from"))
    to_recipients = item.get("toRecipients") or []
    cc_recipients = item.get("ccRecipients") or []
    to_labels = _recipient_labels(to_recipients)
    cc_labels = _recipient_labels(cc_recipients)
    flag_status = _flag_status(item)
    categories = item.get("categories") or []
    received_datetime = item.get("receivedDateTime")
    received_label = _format_datetime_label(received_datetime)
    body_preview = item.get("bodyPreview")
    body_preview_clean = _clean_text_snippet(body_preview)
    subject = item.get("subject") or "(no subject)"
    status_bits: list[str] = []
    if not bool(item.get("isRead", False)):
        status_bits.append("unread")
    if bool(item.get("hasAttachments", False)):
        status_bits.append("attachments")
    # Only mention a flag when there is one -- most mail is unflagged, and
    # saying so on every row is noise that crowds out the body preview.
    if flag_status == "flagged":
        status_bits.append("flagged")
    elif flag_status == "complete":
        status_bits.append("follow-up done")
    summary_parts = [subject]
    meta_bits = [bit for bit in [sender_label, received_label] if bit]
    if meta_bits:
        summary_parts.append("from " + " • ".join(meta_bits) if sender_label else " • ".join(meta_bits))
    # Recipients go in the summary string so a model reading a list result sees
    # them without inspecting fields -- the whole point on a Sent Items scan,
    # where every row shares the same sender.
    if to_labels:
        summary_parts.append("to " + ", ".join(to_labels))
    if cc_labels:
        summary_parts.append("cc " + ", ".join(cc_labels))
    if status_bits:
        summary_parts.append(f"[{', '.join(status_bits)}]")
    if categories:
        summary_parts.append("categories: " + ", ".join(categories))
    if body_preview_clean:
        summary_parts.append(body_preview_clean)
    return MailMessageSummary(
        id=item["id"],
        internet_message_id=item.get("internetMessageId"),
        subject=item.get("subject"),
        sender_name=sender_name,
        sender_email=sender_email,
        received_datetime=received_datetime,
        received_label=received_label,
        sender_label=sender_label,
        to_recipients=to_recipients,
        to_recipient_labels=to_labels,
        cc_recipients=cc_recipients,
        cc_recipient_labels=cc_labels,
        is_read=bool(item.get("isRead", False)),
        has_attachments=bool(item.get("hasAttachments", False)),
        flag_status=flag_status,
        categories=categories,
        conversation_id=item.get("conversationId"),
        body_preview=body_preview,
        summary=" — ".join(summary_parts),
    )



def list_messages(
    account_id: str | None = None,
    folder: str = "inbox",
    limit: int = 10,
    include_body_preview: bool = True,
    since: str | None = None,
    until: str | None = None,
    flag_status: str | None = None,
    category: str | None = None,
    fields: list[str] | None = None,
) -> list[MailMessageSummary]:
    """List recent messages from *folder*, newest first.

    *since* (ISO-8601) applies a server-side ``$filter=receivedDateTime ge …``
    so the time window is narrowed by Graph rather than by over-fetching.
    *fields* overrides the selected columns (``id`` is always forced in) for
    callers that want a leaner or extended payload; when omitted, the default
    summary column set is used. Note that only columns backing
    ``MailMessageSummary`` surface in the result — omitting one (e.g. ``from``)
    yields a summary with that value ``None``.
    """
    client = GraphClient(account_id)
    folder_id = FOLDERS.get(folder.lower(), folder)
    validate_path_segment(folder_id, "folder")
    if fields is not None:
        select_fields = _sanitize_select(fields)
    else:
        select_fields = [
            f for f in _SUMMARY_SELECT if f != "bodyPreview" or include_body_preview
        ]
    params: dict[str, object] = {
        "$top": min(limit, 50),
        "$orderby": "receivedDateTime desc",
        "$select": ",".join(select_fields),
    }
    date_clauses: list[str] = []
    if since is not None:
        date_clauses.append(f"receivedDateTime ge {_validate_iso(since, 'since')}")
    if until is not None:
        date_clauses.append(f"receivedDateTime le {_validate_iso(until, 'until')}")
    if category:
        # Confirmed live: Exchange accepts a category restriction *with* an
        # `$orderby`, unlike the flag restriction below, so the sort stays.
        date_clauses.append(f"categories/any(c:c eq {_odata_string(category)})")
    if flag_status is not None:
        if flag_status not in _VALID_FLAG_STATUSES:
            raise ValueError(f"flag_status must be one of {sorted(_VALID_FLAG_STATUSES)}")
        # Server-side so "show me my flagged mail" is one request rather than a
        # full scan filtered after the fetch.
        date_clauses.append(f"flag/flagStatus eq '{flag_status}'")
        # Exchange refuses a flag restriction combined with a sort: the pairing
        # returns "The restriction or sort order is too complex for this
        # operation". Verified against the live API -- the same filter without
        # `$orderby` succeeds, including alongside a receivedDateTime clause.
        # Ordering is restored below rather than given up.
        params.pop("$orderby", None)
    if date_clauses:
        params["$filter"] = " and ".join(date_clauses)
    items = client.paginate(
        f"/me/mailFolders/{folder_id}/messages",
        params=params,
        limit=limit,
    )
    if flag_status is not None:
        # Graph returned these unsorted, so sort here to keep the newest-first
        # contract. Note this orders the rows that came back; it cannot make
        # `limit` select the newest ones, since the server chose them unsorted.
        items.sort(key=lambda i: i.get("receivedDateTime") or "", reverse=True)
    return [_message_summary(item) for item in items]



def get_message(account_id: str | None, message_id: str) -> MailMessageDetail:
    """Fetch the full detail of a single message including body content."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    item = client.request(
        "GET",
        f"/me/messages/{message_id}",
        params={
            "$select": "id,internetMessageId,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,importance,flag,categories,bodyPreview,body",
        },
    ) or {}
    sender = item.get("from")
    to_recipients = item.get("toRecipients") or []
    cc_recipients = item.get("ccRecipients") or []
    received_datetime = item.get("receivedDateTime")
    received_label = _format_datetime_label(received_datetime)
    body_preview = item.get("bodyPreview")
    body_preview_clean = _clean_text_snippet(body_preview)
    sender_label = _address_label(sender)
    to_labels = _recipient_labels(to_recipients)
    cc_labels = _recipient_labels(cc_recipients)
    summary_parts = [item.get("subject") or "(no subject)"]
    if sender_label:
        summary_parts.append(f"from {sender_label}")
    if received_label:
        summary_parts.append(received_label)
    if body_preview_clean:
        summary_parts.append(body_preview_clean)
    return MailMessageDetail(
        id=item["id"],
        internet_message_id=item.get("internetMessageId"),
        subject=item.get("subject"),
        sender=sender,
        sender_label=sender_label,
        to_recipients=to_recipients,
        to_recipient_labels=to_labels,
        cc_recipients=cc_recipients,
        cc_recipient_labels=cc_labels,
        received_datetime=received_datetime,
        received_label=received_label,
        is_read=bool(item.get("isRead", False)),
        has_attachments=bool(item.get("hasAttachments", False)),
        importance=item.get("importance"),
        flag_status=_flag_status(item),
        categories=item.get("categories") or [],
        body_preview=body_preview,
        body_content_type=(item.get("body") or {}).get("contentType"),
        body_content=(item.get("body") or {}).get("content"),
        body_preview_clean=body_preview_clean,
        summary=" — ".join(summary_parts),
    )



def search_messages(account_id: str | None, query: str, limit: int = 10) -> list[MailMessageSummary]:
    """Search messages via the Graph ``$search`` OData parameter."""
    client = GraphClient(account_id)
    safe_query = query.replace('"', '')
    items = client.paginate(
        "/me/messages",
        params={
            "$search": f'"{safe_query}"',
            "$top": min(limit, 50),
            "$select": ",".join(_SUMMARY_SELECT),
        },
        limit=limit,
    )
    return [_message_summary(item) for item in items]



def mark_message_read(account_id: str | None, message_id: str, is_read: bool = True) -> dict[str, object]:
    """Toggle the read/unread flag on a message. Requires Mail.ReadWrite."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    client.request("PATCH", f"/me/messages/{message_id}", json_body={"isRead": is_read})
    return {
        "ok": True,
        "message_id": message_id,
        "is_read": is_read,
        "action": "mark_read" if is_read else "mark_unread",
    }



def move_message(account_id: str | None, message_id: str, destination: str) -> dict[str, object]:
    """Move a message to *destination* folder. Requires Mail.ReadWrite."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    destination_id = FOLDERS.get(destination.lower(), destination)
    response = client.request(
        "POST",
        f"/me/messages/{message_id}/move",
        json_body={"destinationId": destination_id},
    ) or {}
    return {
        "ok": True,
        "message_id": message_id,
        "destination": destination,
        "destination_id": destination_id,
        "moved_message_id": response.get("id"),
        "action": "move",
    }



def delete_message(account_id: str | None, message_id: str, permanent: bool = False) -> dict[str, object]:
    """Delete a message. Moves to Deleted Items by default; permanently deletes if *permanent* is True."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    if permanent:
        client.request("DELETE", f"/me/messages/{message_id}")
        return {
            "ok": True,
            "message_id": message_id,
            "permanent": True,
            "action": "delete",
        }
    response = client.request(
        "POST",
        f"/me/messages/{message_id}/move",
        json_body={"destinationId": "deleteditems"},
    ) or {}
    return {
        "ok": True,
        "message_id": message_id,
        "permanent": False,
        "destination": "deleteditems",
        "moved_message_id": response.get("id"),
        "action": "move_to_deleted",
    }



def _matches_filters(
    item: MailMessageSummary,
    *,
    sender_contains: str | None = None,
    subject_contains: str | None = None,
    recipient_contains: str | None = None,
    flag_status: str | None = None,
    category: str | None = None,
    received_after: str | datetime | None = None,
    unread_only: bool = False,
) -> bool:
    """Return True if *item* passes all specified filter criteria.

    *received_after* may be an ISO-8601 string or an already-parsed datetime;
    callers scanning a large folder should pre-parse once (see
    ``bulk_manage_messages_multi_pass``) rather than re-parsing per message. It
    is retained as a client-side check even though the scan now also bounds the
    window server-side, so this function stays correct for callers that do not.

    *recipient_contains* matches across **both** ``to`` and ``cc``. Splitting
    them would force a caller who just wants "anything addressed to my vendor
    alias" to make the same query twice; the raw ``to_recipients`` /
    ``cc_recipients`` fields remain available when the distinction matters.
    """
    sender_label = (item.sender_label or "").lower()
    subject = (item.subject or "").lower()

    if sender_contains and sender_contains.lower() not in sender_label:
        return False
    if subject_contains and subject_contains.lower() not in subject:
        return False
    if recipient_contains:
        needle = recipient_contains.lower()
        haystack = " ".join(item.to_recipient_labels + item.cc_recipient_labels).lower()
        if needle not in haystack:
            return False
    if flag_status and item.flag_status != flag_status:
        # A message whose flag column was never selected reports None, which is
        # not the same as notFlagged -- so it cannot satisfy any flag filter.
        return False
    if category:
        needle = category.lower()
        if not any(needle == c.lower() for c in item.categories):
            return False
    if unread_only and item.is_read:
        return False
    if received_after and item.received_datetime:
        cutoff = (
            received_after
            if isinstance(received_after, datetime)
            else _parse_utc(received_after, "received_after")
        )
        if cutoff.tzinfo is None:
            cutoff = cutoff.replace(tzinfo=timezone.utc)
        actual = _parse_utc(item.received_datetime, "receivedDateTime")
        if actual < cutoff:
            return False
    return True



_BULK_ACTIONS = {"delete", "move", "mark_read", "mark_unread"}

#: Largest page Microsoft Graph will return for a messages collection. Scanning
#: at this size keeps a full-folder walk to a handful of round-trips.
GRAPH_MAX_PAGE_SIZE = 1000


def _folder_total_count(client: GraphClient, folder_id: str) -> int | None:
    """Return the folder's ``totalItemCount``, or ``None`` if unavailable.

    Used only to give the caller a scale anchor (scanned vs. total); it is a
    live figure and may already be stale by the time it is read.
    """
    try:
        payload = client.request(
            "GET",
            f"/me/mailFolders/{folder_id}",
            params={"$select": "totalItemCount"},
        ) or {}
    except GraphRequestError:
        return None
    return payload.get("totalItemCount")


#: Bulk actions that cannot be casually undone, and so require a confirm token.
#: ``mark_read`` / ``mark_unread`` are excluded: they are trivially reversible,
#: so gating them would be friction with no safety payoff.
_GATED_BULK_ACTIONS = {"delete", "move"}


def _confirm_token(action: str, destination: str | None, ids: list[str]) -> str:
    """Derive the confirmation token for a specific matched set.

    Bound to the *matched message ids* rather than to the filter arguments, so a
    token can never authorise a set the caller did not actually see previewed.
    Deriving it -- rather than storing it -- means no server state, no salt and
    no clock, which keeps it valid across a Lambda cold start between the
    preview call and the live one.

    The count prefix is load-bearing: the server remembers nothing, so without it
    a mismatch could only report "the set changed" and never "42 became 43". The
    count is not a secret, and carrying it is what makes the error actionable.
    """
    payload = "|".join([action, destination or "", *sorted(ids)])
    digest = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    return f"{len(ids)}-{digest[:12]}"


def _collect_matches(
    client: GraphClient,
    folder_id: str,
    *,
    filter_kwargs: dict[str, object],
    scan_limit: int | None,
    received_after_filter: str | None = None,
    received_before_filter: str | None = None,
    category_filter: str | None = None,
) -> dict[str, object]:
    """Non-mutating newest-first scan collecting messages that pass the filters.

    Walks the whole folder when *scan_limit* is ``None``; otherwise stops once
    *scan_limit* distinct messages have been inspected. Pages at
    ``GRAPH_MAX_PAGE_SIZE`` (or the remaining budget, whichever is smaller) so a
    full-folder scan costs only a handful of round-trips.

    Pagination is anchored to ``receivedDateTime`` via a ``le`` cursor rather
    than to a ``$skip`` offset. On a live mailbox this matters: new mail
    arriving at the top does not shift the cursor, and messages removed by
    rules below the cursor simply drop out -- neither causes the offset-drift
    skips that ``$skip`` paging suffers. Boundary ties on ``receivedDateTime``
    are absorbed by using ``le`` and de-duplicating on message id.

    Date bounds travel to Graph rather than being applied after the fetch. That
    is what makes a date-scoped query O(window) instead of O(folder): asking for
    one month of a 50k-message mailbox used to page all 50k rows. The upper bound
    needs no extra machinery -- paging already anchors on ``receivedDateTime le``,
    so *received_before_filter* is simply the initial cursor and the scan starts
    inside the window instead of at the newest message.

    Returns a dict with ``matches`` (summaries), ``scanned`` (distinct
    messages inspected), ``passes`` (pages fetched), ``exhausted`` (reached the
    end of the folder or window), and ``stop_reason``.
    """
    select = ",".join(_SUMMARY_SELECT)

    matches: list[MailMessageSummary] = []
    seen_ids: set[str] = set()
    scanned = 0
    passes = 0
    # The upper bound *is* the starting cursor: one `le` clause serves both.
    cursor: str | None = received_before_filter
    exhausted = False
    windowed = bool(received_after_filter or received_before_filter)
    stop_reason = "window_exhausted" if windowed else "folder_exhausted"

    while True:
        if scan_limit is not None and scanned >= scan_limit:
            stop_reason = "scan_limit_reached"
            break

        if scan_limit is None:
            top = GRAPH_MAX_PAGE_SIZE
        else:
            top = max(1, min(GRAPH_MAX_PAGE_SIZE, scan_limit - scanned))

        params: dict[str, object] = {
            "$top": top,
            "$orderby": "receivedDateTime desc",
            "$select": select,
        }
        clauses: list[str] = []
        if cursor is not None:
            clauses.append(f"receivedDateTime le {cursor}")
        if received_after_filter is not None:
            # Re-applied every page: the cursor rewrites the `le` half each time,
            # and dropping the `ge` half would let page two scan past the window.
            clauses.append(f"receivedDateTime ge {received_after_filter}")
        if category_filter is not None:
            # Also re-applied per page, and safe to combine with the
            # `receivedDateTime desc` sort the cursor depends on -- which is why
            # categories can be pushed server-side here while flags cannot.
            clauses.append(f"categories/any(c:c eq {_odata_string(category_filter)})")
        if clauses:
            params["$filter"] = " and ".join(clauses)

        payload = client.request(
            "GET", f"/me/mailFolders/{folder_id}/messages", params=params
        ) or {}
        items = payload.get("value", [])
        # Graph may return fewer items than ``$top`` and still have more to
        # give, signalled by ``@odata.nextLink``. A short page is therefore
        # only the folder end when Graph also says there is nothing after it;
        # trusting page length alone would report a partial scan as a true
        # folder total.
        has_more = bool(payload.get("@odata.nextLink"))
        passes += 1

        new_in_page = 0
        oldest = cursor
        for item in items:
            message_id = item.get("id")
            if message_id in seen_ids:
                continue
            seen_ids.add(message_id)
            new_in_page += 1
            scanned += 1
            summary = _message_summary(item)
            received = item.get("receivedDateTime")
            if received and (oldest is None or received < oldest):
                oldest = received
            if _matches_filters(summary, **filter_kwargs):
                matches.append(summary)

        if len(items) < top and not has_more:
            exhausted = True
            stop_reason = "window_exhausted" if windowed else "folder_exhausted"
            break
        if new_in_page == 0 or oldest == cursor:
            # A full page of messages sharing the cursor timestamp: cannot
            # advance without a secondary sort key. Stop rather than loop.
            stop_reason = "cursor_stalled"
            break
        cursor = oldest

    return {
        "matches": matches,
        "scanned": scanned,
        "passes": passes,
        "exhausted": exhausted,
        "stop_reason": stop_reason,
    }


def bulk_manage_messages_multi_pass(
    account_id: str | None = None,
    *,
    folder: str = "inbox",
    sender_contains: str | None = None,
    subject_contains: str | None = None,
    recipient_contains: str | None = None,
    flag_status: str | None = None,
    category: str | None = None,
    received_after: str | None = None,
    received_before: str | None = None,
    unread_only: bool = False,
    action: str = "delete",
    destination: str | None = None,
    scan_limit: int | None = None,
    dry_run: bool = True,
    confirm_token: str | None = None,
) -> dict[str, object]:
    """Scan a folder newest-first and act on matching messages.

    With *scan_limit* ``None`` (the default) the entire folder is scanned, so
    "act on all messages matching X" covers the whole folder rather than a
    window. A positive *scan_limit* caps how many messages are inspected and is
    honored exactly -- never silently clamped.

    Collection and action are two separate phases: the scan completes without
    mutating anything, then the action is applied to the collected message ids.
    This is deliberate. The previous implementation deleted/moved matches while
    still paginating with ``$skip``, so its own mutations shifted the offset and
    silently skipped messages -- and its dry-run preview did not match what the
    live run would touch.

    Because the mailbox is never static (mail arrives, rules move and delete
    messages at any moment), the returned counts are a point-in-time snapshot,
    **not** a guarantee. Re-running will legitimately see a different set.
    Callers should read the reported fields rather than infer completeness:

    * ``scanned`` -- distinct messages inspected this run.
    * ``matched`` -- how many passed the filters.
    * ``total_in_folder`` -- folder size, as a scale anchor.
    * ``truncated`` -- ``False`` only when the scan reached the end of the
      folder; ``True`` if *scan_limit* cut it short, so more may exist deeper.
    * ``stop_reason`` -- ``folder_exhausted`` | ``window_exhausted`` |
      ``scan_limit_reached`` | ``cursor_stalled``. ``window_exhausted`` means a
      date bound was given and the scan covered all of it -- complete coverage of
      what was asked, but not of the folder, so the two are reported separately.

    *received_after* / *received_before* are applied by Graph, not after the
    fetch, so a date-scoped query costs O(window) rather than O(folder). This is
    what makes reaching old mail cheap: without it, one month of a 50k-message
    mailbox means paging all 50k rows.

    In a live (non-dry-run) run, messages that a rule or another client moved
    or deleted between collection and action are reported as ``already_gone``
    rather than raised as errors.

    ``delete`` and ``move`` additionally require *confirm_token*, the value
    returned by a preceding dry-run. The live run re-derives the token from its
    own scan and refuses if it differs, so a preview of 42 messages cannot be
    used to act on a set that has since become something else. Reversible
    actions (``mark_read`` / ``mark_unread``) need no token.
    """
    if scan_limit is not None and scan_limit < 1:
        raise ValueError("scan_limit must be a positive integer, or None to scan the whole folder")
    if action not in _BULK_ACTIONS:
        raise ValueError(
            "action must be one of: delete, move, mark_read, mark_unread"
        )
    if action == "move" and not destination:
        raise ValueError("destination is required when action='move'")
    if flag_status is not None and flag_status not in _VALID_FLAG_STATUSES:
        raise ValueError(f"flag_status must be one of {sorted(_VALID_FLAG_STATUSES)}")

    # Validate up front so a malformed or inverted window fails before the first
    # Graph call. An inverted window would otherwise return zero matches, which
    # reads as "nothing there" rather than "you asked for an empty range".
    after_filter = _validate_iso(received_after, "received_after") if received_after else None
    before_filter = (
        _validate_iso(received_before, "received_before") if received_before else None
    )
    if after_filter and before_filter:
        if _parse_utc(after_filter, "received_after") > _parse_utc(
            before_filter, "received_before"
        ):
            raise ValueError(
                f"received_after ({received_after}) is later than received_before "
                f"({received_before}), so the window is empty."
            )

    client = GraphClient(account_id)
    folder_id = FOLDERS.get(folder.lower(), folder)
    validate_path_segment(folder_id, "folder")

    filter_kwargs: dict[str, object] = {
        "sender_contains": sender_contains,
        "subject_contains": subject_contains,
        "recipient_contains": recipient_contains,
        "flag_status": flag_status,
        "category": category,
        # Parsed once up front so a malformed cutoff fails before any Graph
        # call, and so a whole-folder scan does not re-parse it per message.
        "received_after": (
            _parse_utc(received_after, "received_after") if received_after else None
        ),
        "unread_only": unread_only,
    }

    collected = _collect_matches(
        client,
        folder_id,
        filter_kwargs=filter_kwargs,
        scan_limit=scan_limit,
        received_after_filter=after_filter,
        received_before_filter=before_filter,
        category_filter=category,
    )
    matches: list[MailMessageSummary] = collected["matches"]
    truncated = not collected["exhausted"]

    report: dict[str, object] = {
        "ok": True,
        "dry_run": dry_run,
        "action": action,
        "folder": folder,
        "scanned": collected["scanned"],
        "matched": len(matches),
        "match_count": len(matches),  # backwards-compatible alias
        "total_in_folder": _folder_total_count(client, folder_id),
        "passes": collected["passes"],
        "truncated": truncated,
        "stop_reason": collected["stop_reason"],
    }

    matched_ids = [item.id for item in matches]
    gated = action in _GATED_BULK_ACTIONS

    if dry_run:
        report["matches"] = [
            {
                "id": item.id,
                "subject": item.subject,
                "sender": item.sender_label,
                "received": item.received_datetime,
                "summary": item.summary,
            }
            for item in matches
        ]
        report["results"] = None
        if gated:
            # Only gated actions carry a token, so its presence is itself the
            # signal that confirmation is required.
            report["confirm_token"] = _confirm_token(action, destination, matched_ids)
        return report

    if gated:
        expected = _confirm_token(action, destination, matched_ids)
        if not confirm_token:
            raise ValueError(
                f"Bulk {action} requires confirmation. Re-run with dry_run=True to "
                f"review the {len(matched_ids)} matching message(s), then pass the "
                f"confirm_token it returns. Current token: {expected}"
            )
        if confirm_token != expected:
            previewed = confirm_token.split("-", 1)[0]
            raise ValueError(
                f"confirm_token does not match the current scan: it described "
                f"{previewed} message(s), this scan matched {len(matched_ids)}. "
                f"The mailbox is live, so re-check the preview before acting. "
                f"Confirm with the new token: {expected}"
            )

    results: list[dict[str, object]] = []
    acted = 0
    already_gone = 0
    failed = 0
    for item in matches:
        try:
            if action == "delete":
                result = delete_message(account_id, item.id, permanent=False)
            elif action == "mark_read":
                result = mark_message_read(account_id, item.id, True)
            elif action == "mark_unread":
                result = mark_message_read(account_id, item.id, False)
            else:  # move (destination validated above)
                result = move_message(account_id, item.id, destination)
        except GraphRequestError as exc:
            if exc.status_code in {404, 410}:
                already_gone += 1
                status = "already_gone"
            else:
                failed += 1
                status = "error"
            results.append(
                {
                    "ok": False,
                    "message_id": item.id,
                    "status": status,
                    "error": str(exc),
                    "subject": item.subject,
                    "sender": item.sender_label,
                }
            )
            continue
        result["subject"] = item.subject
        result["sender"] = item.sender_label
        results.append(result)
        acted += 1

    report["matches"] = None
    report["results"] = results
    report["acted"] = acted
    report["already_gone"] = already_gone
    report["failed"] = failed
    report["match_count"] = acted  # live-run compat: count of actions performed
    return report


def create_draft(
    account_id: str | None = None,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    send_as: str | None = None,
) -> dict:
    """Create a draft message without sending it."""
    client = GraphClient(account_id)
    message_body: dict = {
        "subject": subject,
        "body": {"contentType": "text", "content": body},
        "toRecipients": _build_recipients(to),
    }
    if cc:
        message_body["ccRecipients"] = _build_recipients(cc)
    if bcc:
        message_body["bccRecipients"] = _build_recipients(bcc)
    from_field = _build_from(send_as)
    if from_field:
        message_body["from"] = from_field

    draft = client.request("POST", "/me/messages", json_body=message_body) or {}
    preview = _draft_preview(draft)
    return {"ok": True, "draft": preview.model_dump()}


def update_draft(
    account_id: str | None = None,
    *,
    message_id: str,
    to: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    send_as: str | None = None,
) -> dict:
    """Update a draft message.  Only provided fields are changed."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    update: dict = {}
    if subject is not None:
        update["subject"] = subject
    if body is not None:
        update["body"] = {"contentType": "text", "content": body}
    if to is not None:
        update["toRecipients"] = _build_recipients(to)
    if cc is not None:
        update["ccRecipients"] = _build_recipients(cc)
    if bcc is not None:
        update["bccRecipients"] = _build_recipients(bcc)
    from_field = _build_from(send_as)
    if from_field:
        update["from"] = from_field

    result = client.request("PATCH", f"/me/messages/{message_id}", json_body=update) or {}
    preview = _draft_preview(result)
    return {"ok": True, "draft": preview.model_dump()}


def send_draft(
    account_id: str | None = None,
    *,
    message_id: str,
) -> dict:
    """Send a previously created draft message."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    client.request("POST", f"/me/messages/{message_id}/send")
    return {"ok": True, "action": "sent", "message_id": message_id}


def list_attachments(
    account_id: str | None = None,
    *,
    message_id: str,
) -> list[AttachmentSummary]:
    """List attachment metadata for a message."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    payload = client.request(
        "GET",
        f"/me/messages/{message_id}/attachments",
        params={"$select": "id,name,size,contentType,isInline"},
    ) or {"value": []}
    return [
        AttachmentSummary(
            id=item["id"],
            name=item.get("name"),
            size=item.get("size"),
            content_type=item.get("contentType"),
            is_inline=bool(item.get("isInline", False)),
        )
        for item in payload.get("value", [])
    ]


def get_attachment(
    account_id: str | None = None,
    *,
    message_id: str,
    attachment_id: str,
) -> AttachmentDetail:
    """Download a single attachment.  Returns inline base64 if under the size limit."""
    validate_path_segment(message_id, "message_id")
    validate_path_segment(attachment_id, "attachment_id")
    client = GraphClient(account_id)
    item = client.request(
        "GET",
        f"/me/messages/{message_id}/attachments/{attachment_id}",
    ) or {}
    size = item.get("size") or 0
    if size > settings.max_attachment_inline_size:
        return AttachmentDetail(
            id=item.get("id", attachment_id),
            name=item.get("name"),
            size=size,
            content_type=item.get("contentType"),
            is_inline=bool(item.get("isInline", False)),
            content_omitted=True,
            omit_reason=f"Attachment size {size} exceeds limit {settings.max_attachment_inline_size}",
        )
    return AttachmentDetail(
        id=item.get("id", attachment_id),
        name=item.get("name"),
        size=size,
        content_type=item.get("contentType"),
        is_inline=bool(item.get("isInline", False)),
        content_base64=item.get("contentBytes"),
    )


def add_attachment_to_draft(
    account_id: str | None = None,
    *,
    message_id: str,
    name: str,
    content_base64: str,
    content_type: str = "application/octet-stream",
) -> dict:
    """Attach a file to a draft message."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    payload = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentType": content_type,
        "contentBytes": content_base64,
    }
    result = client.request(
        "POST",
        f"/me/messages/{message_id}/attachments",
        json_body=payload,
    ) or {}
    return {
        "ok": True,
        "attachment_id": result.get("id"),
        "name": name,
        "content_type": content_type,
    }


def create_folder(
    account_id: str | None = None,
    *,
    name: str,
    parent_folder_id: str | None = None,
) -> dict:
    """Create a new mail folder, optionally under a parent folder."""
    client = GraphClient(account_id)
    if parent_folder_id:
        validate_path_segment(parent_folder_id, "parent_folder_id")
        path = f"/me/mailFolders/{parent_folder_id}/childFolders"
    else:
        path = "/me/mailFolders"
    result = client.request("POST", path, json_body={"displayName": name}) or {}
    folder = MailFolderSummary(
        id=result.get("id", ""),
        display_name=result.get("displayName", name),
        total_item_count=result.get("totalItemCount"),
        unread_item_count=result.get("unreadItemCount"),
        child_folder_count=result.get("childFolderCount"),
    )
    return {"ok": True, "folder": folder.model_dump()}


def list_child_folders(
    account_id: str | None = None,
    *,
    folder_id: str,
) -> list[MailFolderSummary]:
    """List subfolders of a given mail folder."""
    validate_path_segment(folder_id, "folder_id")
    client = GraphClient(account_id)
    payload = client.request(
        "GET",
        f"/me/mailFolders/{folder_id}/childFolders",
        params={"$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount"},
    ) or {"value": []}
    return [
        MailFolderSummary(
            id=item["id"],
            display_name=item.get("displayName", ""),
            total_item_count=item.get("totalItemCount"),
            unread_item_count=item.get("unreadItemCount"),
            child_folder_count=item.get("childFolderCount"),
        )
        for item in payload.get("value", [])
    ]


def flag_message(
    account_id: str | None = None,
    *,
    message_id: str,
    flag_status: str,
) -> dict:
    """Set follow-up flag on a message.  Values: flagged, complete, notFlagged."""
    if flag_status not in _VALID_FLAG_STATUSES:
        raise ValueError(f"flag_status must be one of {_VALID_FLAG_STATUSES}")
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    client.request("PATCH", f"/me/messages/{message_id}", json_body={"flag": {"flagStatus": flag_status}})
    return {"ok": True, "message_id": message_id, "flag_status": flag_status}


def categorize_message(
    account_id: str | None = None,
    *,
    message_id: str,
    categories: list[str],
) -> dict:
    """Apply color categories to a message."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    client.request("PATCH", f"/me/messages/{message_id}", json_body={"categories": categories})
    return {"ok": True, "message_id": message_id, "categories": categories}


def list_aliases(account_id: str | None = None) -> dict:
    """List email aliases available for the authenticated user.

    Parses ``proxyAddresses`` — ``SMTP:`` (uppercase) is primary,
    ``smtp:`` (lowercase) entries are aliases.
    """
    client = GraphClient(account_id)
    profile = client.request(
        "GET", "/me",
        params={"$select": "mail,proxyAddresses,userPrincipalName"},
    ) or {}
    proxy_addresses = profile.get("proxyAddresses") or []
    primary = profile.get("mail") or profile.get("userPrincipalName")
    aliases: list[str] = []
    all_addresses: list[str] = []
    for addr in proxy_addresses:
        if addr.startswith("SMTP:"):
            primary = addr[5:]
            all_addresses.append(addr[5:])
        elif addr.startswith("smtp:"):
            aliases.append(addr[5:])
            all_addresses.append(addr[5:])
    if primary and primary not in all_addresses:
        all_addresses.insert(0, primary)
    return {
        "primary": primary,
        "aliases": aliases,
        "all_addresses": all_addresses,
    }


def _build_recipients(emails: list[str] | None) -> list[dict]:
    """Convert a list of email strings to Graph recipient format."""
    if not emails:
        return []
    return [{"emailAddress": {"address": e}} for e in emails]


def _build_from(send_as: str | None) -> dict | None:
    """Build a Graph 'from' field from an alias email address."""
    if not send_as:
        return None
    return {"emailAddress": {"address": send_as}}


def _draft_id(payload: dict | None) -> str:
    """Pull the id out of a freshly created draft, or fail with a typed error.

    Graph is expected to echo the created draft, but an empty body would
    otherwise surface as a bare ``KeyError`` from deep inside a dry-run.
    """
    draft_id = (payload or {}).get("id")
    if not draft_id:
        raise GraphRequestError("Microsoft Graph did not return a draft id")
    return draft_id


def _draft_preview(payload: dict) -> DraftPreview:
    """Extract a DraftPreview from a Graph message payload."""
    from_addr = ((payload.get("from") or {}).get("emailAddress") or {}).get("address")
    to_addrs = [
        (r.get("emailAddress") or {}).get("address", "")
        for r in (payload.get("toRecipients") or [])
    ]
    cc_addrs = [
        (r.get("emailAddress") or {}).get("address", "")
        for r in (payload.get("ccRecipients") or [])
    ]
    bcc_addrs = [
        (r.get("emailAddress") or {}).get("address", "")
        for r in (payload.get("bccRecipients") or [])
    ]
    return DraftPreview(
        id=payload.get("id", ""),
        subject=payload.get("subject"),
        from_address=from_addr,
        to_recipients=to_addrs,
        cc_recipients=cc_addrs,
        bcc_recipients=bcc_addrs,
        body_preview=payload.get("bodyPreview"),
    )


def send_message(
    account_id: str | None = None,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    send_as: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Compose and send (or preview) an email message.

    When *dry_run* is True (default), creates a temporary draft, returns
    a preview, then deletes the draft.  Set *dry_run=False* to actually send.
    """
    client = GraphClient(account_id)
    message_body: dict = {
        "subject": subject,
        "body": {"contentType": "text", "content": body},
        "toRecipients": _build_recipients(to),
    }
    if cc:
        message_body["ccRecipients"] = _build_recipients(cc)
    if bcc:
        message_body["bccRecipients"] = _build_recipients(bcc)
    from_field = _build_from(send_as)
    if from_field:
        message_body["from"] = from_field

    if dry_run:
        draft = client.request("POST", "/me/messages", json_body=message_body) or {}
        draft_id = _draft_id(draft)
        try:
            preview = _draft_preview(draft)
            preview.message = "Dry-run: message NOT sent. Set dry_run=False to send."
        finally:
            # The preview draft is scratch state; never leave it in Drafts.
            client.request("DELETE", f"/me/messages/{draft_id}")
        return {"ok": True, "dry_run": True, "preview": preview.model_dump()}

    client.request("POST", "/me/sendMail", json_body={"message": message_body})
    return {
        "ok": True,
        "dry_run": False,
        "action": "sent",
        "to": to,
        "subject": subject,
    }


def reply_to_message(
    account_id: str | None = None,
    *,
    message_id: str,
    body: str,
    reply_all: bool = False,
    send_as: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Reply to a message.  Dry-run creates a temporary reply draft for preview."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    action = "replyAll" if reply_all else "reply"

    if dry_run:
        create_action = "createReplyAll" if reply_all else "createReply"
        draft = client.request("POST", f"/me/messages/{message_id}/{create_action}") or {}
        draft_id = _draft_id(draft)
        try:
            update_body: dict = {"body": {"contentType": "text", "content": body}}
            from_field = _build_from(send_as)
            if from_field:
                update_body["from"] = from_field
            client.request("PATCH", f"/me/messages/{draft_id}", json_body=update_body)
            refreshed = client.request(
                "GET", f"/me/messages/{draft_id}",
                params={"$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,bodyPreview"},
            ) or draft
            preview = _draft_preview(refreshed)
            preview.message = "Dry-run: reply NOT sent. Set dry_run=False to send."
        finally:
            client.request("DELETE", f"/me/messages/{draft_id}")
        return {"ok": True, "dry_run": True, "preview": preview.model_dump()}

    json_body: dict = {"comment": body}
    if send_as:
        json_body["message"] = {"from": _build_from(send_as)}
    client.request("POST", f"/me/messages/{message_id}/{action}", json_body=json_body)
    return {"ok": True, "dry_run": False, "action": action, "message_id": message_id}


def forward_message(
    account_id: str | None = None,
    *,
    message_id: str,
    to: list[str],
    body: str | None = None,
    send_as: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Forward a message.  Dry-run creates a temporary forward draft for preview."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)

    if dry_run:
        draft = client.request("POST", f"/me/messages/{message_id}/createForward") or {}
        draft_id = _draft_id(draft)
        try:
            update_body: dict = {"toRecipients": _build_recipients(to)}
            if body:
                update_body["body"] = {"contentType": "text", "content": body}
            from_field = _build_from(send_as)
            if from_field:
                update_body["from"] = from_field
            client.request("PATCH", f"/me/messages/{draft_id}", json_body=update_body)
            refreshed = client.request(
                "GET", f"/me/messages/{draft_id}",
                params={"$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,bodyPreview"},
            ) or draft
            preview = _draft_preview(refreshed)
            preview.message = "Dry-run: forward NOT sent. Set dry_run=False to send."
        finally:
            client.request("DELETE", f"/me/messages/{draft_id}")
        return {"ok": True, "dry_run": True, "preview": preview.model_dump()}

    json_body: dict = {
        "toRecipients": _build_recipients(to),
    }
    if body:
        json_body["comment"] = body
    if send_as:
        json_body["message"] = {"from": _build_from(send_as)}
    client.request("POST", f"/me/messages/{message_id}/forward", json_body=json_body)
    return {"ok": True, "dry_run": False, "action": "forward", "message_id": message_id, "to": to}
