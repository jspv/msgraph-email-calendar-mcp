"""Outlook mail operations: list, get, search, move, delete, and bulk-manage messages."""

from __future__ import annotations

from datetime import datetime, timezone

from .graph import GraphClient, validate_path_segment
from .models import (
    MailFolderSummary,
    MailMessageDetail,
    MailMessageSummary,
    _address_label,
    _clean_text_snippet,
    _format_datetime_label,
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



def _message_summary(item: dict) -> MailMessageSummary:
    sender_name, sender_email = _sender_parts(item)
    sender_label = _address_label(item.get("from"))
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
    summary_parts = [subject]
    meta_bits = [bit for bit in [sender_label, received_label] if bit]
    if meta_bits:
        summary_parts.append("from " + " • ".join(meta_bits) if sender_label else " • ".join(meta_bits))
    if status_bits:
        summary_parts.append(f"[{', '.join(status_bits)}]")
    if body_preview_clean:
        summary_parts.append(body_preview_clean)
    return MailMessageSummary(
        id=item["id"],
        subject=item.get("subject"),
        sender_name=sender_name,
        sender_email=sender_email,
        received_datetime=received_datetime,
        received_label=received_label,
        sender_label=sender_label,
        is_read=bool(item.get("isRead", False)),
        has_attachments=bool(item.get("hasAttachments", False)),
        body_preview=body_preview,
        summary=" — ".join(summary_parts),
    )



def list_messages(
    account_id: str | None = None,
    folder: str = "inbox",
    limit: int = 10,
    include_body_preview: bool = True,
) -> list[MailMessageSummary]:
    """List recent messages from *folder*, newest first."""
    client = GraphClient(account_id)
    folder_id = FOLDERS.get(folder.lower(), folder)
    validate_path_segment(folder_id, "folder")
    select_fields = [
        "id",
        "subject",
        "from",
        "receivedDateTime",
        "isRead",
        "hasAttachments",
    ]
    if include_body_preview:
        select_fields.append("bodyPreview")
    items = client.paginate(
        f"/me/mailFolders/{folder_id}/messages",
        params={
            "$top": min(limit, 50),
            "$orderby": "receivedDateTime desc",
            "$select": ",".join(select_fields),
        },
        limit=limit,
    )
    return [_message_summary(item) for item in items]



def get_message(account_id: str | None, message_id: str) -> MailMessageDetail:
    """Fetch the full detail of a single message including body content."""
    validate_path_segment(message_id, "message_id")
    client = GraphClient(account_id)
    item = client.request(
        "GET",
        f"/me/messages/{message_id}",
        params={
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,importance,bodyPreview,body",
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
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview",
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
    received_after: str | None = None,
    unread_only: bool = False,
) -> bool:
    """Return True if *item* passes all specified filter criteria."""
    sender_label = (item.sender_label or "").lower()
    subject = (item.subject or "").lower()

    if sender_contains and sender_contains.lower() not in sender_label:
        return False
    if subject_contains and subject_contains.lower() not in subject:
        return False
    if unread_only and item.is_read:
        return False
    if received_after and item.received_datetime:
        cutoff = datetime.fromisoformat(received_after.replace("Z", "+00:00"))
        actual = datetime.fromisoformat(item.received_datetime.replace("Z", "+00:00"))
        if actual < cutoff:
            return False
    return True



def bulk_manage_messages(
    account_id: str | None = None,
    *,
    folder: str = "inbox",
    sender_contains: str | None = None,
    subject_contains: str | None = None,
    received_after: str | None = None,
    unread_only: bool = False,
    action: str = "delete",
    destination: str | None = None,
    limit: int = 100,
    dry_run: bool = True,
) -> dict[str, object]:
    """Single-pass bulk operation on filtered messages.

    When *dry_run* is True (the default), returns a preview of matching
    messages without performing any action.
    """
    candidates = list_messages(account_id=account_id, folder=folder, limit=limit)
    matches = [
        item for item in candidates
        if _matches_filters(
            item,
            sender_contains=sender_contains,
            subject_contains=subject_contains,
            received_after=received_after,
            unread_only=unread_only,
        )
    ]

    preview = [
        {
            "id": item.id,
            "subject": item.subject,
            "sender": item.sender_label,
            "received": item.received_datetime,
            "summary": item.summary,
        }
        for item in matches
    ]

    if dry_run:
        return {
            "ok": True,
            "dry_run": True,
            "action": action,
            "match_count": len(matches),
            "matches": preview,
        }

    results: list[dict[str, object]] = []
    for item in matches:
        if action == "delete":
            result = delete_message(account_id, item.id, permanent=False)
        elif action == "mark_read":
            result = mark_message_read(account_id, item.id, True)
        elif action == "mark_unread":
            result = mark_message_read(account_id, item.id, False)
        elif action == "move":
            if not destination:
                raise ValueError("destination is required when action='move'")
            result = move_message(account_id, item.id, destination)
        else:
            raise ValueError("action must be one of: delete, move, mark_read, mark_unread")
        result["subject"] = item.subject
        result["sender"] = item.sender_label
        results.append(result)

    return {
        "ok": True,
        "dry_run": False,
        "action": action,
        "match_count": len(matches),
        "results": results,
    }



def bulk_manage_messages_multi_pass(
    account_id: str | None = None,
    *,
    folder: str = "inbox",
    sender_contains: str | None = None,
    subject_contains: str | None = None,
    received_after: str | None = None,
    unread_only: bool = False,
    action: str = "delete",
    destination: str | None = None,
    limit_per_pass: int = 50,
    max_passes: int = 5,
    dry_run: bool = True,
) -> dict[str, object]:
    """Multi-pass bulk operation that paginates through messages.

    Follows ``@odata.nextLink`` across up to *max_passes* pages,
    applying filters and the chosen *action* to each match.
    Defaults to dry-run mode.
    """
    client = GraphClient(account_id)
    folder_id = FOLDERS.get(folder.lower(), folder)
    validate_path_segment(folder_id, "folder")
    path = f"/me/mailFolders/{folder_id}/messages"
    params = {
        "$top": min(limit_per_pass, 50),
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview",
    }

    next_path: str | None = path
    next_params: dict[str, object] | None = params
    passes = 0
    aggregate_matches: list[dict[str, object]] = []
    aggregate_results: list[dict[str, object]] = []
    seen_ids: set[str] = set()

    while next_path and passes < max_passes:
        passes += 1
        payload = client.request("GET", next_path, params=next_params) or {}
        next_params = None
        items = payload.get("value", [])
        summaries = [_message_summary(item) for item in items]
        matches = [
            item for item in summaries
            if _matches_filters(
                item,
                sender_contains=sender_contains,
                subject_contains=subject_contains,
                received_after=received_after,
                unread_only=unread_only,
            )
        ]

        preview = [
            {
                "id": item.id,
                "subject": item.subject,
                "sender": item.sender_label,
                "received": item.received_datetime,
                "summary": item.summary,
            }
            for item in matches
            if item.id not in seen_ids
        ]

        if dry_run:
            for item in preview:
                seen_ids.add(item["id"])
            aggregate_matches.extend(preview)
        else:
            for item in matches:
                if item.id in seen_ids:
                    continue
                seen_ids.add(item.id)
                if action == "delete":
                    result = delete_message(account_id, item.id, permanent=False)
                elif action == "mark_read":
                    result = mark_message_read(account_id, item.id, True)
                elif action == "mark_unread":
                    result = mark_message_read(account_id, item.id, False)
                elif action == "move":
                    if not destination:
                        raise ValueError("destination is required when action='move'")
                    result = move_message(account_id, item.id, destination)
                else:
                    raise ValueError("action must be one of: delete, move, mark_read, mark_unread")
                result["subject"] = item.subject
                result["sender"] = item.sender_label
                aggregate_results.append(result)

        next_path = payload.get("@odata.nextLink")

    return {
        "ok": True,
        "dry_run": dry_run,
        "action": action,
        "match_count": len(aggregate_matches) if dry_run else len(aggregate_results),
        "matches": aggregate_matches if dry_run else None,
        "results": aggregate_results if not dry_run else None,
        "passes": passes,
    }
