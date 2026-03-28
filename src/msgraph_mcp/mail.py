"""Outlook mail operations: list, get, search, move, delete, and bulk-manage messages."""

from __future__ import annotations

from datetime import datetime, timezone

from .config import settings
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


_VALID_FLAG_STATUSES = {"flagged", "complete", "notFlagged"}


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
        preview = _draft_preview(draft)
        preview.message = "Dry-run: message NOT sent. Set dry_run=False to send."
        client.request("DELETE", f"/me/messages/{draft['id']}")
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
        update_body: dict = {"body": {"contentType": "text", "content": body}}
        from_field = _build_from(send_as)
        if from_field:
            update_body["from"] = from_field
        client.request("PATCH", f"/me/messages/{draft['id']}", json_body=update_body)
        refreshed = client.request(
            "GET", f"/me/messages/{draft['id']}",
            params={"$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,bodyPreview"},
        ) or draft
        preview = _draft_preview(refreshed)
        preview.message = "Dry-run: reply NOT sent. Set dry_run=False to send."
        client.request("DELETE", f"/me/messages/{draft['id']}")
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
        update_body: dict = {"toRecipients": _build_recipients(to)}
        if body:
            update_body["body"] = {"contentType": "text", "content": body}
        from_field = _build_from(send_as)
        if from_field:
            update_body["from"] = from_field
        client.request("PATCH", f"/me/messages/{draft['id']}", json_body=update_body)
        refreshed = client.request(
            "GET", f"/me/messages/{draft['id']}",
            params={"$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,bodyPreview"},
        ) or draft
        preview = _draft_preview(refreshed)
        preview.message = "Dry-run: forward NOT sent. Set dry_run=False to send."
        client.request("DELETE", f"/me/messages/{draft['id']}")
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
