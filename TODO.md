# TODO

## Attachment support

The `has_attachments` flag is surfaced on messages but there are no tools to work with attachments.

- [ ] `list_attachments(message_id)` — fetch metadata only (`id`, `name`, `size`, `contentType`) via `/me/messages/{id}/attachments?$select=id,name,size,contentType` without pulling content bytes
- [ ] `get_attachment(message_id, attachment_id)` — download a single attachment's content via `/me/messages/{id}/attachments/{attachmentId}/$value`

Note: Graph does not support partial/range fetching of message bodies — `bodyPreview` (~255 chars plain text) is the only built-in truncation. Attachment content must be fetched as a separate call from the message itself.
