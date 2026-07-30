# TODO

## Safety

- [ ] Optional allowlist for `user_id` (shared-calendar targets) and outbound
      recipient domains. See SECURITY_REVIEW.md → "Shared mailboxes and calendars".
- [ ] `MSGRAPH_READ_ONLY=1` kill switch that drops mutating tools regardless of
      token scope.
- [ ] Extend the confirm-token gate to `forward_message` — it forwards
      attachments and is the highest-severity outbound tool, but still acts on a
      single `dry_run=False` call.

## Performance

- [ ] Async tool functions. Everything is synchronous and `time.sleep`s on
      retry backoff, so a throttled whole-folder scan stalls the server process.

## Known limitations

- Graph does not support partial/range fetching of message bodies —
  `bodyPreview` (~255 chars plain text) is the only built-in truncation.
- Bulk paging anchors on `receivedDateTime`. A full page of messages sharing one
  timestamp cannot advance the cursor; the scan stops and reports
  `stop_reason="cursor_stalled"` rather than looping.
- The confirm token and the calendar dry-run defaults make destructive scope
  **visible** before it is applied; they do not stop prompt injection, since an
  injected instruction can walk the two-step flow itself. See SECURITY_REVIEW.md
  → "Remaining risks".

## Done in 0.2.0

- [x] CI on Python 3.11–3.14, plus a no-network guard and a docs-parity test that
      fails the build when the tool tables drift from the registered tools.
- [x] Single pooled `httpx.Client` — a bulk action no longer pays one TLS
      handshake per message.
- [x] `dry_run=True` default for `create_event` / `update_event` / `delete_event`.
- [x] Confirmation token for bulk `delete` / `move` escalation.
