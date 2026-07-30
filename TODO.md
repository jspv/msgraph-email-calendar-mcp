# TODO

## Safety

- [ ] Confirmation token for bulk escalation — require the caller to echo a
      value from the dry-run response before `dry_run=False` acts, so a single
      injected tool call cannot go from preview to folder-wide delete.
- [ ] `dry_run` for calendar writes. `create_event`/`update_event`/`delete_event`
      email attendees the moment they are called, with no preview step.
- [ ] Optional allowlist for `user_id` (shared-calendar targets) and outbound
      recipient domains. See SECURITY_REVIEW.md → "Shared mailboxes and calendars".
- [ ] `MSGRAPH_READ_ONLY=1` kill switch that drops mutating tools regardless of
      token scope.

## Performance

- [ ] Reuse a single `httpx.Client`. `GraphClient.request` constructs a new one
      per call, so a 500-message bulk action pays 500 TCP+TLS handshakes.
- [ ] Async tool functions. Everything is synchronous and `time.sleep`s on
      retry backoff, so a throttled whole-folder scan stalls the server process.

## Known limitations

- Graph does not support partial/range fetching of message bodies —
  `bodyPreview` (~255 chars plain text) is the only built-in truncation.
- Bulk paging anchors on `receivedDateTime`. A full page of messages sharing one
  timestamp cannot advance the cursor; the scan stops and reports
  `stop_reason="cursor_stalled"` rather than looping.
