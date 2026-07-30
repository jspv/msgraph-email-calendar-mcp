# Changelog

## 0.5.0

Follow-up flags and categories are readable, closing a write-only asymmetry.

### Added

- **`flag_status` and `categories` on every list and detail result.**
  `flag_message` and `categorize_message` could always set these, but no read
  path returned them, so an agent could flag a message and then had no way to
  tell which messages were flagged — including one it had flagged itself moments
  earlier. "Show me my flagged mail" was unanswerable.

  The `fields` override was not a workaround: `$select` widened and Graph did
  return `flag`, but the model dropped it on the way back — the same failure
  mode recipients had before #11.

  `flag_status` is kept as Graph's raw three-state string (`flagged` /
  `complete` / `notFlagged`) rather than a bool, which would lose the
  distinction between "never flagged" and "followed up and done", and which
  feeds straight back into `flag_message`. `None` means the column was not
  selected — distinct from `notFlagged`, which means Graph was asked.

  Flagged state and categories also appear in the generated `summary` string.
  Unflagged rows say nothing, since most mail is unflagged and the note would
  crowd out the body preview.

- **`flag_status` filter on `list_messages`**, applied server-side, so "show me
  my flagged mail" is one request rather than a folder scan.
- **`flag_status` and `category` filters on `bulk_manage_messages`**, applied
  client-side — its cursor pagination depends on a `receivedDateTime` sort that
  a server-side flag filter would force it to drop (see below).

### Known Graph constraint

Exchange rejects a `flag/flagStatus` restriction combined with a sort:

> The restriction or sort order is too complex for this operation.

Confirmed against the live API, not inferred — every shape carrying an
`$orderby` fails, while the same filter without one succeeds, including
alongside a `receivedDateTime` clause. `list_messages` therefore drops
`$orderby` for flag queries and re-sorts the returned rows client-side to keep
its newest-first contract.

Consequence worth knowing: because the server selected those rows unsorted,
`limit` picks an **arbitrary** subset rather than the newest N. Raise `limit`
above the expected flagged count, or narrow with `since`/`until`, which does
combine with the flag filter.

Mocked tests could not have caught this — they assert the `$filter` string is
built, not that Graph accepts it. The constraint is now pinned by name in the
test suite.

## 0.4.0

Closes #4 and #11. Both are about list-level results carrying enough to act on
without a follow-up call per row.

### Added

- **Date windows are applied by Graph** (#4). `bulk_manage_messages` gains
  `received_before`, and both bounds now travel in the `$filter` instead of being
  compared after the fetch. A date-scoped query costs O(window) rather than
  O(folder) — on a 50k-message mailbox, "what did this sender send me last March"
  went from paging all 50k rows to reading one slice.

  The upper bound needed no new machinery: paging already anchors on
  `receivedDateTime le`, so `received_before` is simply the initial cursor and
  the scan starts inside the window rather than at the newest message.

  `list_messages` gains the matching `until` to pair with its existing `since`.

  Note on #4's original framing: old mail was not *unreachable* — `0c4f542` (#6)
  had already made a full-folder scan the default. It was expensive. That is a
  narrower claim than the issue made, but a sufficient one, and the issue has
  been retitled accordingly.

- **Recipients at list level** (#11). `MailMessageSummary` gains
  `to_recipients`, `to_recipient_labels`, `cc_recipients`, and
  `cc_recipient_labels`, and `_SUMMARY_SELECT` requests `toRecipients` /
  `ccRecipients` for **every** folder.

  Unconditional is deliberate. Fetching them only for Sent Items would skip the
  Inbox, which is exactly where the two motivating signals live: which alias a
  message was delivered to when the owner uses a per-vendor address, and whether
  the owner was addressed directly or merely copied. That costs two extra arrays
  per row on a whole-folder scan; if it ever hurts, the answer is a leaner
  projection, not dropping them on Inbox.

  Recipients also appear in the generated `summary` string, so a model reading a
  Sent Items list can tell the rows apart — previously every row showed the
  mailbox owner as sender and nothing else distinguishing.

- **`recipient_contains`** on `bulk_manage_messages`, matching across both To and
  Cc. This is what makes "everything sent to my *vendor* alias" one call.

### Changed

- **New `stop_reason`: `window_exhausted`.** Reported instead of
  `folder_exhausted` when a date bound was given. The scan covered all of what
  was asked but not all of the folder, and reporting the latter would overclaim.
- An inverted window (`received_after` later than `received_before`) is rejected
  up front rather than silently returning zero matches, which would read as
  "nothing there".

## 0.3.0

Fixes the two open calendar-timezone bugs, #9 and #10. Both reproduce exactly as
their issues describe; both are in code paths a model exercises by default.

### Breaking

- **Calendar times without a UTC offset are now refused unless a zone is
  available.** Previously `start_iso="2026-08-01T14:00:00"` was read as UTC, so
  an Eastern caller who meant 2pm booked 10:00 EDT — silently, with invitations
  already sent (#9). Three ways to be explicit:

  | Input | Result |
  |---|---|
  | `2026-08-01T14:00:00-04:00` | converted to `18:00:00` UTC (unchanged) |
  | `2026-08-01T14:00:00` + `timezone="America/New_York"` | handed to Graph unconverted with that zone |
  | `2026-08-01T14:00:00`, nothing configured | `ValueError` naming all three ways out |

  Passing the wall-clock time through with its zone, rather than flattening to
  UTC, is also what keeps a recurring event at 14:00 local across a DST
  boundary.

  Affects `create_event`, `update_event`, and `check_availability`
  (`find_meeting_times` / `get_schedule`).

- **New `timezone` parameter** on those tools, and a new
  `MSGRAPH_DEFAULT_TIMEZONE` environment variable as the server-side fallback.
  IANA names are validated locally so a typo fails immediately instead of as a
  Graph 400; Windows zone ids pass through.

### Fixed

- **All-day events with an offset-bearing start/end no longer emit a non-midnight
  time** (#10). `2026-04-01T00:00:00-04:00` with `is_all_day=True` was being
  converted to `04:00:00`, which Graph's documented contract rejects. All-day
  times now keep the caller's calendar *date* and emit midnight, so
  `2026-04-01T23:00:00-04:00` books April 1 rather than April 2.

  This was the one case the pre-`fcc17cc` relabel got accidentally right, so the
  general offset fix regressed it. `test_all_day_event` asserted only
  `isAllDay is True` and could not have caught it; it now asserts the emitted
  `start`/`end`.

  **Confirmed against the live API**, not just the documented contract: posting
  the pre-fix payload returns a hard 400 — *"The Event.Start property for an
  all-day event needs to be set to midnight."* So between `fcc17cc` and this
  release, every all-day event created with an offset-bearing time failed
  outright rather than being merely misplaced. `scripts/smoke_test.py
  check-all-day-event` re-settles this if the all-day path is ever touched
  again; without `--apply` it only prints the payloads.

## 0.2.0

Safety-first release. Three tool contracts change in ways that will break
existing callers — deliberately, while the project is pre-1.0 and the change is
still cheap.

### Breaking

- **`create_event`, `update_event`, `delete_event` now default to `dry_run=True`.**
  Graph mails invitations and cancellations the moment these are called, with no
  undo. Callers that relied on the old immediate behaviour must pass
  `dry_run=False` explicitly.

  `create_event`'s preview costs no Graph call and echoes the exact body that
  would be sent, including times converted to UTC. `update_event` spends one read
  to show current state beside the proposed changes. `delete_event` spends one
  read to name the event — subject, time, organizer, attendee count — and states
  which outcome applies: with `cancel_message` attendees are **notified**;
  without it the event is destroyed and **nobody is told**.

  `respond_to_event` is intentionally **not** gated; responding again reverses it.

- **`bulk_manage_messages` requires a `confirm_token` for `delete` and `move`.**
  The dry run returns a token derived from the message ids it actually matched;
  the live run rescans, re-derives the token, and refuses if it differs — handing
  back the fresh token and the count delta so recovery is one call:

  ```
  confirm_token does not match the current scan: it described 42 message(s),
  this scan matched 43. The mailbox is live, so re-check the preview before
  acting. Confirm with the new token: 43-9f1c2ae5b7d0
  ```

  Nothing is stored server-side, so this works across a Lambda cold start between
  the two calls. `mark_read` / `mark_unread` are reversible and need no token.

### Added

- **CI** (`.github/workflows/ci.yml`) — test suite on Python 3.11–3.14 for every
  push and pull request to `main`.
- **No-network guard** (`tests/conftest.py`) — any real outbound HTTP request
  from a test now raises. Added because a test once reached the live Graph API
  using a developer's cached token.
- **Docs-parity test** (`tests/test_docs_parity.py`) — the tool tables in
  `README.md` and `AGENTS.md` must match the registered tool set exactly, in both
  directions. Undocumented tools and phantom rows both fail the build.

### Changed

- **One pooled `httpx.Client` per process** instead of one per request. A bulk
  action builds a fresh `GraphClient` per message, so the old code paid a TCP+TLS
  handshake per message; a 500-message run now reuses one pool.

### Security

The confirmation gates make destructive scope **visible** in the transcript
before it is applied. They do **not** stop prompt injection — injected text can
walk the two-step flow itself. See `SECURITY_REVIEW.md` → "Remaining risks".

## 0.1.0

Initial release.
