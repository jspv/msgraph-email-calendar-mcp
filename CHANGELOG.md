# Changelog

## 0.8.0

Three capability gaps: recurring events, immutable ids (#14), and delta sync (#13).
All three probed against the live API before implementing.

### Added

- **Recurring events.** `create_event` gains `repeat`
  (`daily`/`weekly`/`monthly`/`yearly`), `repeat_interval`, `repeat_days`,
  `repeat_count` and `repeat_until`. This is a compact vocabulary over Graph's
  six pattern types and three range types: a weekly repeat needs `repeat_days`,
  monthly and yearly derive their day from `start_iso`, and with neither
  `repeat_count` nor `repeat_until` the series never ends. The dry-run states
  the repeat in words — *"repeats weekly on monday, wednesday, 4 times"* — since
  a raw pattern dict is not something a human can review.

  Reads now expose `type` (`singleInstance` / `occurrence` / `exception` /
  `seriesMaster`) and `series_master_id` on both calendar models, plus
  `recurrence` on the detail. Without those, a repeat was indistinguishable from
  a one-off and there was no way to find the master in order to edit the series.

  `delete_event`'s preview now warns when the target is a `seriesMaster`:
  removing it deletes **every** occurrence, which the subject and time alone
  give no hint of.

  Worth recording: `list_events` already used `calendarView`, which *expands* a
  series into occurrences. Plain `/events` returns only the master — so the
  existing endpoint choice was correct and must not be changed.

- **`sync_messages` — delta query** (#13). Returns what *changed* in a folder
  since a token, including deletions as explicit `removed: true` entries rather
  than as absences to be inferred. A date window cannot do this:
  `receivedDateTime` never moves after delivery, so a message flagged or moved
  yesterday still carries its original date.

  `GraphClient.paginate_delta` follows `@odata.nextLink` and returns the
  terminating `@odata.deltaLink`, which `paginate` discards. A `limit` that cuts
  the walk short returns **no** token — one from a partial read would not cover
  the unfetched rows and would silently skip them on the next sync. An expired
  token (`410 Gone`) surfaces as an instruction to re-sync, not a generic error.

  **The baseline sync is expensive.** Establishing a token means enumerating the
  whole folder: a live run returned 4,912 changes for this Inbox and took several
  minutes. Every *subsequent* sync is cheap — resuming with the token returned 0
  changes immediately — but the first call is not something to do per request.
  On the folders this matters most for (Deleted Items ~24k, Archive ~19.8k) plan
  for a slow first pass and persist the token.

  Untested: the `410 Gone` path. A malformed token returns `400 Badly formed
  token`, and an expired one could not be manufactured on demand, so the re-sync
  message is covered by unit tests but has not been seen against live Graph.

- **`GRAPH_IMMUTABLE_IDS`** (#14). Requests `Prefer: IdType="ImmutableId"`, so a
  message id survives folder moves. All-or-nothing per client, never per call —
  mixing id types in a session is how an id resolved under one regime reaches a
  call using the other. Flipping it invalidates any id already persisted.

  The setting lives in `_headers()`, not in a per-call branch, specifically so it
  rides `paginate`'s continuation requests — which re-enter `request` with a full
  URL and no params. Keyed off params, page 1 and page 2 would return different
  id types.

### Fixed

- **`Prefer` was assigned rather than appended.** It is a comma-separated list,
  so adding `IdType` naively would have dropped
  `outlook.body-content-type="text"` on exactly the calls that fetch a body —
  silently, since the response is still valid, just HTML. `GraphClient.request`
  now takes an optional `headers` argument merged over the defaults, with
  `Prefer` combined via `_merge_prefer`.

## 0.7.0

### Fixed

- **`list_events` never got the timezone rule #9 established for writes.** The
  same bare wall-clock string was refused by `create_event` and silently
  accepted by `list_events`, shifting the query window by the caller's offset.
  That is precisely the "selective rather than uniform" failure #9 criticised in
  `fcc17cc`, reproduced by fixing the writes and leaving the reads. `list_events`
  now takes `timezone` and applies the same refusal.

  One deliberate difference: a `$filter` compares instants, so an offsetless read
  time is *resolved through* its zone to UTC, whereas a write hands Graph the
  wall-clock string plus the zone name — which is what preserves intent across
  DST for recurring events.

- **Datetime labels had trailing and doubled spaces, and never named the zone.**
  `%Z` is empty for a naive datetime, so an event range rendered as
  `'2026-08-04 14:00  → 2026-08-04 15:00 '`. Labels now take the zone from the
  `dateTimeTimeZone` payload beside the value. Since 0.3.0 these times are no
  longer reliably UTC, so an unlabelled one was ambiguous rather than untidy.

### Added

- **`internet_message_id` on every list and detail result** (#12). Graph remints
  a message `id` on every folder move, and — per the live test in that issue — a
  round trip does not restore the original, so there is no cached id to fall back
  on. Since moving mail is a core feature here, any caller persisting per-message
  state previously had no key that survived its own operations. The RFC 5322
  Message-ID does, and it is meaningful outside this mailbox.

  Not included, per the issue's own scoping: accepting it as an input identifier,
  which would need a `$filter` lookup per call.

## 0.6.0

### Added

- **Server-side `category` filtering** on `list_messages` and
  `bulk_manage_messages`. Previously `category` matched client-side, so
  "everything I filed under Rent" still paged the whole folder. Now Graph
  returns only the matching rows — a live check scanned 1 message instead of the
  folder.

  Probed against the live API before implementing, with a real message tagged
  and a negative control, because a mocked test proves only that a `$filter`
  string was built — the gap that let a broken flag filter ship in 0.5.0.
  Findings: `categories/any(c:c eq 'X')` matches correctly, a non-matching
  category returns nothing, and unlike `flag/flagStatus` it is accepted
  **alongside `$orderby`**. That is why categories can go into the bulk scan,
  whose cursor pagination depends on the `receivedDateTime` sort, while flags
  cannot.

### Security

- Category names are the only caller-supplied value that reaches a `$filter`
  without a parser or enum in front of it — dates go through `_validate_iso`,
  flag states through an enum. They now go through `_odata_string`, which quotes
  the value and doubles embedded single quotes per OData. Without it a name like
  `Bob's stuff` closes the literal early and the remainder is parsed as
  operators.

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
