# Changelog

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

  The Graph-rejects-non-midnight claim comes from the documented contract, not
  from a live call. `scripts/smoke_test.py check-all-day-event` prints both
  payloads and, with `--apply`, POSTs the pre-fix one to settle whether Graph
  400s, rounds, or accepts.

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
