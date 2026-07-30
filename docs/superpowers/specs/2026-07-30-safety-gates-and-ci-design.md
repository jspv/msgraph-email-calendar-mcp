# Safety gates and CI — design

**Date:** 2026-07-30
**Status:** approved
**Scope:** roadmap items 1–4 (CI, httpx client reuse, calendar dry-run, bulk confirm token). Async is explicitly out of scope.
**Version impact:** `0.1.0` → `0.2.0` (three tool contracts change).

## Decisions taken

Two questions were settled before design:

1. **Safety posture — safety-first, breaking.** Calendar writes default to `dry_run=True`;
   bulk destructive actions require a confirm token echoed from a dry-run. Accepted that
   this breaks existing callers; the project is pre-1.0, so breaking changes are cheap now
   and expensive later.
2. **Token binding — the exact matched message set.** The token is derived from the
   matched ids rather than from the filter arguments, so a token cannot authorise a set
   the caller never saw.

## 1. CI

`.github/workflows/ci.yml`, on push and pull_request to `main`.

- Matrix: Python 3.11 and 3.12 (`requires-python = ">=3.11"`).
- `astral-sh/setup-uv@v5` with caching, `uv python install`, `uv sync --dev`, `uv run pytest -q`.
- `scripts/smoke_test.py` reaches the network and is not collected — `testpaths = ["tests"]`
  already excludes it. Keep it that way.

Two things ship with CI because CI is what makes them enforceable:

### `tests/conftest.py` — network guard

Autouse session fixture patching `httpx.Client.request` to raise. During the previous
batch a RED test fired a live Graph request against the developer's cached token.
`AGENTS.md` now forbids network access in tests; this makes the rule fail loudly instead
of relying on discipline. Existing tests patch `GraphClient`, which sits above this layer,
so they are unaffected.

### `tests/test_docs_parity.py` — drift-catcher

A test rather than a script, so it runs locally on every `pytest`. Reloads `tools` under
full scopes (the `test_scope_filtering` pattern), takes names from `mcp.list_tools()`, and
asserts set-equality against tool names parsed from the markdown tables in `README.md` and
`AGENTS.md`. Both directions matter: an undocumented tool fails, and a documented tool
that does not exist fails. The second direction is what caught the phantom
`list_attachments` row.

## 2. httpx.Client reuse — `graph.py`

Replace the per-request `with httpx.Client(...)` with a lazily-built module-level client.

```python
_http: httpx.Client | None = None

def _http_client() -> httpx.Client:
    global _http
    if _http is None or _http.is_closed:
        _http = httpx.Client(
            timeout=settings.timeout_seconds,
            limits=httpx.Limits(max_keepalive_connections=10, max_connections=20),
        )
    return _http

def close_http_client() -> None: ...  # registered via atexit
```

**Placement rationale.** The motivating workload is `bulk_manage_messages_multi_pass`,
whose action loop calls `delete_message(account_id, item.id)` / `move_message(...)` per
message — each constructing a *fresh* `GraphClient`. Per-instance pooling would buy
nothing for the only case that needs it. Module-level scope is what turns N handshakes
into one pool.

No `base_url` on the client: `_normalize_path` keeps its current semantics and
`tests/test_graph.py` (direct `GraphClient(account_id=None)`, no httpx involvement) is
untouched.

`httpx.Client` is safe for concurrent use and its pool evicts broken connections, so
sharing does not introduce shared-failure state.

**Tests** (`tests/test_graph_client_reuse.py`): two `GraphClient` instances resolve to the
same underlying client; the client is not closed after a request; `close_http_client()`
resets the global and the next call rebuilds.

## 3. Calendar dry-run

`dry_run: bool = True` on `create_event`, `update_event`, `delete_event` in `calendar.py`
and the `tools.py` wrappers. Response mirrors mail's existing
`{"ok", "dry_run", "preview", "message"}`.

### `create_event` — zero Graph calls

Unlike mail (which POSTs a real draft to get Graph's rendering), the event body is fully
known client-side.

```python
{"ok": True, "dry_run": True, "action": "create",
 "preview": {"path": path, "event": event_body},
 "message": "Dry-run: event NOT created. Set dry_run=False to create."}
```

The preview echoes the *converted* UTC body, so it also surfaces the timezone conversion:
a `14:00-07:00` input visibly becomes `21:00:00 UTC` before anything is booked.

### `update_event` — one GET

Preview shows current-vs-proposed rather than a bare patch dict:
`"preview": {"current": {...}, "changes": update}`. If the GET fails the error propagates
rather than degrading to a partial preview — a caller who cannot read the event cannot
patch it either, so there is no state worth inventing.

### `delete_event` — one GET via `get_event()`

`CalendarEventDetail` already carries `time_label`, `attendee_labels`, `organizer_label`.

```python
"preview": {"subject", "time_label", "attendee_count", "organizer", "is_cancelled"}
```

The message distinguishes the two paths, which differ in a way the API shape understates:

- `cancel_message` set → `POST /cancel` → **attendees are notified**
- otherwise → `DELETE` → attendees are **not** notified

e.g. `"Dry-run: event NOT cancelled. Set dry_run=False to cancel and notify 14 attendees."`

### `respond_to_event` is excluded

Accept/decline is reversible by responding again. A gate costs friction and buys nothing.

## 4. Bulk confirm token

```python
def _confirm_token(action, destination, ids) -> str:
    digest = sha256("|".join([action, destination or "", *sorted(ids)]).encode()).hexdigest()
    return f"{len(ids)}-{digest[:12]}"
```

The count prefix is deliberate. A bare digest leaves the server unable to report what the
preview matched — it keeps no state — so a mismatch could only say "changed", not
"42 → 43". The count is not a secret, and carrying it makes the error actionable.

Dry-run adds `confirm_token` to its report. The live run takes
`confirm_token: str | None = None` and recomputes the digest from **its own** scan:

- `dry_run=False` with no token → `ValueError` directing the caller to preview first.
- Mismatch → `ValueError`: *"The matched set changed since the preview (42 → 43 messages).
  Re-check the preview, then confirm with: `43-9f1c...`"*

No server state, no salt, no clock; a Lambda cold start between calls is fine. Forging a
token requires the matched id set, which requires running the scan, which *is* the dry run.

**Scoped to `delete` and `move`.** `mark_read` / `mark_unread` are trivially reversible;
gating them is friction with no safety payoff — the same reasoning that excludes
`respond_to_event`.

## Security note to carry into SECURITY_REVIEW.md

The confirm token forces a preview to exist in the transcript, but it does **not** stop a
determined prompt injection: injected text can say "run the dry-run, then call again with
the token," and a compliant model will. Its real value is that destructive scope becomes
*visible* to the client and the human before anything happens. Documenting it as an
anti-injection defense would be false.

## Order of work

1. CI (+ conftest network guard, + docs-parity test) — so everything after is verified.
2. httpx client reuse — no contract change.
3. Calendar dry-run.
4. Bulk confirm token.
5. Docs sweep + version bump.

All steps TDD (red → green → refactor). Docs land in the same commit as the code that
changes the tool surface, per the `AGENTS.md` rule.
