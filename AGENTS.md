# AGENTS.md — msgraph-email-calendar-mcp

This file guides AI agents working in this repository.

---

## Project Overview

`msgraph-mcp` is a [Model Context Protocol](https://modelcontextprotocol.io/) server that exposes Microsoft Outlook mail and calendar operations as MCP tools. It authenticates via Microsoft device-code flow (delegated, user-context permissions) and wraps the Microsoft Graph v1.0 REST API.

**Entry point:** `uv run msgraph-mcp` → starts the FastMCP server over stdio
**MCP config:** `.mcp.json` (consumed by Claude Code or any MCP client)

This server **reads, writes, sends, and deletes** real user data. Treat it accordingly.

---

## Repository Layout

```
src/msgraph_mcp/
  server.py       Entry point: validates config, starts FastMCP
  config.py       Settings loaded from .env (CLIENT_ID, TENANT_ID, SCOPES, CACHE_PATH)
  auth.py         Device-code flow, MSAL token cache, account listing
  graph.py        Authenticated HTTP client with retry/backoff and security validation
  mail.py         Mail: list, get, search, organize, bulk, drafts, attachments, send
  calendar.py     Calendar: list/get/create/update/delete events, scheduling, free-busy
  contacts.py     People search
  models.py       Pydantic response models, formatting and datetime helpers
  errors.py       Exception hierarchy (MsGraphMcpError, AuthError, GraphRequestError)
  tools.py        FastMCP tool registrations (29 tools: 3 auth + 26 scope-gated)

tests/            Unit tests (pytest); no integration tests — use smoke_test.py
scripts/
  smoke_test.py   CLI harness for manual end-to-end testing without an MCP client
.env.example      Configuration template — copy to .env and fill in CLIENT_ID
SECURITY_REVIEW.md  Threat model and security analysis — keep in sync (see below)
```

---

## Environment Setup

```bash
cp .env.example .env          # Fill in MICROSOFT_CLIENT_ID (required)
uv sync                       # Install dependencies
uv sync --dev                 # Also install dev/test dependencies
```

Required `.env` keys:
- `MICROSOFT_CLIENT_ID` — Azure app registration client ID (required)
- `MICROSOFT_TENANT_ID` — defaults to `common`
- `MICROSOFT_SCOPES` — space-separated delegated scopes. Also gates **tool registration**: a tool is exposed only if a scope satisfying it is present (see `_requires_scope` in `tools.py`), so a read-only scope set never advertises write/send tools. Auth tools always register.
- `MICROSOFT_TOKEN_CACHE_PATH` — defaults to `.data/msal_token_cache.json`
- `MAX_LIST_LIMIT`, `MAX_ATTACHMENT_INLINE_SIZE` — optional bounds

Default scopes are **read/write/send**: `User.Read Mail.ReadWrite Mail.Send Calendars.ReadWrite Calendars.ReadWrite.Shared People.Read`.

---

## Running and Testing

```bash
# Start the MCP server
uv run msgraph-mcp

# Unit tests
uv run pytest

# Manual smoke test (no MCP client needed)
python3 scripts/smoke_test.py status
python3 scripts/smoke_test.py list-messages --folder inbox --limit 5
python3 scripts/smoke_test.py list-events --limit 10
```

Tests must not reach the network. Patch `GraphClient` (or the `mail.` /
`calendar.` function under test) — a test that hits Graph will use the
developer's real cached token against their real mailbox.

---

## MCP Tools Reference

29 tools. Auth always registers; the other 26 are scope-gated.

### Authentication (must complete before other tools work)

| Tool | Description |
|------|-------------|
| `auth_status` | Show configuration and cached accounts (or framework-managed state) |
| `start_auth` | Begin device-code flow (returns URL + one-time code) |
| `finish_auth` | Complete flow after user approves at `microsoft.com/devicelogin` |

### Mail — read (`Mail.Read` / `Mail.ReadWrite`)

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `list_folders` | `parent_folder_id` | Returns folder IDs; pass a parent to list subfolders |
| `list_messages` | `folder`, `limit` (≤1000), `since`, `fields` | Newest first; `since` filters server-side, `fields` overrides `$select`, results carry `conversation_id` |
| `get_message` | `message_id` | Full body + recipients |
| `search_messages` | `query`, `limit` (≤1000) | Graph `$search` OData |
| `get_attachments` | `message_id`, `attachment_id` | Without `attachment_id`: list metadata. With: download (base64 under 1.5 MB) |

### Mail — write (`Mail.ReadWrite`)

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `update_message` | `is_read`, `flag_status`, `categories` | Only supplied fields change |
| `move_message` | `message_id`, `destination` | Well-known names: `inbox`, `drafts`, `sent`, `archive`, `deleted`, `junk` |
| `delete_message` | `message_id`, `permanent=False` | Soft-delete by default |
| `bulk_manage_messages` | filters + `action`, optional `limit`, `dry_run=True` | Dry-run **on** by default. Scans the whole folder unless `limit` is given. Check `truncated`/`stop_reason` — `truncated=False` means the count is a true folder total |
| `create_draft` | `to`, `subject`, `body` | Saved to Drafts, not sent |
| `add_attachment_to_draft` | `message_id`, `name`, `content_base64` | — |
| `create_folder` | `name`, `parent_folder_id` | — |

### Mail — send (`Mail.Send`) — irreversible

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `send_message` | `to`, `subject`, `body`, `dry_run=True` | Dry-run builds a preview draft and deletes it |
| `reply_to_message` | `message_id`, `body`, `reply_all`, `dry_run=True` | — |
| `forward_message` | `message_id`, `to`, `dry_run=True` | Forwards attachments too — highest-risk tool here |
| `manage_draft` | `message_id`, fields…, `send=False` | Update needs `Mail.ReadWrite`; sending needs `Mail.Send` |

### Calendar — read (`Calendars.Read*`)

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `list_calendars` | `user_id` | Pass `user_id` for another user's calendars |
| `list_events` | `calendar_id`, `start_iso`, `end_iso`, `limit` (≤100), `user_id` | Default window: −1 day to +14 days |
| `get_event` | `event_id`, `user_id` | Body, attendees, organizer |
| `check_availability` | `emails`, `start_iso`, `end_iso`, `mode` | `free_busy` or `suggest` |

### Calendar — write (`Calendars.ReadWrite*`) — notifies attendees

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `create_event` | `subject`, `start_iso`, `end_iso`, `attendees` | Graph emails invitations immediately |
| `update_event` | `event_id`, fields… | Updates notify attendees |
| `delete_event` | `event_id`, `cancel_message` | No `cancel_message` → hard delete, no notice |
| `respond_to_event` | `event_id`, `response` | Always notifies the organizer |

### Profile (`User.Read`) and People (`People.Read`)

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `list_aliases` | `account_id` | Send-from addresses, parsed from `proxyAddresses`. Use to pick a valid `send_as` |
| `search_people` | `query`, `limit` | Resolve names to email addresses |

---

## Security Constraints

These are load-bearing constraints — do not weaken them:

- **Path segment validation** (`graph.py`): User/message/event/calendar/folder IDs and `user_id` are validated against `[A-Za-z0-9_\-=+.]+` before being interpolated into URL paths. Never bypass this.
- **Next-link validation** (`graph.py`): Pagination only follows `@odata.nextLink` URLs that are HTTPS on the Microsoft Graph domain. Do not relax this.
- **OData parameter validation** (`mail.py`): `$select` fields via `_sanitize_select`, `$filter` datetimes via `_validate_iso`, `$search` quotes stripped. Any new OData parameter built from caller input needs equivalent validation.
- **Datetime normalization** (`models.py:_parse_utc`): Caller datetimes are parsed to tz-aware UTC. Calendar writes must go through `calendar._graph_datetime` — Graph's `dateTime` field carries no offset, so an offset-bearing string paired with `timeZone: "UTC"` books the event at the wrong hour.
- **Token cache permissions** (`auth.py`): Cache file `0600`, parent directory `0700`. Preserve these.
- **Soft-delete default**: `delete_message(permanent=False)`. The `permanent=True` path is irreversible — keep the default.
- **Dry-run defaults**: `bulk_manage_messages`, `send_message`, `reply_to_message`, and `forward_message` all default to preview. Never flip a default to the acting value.
- **Dry-run drafts are cleaned up in `finally`** (`mail.py`): the preview path creates a real draft. Keep the deletion unconditional.
- **Input caps**: `list_messages`/`search_messages` at `max_list_limit` (default 1000), `list_events` at 100. `bulk_manage_messages` is **not** capped — it scans the whole folder by default; pass `limit` to bound it.
- **Scope gating** (`tools.py`): `_requires_scope` is surface reduction, not authorization. A tool spanning two scopes must check the specific scope its branch needs at call time (see `manage_draft`).

---

## Development Guidelines

- **Any change to the tool surface requires updating `SECURITY_REVIEW.md` in the same change.** This applies to new tools, new parameters that widen reach (a new `user_id`-style target, a new outbound path), and changes to a safety default. This rule was previously ignored and the threat model drifted badly behind the code — do not let that recur.
- **New tools** go in `tools.py` (registration) and the appropriate domain module (`mail.py`, `calendar.py`, `contacts.py`). Add a Pydantic model to `models.py` if a new response shape is needed.
- **Error handling**: Raise from the custom hierarchy in `errors.py` for Graph/auth failures; plain `ValueError` is the established convention for caller-input validation. Never surface raw Graph error payloads.
- **Validate caller input up front**, before the first Graph call, so a bad argument fails cheaply and cannot half-execute a bulk operation.
- **Tests**: Unit tests use `unittest.mock` to patch `GraphClient`. There are no live-API integration tests. Use `scripts/smoke_test.py` for live testing.
- **No new dependencies** without updating `pyproject.toml` and running `uv lock`.
- **Python 3.11+** required.

---

## Common Agent Workflows

**Check auth before doing anything:**
```
auth_status → if no accounts → start_auth → (user approves) → finish_auth
```

**Read recent inbox:**
```
list_folders → find inbox folder id → list_messages(folder, limit=10)
```

**Find and archive newsletters (preview first):**
```
bulk_manage_messages(sender_contains="newsletter", action="move",
                     destination="archive", dry_run=True)
# Review count and truncated/stop_reason, then re-run with dry_run=False
```

**Send with an attachment:**
```
create_draft → add_attachment_to_draft → manage_draft(send=True)
```

**Check today's calendar:**
```
list_calendars → list_events(calendar_id, start_iso=today, end_iso=today+1day)
```
