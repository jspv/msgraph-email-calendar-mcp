# AGENTS.md — msgraph-email-calendar-mcp

This file guides AI agents working in this repository.

---

## Project Overview

`msgraph-mcp` is a [Model Context Protocol](https://modelcontextprotocol.io/) server that exposes Microsoft Outlook mail and calendar operations as MCP tools. It authenticates via Microsoft device-code flow (delegated, user-context permissions) and wraps the Microsoft Graph v1.0 REST API.

**Entry point:** `uv run msgraph-mcp` → starts the FastMCP server over stdio
**MCP config:** `.mcp.json` (consumed by Claude Code or any MCP client)

---

## Repository Layout

```
src/msgraph_mcp/
  server.py       Entry point: validates config, starts FastMCP
  config.py       Settings loaded from .env (CLIENT_ID, TENANT_ID, SCOPES, CACHE_PATH)
  auth.py         Device-code flow, MSAL token cache, account listing
  graph.py        Authenticated HTTP client with retry/backoff and security validation
  mail.py         Email operations: list, get, search, mark, move, delete, bulk
  calendar.py     Calendar operations: list calendars, list/get events
  models.py       Pydantic response models and text-formatting helpers
  errors.py       Exception hierarchy (MsGraphMcpError, AuthError, GraphRequestError)
  tools.py        FastMCP @mcp.tool() registrations (14 tools total)

tests/            Unit tests (pytest); no integration tests — use smoke_test.py
scripts/
  smoke_test.py   CLI harness for manual end-to-end testing without an MCP client
.env.example      Configuration template — copy to .env and fill in CLIENT_ID
SECURITY_REVIEW.md  Threat model and security analysis
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
- `MICROSOFT_TENANT_ID` — defaults to `organizations`
- `MICROSOFT_SCOPES` — space-separated delegated scopes. Also gates **tool registration**: a tool is exposed only if a scope satisfying it is present (see `_requires_scope` in `tools.py`), so a read-only scope set never advertises write/send tools. Auth tools always register.
- `MICROSOFT_TOKEN_CACHE_PATH` — defaults to `.data/msal_token_cache.json`

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

---

## MCP Tools Reference

### Authentication (must complete before mail/calendar tools work)

| Tool | Description |
|------|-------------|
| `auth_status` | Show configuration and cached accounts |
| `start_auth` | Begin device-code flow (returns URL + one-time code) |
| `finish_auth` | Complete flow after user approves at `microsoft.com/devicelogin` |
| `list_accounts` | List all locally cached accounts |

### Mail

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `list_folders` | `account_id` | Returns folder IDs needed by other tools |
| `list_messages` | `folder_id`, `limit` (≤50) | Newest first |
| `get_message` | `message_id` | Full body + recipients |
| `search_messages` | `query`, `limit` (≤50) | Graph `$search` OData |
| `mark_message_read` | `message_id`, `is_read` | Toggle read/unread |
| `move_message` | `message_id`, `destination_folder` | Accepts well-known names: `inbox`, `drafts`, `sent`, `archive`, `deleted`, `junk` |
| `delete_message` | `message_id`, `permanent=False` | Soft-delete by default |
| `bulk_manage_messages` | filters + `action`, optional `limit`, `dry_run=True` | Dry-run **on** by default. Scans the whole folder unless `limit` is given. Check `truncated`/`stop_reason` in the result — `truncated=False` means the count is a true folder total |

### Calendar

| Tool | Key parameters | Notes |
|------|---------------|-------|
| `list_calendars` | `account_id` | Lists all readable calendars |
| `list_events` | `calendar_id`, `start_time`, `end_time`, `limit` (≤100) | Default window: −1 day to +14 days |
| `get_event` | `event_id` | Full details: body, attendees, organizer |

---

## Security Constraints

These are load-bearing constraints — do not weaken them:

- **Path segment validation** (`graph.py`): User/message IDs are validated against `[A-Za-z0-9_\-=+.]+` before being interpolated into URL paths. Never bypass this.
- **Next-link validation** (`graph.py`): Pagination only follows `@odata.nextLink` URLs that are HTTPS on the Microsoft Graph domain. Do not relax this.
- **Token cache permissions** (`auth.py`): Cache file is written `0600`, parent directory `0700`. Preserve these when modifying auth code.
- **Soft-delete default**: `delete_message(permanent=False)` moves to Deleted Items. The `permanent=True` path is irreversible — keep the default.
- **Bulk dry-run default**: `bulk_manage_messages(dry_run=True)` previews without acting. Always confirm intent before setting `dry_run=False`.
- **Input caps**: `list_messages` and `search_messages` are capped at 50, `list_events` at 100 (in `tools.py`). `bulk_manage_messages` is **not** capped — it scans the whole folder by default; pass `limit` to bound it. Do not raise the list caps without understanding Graph API rate limits.

---

## Development Guidelines

- **Do not add write operations** (send, create, update) without updating `SECURITY_REVIEW.md` first.
- **New tools** go in `tools.py` (registration) and the appropriate domain module (`mail.py`, `calendar.py`). Add a Pydantic model to `models.py` if a new response shape is needed.
- **Error handling**: Raise from the custom hierarchy in `errors.py`. Never surface raw Graph API error payloads to the caller.
- **Tests**: Unit tests use `unittest.mock` to patch `GraphClient`. There are no live-API integration tests. Use `scripts/smoke_test.py` for live testing.
- **No new dependencies** without updating `pyproject.toml` and running `uv lock` to regenerate `uv.lock`.
- **Python 3.11+** required.

---

## Common Agent Workflows

**Check auth before doing anything:**
```
auth_status → if no accounts → start_auth → (user approves) → finish_auth
```

**Read recent inbox:**
```
list_folders → find inbox folder_id → list_messages(folder_id, limit=10)
```

**Find and archive newsletters (preview first):**
```
bulk_manage_messages(sender_contains="newsletter", action="move",
                     destination_folder="archive", dry_run=True)
# Review count, then re-run with dry_run=False
```

**Check today's calendar:**
```
list_calendars → list_events(calendar_id, start_time=today, end_time=today+1day)
```
