# msgraph-mcp

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that gives AI assistants full access to Microsoft Outlook email and calendar through the Microsoft Graph API. Built on [FastMCP](https://github.com/jlowin/fastmcp), it supports delegated authentication via device-code flow and can run locally or in a framework-managed cloud environment.

## Features

**26 tools** across mail, calendar, contacts, and scheduling:

### Mail

- **Read** — list folders, messages (with `since` time filter and `conversation_id` threading), search (OData `$search`), attachments (inline base64 for files under 1.5 MB)
- **Compose** — send, reply, reply-all, forward with dry-run preview by default
- **Drafts** — create, update, attach files, then send when ready
- **Organize** — mark read/unread, flag, categorize, move to folder, soft- or hard-delete
- **Bulk** — filtered operations (delete, mark read/unread, move) with dry-run preview; scans the **entire folder** by default, or pass `limit` to cap how many messages are scanned
- **Folders & aliases** — create mail folders, list send-from addresses

### Calendar

- **Read** — list calendars, events (default window: yesterday through 14 days out), full event details
- **Write** — create, update, delete/cancel events with attendees, body, location, all-day support
- **Shared calendars** — full read/write access to other users' calendars via `user_id` parameter
- **Scheduling** — check free/busy status for multiple users, or let Graph suggest optimal meeting times
- **Responses** — accept, decline, or tentatively accept meeting invitations

### Contacts

- **People search** — resolve display names to email addresses using the People API

### Authentication

- **Device-code flow** — interactive three-step auth (`start_auth` → user approves → `finish_auth`)
- **Multi-account** — cache and switch between multiple Microsoft accounts
- **Framework mode** — accept pre-authenticated tokens via environment variables for serverless deployments

## Tool reference

| Area | Tool | Description |
|------|------|-------------|
| Auth | `auth_status` | Show configuration and cached accounts |
| Auth | `start_auth` | Begin device-code flow (returns URL + code) |
| Auth | `finish_auth` | Complete device-code flow after user approval |
| Mail | `list_folders` | List mail folders with item/unread counts |
| Mail | `list_messages` | List messages in a folder (limit 1000; `since`/`until` window, `flag_status` and `category` filters, `fields` override; results carry `conversation_id`, recipients, `flag_status`, `categories`) |
| Mail | `get_message` | Full message details including body |
| Mail | `search_messages` | Search via OData `$search` (limit 1000; `conversation_id` in results) |
| Mail | `get_attachments` | List attachment metadata, or download one by `attachment_id` |
| Mail | `send_message` | Send a new email (dry-run by default) |
| Mail | `reply_to_message` | Reply or reply-all (dry-run by default) |
| Mail | `forward_message` | Forward a message (dry-run by default) |
| Mail | `create_draft` | Create a draft without sending |
| Mail | `manage_draft` | Update or send an existing draft |
| Mail | `add_attachment_to_draft` | Attach a file to a draft |
| Mail | `update_message` | Mark read/unread, flag, or categorize |
| Mail | `move_message` | Move to a folder (supports well-known names) |
| Mail | `delete_message` | Soft-delete or permanently delete |
| Mail | `bulk_manage_messages` | Bulk filtered actions with dry-run (whole folder by default; `received_after`/`received_before` bound it server-side; `recipient_contains`, `flag_status`, `category` filters; `delete`/`move` need a `confirm_token`) |
| Mail | `create_folder` | Create a new mail folder |
| Mail | `list_aliases` | List email aliases / send-from addresses |
| Calendar | `list_calendars` | List calendars (own or shared via `user_id`) |
| Calendar | `list_events` | List events in a time range (limit 100) |
| Calendar | `get_event` | Full event details with attendees |
| Calendar | `create_event` | Create a calendar event (dry-run by default; needs a `timezone` for offsetless times) |
| Calendar | `update_event` | Update an existing event (dry-run by default; needs a `timezone` for offsetless times) |
| Calendar | `delete_event` | Delete or cancel an event (dry-run by default) |
| Calendar | `respond_to_event` | Accept, decline, or tentatively accept |
| Calendar | `check_availability` | Free/busy lookup or meeting time suggestions (needs a `timezone` for offsetless times) |
| Contacts | `search_people` | Search contacts by name (limit 50; returns `job_title`) |

## Prerequisites

- **Python 3.11+**
- **An Azure app registration** with delegated Microsoft Graph permissions (see below)
- **uv** (recommended) or pip

## Azure app registration

Create an app registration in [Microsoft Entra admin center](https://entra.microsoft.com/) (Azure AD).

### 1. Supported account types

Choose one:
- **Accounts in this organizational directory only** — single tenant
- **Accounts in any organizational directory** — multi-tenant work/school accounts

### 2. Authentication

- Enable **Allow public client flows** (required for device-code flow)

### 3. API permissions

Add **delegated** Microsoft Graph permissions:

| Permission | Purpose |
|---|---|
| `User.Read` | Read signed-in user profile |
| `Mail.ReadWrite` | Read, move, flag, categorize, delete mail |
| `Mail.Send` | Send mail, reply, forward |
| `Calendars.ReadWrite` | Read and write calendar events |
| `Calendars.ReadWrite.Shared` | Access shared / delegated calendars |
| `People.Read` | Search contacts by name |

> **`Calendars.ReadWrite.Shared` is the widest permission here.** With it, every
> calendar tool accepts a `user_id` and can create, modify, delete, or cancel
> events on any calendar the signed-in user has been granted access to —
> cancelling someone else's meeting emails all of its attendees. There is no
> allowlist of targetable users; authorization rests entirely with Graph. Drop
> this scope if you don't need it. See [SECURITY_REVIEW.md](SECURITY_REVIEW.md).

For **read-only** use, replace `Mail.ReadWrite` and `Mail.Send` with `Mail.Read`, and `Calendars.ReadWrite` / `Calendars.ReadWrite.Shared` with `Calendars.Read`, and set `MICROSOFT_SCOPES` to match.

**Scope-based tool registration:** tools are exposed to the client only when a scope that satisfies them is present in `MICROSOFT_SCOPES`. A read-only scope set never advertises `delete_message`, `send_message`, `create_event`, etc. — the model can't attempt actions the token could not perform. With the default (full) scope set, all tools are available. The auth tools (`auth_status`, `start_auth`, `finish_auth`) are always registered so you can authenticate before any scope is granted.

### 4. Admin consent

Grant admin consent for the tenant if required by your organization's policies.

## Configuration

Copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
```

| Variable | Default | Description |
|---|---|---|
| `MICROSOFT_CLIENT_ID` | *(required)* | Azure app registration client ID |
| `MICROSOFT_TENANT_ID` | `common` | `organizations` (work/school only), `common` (any), or a specific tenant GUID |
| `MICROSOFT_SCOPES` | `User.Read Mail.ReadWrite Mail.Send Calendars.ReadWrite Calendars.ReadWrite.Shared People.Read` | Space-separated delegated permissions |
| `MICROSOFT_TOKEN_CACHE_PATH` | `.data/msal_token_cache.json` | Path to the local MSAL token cache |
| `MAX_ATTACHMENT_INLINE_SIZE` | `1572864` | Max attachment size (bytes) for inline base64 (default 1.5 MB) |
| `MAX_LIST_LIMIT` | `1000` | Max items returned by `list_messages` / `search_messages` |
| `MSGRAPH_DEFAULT_TIMEZONE` | *(unset)* | IANA or Windows zone applied to calendar times written without a UTC offset. Unset → such times are refused, not guessed |

**Recommended tenant values:**
- `organizations` — work/school accounts only (most common for enterprise)
- A specific tenant GUID — locks authentication to a single organization
- `common` — any Microsoft account (work, school, or personal)

## Deployment

### Local (stdio)

The default transport is stdio, suitable for desktop MCP clients like Claude Code, Claude Desktop, Cursor, and VS Code.

```bash
# Install dependencies
uv sync

# Run the server
uv run msgraph-mcp
```

Or with pip:

```bash
pip install -e .
msgraph-mcp
```

#### MCP client configuration

Add to your MCP client's configuration (e.g. Claude Desktop `claude_desktop_config.json`, `.mcp.json` for Claude Code, etc.):

```json
{
  "mcpServers": {
    "msgraph-mcp": {
      "type": "stdio",
      "command": "uv",
      "args": ["run", "msgraph-mcp"],
      "env": {
        "MICROSOFT_CLIENT_ID": "your-client-id",
        "MICROSOFT_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

If you use a `.env` file in the project directory, the `env` block can be omitted.

### Cloud — AWS Lambda with mcp-lambda-wrappers (ChatGPT, Claude.ai)

For use with **remote MCP clients** like ChatGPT and Claude.ai, this server can be deployed as a serverless AWS Lambda function using [mcp-cloud-wrappers](https://github.com/jspv/mcp-cloud-wrappers). That framework wraps any stdio-based MCP server behind Amazon Bedrock AgentCore Gateway with full OAuth 2.0 and Dynamic Client Registration (RFC 7591) support — no code changes required in this project.

**What the framework provides:**

- **Serverless deployment** — runs this MCP server as a Lambda subprocess behind AgentCore Gateway
- **Per-user OAuth** — each user authenticates with their own Microsoft account; tokens are stored in AWS Secrets Manager with automatic refresh
- **Caller authentication** — Cognito JWT validation for all inbound requests
- **Dynamic Client Registration** — MCP clients (ChatGPT, Claude.ai) self-register via a standard `/register` endpoint
- **Zero idle cost** — Lambda functions spin up on demand

**How it works:**

1. An MCP client sends a tool call to the AgentCore Gateway endpoint
2. The framework validates the caller's JWT, extracts their identity, and loads their Microsoft Graph OAuth token from Secrets Manager
3. The token is injected as `GRAPH_ACCESS_TOKEN` into this server's environment
4. This server runs as a subprocess, reads the token, and executes the tool against Microsoft Graph
5. If the user hasn't authenticated yet, `start_auth` returns the framework's OAuth URL instead of a device code

This project is used as the **reference example service** in mcp-lambda-wrappers — see `infra/lambda/services/msgraph/` in that repo for the full configuration.

**Quick deploy (from the mcp-lambda-wrappers repo):**

```bash
# One-time: deploy shared infrastructure (Cognito, DCR, OAuth callback)
make deploy-shared

# Create the Azure app secret
aws secretsmanager create-secret \
  --name mcp-wrappers-msgraph-service-secrets \
  --secret-string '{"MICROSOFT_CLIENT_ID": "your-client-id"}'

# Generate tool definitions and deploy
make gen-tools SERVICE=msgraph
make deploy-service SERVICE=msgraph
```

#### Framework environment variables

When running inside the framework, this server auto-detects Lambda mode via these injected environment variables:

| Variable | Description |
|---|---|
| `GRAPH_ACCESS_TOKEN` | Pre-authenticated Microsoft Graph access token (per-user) |
| `OAUTH_AUTHENTICATED` | Set to `true` when auth is complete |
| `OAUTH_USER_ID` | Authenticated user identifier |
| `OAUTH_AUTH_URL` | OAuth authorization URL (shown when user needs to authenticate) |
| `SERVICE_NAME` | Service identifier for the framework |

In this mode:
- The MSAL device-code flow is bypassed — tokens are injected by the framework
- No local token cache is used (compatible with read-only filesystems like Lambda's `/var/task`)
- `auth_status` reports the framework-managed token state
- `start_auth` / `finish_auth` return guidance to authenticate through the framework's OAuth flow instead

### Docker

While no Dockerfile is included, the server can be containerized:

```dockerfile
FROM python:3.12-slim
WORKDIR /app
COPY . .
RUN pip install --no-cache-dir .
ENV MICROSOFT_CLIENT_ID=""
ENV MICROSOFT_TENANT_ID="organizations"
EXPOSE 8000
CMD ["msgraph-mcp"]
```

For persistent authentication, mount a volume for the token cache:

```bash
docker run -v msgraph-data:/app/.data \
  -e MICROSOFT_CLIENT_ID=your-id \
  -e MICROSOFT_TENANT_ID=your-tenant \
  msgraph-mcp
```

## Safety defaults

Write operations default to safe behavior:

| Feature | Default | Notes |
|---|---|---|
| `send_message` | `dry_run=True` | Creates a temporary draft for preview, then deletes it |
| `reply_to_message` | `dry_run=True` | Preview before sending |
| `forward_message` | `dry_run=True` | Preview before sending |
| `bulk_manage_messages` | `dry_run=True` | Shows matches without executing; `delete`/`move` also need a `confirm_token` |
| `delete_message` | `permanent=False` | Moves to Deleted Items (recoverable) |
| `create_event` | `dry_run=True` | Graph mails invitations immediately, so preview first |
| `update_event` | `dry_run=True` | Preview shows current state next to the proposed changes |
| `delete_event` | `dry_run=True` | Preview names the event and says whether attendees are notified |

`respond_to_event` is deliberately **not** gated — accepting or declining is
reversible by responding again, so a confirmation step would be friction with no
safety payoff.

### Calendar times and timezones

Graph's `dateTimeTimeZone` pairs a *naive* wall-clock string with a separate zone
name, so how a time is written matters:

| Input | Sent to Graph |
|---|---|
| `2026-08-01T14:00:00-04:00` | `{"dateTime": "2026-08-01T18:00:00", "timeZone": "UTC"}` — the instant is unambiguous, so it is converted |
| `2026-08-01T14:00:00` + `timezone="America/New_York"` | `{"dateTime": "2026-08-01T14:00:00", "timeZone": "America/New_York"}` — handed to Graph unconverted, which keeps recurring events correct across DST |
| `2026-08-01T14:00:00`, no zone anywhere | **Refused** |

That last row is deliberate. Reading a bare local time as UTC is how a 2pm
Eastern meeting silently becomes 10:00 EDT, with invitations already sent — and
on an MCP server the caller is usually a model turning "book me 2pm Thursday"
into exactly that string. Set `MSGRAPH_DEFAULT_TIMEZONE` if you want a
server-side default instead of passing `timezone` per call.

**All-day events** are a separate contract: Graph wants midnight in the stated
zone, so the calendar *date* is preserved and the instant is not.
`2026-04-01T23:00:00-04:00` with `is_all_day=True` books April 1, not April 2.

### Confirming a bulk delete or move

Destructive bulk actions are two-step. The dry run returns a `confirm_token`
derived from the ids it actually matched:

```jsonc
// 1. preview
{"action": "delete", "dry_run": true}
// -> {"matched": 42, "confirm_token": "42-b7e2d4a1c3f9", "matches": [...]}

// 2. act
{"action": "delete", "dry_run": false, "confirm_token": "42-b7e2d4a1c3f9"}
```

The live run rescans and re-derives the token from its own results. If the
mailbox changed in between, it refuses rather than acting on a set you never saw:

```
confirm_token does not match the current scan: it described 42 message(s),
this scan matched 43. The mailbox is live, so re-check the preview before
acting. Confirm with the new token: 43-9f1c2ae5b7d0
```

Nothing is stored server-side — the token is recomputed each time — so this works
unchanged across a Lambda cold start between the two calls.

### Bulk scan semantics

`bulk_manage_messages` scans newest-first and applies its filters client-side. By
default (`limit=None`) it scans the **entire folder**, so "find/act on all messages
matching X" returns a true total. Pass `limit=N` to scan at most the newest N
messages; the value is honored exactly (paged internally at up to 1000/request),
never silently clamped. The response reports coverage explicitly:

| Field | Meaning |
|---|---|
| `scanned` | Distinct messages inspected this run |
| `matched` | How many passed the filters |
| `total_in_folder` | Folder size, as a scale anchor |
| `truncated` | `False` only when the scan reached the folder end (count is a true total) |
| `stop_reason` | `folder_exhausted` \| `window_exhausted` \| `scan_limit_reached` \| `cursor_stalled` |
| `acted` / `already_gone` / `failed` | Per-message outcomes of a live (non-dry-run) run |

#### Follow-up flags and categories

`flag_message` and `categorize_message` could always *write* these; nothing read
them back, so an agent could flag a message and then had no way to report which
messages were flagged. Both are now on every list and detail result
(`flag_status`, `categories`) and appear in the generated `summary` string.

`list_messages(flag_status="flagged")` filters server-side, so "show me my
flagged mail" is one request. Two constraints worth knowing, both confirmed
against the live API rather than inferred:

- Exchange rejects a flag restriction combined with a sort — the pairing returns
  *"The restriction or sort order is too complex for this operation."* So
  `$orderby` is dropped for these queries and the rows are re-sorted client-side
  to keep the newest-first contract.
- Because the server chose those rows unsorted, `limit` selects an *arbitrary*
  subset rather than the newest N. Raise `limit` above your expected flagged
  count, or narrow with `since`/`until` — a date clause **does** combine with the
  flag filter.

**Categories behave differently, and better.** Exchange accepts a category
restriction *with* a sort, so `category` is pushed server-side on both
`list_messages` and `bulk_manage_messages`, keeps newest-first ordering, and
imposes none of the `limit` caveat above. Filtering a 50k-message folder by
category reads only the matching rows — a live check scanned 1 message rather
than the folder.

`bulk_manage_messages` filters *flags* client-side, since its cursor pagination
depends on the `receivedDateTime` sort a flag restriction would force it to drop.
Category names are escaped as OData string literals, so a name containing a
quote (`Bob's stuff`) is handled rather than breaking the filter.

#### Date windows

`received_after` and `received_before` are applied by **Graph**, not after the
fetch, so a date-scoped query reads only its window. This is what makes old mail
cheap to reach: without a bound, "what did this sender send me last March" pages
the entire folder to match a handful of rows. On a 50k-message mailbox that is
the difference between ~50 round-trips and one.

The upper bound needs no extra machinery — paging already anchors on
`receivedDateTime le`, so `received_before` *is* the starting cursor and the scan
begins inside the window rather than at the newest message.

When either bound is given, `stop_reason` is `window_exhausted` rather than
`folder_exhausted`. That distinction is deliberate: the scan covered all of what
you asked for, but not all of the folder, and reporting the latter would
overclaim.

The mailbox is live, so counts are a point-in-time snapshot: re-running may
legitimately see a different set. Collection and action are separate phases —
messages moved or deleted (e.g. by a rule) between the two are reported as
`already_gone`, not errors. `truncated=True` means matches may exist deeper than
this call reached; it is **not** safe to read a converged count of 0 as "the folder
is clean."

## Security

- **Path segment validation** — all user-supplied IDs are validated against a safe-character pattern before URL interpolation, blocking path traversal
- **Next-link hardening** — pagination only follows HTTPS URLs on the configured Graph host
- **Search sanitization** — double-quotes stripped from OData `$search` queries
- **Retry with backoff** — automatic retry for HTTP 429 and transient 5xx errors (3 attempts, respects `Retry-After`)
- **Error translation** — raw Graph API payloads are never exposed to callers
- **Token cache permissions** — cache file `0600`, parent directory `0700`, symlinks rejected

See [SECURITY_REVIEW.md](SECURITY_REVIEW.md) for the full threat model and remaining risks.

## Development

```bash
# Install with dev dependencies
uv sync --dev

# Run tests
uv run pytest

# Or with pip
pip install -e '.[dev]'
pytest
```

CI runs the suite on Python 3.11–3.14 for every push and pull request to `main`
(`.github/workflows/ci.yml`).

Two guardrails run as part of the ordinary test suite:

- **`tests/conftest.py`** blocks real outbound HTTP. Tests that need it patch
  `GraphClient`; anything that reaches `httpx` fails loudly. This exists because
  a test once made a live Graph call against a developer's cached token.
- **`tests/test_docs_parity.py`** asserts the tool tables in `README.md` and
  `AGENTS.md` match the registered tools exactly, in both directions. Adding a
  tool without documenting it — or documenting one that does not exist — fails
  the build.

### Smoke test harness

A CLI harness for manual testing without a full MCP client:

```bash
# Auth
python3 scripts/smoke_test.py status
python3 scripts/smoke_test.py start-auth
python3 scripts/smoke_test.py finish-auth
python3 scripts/smoke_test.py list-accounts

# Mail
python3 scripts/smoke_test.py list-folders
python3 scripts/smoke_test.py list-messages --folder inbox --limit 5
python3 scripts/smoke_test.py get-message MESSAGE_ID
python3 scripts/smoke_test.py search-messages "search term"
python3 scripts/smoke_test.py mark-message-read MESSAGE_ID
python3 scripts/smoke_test.py move-message MESSAGE_ID archive
python3 scripts/smoke_test.py delete-message MESSAGE_ID
python3 scripts/smoke_test.py bulk-manage-messages --sender-contains "newsletters" --limit 50

# Calendar
python3 scripts/smoke_test.py list-calendars
python3 scripts/smoke_test.py list-events --limit 10
python3 scripts/smoke_test.py get-event EVENT_ID
```

## License

See [LICENSE](LICENSE) for details.
