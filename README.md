# msgraph-mcp

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that gives AI assistants full access to Microsoft Outlook email and calendar through the Microsoft Graph API. Built on [FastMCP](https://github.com/jlowin/fastmcp), it supports delegated authentication via device-code flow and can run locally or in a framework-managed cloud environment.

## Features

**26 tools** across mail, calendar, contacts, and scheduling:

### Mail

- **Read** — list folders, messages, search (OData `$search`), attachments (inline base64 for files under 1.5 MB)
- **Compose** — send, reply, reply-all, forward with dry-run preview by default
- **Drafts** — create, update, attach files, then send when ready
- **Organize** — mark read/unread, flag, categorize, move to folder, soft- or hard-delete
- **Bulk** — multi-pass filtered operations (delete, mark read/unread, move) with dry-run preview, up to 1 000 messages per call
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
| Mail | `list_messages` | List messages in a folder (limit 50) |
| Mail | `get_message` | Full message details including body |
| Mail | `search_messages` | Search via OData `$search` (limit 50) |
| Mail | `list_attachments` | List attachment metadata for a message |
| Mail | `get_attachments` | Download a single attachment |
| Mail | `send_message` | Send a new email (dry-run by default) |
| Mail | `reply_to_message` | Reply or reply-all (dry-run by default) |
| Mail | `forward_message` | Forward a message (dry-run by default) |
| Mail | `create_draft` | Create a draft without sending |
| Mail | `manage_draft` | Update or send an existing draft |
| Mail | `add_attachment_to_draft` | Attach a file to a draft |
| Mail | `update_message` | Mark read/unread, flag, or categorize |
| Mail | `move_message` | Move to a folder (supports well-known names) |
| Mail | `delete_message` | Soft-delete or permanently delete |
| Mail | `bulk_manage_messages` | Bulk filtered actions with dry-run (limit 1 000) |
| Mail | `create_folder` | Create a new mail folder |
| Mail | `list_aliases` | List email aliases / send-from addresses |
| Calendar | `list_calendars` | List calendars (own or shared via `user_id`) |
| Calendar | `list_events` | List events in a time range (limit 100) |
| Calendar | `get_event` | Full event details with attendees |
| Calendar | `create_event` | Create a calendar event |
| Calendar | `update_event` | Update an existing event |
| Calendar | `delete_event` | Delete or cancel an event |
| Calendar | `respond_to_event` | Accept, decline, or tentatively accept |
| Calendar | `check_availability` | Free/busy lookup or meeting time suggestions |
| Contacts | `search_people` | Search contacts by name (limit 50) |

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

For **read-only** use, replace `Mail.ReadWrite` and `Mail.Send` with `Mail.Read`, and `Calendars.ReadWrite` / `Calendars.ReadWrite.Shared` with `Calendars.Read`. Write tools will return permission errors but everything else works.

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
| `bulk_manage_messages` | `dry_run=True` | Shows matches without executing |
| `delete_message` | `permanent=False` | Moves to Deleted Items (recoverable) |

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
