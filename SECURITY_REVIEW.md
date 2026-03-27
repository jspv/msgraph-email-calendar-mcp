# Security Review - msgraph-mcp

## Authentication

- Uses Microsoft device-code flow with delegated permissions via MSAL.
- Scopes: `User.Read`, `Mail.ReadWrite`, `Calendars.Read`.
- Token cache is stored locally at `.data/msal_token_cache.json` (configurable).
  - Parent directory is created with `0700` permissions; cache file with `0600`.
  - Symlinked cache directories are rejected.
  - Cache contents grant mailbox access — anyone with filesystem access to the runtime can read/write mail.
- `auth_status` does not expose the token cache path.
- `finish_auth` returns a minimal account summary, not the full `id_token_claims` object.

## Graph API request safety

- **Next-link hardening**: Pagination only follows `@odata.nextLink` URLs that are HTTPS on the configured Microsoft Graph host. The base path prefix is stripped during normalization to prevent doubled path segments.
- **Path segment validation**: All user-supplied IDs (message, event, calendar, folder) are validated against a safe-character pattern (`[A-Za-z0-9_\-=+.]+`) before interpolation into URL paths. This blocks path traversal via `/`, `..`, `?`, `#`, or spaces.
- **Search query sanitization**: Double-quote characters are stripped from OData `$search` queries to prevent malformed expressions.
- **Retry/backoff**: Automatic retry with backoff for 429 (rate limit) and transient 5xx responses. `Retry-After` header parsing handles non-integer values safely.
- **Error translation**: Graph API error responses are translated to `GraphRequestError` with extracted error messages/codes. Raw response payloads are not surfaced to MCP tool callers.

## Input bounds

- `list_messages` limit: capped at `max_list_limit` (default 50).
- `list_events` limit: capped at `max_event_limit` (default 100).
- `search_messages` limit: capped at `max_list_limit` (default 50). Empty queries rejected.
- `bulk_manage_messages` limit per pass: capped at `max_event_limit` (default 100). Max passes: capped at 10.

## Write operations

Write actions require `Mail.ReadWrite` scope and a fresh auth flow that includes it.

| Tool | What it does | Safety defaults |
|---|---|---|
| `mark_message_read` | Toggles read/unread flag | `is_read=True` |
| `move_message` | Moves a message to another folder | Folder name is resolved through a known-folder map or validated as a safe ID |
| `delete_message` | Soft-deletes (moves to Deleted Items) or permanently deletes | `permanent=False` — defaults to soft delete |
| `bulk_manage_messages` | Applies an action to filtered messages across multiple pages | `dry_run=True` — defaults to preview only |

### Bulk operation risk

`bulk_manage_messages` can process up to **1000 messages** in a single call (100/pass x 10 passes at maximum bounds). Safeguards:

- **Dry-run by default**: `dry_run=True` returns a preview of matched messages without acting.
- **No automatic escalation**: There is no built-in confirmation step between dry-run and live execution. The MCP client or calling agent decides whether to proceed. A misbehaving or prompt-injected agent could call with `dry_run=False` directly.
- **Soft delete by default**: The `delete` action moves messages to Deleted Items rather than permanently removing them.

## Remaining risks

- **Token cache at rest**: The cache file is not encrypted. Local filesystem access grants full delegated mailbox access for the cached account's scopes.
- **No per-operation confirmation**: Write and bulk tools execute immediately when called. Safety depends entirely on the MCP client or agent gating calls appropriately.
- **Single-tenant by default**: The `MICROSOFT_TENANT_ID` setting controls which tenants can authenticate, but there is no per-account allowlist once a token is cached.

## Suggested improvements

1. Encrypt token cache at rest when the deployment environment supports it.
2. Add optional account or tenant allowlist checks.
3. Add a `Calendars.ReadWrite` scope and calendar write operations only behind an explicit opt-in.
