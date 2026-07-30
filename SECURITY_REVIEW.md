# Security Review - msgraph-mcp

Scope of this review: the MCP tool surface in `tools.py` and the Graph client
beneath it. Last reconciled against the code on 2026-07-29.

## Authentication

- Microsoft device-code flow with **delegated** permissions via MSAL. Every
  operation runs as the signed-in user; there are no application permissions.
- Default scopes (`config.py`): `User.Read`, `Mail.ReadWrite`, `Mail.Send`,
  `Calendars.ReadWrite`, `Calendars.ReadWrite.Shared`, `People.Read`.
  This is a **read/write/send** default, not a read-only one.
- Token cache stored locally at `.data/msal_token_cache.json` (configurable).
  - Parent directory created `0700`, cache file `0600`.
  - Symlinked cache directories are rejected.
  - Cache contents grant mailbox access — anyone with filesystem access to the
    runtime can read, send, and delete mail.
- `auth_status` does not expose the token cache path.
- `finish_auth` returns a minimal account summary, not full `id_token_claims`.

### Framework-managed mode

When `GRAPH_ACCESS_TOKEN`, `OAUTH_AUTH_URL`, `OAUTH_AUTHENTICATED`, or
`SERVICE_NAME` is set, the server assumes it is running inside the MCP Lambda
wrapper and takes the access token from the environment, bypassing MSAL and
skipping local cache creation. Consequences:

- The token's real scopes are whatever the framework granted. The server does
  not verify them; `MICROSOFT_SCOPES` only controls which tools are *offered*.
- Any process able to set `GRAPH_ACCESS_TOKEN` in the server's environment
  controls the mailbox the server acts on.

## Tool exposure is gated by scope

`_requires_scope` in `tools.py` registers a tool only when `settings.scopes`
contains a scope that satisfies it. A read-only scope set never advertises
`delete_message`, `send_message`, or `create_event`, so a prompt-injected model
cannot invoke what was never registered.

This is a **surface-reduction** control, not an authorization control. It
narrows what the model can attempt; the token still bounds what Graph permits.
Narrowing `MICROSOFT_SCOPES` without re-authenticating does not shrink an
already-cached token's authority.

`manage_draft` spans two scopes (updating needs `Mail.ReadWrite`, sending needs
`Mail.Send`) and registers under either; its update path checks for
`Mail.ReadWrite` at call time and refuses rather than issuing a doomed request.

## Graph API request safety

- **Next-link hardening**: Pagination only follows `@odata.nextLink` URLs that
  are HTTPS on the configured Microsoft Graph host. The base path prefix is
  stripped during normalization to prevent doubled path segments.
- **Path segment validation**: All user-supplied IDs (message, event, calendar,
  folder, attachment, `user_id`) are validated against `[A-Za-z0-9_\-=+.]+`
  before interpolation into URL paths. This blocks path traversal via `/`,
  `..`, `?`, `#`, or spaces.
- **OData parameter validation**:
  - `$search` — double quotes stripped from the query.
  - `$select` — caller-supplied `fields` validated against
    `[A-Za-z][A-Za-z0-9./]*`; `id` always forced in.
  - `$filter` — `since` is parsed as ISO-8601 before interpolation; a value
    that parses as a datetime cannot smuggle OData operators.
- **Datetime normalization**: `received_after` and calendar event times are
  parsed to tz-aware UTC. Offset-bearing event times are *converted*, not
  relabelled, so an event cannot be silently booked at the wrong hour.
- **Retry/backoff**: Automatic retry for 429 and transient 5xx. `Retry-After`
  parsing handles non-integer values safely.
- **Error translation**: Graph errors become `GraphRequestError` with extracted
  messages/codes. Raw response payloads are not surfaced to tool callers.

## Input bounds

- `list_messages` / `search_messages`: capped at `max_list_limit` (default
  1000, override `MAX_LIST_LIMIT`). Empty search queries rejected.
- `list_events`: capped at `max_event_limit` (default 100).
- `get_attachments`: attachment content is returned inline as base64 only below
  `MAX_ATTACHMENT_INLINE_SIZE` (default 1.5 MB); larger attachments return
  metadata with `content_omitted=True`.
- `bulk_manage_messages`: **not capped**. It scans the whole folder by default;
  an optional positive `limit` bounds the scan and is honored exactly. Paged
  internally at up to 1000 messages/request. An unfiltered whole-folder scan
  reads every message in the folder — bounded by folder size, not a fixed cap.

## Write operations

All of the following mutate user data. Each requires the corresponding scope
*and* a token that actually carries it.

### Mail — mutation (`Mail.ReadWrite`)

| Tool | What it does | Safety defaults |
|---|---|---|
| `update_message` | Read/unread, follow-up flag, categories | Only supplied fields change |
| `move_message` | Moves a message to another folder | Destination resolved through a known-folder map or validated as a safe ID |
| `delete_message` | Soft- or hard-deletes | `permanent=False` — soft delete |
| `bulk_manage_messages` | Applies an action to filtered messages folder-wide | `dry_run=True` — preview only; `delete`/`move` additionally require a `confirm_token` |
| `create_draft` | Creates a draft in Drafts | Not sent |
| `add_attachment_to_draft` | Uploads base64 content onto a draft | Size bounded only by Graph |
| `create_folder` | Creates a mail folder | — |

### Mail — outbound (`Mail.Send`)

Outbound tools transmit to third parties and are **not reversible**.

| Tool | What it does | Safety defaults |
|---|---|---|
| `send_message` | Sends a new email | `dry_run=True` — builds a preview draft, then deletes it |
| `reply_to_message` | Replies / replies-all | `dry_run=True` |
| `forward_message` | Forwards a message, including its attachments | `dry_run=True` |
| `manage_draft` | Updates and/or sends an existing draft | `send=False` |

Notes:

- The dry-run preview creates a **real draft** in the user's mailbox, reads it
  back, and deletes it in a `finally` block. A crash between create and delete
  no longer strands a draft, but a hard process kill still could.
- `send_as` sets the `from` address. Graph rejects addresses the account has no
  send-as right to, so this is not an impersonation primitive on its own.
- `forward_message` forwards the original message *with its attachments*. A
  prompt-injected agent calling it with `dry_run=False` is a data-exfiltration
  path — this is the highest-severity tool in the server.

### Calendar (`Calendars.ReadWrite`)

| Tool | What it does | Safety defaults |
|---|---|---|
| `create_event` | Creates an event, optionally inviting attendees | `dry_run=True`; invitations are sent by Graph on creation |
| `update_event` | Updates an event | `dry_run=True`; only supplied fields change; updates notify attendees |
| `delete_event` | Deletes, or cancels with a message to attendees | `dry_run=True`; `cancel_message=None` → hard delete, no notice |
| `respond_to_event` | Accepts / declines / tentatively accepts | `sendResponse` is always `True` — the organizer is always notified |

All four generate outbound email to attendees as a side effect. The three
destructive ones now preview by default. `respond_to_event` does not, and that is
deliberate: responding again reverses it, so gating it would add friction without
reducing risk.

`delete_event`'s preview names the event — subject, time, organizer, attendee
count — rather than echoing the id it was given, and states which of two very
different outcomes applies: with `cancel_message` attendees are **notified**;
without it the event is destroyed and **nobody is told**.

### Shared mailboxes and calendars (`Calendars.ReadWrite.Shared`)

Every calendar tool accepts an optional `user_id`. When supplied, requests
target `/users/{user_id}/…` instead of `/me/…`, so the server can **read,
create, modify, delete, and cancel events on another user's calendar** to the
extent the signed-in user has been granted access.

This is the widest blast radius in the server:

- `user_id` is validated as a safe path segment, but is otherwise unconstrained
  — there is no allowlist of which users may be targeted.
- Authorization is enforced entirely by Graph, based on delegated sharing
  permissions. The server performs no check of its own.
- Cancelling another user's meeting emails every attendee on it.

Deployments that do not need this should drop `Calendars.ReadWrite.Shared` from
`MICROSOFT_SCOPES` and re-authenticate.

### People (`People.Read`)

`search_people` reads the directory/contacts graph to resolve names to email
addresses. Read-only, but it is a directory-enumeration surface.

## Remaining risks

- **Prompt injection reaching outbound tools.** Message bodies and calendar
  bodies are attacker-controlled text delivered straight into a model's
  context. `forward_message`, `send_message`, and `reply_to_message` turn that
  into an exfiltration channel. Dry-run defaults help only if the client does
  not blindly re-invoke with `dry_run=False`.
- **Confirmation gates make scope visible; they do not stop injection.** Bulk
  `delete`/`move` now require a `confirm_token` derived from the matched message
  ids, and destructive calendar writes preview by default. Both force the
  affected set to appear in the transcript before anything happens. Neither
  stops a determined prompt injection: injected text can simply say "run the
  dry-run, then call again with the token," and a compliant model will. The
  value is **visibility to the human and the client**, not authorization.
  Documenting these as anti-injection defenses would be false.
- **Bulk operations are unbounded by default.** A whole-folder
  `bulk_manage_messages` with a loose filter can still move or delete every
  message in a folder — the confirm token bounds *surprise*, not *scale*. It
  guarantees the acting run touches the set the preview reported, and refuses if
  the mailbox shifted underneath. Soft-delete remains the only cushion.
- **Other write tools still execute immediately.** `update_message`,
  `move_message`, `delete_message`, `create_folder`, `manage_draft`, and
  `respond_to_event` act on the first call. Safety there depends on the MCP
  client or agent gating them.
- **Token cache at rest is unencrypted.** Local filesystem access grants full
  delegated mailbox access for the cached account's scopes.
- **No per-account or per-target allowlist.** `MICROSOFT_TENANT_ID` controls
  which tenants can authenticate, but once a token is cached there is no
  restriction on which `account_id` or `user_id` a tool may target.
- **Blocking retries.** Tools are synchronous and `time.sleep` on backoff, so a
  throttled bulk scan stalls the whole server process.

## Suggested improvements

1. Encrypt the token cache at rest where the deployment environment supports it.
2. Add an optional allowlist for `user_id` (shared-calendar targets) and for
   outbound recipient domains.
3. Add a `MSGRAPH_READ_ONLY=1` kill switch that drops all mutating tools
   regardless of token scope.
4. Extend the confirm-token gate to `forward_message`, the highest-severity
   outbound tool, so exfiltration also requires a two-step escalation.

Shipped in 0.2.0: the bulk confirmation token (previously item 4) and calendar
dry-run defaults (previously item 5).
