# msgraph-mcp

Microsoft Graph MCP server for Outlook mail and calendar.

## Scope

- Delegated auth via Microsoft device code flow
- Enterprise/work account friendly by default
- List authenticated accounts and show auth status
- Read mail folders, messages, and message details
- Search messages
- Mark messages read/unread, move, and delete
- Bulk message management with filtering and dry-run support
- Read calendar list, events, and event details

## Azure app registration checklist

Create an app registration in Microsoft Entra / Azure for your work account tenant.

Recommended settings:

1. **Supported account types**
   - Accounts in this organizational directory only, or
   - Accounts in any organizational directory if you want broader org support

2. **Authentication**
   - Enable **Allow public client flows**
   - Device code flow should be allowed for public client usage

3. **Delegated Microsoft Graph permissions**

   | Permission | Why | Used by |
   |---|---|---|
   | `User.Read` | Read basic profile of the signed-in user | `auth_status`, `list_accounts` |
   | `Mail.Read` | Read mail folders and messages | `list_folders`, `list_messages`, `get_message`, `search_messages` |
   | `Mail.ReadWrite` | Mark read/unread, move, and delete messages | `mark_message_read`, `move_message`, `delete_message`, `bulk_manage_messages` |
   | `Calendars.Read` | Read calendars and events | `list_calendars`, `list_events`, `get_event` |

   If you only need read access, you can replace `Mail.ReadWrite` with `Mail.Read` in your `.env` `MICROSOFT_SCOPES` — the write tools will return permission errors but everything else works.

4. **Admin/user consent**
   - Grant consent as needed for the tenant

## Local setup

1. Copy `.env.example` to `.env`
2. Fill in:
   - `MICROSOFT_CLIENT_ID`
   - `MICROSOFT_TENANT_ID`
3. Install dependencies
4. Run the MCP server

Suggested tenant values:
- `organizations` for work/school accounts only
- specific tenant GUID for a single enterprise tenant

## Notes

- Uses device code flow with delegated permissions.
- Token cache is stored locally and excluded from git.
- Includes retry/backoff for 429 and transient 5xx Microsoft Graph failures.
- Mail and calendar tool outputs include extra agent-friendly summary/label fields so downstream consumers do less formatting work.
- Path parameters and search queries are validated/sanitized before use.
- See `SECURITY_REVIEW.md` for security notes.

## Smoke test harness

A lightweight local harness is included for first-run validation without full MCP client wiring.

Install dependencies first, then run examples like:

```bash
pip install -e .

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
