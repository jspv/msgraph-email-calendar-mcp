#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from msgraph_mcp import auth, calendar, mail  # noqa: E402
from msgraph_mcp.errors import MsGraphMcpError  # noqa: E402


def print_json(data: object) -> None:
    print(json.dumps(data, indent=2, ensure_ascii=False))



def cmd_status(_: argparse.Namespace) -> int:
    print_json(auth.auth_status())
    return 0



def cmd_start_auth(_: argparse.Namespace) -> int:
    flow = auth.begin_device_flow()
    print_json(
        {
            "verification_uri": flow.verification_uri,
            "user_code": flow.user_code,
            "expires_in": flow.expires_in,
            "message": flow.message,
        }
    )
    return 0



def cmd_finish_auth(_: argparse.Namespace) -> int:
    result = auth.complete_device_flow()
    claims = result.get("id_token_claims", {}) or {}
    print_json(
        {
            "preferred_username": claims.get("preferred_username"),
            "name": claims.get("name"),
            "tenant_id": claims.get("tid"),
            "scope": result.get("scope"),
        }
    )
    return 0



def cmd_list_accounts(_: argparse.Namespace) -> int:
    print_json([item.model_dump() for item in auth.list_accounts()])
    return 0



def cmd_list_folders(args: argparse.Namespace) -> int:
    items = mail.list_folders(account_id=args.account_id, include_hidden=args.include_hidden)
    print_json([item.model_dump() for item in items])
    return 0



def cmd_list_messages(args: argparse.Namespace) -> int:
    items = mail.list_messages(account_id=args.account_id, folder=args.folder, limit=args.limit)
    print_json([item.model_dump() for item in items])
    return 0



def cmd_get_message(args: argparse.Namespace) -> int:
    item = mail.get_message(account_id=args.account_id, message_id=args.message_id)
    print_json(item.model_dump())
    return 0



def cmd_mark_message_read(args: argparse.Namespace) -> int:
    print_json(mail.mark_message_read(account_id=args.account_id, message_id=args.message_id, is_read=not args.unread))
    return 0



def cmd_move_message(args: argparse.Namespace) -> int:
    print_json(mail.move_message(account_id=args.account_id, message_id=args.message_id, destination=args.destination))
    return 0



def cmd_delete_message(args: argparse.Namespace) -> int:
    print_json(mail.delete_message(account_id=args.account_id, message_id=args.message_id, permanent=args.permanent))
    return 0



def cmd_search_messages(args: argparse.Namespace) -> int:
    items = mail.search_messages(account_id=args.account_id, query=args.query, limit=args.limit)
    print_json([item.model_dump() for item in items])
    return 0



def cmd_bulk_manage_messages(args: argparse.Namespace) -> int:
    print_json(
        mail.bulk_manage_messages_multi_pass(
            account_id=args.account_id,
            folder=args.folder,
            sender_contains=args.sender_contains,
            subject_contains=args.subject_contains,
            received_after=args.received_after,
            unread_only=args.unread_only,
            action=args.action,
            destination=args.destination,
            scan_limit=args.limit,
            dry_run=not args.apply,
            confirm_token=args.confirm_token,
        )
    )
    return 0



def cmd_check_all_day_event(args: argparse.Namespace) -> int:
    """Settle issue #10 against the live API: does Graph accept a non-midnight all-day event?

    The fix in `_graph_datetime` follows Graph's *documented* contract -- an
    all-day event must start and end at midnight in its stated zone -- but that
    was never confirmed by a real call. This probe sends the pre-fix payload
    (04:00, not midnight) and reports what actually happens.

    Creates a real event on --apply. Without it, only the payloads are printed.
    """
    offset = args.offset
    bad = {
        "subject": args.subject,
        "start": {"dateTime": f"{args.date}T04:00:00", "timeZone": "UTC"},
        "end": {"dateTime": f"{args.end_date}T04:00:00", "timeZone": "UTC"},
        "isAllDay": True,
    }
    print("Pre-fix payload (what the bug emitted for "
          f"{args.date}T00:00:00{offset}):")
    print_json(bad)
    print()
    print("Post-fix payload (what this repo now emits):")
    print_json(
        {
            "subject": args.subject,
            "start": calendar._graph_datetime(
                f"{args.date}T00:00:00{offset}", "start_iso", all_day=True
            ),
            "end": calendar._graph_datetime(
                f"{args.end_date}T00:00:00{offset}", "end_iso", all_day=True
            ),
            "isAllDay": True,
        }
    )
    print()

    if not args.apply:
        print("Dry run. Re-run with --apply to POST the pre-fix payload and see "
              "whether Graph 400s, silently rounds, or accepts it.")
        return 0

    from msgraph_mcp.graph import GraphClient

    client = GraphClient(args.account_id)
    try:
        created = client.request("POST", "/me/calendar/events", json_body=bad)
    except MsGraphMcpError as exc:
        print(f"RESULT: Graph REJECTED the non-midnight all-day event -> {exc}")
        print("=> Issue #10 confirmed; the midnight fix is required.")
        return 0

    print("RESULT: Graph ACCEPTED it. What it stored:")
    print_json({k: created.get(k) for k in ("id", "start", "end", "isAllDay")})
    stored = (created.get("start") or {}).get("dateTime", "")
    if "T00:00:00" in stored:
        print("=> Graph silently ROUNDED to midnight. The fix is still correct "
              "(it makes the stored value predictable), but #10 is not a 400.")
    else:
        print("=> Graph stored it as sent. Issue #10's premise does not hold; "
              "reconsider the fix.")
    print(f"\nClean up: delete_event(event_id={created.get('id')!r})")
    return 0


def cmd_list_calendars(args: argparse.Namespace) -> int:
    items = calendar.list_calendars(account_id=args.account_id)
    print_json([item.model_dump() for item in items])
    return 0



def cmd_list_events(args: argparse.Namespace) -> int:
    items = calendar.list_events(
        account_id=args.account_id,
        start_iso=args.start_iso,
        end_iso=args.end_iso,
        calendar_id=args.calendar_id,
        limit=args.limit,
    )
    print_json([item.model_dump() for item in items])
    return 0



def cmd_get_event(args: argparse.Namespace) -> int:
    item = calendar.get_event(account_id=args.account_id, event_id=args.event_id)
    print_json(item.model_dump())
    return 0



def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Smoke test harness for msgraph-mcp")
    subparsers = parser.add_subparsers(dest="command", required=True)

    status = subparsers.add_parser("status")
    status.set_defaults(func=cmd_status)

    start_auth = subparsers.add_parser("start-auth")
    start_auth.set_defaults(func=cmd_start_auth)

    finish_auth = subparsers.add_parser("finish-auth")
    finish_auth.set_defaults(func=cmd_finish_auth)

    list_accounts = subparsers.add_parser("list-accounts")
    list_accounts.set_defaults(func=cmd_list_accounts)

    list_folders = subparsers.add_parser("list-folders")
    list_folders.add_argument("--account-id")
    list_folders.add_argument("--include-hidden", action="store_true")
    list_folders.set_defaults(func=cmd_list_folders)

    list_messages = subparsers.add_parser("list-messages")
    list_messages.add_argument("--account-id")
    list_messages.add_argument("--folder", default="inbox")
    list_messages.add_argument("--limit", type=int, default=10)
    list_messages.set_defaults(func=cmd_list_messages)

    get_message = subparsers.add_parser("get-message")
    get_message.add_argument("message_id")
    get_message.add_argument("--account-id")
    get_message.set_defaults(func=cmd_get_message)

    mark_message_read = subparsers.add_parser("mark-message-read")
    mark_message_read.add_argument("message_id")
    mark_message_read.add_argument("--account-id")
    mark_message_read.add_argument("--unread", action="store_true")
    mark_message_read.set_defaults(func=cmd_mark_message_read)

    move_message = subparsers.add_parser("move-message")
    move_message.add_argument("message_id")
    move_message.add_argument("destination")
    move_message.add_argument("--account-id")
    move_message.set_defaults(func=cmd_move_message)

    delete_message = subparsers.add_parser("delete-message")
    delete_message.add_argument("message_id")
    delete_message.add_argument("--account-id")
    delete_message.add_argument("--permanent", action="store_true")
    delete_message.set_defaults(func=cmd_delete_message)

    search_messages = subparsers.add_parser("search-messages")
    search_messages.add_argument("query")
    search_messages.add_argument("--account-id")
    search_messages.add_argument("--limit", type=int, default=10)
    search_messages.set_defaults(func=cmd_search_messages)

    bulk_manage = subparsers.add_parser("bulk-manage-messages")
    bulk_manage.add_argument("--account-id")
    bulk_manage.add_argument("--folder", default="inbox")
    bulk_manage.add_argument("--sender-contains")
    bulk_manage.add_argument("--subject-contains")
    bulk_manage.add_argument("--received-after")
    bulk_manage.add_argument("--unread-only", action="store_true")
    bulk_manage.add_argument("--action", default="delete", choices=["delete", "move", "mark_read", "mark_unread"])
    bulk_manage.add_argument("--destination")
    bulk_manage.add_argument(
        "--limit", type=int, default=None,
        help="Max messages to scan; omit to scan the entire folder",
    )
    bulk_manage.add_argument("--apply", action="store_true")
    bulk_manage.add_argument(
        "--confirm-token",
        help="Token from a preceding dry run; required with --apply for delete/move",
    )
    bulk_manage.set_defaults(func=cmd_bulk_manage_messages)

    all_day_probe = subparsers.add_parser(
        "check-all-day-event",
        help="Issue #10: does Graph reject a non-midnight all-day event?",
    )
    all_day_probe.add_argument("--account-id")
    all_day_probe.add_argument("--subject", default="all-day probe (issue #10)")
    all_day_probe.add_argument("--date", default="2027-04-01")
    all_day_probe.add_argument("--end-date", default="2027-04-02")
    all_day_probe.add_argument("--offset", default="-04:00")
    all_day_probe.add_argument(
        "--apply",
        action="store_true",
        help="Actually POST the pre-fix payload (creates a real event)",
    )
    all_day_probe.set_defaults(func=cmd_check_all_day_event)

    list_calendars = subparsers.add_parser("list-calendars")
    list_calendars.add_argument("--account-id")
    list_calendars.set_defaults(func=cmd_list_calendars)

    list_events = subparsers.add_parser("list-events")
    list_events.add_argument("--account-id")
    list_events.add_argument("--start-iso")
    list_events.add_argument("--end-iso")
    list_events.add_argument("--calendar-id")
    list_events.add_argument("--limit", type=int, default=25)
    list_events.set_defaults(func=cmd_list_events)

    get_event = subparsers.add_parser("get-event")
    get_event.add_argument("event_id")
    get_event.add_argument("--account-id")
    get_event.set_defaults(func=cmd_get_event)

    return parser



def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return args.func(args)
    except MsGraphMcpError as exc:
        print_json({"error": str(exc)})
        return 2
    except Exception as exc:  # pragma: no cover
        print_json({"error": f"Unexpected error: {exc}"})
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
