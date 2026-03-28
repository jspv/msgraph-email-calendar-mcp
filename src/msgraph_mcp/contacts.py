"""People search via Microsoft Graph."""

from __future__ import annotations

from .graph import GraphClient
from .models import PersonResult


def search_people(
    account_id: str | None = None,
    *,
    query: str,
    limit: int = 10,
) -> list[PersonResult]:
    """Search for people by name to resolve email addresses."""
    client = GraphClient(account_id)
    payload = client.request(
        "GET",
        "/me/people",
        params={
            "$search": f'"{query.replace(chr(34), "")}"',
            "$top": min(limit, 50),
            "$select": "displayName,scoredEmailAddresses",
        },
    ) or {"value": []}
    results: list[PersonResult] = []
    for item in payload.get("value", []):
        emails = item.get("scoredEmailAddresses") or []
        email = emails[0].get("address") if emails else None
        results.append(PersonResult(
            name=item.get("displayName"),
            email=email,
        ))
    return results
