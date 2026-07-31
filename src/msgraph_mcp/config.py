"""Application settings loaded from environment variables and .env file."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    """Immutable configuration for the MCP server.

    All values are populated from environment variables at import time
    via ``load_settings()``.  See ``.env.example`` for the full list.
    """
    client_id: str
    tenant_id: str = "common"
    scopes: tuple[str, ...] = (
        "User.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.ReadWrite",
        "Calendars.ReadWrite.Shared",
        "People.Read",
    )
    token_cache_path: Path = Path(".data/msal_token_cache.json")
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    timeout_seconds: float = 30.0
    max_list_limit: int = 1000
    max_event_limit: int = 100
    max_attachment_inline_size: int = 1_572_864  # 1.5 MB
    #: IANA or Windows zone name applied to calendar times written without a UTC
    #: offset. Left unset, such times are *refused* rather than guessed at: this
    #: is an MCP server, so the caller is usually a model turning "book me 2pm
    #: Thursday" into a bare wall-clock string, and reading that as UTC books a
    #: real meeting at the wrong hour with invitations already sent.
    default_timezone: str | None = None
    #: Request Graph's immutable message ids (``Prefer: IdType="ImmutableId"``).
    #: All-or-nothing per client, never per call: mixing id types in one session
    #: is how an id resolved under one regime gets passed to a call using the
    #: other. Flipping this invalidates any id a caller has already persisted.
    immutable_ids: bool = False



def load_settings() -> Settings:
    """Build a ``Settings`` instance from the current environment."""
    client_id = os.getenv("MICROSOFT_CLIENT_ID", "").strip()
    tenant_id = os.getenv("MICROSOFT_TENANT_ID", "common").strip() or "common"
    scopes_raw = os.getenv(
        "MICROSOFT_SCOPES",
        "User.Read Mail.ReadWrite Mail.Send Calendars.ReadWrite Calendars.ReadWrite.Shared People.Read",
    )
    scopes = tuple(part for part in scopes_raw.split() if part)
    token_cache_path = Path(
        os.getenv("MICROSOFT_TOKEN_CACHE_PATH", ".data/msal_token_cache.json")
    ).expanduser()
    max_attachment_inline_size = int(
        os.getenv("MAX_ATTACHMENT_INLINE_SIZE", "1572864")
    )
    max_list_limit = int(os.getenv("MAX_LIST_LIMIT", "1000"))
    default_timezone = os.getenv("MSGRAPH_DEFAULT_TIMEZONE", "").strip() or None
    immutable_ids = os.getenv("GRAPH_IMMUTABLE_IDS", "").strip().lower() in {"1", "true", "yes"}
    return Settings(
        client_id=client_id,
        tenant_id=tenant_id,
        scopes=scopes,
        token_cache_path=token_cache_path,
        max_attachment_inline_size=max_attachment_inline_size,
        max_list_limit=max_list_limit,
        default_timezone=default_timezone,
        immutable_ids=immutable_ids,
    )


settings = load_settings()
