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
        "Calendars.Read",
    )
    token_cache_path: Path = Path(".data/msal_token_cache.json")
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    timeout_seconds: float = 30.0
    max_list_limit: int = 50
    max_event_limit: int = 100
    max_attachment_inline_size: int = 1_572_864  # 1.5 MB



def load_settings() -> Settings:
    """Build a ``Settings`` instance from the current environment."""
    client_id = os.getenv("MICROSOFT_CLIENT_ID", "").strip()
    tenant_id = os.getenv("MICROSOFT_TENANT_ID", "common").strip() or "common"
    scopes_raw = os.getenv(
        "MICROSOFT_SCOPES",
        "User.Read Mail.ReadWrite Calendars.Read",
    )
    scopes = tuple(part for part in scopes_raw.split() if part)
    token_cache_path = Path(
        os.getenv("MICROSOFT_TOKEN_CACHE_PATH", ".data/msal_token_cache.json")
    ).expanduser()
    max_attachment_inline_size = int(
        os.getenv("MAX_ATTACHMENT_INLINE_SIZE", "1572864")
    )
    return Settings(
        client_id=client_id,
        tenant_id=tenant_id,
        scopes=scopes,
        token_cache_path=token_cache_path,
        max_attachment_inline_size=max_attachment_inline_size,
    )


settings = load_settings()
