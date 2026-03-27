"""Entry point for the MCP server process."""

from __future__ import annotations

import sys

from .config import settings
from .tools import mcp


def main() -> None:
    """Validate configuration and start the FastMCP server."""
    if not settings.client_id:
        print("MICROSOFT_CLIENT_ID is required", file=sys.stderr)
        raise SystemExit(1)
    mcp.run()


if __name__ == "__main__":
    main()
