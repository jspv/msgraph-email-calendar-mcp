"""Exception hierarchy for Microsoft Graph MCP errors."""

from __future__ import annotations


class MsGraphMcpError(RuntimeError):
    """Base application error."""


class AuthError(MsGraphMcpError):
    """Authentication or token acquisition failed."""


class GraphRequestError(MsGraphMcpError):
    """Microsoft Graph request failed."""

    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code
