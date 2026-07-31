"""Low-level HTTP client for Microsoft Graph API requests.

Provides ``GraphClient`` for making authenticated requests with automatic
retry/backoff, pagination, and next-link security validation.

When running inside the MCP Lambda wrapper framework, the environment
variable ``GRAPH_ACCESS_TOKEN`` is set by the framework's credential
manager.  If present, it is used directly and MSAL is bypassed entirely.
"""

from __future__ import annotations

import atexit
import os
import re
import time
from typing import Any
from urllib.parse import urlparse

import httpx

from .auth import get_access_token
from .config import settings
from .errors import GraphRequestError

_SAFE_ID_PATTERN = re.compile(r"^[A-Za-z0-9_\-=+.]+$")

# One connection pool for the whole process. A client per request would mean a
# fresh TCP + TLS handshake per request, and the workload where that hurts is
# ``bulk_manage_messages_multi_pass``: its action loop calls ``delete_message``
# (or ``move_message``) once per message, each of which builds its own
# ``GraphClient``. Pooling inside the instance would therefore save nothing in
# the one case that needs it -- the pool has to outlive the instance.
#
# ``httpx.Client`` is safe for concurrent use and evicts broken connections from
# its pool on its own, so sharing does not introduce shared failure state.
_http: httpx.Client | None = None


def _http_client() -> httpx.Client:
    """Return the shared HTTP client, building it on first use."""
    global _http
    if _http is None or _http.is_closed:
        _http = httpx.Client(
            timeout=settings.timeout_seconds,
            limits=httpx.Limits(max_keepalive_connections=10, max_connections=20),
        )
    return _http


def close_http_client() -> None:
    """Dispose of the shared client. The next request builds a fresh one."""
    global _http
    if _http is not None:
        _http.close()
        _http = None


atexit.register(close_http_client)


def _merge_prefer(headers: dict[str, str], value: str) -> None:
    """Append to ``Prefer`` instead of overwriting it.

    ``Prefer`` is a comma-separated list. Assigning over it is how
    ``IdType="ImmutableId"`` would silently drop
    ``outlook.body-content-type="text"`` on exactly the calls that fetch a body
    -- the response is still valid, just HTML.
    """
    existing = headers.get("Prefer")
    if not existing:
        headers["Prefer"] = value
        return
    parts = [p.strip() for p in existing.split(",") if p.strip()]
    if value not in parts:
        parts.append(value)
    headers["Prefer"] = ", ".join(parts)


def validate_path_segment(value: str, label: str = "id") -> str:
    """Ensure *value* is safe for interpolation into a Graph API URL path.

    Raises ``ValueError`` if the value is empty or contains characters
    outside the safe set (alphanumeric, ``-``, ``_``, ``=``, ``+``, ``.``).
    """
    if not value or not _SAFE_ID_PATTERN.match(value):
        raise ValueError(f"Invalid {label}: must be non-empty and contain only safe characters")
    return value


class GraphClient:
    """Authenticated HTTP client for Microsoft Graph v1.0.

    Each instance is bound to a single account (or the default account
    when *account_id* is ``None``).  Requests are retried up to three
    times on 429 / 5xx responses with exponential backoff.
    """

    def __init__(self, account_id: str | None = None) -> None:
        self.account_id = account_id

    def _headers(self) -> dict[str, str]:
        # Framework-injected token takes precedence (Lambda mode)
        token = os.environ.get("GRAPH_ACCESS_TOKEN")
        if not token:
            # Fall back to MSAL (local/standalone mode)
            token = get_access_token(self.account_id)
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }
        if settings.immutable_ids:
            # Set here rather than per-call so it also rides on `paginate`'s
            # continuation requests, which re-enter `request` with a full URL and
            # no params. Anything keyed off params would apply to page 1 only,
            # and page 2 would come back with the other id type.
            headers["Prefer"] = 'IdType="ImmutableId"'
        return headers

    def _normalize_path(self, path: str) -> str:
        """Convert *path* to a relative path suitable for ``graph_base_url``.

        Full ``@odata.nextLink`` URLs are validated (must be HTTPS on the
        configured Graph host) and have the base path prefix stripped so
        the caller can prepend ``graph_base_url`` without duplication.
        """
        if path.startswith("http://") or path.startswith("https://"):
            parsed = urlparse(path)
            base = urlparse(settings.graph_base_url)
            if parsed.scheme != "https" or parsed.netloc != base.netloc:
                raise ValueError("Refusing to follow nextLink outside Microsoft Graph")
            result_path = parsed.path
            if base.path and result_path.startswith(base.path):
                result_path = result_path[len(base.path):]
            if not result_path.startswith("/"):
                result_path = "/" + result_path
            return result_path + (("?" + parsed.query) if parsed.query else "")
        if not path.startswith("/"):
            raise ValueError("Graph path must start with '/'")
        return path

    def request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        json_body: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
    ) -> dict[str, Any] | None:
        """Execute a single Graph API request with retry/backoff.

        Returns the parsed JSON response, or ``None`` for empty bodies
        (e.g. successful DELETE).  Raises ``GraphRequestError`` on failure.
        """
        request_headers = self._headers()
        if params and "$search" in params:
            request_headers["ConsistencyLevel"] = "eventual"
            _merge_prefer(request_headers, 'outlook.body-content-type="text"')
        elif params and "body" in str(params.get("$select", "")):
            _merge_prefer(request_headers, 'outlook.body-content-type="text"')
        for key, value in (headers or {}).items():
            if key.lower() == "prefer":
                _merge_prefer(request_headers, value)
            else:
                request_headers[key] = value
        headers = request_headers

        normalized_path = self._normalize_path(path)
        client = _http_client()
        retries = 3
        for attempt in range(retries):
            try:
                response = client.request(
                    method=method,
                    url=f"{settings.graph_base_url}{normalized_path}" if normalized_path.startswith("/") else normalized_path,
                    headers=headers,
                    params=params,
                    json=json_body,
                )
                if response.status_code == 429 and attempt < retries - 1:
                    try:
                        retry_after = int(response.headers.get("Retry-After", "2"))
                    except (ValueError, TypeError):
                        retry_after = 2
                    time.sleep(min(retry_after, 10))
                    continue
                if response.status_code >= 500 and attempt < retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                response.raise_for_status()
                if not response.content:
                    return None
                return response.json()
            except httpx.HTTPStatusError as exc:
                status = exc.response.status_code
                if status in {429, 500, 502, 503, 504} and attempt < retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                detail = "Microsoft Graph request failed"
                try:
                    payload = exc.response.json()
                    error = payload.get("error", {})
                    detail = error.get("message") or error.get("code") or detail
                except Exception:
                    pass
                raise GraphRequestError(detail, status_code=status) from exc
            except httpx.HTTPError as exc:
                if attempt < retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                raise GraphRequestError("Network error while calling Microsoft Graph") from exc
        raise GraphRequestError("Microsoft Graph request failed after retries")

    def paginate_delta(
        self,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        limit: int | None = None,
    ) -> tuple[list[dict[str, Any]], str | None]:
        """Walk a delta collection, returning its rows and the delta token.

        ``paginate`` discards ``@odata.deltaLink``, which is the only part of a
        delta response worth keeping -- it is the cursor for the next sync.

        Returns ``(rows, delta_link)``. The token is ``None`` when *limit* cut
        the walk short: a token from a partial read would not cover the rows
        that were never fetched, so handing one back would silently skip them
        on the following sync.
        """
        rows: list[dict[str, Any]] = []
        next_url: str | None = None
        while True:
            payload = (
                self.request("GET", next_url)
                if next_url
                else self.request("GET", path, params=params)
            ) or {}
            rows.extend(payload.get("value", []))
            if limit is not None and len(rows) >= limit:
                return rows[:limit], None
            next_url = payload.get("@odata.nextLink")
            if not next_url:
                return rows, payload.get("@odata.deltaLink")

    def paginate(
        self,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        limit: int | None = None,
    ) -> list[dict[str, Any]]:
        """Follow ``@odata.nextLink`` pages and return up to *limit* items."""
        results: list[dict[str, Any]] = []
        next_url: str | None = None
        first_params = params

        while True:
            if next_url:
                payload = self.request("GET", next_url)
            else:
                payload = self.request("GET", path, params=first_params)
            if not payload:
                break
            for item in payload.get("value", []):
                results.append(item)
                if limit and len(results) >= limit:
                    return results
            next_url = payload.get("@odata.nextLink")
            if not next_url:
                break
        return results
