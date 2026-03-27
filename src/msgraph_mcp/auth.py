"""Microsoft device-code authentication via MSAL.

Manages the two-step device-code flow (start -> user approves -> finish),
token caching on disk, and silent token refresh.
"""

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import msal

from .config import settings
from .errors import AuthError
from .models import AccountInfo


@dataclass
class PendingDeviceFlow:
    """Snapshot of a device-code flow returned to the caller."""
    flow: dict[str, Any]
    verification_uri: str
    user_code: str
    expires_in: int
    message: str


_pending_flow: dict[str, Any] | None = None



def _pending_flow_path() -> Path:
    return _cache_path().with_name("device_flow.json")



def _load_pending_flow() -> dict[str, Any] | None:
    global _pending_flow
    if _pending_flow is not None:
        return _pending_flow
    path = _pending_flow_path()
    try:
        data = json.loads(path.read_text())
    except FileNotFoundError:
        return None
    except json.JSONDecodeError:
        path.unlink(missing_ok=True)
        return None

    expires_at = data.get("expires_at")
    if expires_at and time.time() >= expires_at:
        path.unlink(missing_ok=True)
        return None

    flow = data.get("flow")
    if isinstance(flow, dict):
        _pending_flow = flow
        return flow
    return None



def _save_pending_flow(flow: dict[str, Any]) -> None:
    global _pending_flow
    _pending_flow = flow
    payload = {
        "flow": flow,
        "expires_at": time.time() + int(flow.get("expires_in", 900)),
    }
    path = _pending_flow_path()
    path.write_text(json.dumps(payload))
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass



def _clear_pending_flow() -> None:
    global _pending_flow
    _pending_flow = None
    _pending_flow_path().unlink(missing_ok=True)



def _cache_path() -> Path:
    path = settings.token_cache_path.resolve()
    parent = path.parent.resolve()
    parent.mkdir(parents=True, exist_ok=True)
    if parent.is_symlink():
        raise AuthError("Token cache directory must not be a symlink")
    return path



def _read_cache() -> str | None:
    path = _cache_path()
    try:
        return path.read_text()
    except FileNotFoundError:
        return None



def _write_cache(content: str) -> None:
    path = _cache_path()
    path.write_text(content)
    try:
        os.chmod(path.parent, 0o700)
        os.chmod(path, 0o600)
    except OSError:
        pass



def get_app() -> msal.PublicClientApplication:
    """Create an MSAL ``PublicClientApplication`` with the local token cache."""
    if not settings.client_id:
        raise AuthError("MICROSOFT_CLIENT_ID is required")

    authority = f"https://login.microsoftonline.com/{settings.tenant_id}"
    cache = msal.SerializableTokenCache()
    cache_content = _read_cache()
    if cache_content:
        cache.deserialize(cache_content)

    return msal.PublicClientApplication(
        settings.client_id,
        authority=authority,
        token_cache=cache,
    )



def _save_cache(app: msal.PublicClientApplication) -> None:
    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())



def list_accounts() -> list[AccountInfo]:
    """Return all accounts present in the local token cache."""
    app = get_app()
    return [
        AccountInfo(
            username=account.get("username", "unknown"),
            account_id=account.get("home_account_id", "unknown"),
        )
        for account in app.get_accounts()
    ]



def begin_device_flow() -> PendingDeviceFlow:
    """Initiate a device-code flow and persist it to disk for later completion."""
    app = get_app()
    flow = app.initiate_device_flow(scopes=list(settings.scopes))
    if "user_code" not in flow:
        raise AuthError(flow.get("error_description", "Failed to start device flow"))
    _save_pending_flow(flow)
    return PendingDeviceFlow(
        flow=flow,
        verification_uri=flow.get("verification_uri") or flow.get("verification_url") or "https://microsoft.com/devicelogin",
        user_code=flow["user_code"],
        expires_in=int(flow.get("expires_in", 900)),
        message=flow.get("message", "Authenticate with the provided code."),
    )



def complete_device_flow() -> dict[str, Any]:
    """Finish a pending device-code flow after the user has approved it."""
    flow = _load_pending_flow()
    if not flow:
        raise AuthError("No pending device flow. Start authentication first.")

    app = get_app()
    result = app.acquire_token_by_device_flow(flow)
    if "error" in result:
        raise AuthError(result.get("error_description", result["error"]))
    _save_cache(app)
    _clear_pending_flow()
    return result



def get_access_token(account_id: str | None = None) -> str:
    """Silently acquire a valid access token for *account_id*.

    Uses the first cached account when *account_id* is ``None``,
    ``"default"``, ``"me"``, or ``"primary"``.  Raises ``AuthError``
    if no matching account is found or silent refresh fails.
    """
    app = get_app()
    accounts = app.get_accounts()

    chosen = None
    if accounts:
        if not account_id or account_id in {"default", "me", "primary"}:
            chosen = accounts[0]
        else:
            for account in accounts:
                if account.get("home_account_id") == account_id or account.get("username", "").lower() == account_id.lower():
                    chosen = account
                    break

    if not chosen:
        raise AuthError("No authenticated account found. Authenticate first.")

    result = app.acquire_token_silent(list(settings.scopes), account=chosen)
    if not result:
        raise AuthError("Silent token acquisition failed. Re-authenticate.")
    if "error" in result:
        raise AuthError(result.get("error_description", result["error"]))

    _save_cache(app)
    return result["access_token"]



def auth_status() -> dict[str, Any]:
    """Return a summary of current auth configuration and cached accounts."""
    accounts = list_accounts()
    return {
        "configured": bool(settings.client_id),
        "tenant_id": settings.tenant_id,
        "scopes": list(settings.scopes),
        "accounts": [account.model_dump() for account in accounts],
        "pending_device_flow": _load_pending_flow() is not None,
        "token_cache_present": _cache_path().exists(),
    }
