"""
Device-code OAuth against Microsoft identity platform for a *personal*
Microsoft account (outlook.com / hotmail.com / live.com), using Microsoft
Graph. Tokens are cached on disk (encrypted-at-rest is left to the OS disk;
this is a personal single-user tool, not a multi-tenant service) so you only
have to log in once and then silently refresh.

Requires an app registration you create yourself in the Azure Portal
(free, no Azure subscription cost) with:
  - Supported account types: "Accounts in any organizational directory
    and personal Microsoft accounts"
  - Platform: "Mobile and desktop applications" -> allow the
    `https://login.microsoftonline.com/common/oauth2/nativeclient` redirect
    URI (needed for device code flow) OR just enable "Allow public client
    flows" = Yes under Authentication.
  - No client secret needed (public client).
  - API permissions (delegated, Microsoft Graph): Mail.ReadWrite,
    Mail.Read, MailboxSettings.Read, offline_access, User.Read
"""
import json
import os
import sys
import time
from pathlib import Path

import msal

# How long to wait for the user to complete the device-code login in their
# browser before giving up and returning an error to the caller, instead of
# blocking the MCP tool call for the full ~15 minute code lifetime.
DEVICE_FLOW_TIMEOUT_SECONDS = 120

CACHE_PATH = Path(os.environ.get("OUTLOOK_MCP_TOKEN_CACHE", str(Path.home() / ".outlook_mcp_token_cache.json")))

AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = [
    "Mail.ReadWrite",
    "MailboxSettings.Read",
    "User.Read",
]


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if CACHE_PATH.exists():
        cache.deserialize(CACHE_PATH.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        CACHE_PATH.write_text(cache.serialize())
        try:
            os.chmod(CACHE_PATH, 0o600)
        except OSError:
            pass


def get_access_token() -> str:
    """Return a valid Graph access token, running the interactive device
    code flow the first time and silently refreshing afterwards."""
    client_id = os.environ.get("OUTLOOK_MCP_CLIENT_ID")
    if not client_id:
        raise RuntimeError(
            "OUTLOOK_MCP_CLIENT_ID is not set. Create a free Azure app "
            "registration (personal Microsoft accounts) and set this env var "
            "to its Application (client) ID."
        )

    cache = _load_cache()
    app = msal.PublicClientApplication(client_id, authority=AUTHORITY, token_cache=cache)

    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to create device flow: {json.dumps(flow, indent=2)}")
        print(flow["message"], file=sys.stderr, flush=True)

        # Cap how long we block waiting for the browser login, so a
        # forgotten/abandoned login returns an error to the caller (and the
        # instructions below) instead of hanging for the code's full ~15
        # minute lifetime.
        flow["expires_at"] = min(flow["expires_at"], time.time() + DEVICE_FLOW_TIMEOUT_SECONDS)
        result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            _save_cache(cache)
            raise RuntimeError(
                f"{flow['message']}\n\n"
                "Login was not completed in time (or failed: "
                f"{result.get('error_description', result.get('error', 'unknown error'))}). "
                "Call this tool again to get a fresh code."
            )

    _save_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    return result["access_token"]
