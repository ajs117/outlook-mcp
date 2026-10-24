"""One-time device-code login for outlook-mcp.

Run this in a terminal you can see:

    python login.py

It prints a URL and a code. Open the URL, enter the code, sign in with your
outlook.com account and approve. The resulting token is written to
~/.outlook_mcp_token_cache.json, which the MCP server reads on startup — so
after this completes once, tool calls work silently and never block.

This exists because the server's own login prompt is written to stderr, which
some MCP hosts capture into a log file rather than showing to the user. When
that happens the first tool call just appears to hang until it times out.
"""
import os
import sys
from pathlib import Path

import msal

CLIENT_ID = os.environ.get(
    "OUTLOOK_MCP_CLIENT_ID", "5a059d84-b9e8-41be-b5f2-8be1ab403176"
)
CACHE_PATH = Path(
    os.environ.get(
        "OUTLOOK_MCP_TOKEN_CACHE", str(Path.home() / ".outlook_mcp_token_cache.json")
    )
)
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Mail.ReadWrite", "MailboxSettings.Read", "User.Read"]


def main() -> int:
    cache = msal.SerializableTokenCache()
    if CACHE_PATH.exists():
        cache.deserialize(CACHE_PATH.read_text())

    app = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache
    )

    # If a usable token is already cached, don't make the user log in again.
    for account in app.get_accounts():
        result = app.acquire_token_silent(SCOPES, account=account)
        if result and "access_token" in result:
            print("Already logged in as", account.get("username", "?"))
            return 0

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("Could not start device login:", flow.get("error_description", flow))
        return 1

    print()
    print("  1. Open:", flow["verification_uri"])
    print("  2. Enter code:", flow["user_code"])
    print("  3. Sign in with your outlook.com account and approve.")
    print()
    print("Waiting for sign-in...")
    sys.stdout.flush()

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print("Login failed:", result.get("error"), result.get("error_description", ""))
        return 1

    CACHE_PATH.write_text(cache.serialize())
    try:
        os.chmod(CACHE_PATH, 0o600)
    except OSError:
        pass  # best-effort; Windows ACLs don't map onto POSIX modes

    username = result.get("id_token_claims", {}).get("preferred_username", "?")
    print()
    print("Login OK — signed in as", username)
    print("Token cached at", CACHE_PATH)
    print("The MCP server will now authenticate silently.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
