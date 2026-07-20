# outlook-mcp

A local MCP server for a **personal** Outlook.com / Hotmail / Live account
(not a work/school Microsoft 365 account). It talks to Microsoft Graph
directly using an OAuth device-code login, so it runs entirely on your own
machine — no third-party service sees your mail.

Tools it exposes: `list_folders`, `inbox_summary`, `list_messages`,
`get_message`, `archive_message`, `move_message`, `delete_message`,
`mark_read`, `create_reply_draft`, and bulk variants `bulk_mark_read`,
`bulk_move_messages`, `bulk_archive_messages`, `bulk_delete_messages`
(all backed by Graph's `$batch` endpoint, chunked to 20 ops/call).
`list_messages` also takes `fetch_all=True` to page past the default
25-message limit (up to `max_results`, default 500) instead of returning
a single page.

### Token efficiency

Graph message ids are ~150 characters of opaque base64 each, which dominates
token cost when listing or bulk-triaging a mailbox. To keep tool traffic small:

- **Short refs.** Listings return a compact `ref` (e.g. `m12`) per message
  instead of the raw id. Every tool accepts a ref in place of an id and
  resolves it server-side, so you pass `m12`, not a 150-char string. Refs are
  stable for the life of the server process; after a restart, just re-list.
- **Lean listings.** `list_messages` omits body previews by default
  (`include_preview=True` to opt in, truncated to `preview_chars`) and returns
  only the fields needed for triage.
- **Plaintext, truncated bodies.** `get_message` asks Graph for the plaintext
  body (not raw HTML) and truncates to `max_chars` (default 4000);
  `include_body=False` gives metadata only.
- **Compact bulk results.** The `bulk_*` tools return `{"ok": <count>,
  "failed_count": <n>, "failed": [...]}` — succeeded ids are counted, not
  echoed back.

## 1. Create a free Azure app registration (one-time, ~10 min)

You need this to get a `client_id`. It costs nothing — no Azure subscription
or credit card required for this.

1. Go to https://portal.azure.com and sign in with your Microsoft account
   (your outlook.com account works fine as the sign-in).
2. Search for **"App registrations"** and click **New registration**.
3. Name it anything, e.g. `outlook-mcp-personal`.
4. Under **Supported account types**, choose:
   **"Personal Microsoft accounts only"** (or the "any organizational
   directory and personal Microsoft accounts" option — either works).
5. Leave Redirect URI blank for now. Click **Register**.
6. Copy the **Application (client) ID** shown on the Overview page — you'll
   need it below.
7. In the left nav, go to **Authentication** → **Add a platform** →
   **Mobile and desktop applications** → check the
   `https://login.microsoftonline.com/common/oauth2/nativeclient` box →
   **Configure**.
8. Still on the Authentication page, scroll down to **Advanced settings**
   and set **"Allow public client flows"** to **Yes**. Save.
9. In the left nav, go to **API permissions** → **Add a permission** →
   **Microsoft Graph** → **Delegated permissions** → add `Mail.ReadWrite`,
   `MailboxSettings.Read`, `User.Read`, `offline_access` (offline_access is
   usually pre-added). Click **Add permissions**. (No admin consent needed
   for a personal account signing into their own app.)

That's it — no client secret needed, this is a "public client."

## 2. Install

```bash
cd outlook-mcp
pip install -e .
```

## 3. Configure

Set the client ID from step 1:

```bash
export OUTLOOK_MCP_CLIENT_ID="<your-application-client-id>"
```

(Put this in your shell profile so it persists.)

## 4. Add to Claude Desktop / Cowork

Edit your Claude Desktop MCP config (Settings → Developer → Edit Config, or
directly at `~/Library/Application Support/Claude/claude_desktop_config.json`
on macOS) and add:

```json
{
  "mcpServers": {
    "outlook-personal": {
      "command": "outlook-mcp",
      "env": {
        "OUTLOOK_MCP_CLIENT_ID": "<your-application-client-id>"
      }
    }
  }
}
```

Restart Claude Desktop. The first time you call a tool, the server prints a
device-login URL and code — open the URL in a browser, enter the code, sign
in with your outlook.com account, and approve. After that, tokens refresh
silently and you won't be prompted again (cached in
`~/.outlook_mcp_token_cache.json`, permissions 600).

## Notes / limitations

- `delete_message` is a soft delete (moves to Deleted Items), matching how
  Graph's message DELETE actually behaves — nothing is purged permanently
  by this tool.
- `create_reply_draft` only creates a draft; it never sends mail on its own.
- This is a personal single-user tool: the token cache is a plain JSON file
  on disk, meant for your own machine, not a shared/multi-user deployment.
- If you ever want to turn this into a shareable connector, the gap to
  close is: move off device-code flow to a proper auth-code + PKCE flow
  with a redirect, host it somewhere, get Microsoft's app verification for
  the requested scopes, and register it in Claude's MCP directory.
