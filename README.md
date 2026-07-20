# outlook-mcp

A local MCP server for a **personal** Outlook.com / Hotmail / Live account
(not a work/school Microsoft 365 account). It talks to Microsoft Graph
directly using an OAuth device-code login, so it runs entirely on your own
machine — no third-party service sees your mail.

Tools it exposes: `list_folders`, `inbox_summary`, `list_messages`,
`get_message`, `archive_message`, `move_message`, `delete_message`,
`mark_read`, `create_reply_draft`, bulk variants `bulk_mark_read`,
`bulk_move_messages`, `bulk_archive_messages`, `bulk_delete_messages`
(all backed by Graph's `$batch` endpoint, chunked to 20 ops/call),
`bulk_by_query` (match + act in one call, ids never leave the server),
`create_draft` (compose a new draft), `create_folder`, `find_newsletters`
and `unsubscribe`.

### Common workflows

- **Triage a big folder / clear old mail.** Use `bulk_by_query` with a
  `search` term and `received_before` to sweep, e.g. delete old shipping
  notifications: `action="delete", search="delivery", received_before="2026-06-01"`.
  It runs `dry_run` first (count + sample), then executes on the second call.
  Custom folders work by name — `action="move", destination_folder="Receipts"`
  resolves the name to its id (create it first with `create_folder` if needed).
- **Unsubscribe from newsletters.** `find_newsletters` lists bulk senders and
  how each can be unsubscribed; call `unsubscribe(ref, execute=True)` per sender.
  RFC 8058 one-click senders are unsubscribed with a single POST (no browser);
  `mailto` senders get an unsubscribe draft to send; link-only senders return a
  URL to open yourself (never auto-fetched). Then `bulk_by_query` to clear them.
  Note: unsubscribing confirms your address is live — do it for legitimate
  senders; for real spam, just delete/block.
- **Draft an email.** `create_draft(to, subject, body)` saves to Drafts and
  never sends — you review and send from Outlook. (This server holds
  `Mail.ReadWrite`, which can create drafts but deliberately cannot send.)
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
- **Lean listings.** `list_messages` returns only triage fields by default.
  For content, prefer `include_keywords=True` (a short salient-word list per
  message — denser than a preview) and reach for `get_message` for a full read.
  `include_preview=True` still attaches the full plaintext preview when wanted.
- **Plaintext, truncated bodies.** `get_message` asks Graph for the plaintext
  body (not raw HTML) and truncates to `max_chars` (default 4000);
  `include_body=False` gives metadata only.
- **Compact bulk results.** The `bulk_*` tools return `{"ok": <count>,
  "failed_count": <n>, "failed": [...]}` — succeeded ids are counted, not
  echoed back.
- **Filter-and-act — ids never touch the conversation.** `bulk_by_query`
  matches messages by folder/`filter_query`/`search` and applies one action
  (`archive`/`delete`/`mark_read`/`mark_unread`/`move`) to all of them
  server-side, returning just a count. This is the cheapest way to triage:
  no id list is ever streamed through the model. It runs `dry_run=True` by
  default (returns a count + small sample and does nothing) so a bad filter
  can't silently nuke the wrong mail — call again with `dry_run=False` to
  execute.

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
