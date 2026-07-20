"""
Personal Outlook.com MCP server.

Exposes read/triage tools for a personal Microsoft account's mailbox over
Microsoft Graph. Meant to run locally on your own machine (stdio transport)
and be added to Claude Desktop / Cowork as a local MCP connector.

Setup:
  1. Create a free Azure app registration (see auth.py docstring).
  2. export OUTLOOK_MCP_CLIENT_ID=<your app's client id>
  3. Add this server to your MCP config (see README.md).
  4. First tool call triggers a one-time device-code login in your terminal.
"""
from typing import Optional

from mcp.server.fastmcp import FastMCP

from . import graph_client as graph

mcp = FastMCP("outlook-personal")


@mcp.tool()
def list_folders() -> list[dict]:
    """List all mail folders (Inbox, Archive, Deleted Items, custom folders, etc.) with their ids and display names."""
    folders = graph.list_folders()
    return [
        {
            "id": f["id"],
            "displayName": f["displayName"],
            "totalItemCount": f.get("totalItemCount"),
            "unreadItemCount": f.get("unreadItemCount"),
        }
        for f in folders
    ]


@mcp.tool()
def inbox_summary() -> dict:
    """Quick counts for the inbox: total messages and unread messages."""
    return graph.get_folder_counts("inbox")


def _shape_list_item(m: dict, preview_chars: int) -> dict:
    """Compact per-message row for listings. `ref` is a short opaque handle
    (e.g. "m12") accepted by every other tool in place of the ~150-char Graph
    id; pass refs, not raw ids, to keep calls small."""
    sender = (m.get("from") or {}).get("emailAddress", {})
    item = {
        "ref": graph.ref_for(m["id"]),
        "subject": m.get("subject"),
        "from": sender.get("address"),
        "date": m.get("receivedDateTime"),
        "unread": not m.get("isRead", True),
    }
    if m.get("hasAttachments"):
        item["hasAttachments"] = True
    preview = m.get("bodyPreview")
    if preview:
        item["preview"] = preview[:preview_chars]
    return item


@mcp.tool()
def list_messages(
    folder: str = "inbox",
    top: int = 25,
    skip: int = 0,
    filter_query: Optional[str] = None,
    search: Optional[str] = None,
    fetch_all: bool = False,
    max_results: int = 500,
    include_preview: bool = False,
    preview_chars: int = 140,
) -> list[dict]:
    """List messages in a folder, newest first (unless searching).

    Returns compact rows keyed by a short `ref` handle (not the raw Graph id) —
    pass that ref to get_message / move / delete / the bulk_* tools.

    Args:
        folder: folder id or well-known name (inbox, archive, deleteditems, sentitems, drafts, junkemail).
        top: max number of messages to return per page (ignored if fetch_all is True).
        skip: number of messages to skip, for manual paging (ignored if fetch_all is True).
        filter_query: raw OData $filter, e.g. "receivedDateTime le 2026-01-01T00:00:00Z" or "isRead eq false".
        search: free-text search (Graph $search); when set, filter_query/order are ignored.
        fetch_all: if True, follow pagination automatically and return up to max_results messages
            instead of a single page of `top`. Use this for bulk operations over a whole folder.
        max_results: cap on messages returned when fetch_all is True (safety limit, default 500).
        include_preview: include a short body preview per message (off by default to save tokens).
        preview_chars: max characters of preview per message when include_preview is True.
    """
    msgs = graph.list_messages(
        folder=folder,
        top=top,
        skip=skip,
        filter_query=filter_query,
        search=search,
        fetch_all=fetch_all,
        max_results=max_results,
        include_preview=include_preview,
    )
    return [_shape_list_item(m, preview_chars) for m in msgs]


@mcp.tool()
def get_message(message_id: str, include_body: bool = True, max_chars: int = 4000) -> dict:
    """Get one message by ref (or raw id). Body is returned as plaintext and
    truncated to max_chars to keep it token-cheap.

    Args:
        message_id: a short ref from list_messages (e.g. "m12"), or a raw Graph id.
        include_body: set False to get metadata only (subject/from/date), no body.
        max_chars: truncate the plaintext body to this many characters (0 = no limit).
    """
    m = graph.get_message(message_id, include_body=include_body)
    out = {
        "ref": graph.ref_for(m["id"]),
        "subject": m.get("subject"),
        "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "date": m.get("receivedDateTime"),
        "unread": not m.get("isRead", True),
        "webLink": m.get("webLink"),
    }
    if include_body:
        body = (m.get("body") or {}).get("content") or ""
        if max_chars and len(body) > max_chars:
            out["body"] = body[:max_chars]
            out["body_truncated"] = True
        else:
            out["body"] = body
    return out


@mcp.tool()
def archive_message(message_id: str) -> dict:
    """Move a message (by ref or raw id) to the Archive folder."""
    m = graph.move_message(message_id, "archive")
    return {"ref": graph.ref_for(m["id"]), "movedTo": "archive"}


@mcp.tool()
def move_message(message_id: str, destination_folder: str) -> dict:
    """Move a message (by ref or raw id) to a folder (id or well-known name)."""
    m = graph.move_message(message_id, destination_folder)
    return {"ref": graph.ref_for(m["id"]), "movedTo": destination_folder}


@mcp.tool()
def delete_message(message_id: str) -> dict:
    """Move a message (by ref or raw id) to Deleted Items (soft delete, recoverable until emptied)."""
    graph.delete_message(message_id)
    return {"ref": message_id, "deleted": True}


@mcp.tool()
def mark_read(message_id: str, is_read: bool = True) -> dict:
    """Mark a message (by ref or raw id) as read or unread."""
    m = graph.mark_read(message_id, is_read)
    return {"ref": graph.ref_for(m["id"]), "isRead": is_read}


@mcp.tool()
def create_reply_draft(message_id: str, comment: str = "") -> dict:
    """Create a draft reply to a message (by ref or raw id). Does NOT send it; a human reviews/sends from Outlook."""
    d = graph.create_reply_draft(message_id, comment)
    return {"draft_ref": graph.ref_for(d["id"]), "webLink": d.get("webLink")}


# Bulk tools take a list of refs (from list_messages) or raw ids, and return a
# compact {"ok": <count>, "failed_count": <n>, "failed": [{ref,status,error}...]}
# — the succeeded ids are NOT echoed back, only counted, to keep results small.
@mcp.tool()
def bulk_mark_read(message_ids: list[str], is_read: bool = True) -> dict:
    """Mark many messages (refs or raw ids) read/unread in one call (Graph $batch)."""
    return graph.bulk_mark_read(message_ids, is_read)


@mcp.tool()
def bulk_move_messages(message_ids: list[str], destination_folder: str) -> dict:
    """Move many messages (refs or raw ids) to a folder (id or well-known name) in one call."""
    return graph.bulk_move_messages(message_ids, destination_folder)


@mcp.tool()
def bulk_archive_messages(message_ids: list[str]) -> dict:
    """Move many messages (refs or raw ids) to the Archive folder in one call."""
    return graph.bulk_archive_messages(message_ids)


@mcp.tool()
def bulk_delete_messages(message_ids: list[str]) -> dict:
    """Soft-delete many messages (refs or raw ids) — move to Deleted Items — in one call."""
    return graph.bulk_delete_messages(message_ids)


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
