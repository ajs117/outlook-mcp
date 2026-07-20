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


@mcp.tool()
def list_messages(
    folder: str = "inbox",
    top: int = 25,
    skip: int = 0,
    filter_query: Optional[str] = None,
    search: Optional[str] = None,
    fetch_all: bool = False,
    max_results: int = 500,
) -> list[dict]:
    """List messages in a folder, newest first (unless searching).

    Args:
        folder: folder id or well-known name (inbox, archive, deleteditems, sentitems, drafts, junkemail).
        top: max number of messages to return per page (ignored if fetch_all is True).
        skip: number of messages to skip, for manual paging (ignored if fetch_all is True).
        filter_query: raw OData $filter, e.g. "receivedDateTime le 2026-01-01T00:00:00Z" or "isRead eq false".
        search: free-text search (Graph $search); when set, filter_query/order are ignored.
        fetch_all: if True, follow pagination automatically and return up to max_results messages
            instead of a single page of `top`. Use this for bulk operations over a whole folder.
        max_results: cap on messages returned when fetch_all is True (safety limit, default 500).
    """
    msgs = graph.list_messages(
        folder=folder,
        top=top,
        skip=skip,
        filter_query=filter_query,
        search=search,
        fetch_all=fetch_all,
        max_results=max_results,
    )
    return [
        {
            "id": m["id"],
            "subject": m.get("subject"),
            "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
            "fromName": (m.get("from") or {}).get("emailAddress", {}).get("name"),
            "receivedDateTime": m.get("receivedDateTime"),
            "isRead": m.get("isRead"),
            "hasAttachments": m.get("hasAttachments"),
            "preview": m.get("bodyPreview"),
        }
        for m in msgs
    ]


@mcp.tool()
def get_message(message_id: str) -> dict:
    """Get full details (including body) for a single message by id."""
    m = graph.get_message(message_id)
    return {
        "id": m["id"],
        "subject": m.get("subject"),
        "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "receivedDateTime": m.get("receivedDateTime"),
        "isRead": m.get("isRead"),
        "body": (m.get("body") or {}).get("content"),
        "bodyType": (m.get("body") or {}).get("contentType"),
        "webLink": m.get("webLink"),
    }


@mcp.tool()
def archive_message(message_id: str) -> dict:
    """Move a message to the Archive folder."""
    return graph.move_message(message_id, "archive")


@mcp.tool()
def move_message(message_id: str, destination_folder: str) -> dict:
    """Move a message to an arbitrary folder (id or well-known name)."""
    return graph.move_message(message_id, destination_folder)


@mcp.tool()
def delete_message(message_id: str) -> str:
    """Move a message to Deleted Items (soft delete, recoverable from Deleted Items until it's emptied)."""
    graph.delete_message(message_id)
    return f"Moved {message_id} to Deleted Items."


@mcp.tool()
def mark_read(message_id: str, is_read: bool = True) -> dict:
    """Mark a message as read or unread."""
    return graph.mark_read(message_id, is_read)


@mcp.tool()
def create_reply_draft(message_id: str, comment: str = "") -> dict:
    """Create a draft reply to a message (does NOT send it) so a human can review/send from Outlook."""
    d = graph.create_reply_draft(message_id, comment)
    return {"draft_id": d.get("id"), "webLink": d.get("webLink")}


@mcp.tool()
def bulk_mark_read(message_ids: list[str], is_read: bool = True) -> dict:
    """Mark many messages as read/unread in one call (Graph $batch, up to hundreds of ids).

    Returns {"succeeded": [ids...], "failed": [{"id", "status", "error"}...]}.
    """
    return graph.bulk_mark_read(message_ids, is_read)


@mcp.tool()
def bulk_move_messages(message_ids: list[str], destination_folder: str) -> dict:
    """Move many messages to a folder in one call.

    Args:
        message_ids: message ids to move.
        destination_folder: folder id or well-known name (archive, deleteditems, inbox, etc.).

    Returns {"succeeded": [ids...], "failed": [{"id", "status", "error"}...]}.
    """
    return graph.bulk_move_messages(message_ids, destination_folder)


@mcp.tool()
def bulk_archive_messages(message_ids: list[str]) -> dict:
    """Move many messages to the Archive folder in one call.

    Returns {"succeeded": [ids...], "failed": [{"id", "status", "error"}...]}.
    """
    return graph.bulk_archive_messages(message_ids)


@mcp.tool()
def bulk_delete_messages(message_ids: list[str]) -> dict:
    """Soft-delete many messages (move to Deleted Items) in one call.

    Returns {"succeeded": [ids...], "failed": [{"id", "status", "error"}...]}.
    """
    return graph.bulk_delete_messages(message_ids)


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
