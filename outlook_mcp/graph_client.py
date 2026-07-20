"""Thin wrapper around the Microsoft Graph mail endpoints."""
from urllib.parse import quote

import httpx

from .auth import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

MESSAGE_FIELDS = (
    "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,"
    "parentFolderId,webLink"
)


def _client() -> httpx.Client:
    token = get_access_token()
    return httpx.Client(
        base_url=GRAPH_BASE,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=30.0,
    )


def list_folders() -> list[dict]:
    with _client() as c:
        r = c.get("/me/mailFolders", params={"$top": 100})
        r.raise_for_status()
        return r.json().get("value", [])


def get_folder_counts(folder_id_or_name: str = "inbox") -> dict:
    with _client() as c:
        r = c.get(f"/me/mailFolders/{quote(folder_id_or_name, safe='')}")
        r.raise_for_status()
        data = r.json()
        return {
            "displayName": data.get("displayName"),
            "totalItemCount": data.get("totalItemCount"),
            "unreadItemCount": data.get("unreadItemCount"),
        }


def list_messages(
    folder: str = "inbox",
    top: int = 25,
    skip: int = 0,
    filter_query: str | None = None,
    order_by: str = "receivedDateTime desc",
    search: str | None = None,
) -> list[dict]:
    params = {"$top": top, "$skip": skip, "$select": MESSAGE_FIELDS}
    if search:
        # $search cannot be combined with $orderby or most $filter usage
        params["$search"] = f'"{search}"'
    else:
        params["$orderby"] = order_by
        if filter_query:
            params["$filter"] = filter_query
    with _client() as c:
        r = c.get(f"/me/mailFolders/{quote(folder, safe='')}/messages", params=params)
        r.raise_for_status()
        return r.json().get("value", [])


def get_message(message_id: str) -> dict:
    with _client() as c:
        r = c.get(f"/me/messages/{quote(message_id, safe='')}", params={"$select": MESSAGE_FIELDS + ",body"})
        r.raise_for_status()
        return r.json()


def move_message(message_id: str, destination_folder: str) -> dict:
    with _client() as c:
        r = c.post(f"/me/messages/{quote(message_id, safe='')}/move", json={"destinationId": destination_folder})
        r.raise_for_status()
        return r.json()


def delete_message(message_id: str) -> None:
    """Moves the message to Deleted Items (Graph DELETE on /messages is a
    soft delete to the Deleted Items folder, not a permanent purge)."""
    with _client() as c:
        r = c.delete(f"/me/messages/{quote(message_id, safe='')}")
        r.raise_for_status()


def mark_read(message_id: str, is_read: bool = True) -> dict:
    with _client() as c:
        r = c.patch(f"/me/messages/{quote(message_id, safe='')}", json={"isRead": is_read})
        r.raise_for_status()
        return r.json()


def create_reply_draft(message_id: str, comment: str = "") -> dict:
    with _client() as c:
        r = c.post(f"/me/messages/{quote(message_id, safe='')}/createReply", json={"comment": comment})
        r.raise_for_status()
        return r.json()
