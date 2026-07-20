"""Thin wrapper around the Microsoft Graph mail endpoints."""
from urllib.parse import quote

import httpx

from .auth import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Graph's $batch endpoint accepts at most 20 sub-requests per call.
BATCH_CHUNK_SIZE = 20

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


def _paginated_get(c: httpx.Client, path: str, params: dict, max_results: int) -> list[dict]:
    """Follow @odata.nextLink until max_results items are collected or the
    folder is exhausted. nextLink is an absolute URL, so subsequent requests
    are made without base_url/params."""
    items: list[dict] = []
    r = c.get(path, params=params)
    r.raise_for_status()
    data = r.json()
    items.extend(data.get("value", []))
    next_link = data.get("@odata.nextLink")
    while next_link and len(items) < max_results:
        r = c.get(next_link)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
    return items[:max_results]


def _batch(requests: list[dict]) -> list[dict]:
    """Run Graph $batch sub-requests (each a dict with id/method/url/[body]),
    chunked to Graph's 20-per-call limit. Returns the combined list of
    response dicts (each with id/status/body), in request order."""
    with _client() as c:
        responses: dict[str, dict] = {}
        for i in range(0, len(requests), BATCH_CHUNK_SIZE):
            chunk = requests[i : i + BATCH_CHUNK_SIZE]
            r = c.post("/$batch", json={"requests": chunk})
            r.raise_for_status()
            for resp in r.json().get("responses", []):
                responses[resp["id"]] = resp
    return [responses[req["id"]] for req in requests]


def _bulk_result(message_ids: list[str], responses: list[dict]) -> dict:
    succeeded, failed = [], []
    for message_id, resp in zip(message_ids, responses):
        if 200 <= resp.get("status", 0) < 300:
            succeeded.append(message_id)
        else:
            failed.append({"id": message_id, "status": resp.get("status"), "error": resp.get("body")})
    return {"succeeded": succeeded, "failed": failed}


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
    fetch_all: bool = False,
    max_results: int = 500,
) -> list[dict]:
    """List messages. By default returns a single page of `top` messages
    (Graph caps a single page at 999). Set fetch_all=True to follow
    pagination and collect up to max_results messages across pages."""
    params = {"$top": min(top, 999) if not fetch_all else min(max_results, 999), "$select": MESSAGE_FIELDS}
    if not fetch_all:
        params["$skip"] = skip
    if search:
        # $search cannot be combined with $orderby or most $filter usage
        params["$search"] = f'"{search}"'
    else:
        params["$orderby"] = order_by
        if filter_query:
            params["$filter"] = filter_query
    with _client() as c:
        path = f"/me/mailFolders/{quote(folder, safe='')}/messages"
        if fetch_all:
            return _paginated_get(c, path, params, max_results)
        r = c.get(path, params=params)
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


def bulk_mark_read(message_ids: list[str], is_read: bool = True) -> dict:
    requests = [
        {
            "id": str(i),
            "method": "PATCH",
            "url": f"/me/messages/{quote(mid, safe='')}",
            "headers": {"Content-Type": "application/json"},
            "body": {"isRead": is_read},
        }
        for i, mid in enumerate(message_ids)
    ]
    return _bulk_result(message_ids, _batch(requests))


def bulk_move_messages(message_ids: list[str], destination_folder: str) -> dict:
    requests = [
        {
            "id": str(i),
            "method": "POST",
            "url": f"/me/messages/{quote(mid, safe='')}/move",
            "headers": {"Content-Type": "application/json"},
            "body": {"destinationId": destination_folder},
        }
        for i, mid in enumerate(message_ids)
    ]
    return _bulk_result(message_ids, _batch(requests))


def bulk_archive_messages(message_ids: list[str]) -> dict:
    return bulk_move_messages(message_ids, "archive")


def bulk_delete_messages(message_ids: list[str]) -> dict:
    """Soft-deletes each message (moves to Deleted Items), same semantics as
    delete_message()."""
    requests = [
        {"id": str(i), "method": "DELETE", "url": f"/me/messages/{quote(mid, safe='')}"}
        for i, mid in enumerate(message_ids)
    ]
    return _bulk_result(message_ids, _batch(requests))
