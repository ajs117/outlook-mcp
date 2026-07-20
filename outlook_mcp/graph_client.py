"""Thin wrapper around the Microsoft Graph mail endpoints."""
import itertools
from urllib.parse import quote

import httpx

from .auth import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Graph's $batch endpoint accepts at most 20 sub-requests per call.
BATCH_CHUNK_SIZE = 20

# Lean field set: only what the shaped tool output actually returns. Notably
# excludes bodyPreview by default (see list_messages(include_preview=...)),
# since a full 255-char preview per message across a large folder dominates
# token cost.
MESSAGE_FIELDS = "id,subject,from,receivedDateTime,isRead,hasAttachments,webLink"

# --- Opaque short refs -------------------------------------------------------
# Real Graph message ids are ~150 chars of opaque base64. Emitting them for
# every listed message (and echoing them back on bulk ops) is the single
# largest token sink in this server. We hand the model a tiny stable ref
# ("m1", "m2", ...) instead and resolve it back to the real id server-side.
# The map lives for the process lifetime; if the server restarts, unknown refs
# are passed through unchanged (Graph will 400 and the model just re-lists).
_REF_TO_ID: dict[str, str] = {}
_ID_TO_REF: dict[str, str] = {}
_ref_counter = itertools.count(1)


def ref_for(real_id: str) -> str:
    """Return a stable short ref for a real Graph id, minting one if needed."""
    ref = _ID_TO_REF.get(real_id)
    if ref is None:
        ref = f"m{next(_ref_counter)}"
        _ID_TO_REF[real_id] = ref
        _REF_TO_ID[ref] = real_id
    return ref


def resolve_id(ref_or_id: str) -> str:
    """Map a short ref back to its real Graph id; pass through anything else."""
    return _REF_TO_ID.get(ref_or_id, ref_or_id)


def _client(extra_headers: dict | None = None) -> httpx.Client:
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    if extra_headers:
        headers.update(extra_headers)
    return httpx.Client(base_url=GRAPH_BASE, headers=headers, timeout=30.0)


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
    """Compact result: a success count plus only the failures (as refs, with a
    short reason). Echoing every succeeded id back is pure token waste, so we
    return just the number that worked."""
    ok = 0
    failed = []
    for message_id, resp in zip(message_ids, responses):
        if 200 <= resp.get("status", 0) < 300:
            ok += 1
        else:
            body = resp.get("body") or {}
            reason = (body.get("error") or {}).get("message") if isinstance(body, dict) else str(body)
            failed.append({"ref": ref_for(message_id), "status": resp.get("status"), "error": reason})
    return {"ok": ok, "failed_count": len(failed), "failed": failed}


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
    include_preview: bool = False,
) -> list[dict]:
    """List messages. By default returns a single page of `top` messages
    (Graph caps a single page at 999). Set fetch_all=True to follow
    pagination and collect up to max_results messages across pages.

    bodyPreview is only fetched when include_preview is True, since a full
    preview per message across a large folder is a major token cost."""
    select = MESSAGE_FIELDS + ",bodyPreview" if include_preview else MESSAGE_FIELDS
    params = {"$top": min(top, 999) if not fetch_all else min(max_results, 999), "$select": select}
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


def get_message(message_id: str, include_body: bool = True, text_body: bool = True) -> dict:
    """Fetch one message. When text_body is True we ask Graph for the plaintext
    body (via the Prefer header) instead of the raw HTML, which is far smaller
    and cleaner. Set include_body=False to skip the body entirely."""
    message_id = resolve_id(message_id)
    select = MESSAGE_FIELDS + ",bodyPreview,body" if include_body else MESSAGE_FIELDS + ",bodyPreview"
    headers = {"Prefer": 'outlook.body-content-type="text"'} if (include_body and text_body) else None
    with _client(headers) as c:
        r = c.get(f"/me/messages/{quote(message_id, safe='')}", params={"$select": select})
        r.raise_for_status()
        return r.json()


def move_message(message_id: str, destination_folder: str) -> dict:
    message_id = resolve_id(message_id)
    with _client() as c:
        r = c.post(f"/me/messages/{quote(message_id, safe='')}/move", json={"destinationId": destination_folder})
        r.raise_for_status()
        return r.json()


def delete_message(message_id: str) -> None:
    """Moves the message to Deleted Items (Graph DELETE on /messages is a
    soft delete to the Deleted Items folder, not a permanent purge)."""
    message_id = resolve_id(message_id)
    with _client() as c:
        r = c.delete(f"/me/messages/{quote(message_id, safe='')}")
        r.raise_for_status()


def mark_read(message_id: str, is_read: bool = True) -> dict:
    message_id = resolve_id(message_id)
    with _client() as c:
        r = c.patch(f"/me/messages/{quote(message_id, safe='')}", json={"isRead": is_read})
        r.raise_for_status()
        return r.json()


def create_reply_draft(message_id: str, comment: str = "") -> dict:
    message_id = resolve_id(message_id)
    with _client() as c:
        r = c.post(f"/me/messages/{quote(message_id, safe='')}/createReply", json={"comment": comment})
        r.raise_for_status()
        return r.json()


def bulk_mark_read(message_ids: list[str], is_read: bool = True) -> dict:
    message_ids = [resolve_id(m) for m in message_ids]
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
    message_ids = [resolve_id(m) for m in message_ids]
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
    message_ids = [resolve_id(m) for m in message_ids]
    requests = [
        {"id": str(i), "method": "DELETE", "url": f"/me/messages/{quote(mid, safe='')}"}
        for i, mid in enumerate(message_ids)
    ]
    return _bulk_result(message_ids, _batch(requests))
