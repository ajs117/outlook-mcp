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


# Sender addresses seen on any message this process has fetched. Lets the
# keep-list be checked for a list of bare refs without spending a round trip
# re-fetching mail we already read.
_ID_TO_SENDER: dict[str, str] = {}


def _note_senders(messages: list[dict]) -> None:
    for m in messages:
        addr = ((m.get("from") or {}).get("emailAddress") or {}).get("address")
        if addr and m.get("id"):
            _ID_TO_SENDER[m["id"]] = addr.lower()


def sender_for(ref_or_id: str) -> str | None:
    """Best-known sender for a ref/id, or None if this process never saw it."""
    return _ID_TO_SENDER.get(resolve_id(ref_or_id))


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


# Graph accepts these well-known names directly in place of a folder id.
WELL_KNOWN_FOLDERS = frozenset(
    "archive clutter conflicts conversationhistory deleteditems drafts inbox "
    "junkemail localfailures msgfolderroot outbox recoverableitemsdeletions "
    "scheduled searchfolders sentitems serverfailures syncissues".split()
)


def create_folder(display_name: str, parent_id: str | None = None) -> dict:
    """Create a mail folder (top-level, or under parent_id if given)."""
    path = "/me/mailFolders"
    if parent_id:
        path = f"/me/mailFolders/{quote(parent_id, safe='')}/childFolders"
    with _client() as c:
        r = c.post(path, json={"displayName": display_name})
        r.raise_for_status()
        return r.json()


def resolve_folder(name_or_id: str) -> str:
    """Turn a human folder name into something Graph's move endpoint accepts.
    Well-known names pass through; a display name is looked up (case-insensitive)
    against the user's folders and swapped for its id; anything else (assumed to
    already be an id) is returned unchanged."""
    key = (name_or_id or "").strip()
    if key.lower() in WELL_KNOWN_FOLDERS:
        return key.lower()
    for f in list_folders():
        if (f.get("displayName") or "").lower() == key.lower():
            return f["id"]
    return name_or_id


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


def _build_list_params(
    *,
    top: int,
    skip: int,
    filter_query: str | None,
    order_by: str,
    search: str | None,
    fetch_all: bool,
    max_results: int,
    include_preview: bool,
) -> dict:
    """Build the OData query params for a message listing.

    Kept pure (no I/O) so the query-shaping rules are unit-testable without a
    network/auth round-trip. The critical rule: Graph rejects `$search` when it
    is combined with `$orderby`, `$filter`, or `$skip` — a `$search` request is
    relevance-ordered and paged via `$skiptoken` (the `@odata.nextLink`) only.
    Sending `$skip` alongside `$search` is a 400, so `$skip` lives strictly in
    the non-search branch."""
    select = MESSAGE_FIELDS + ",bodyPreview" if include_preview else MESSAGE_FIELDS
    page_size = min(max_results, 999) if fetch_all else min(top, 999)
    params: dict = {"$top": page_size, "$select": select}
    if search:
        params["$search"] = f'"{search}"'
    else:
        params["$orderby"] = order_by
        if filter_query:
            params["$filter"] = filter_query
        if not fetch_all:
            params["$skip"] = skip
    return params


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
    preview per message across a large folder is a major token cost.

    Note: when `search` is set, `skip`, `order_by` and `filter_query` are
    ignored — Graph does not allow them alongside `$search`."""
    params = _build_list_params(
        top=top,
        skip=skip,
        filter_query=filter_query,
        order_by=order_by,
        search=search,
        fetch_all=fetch_all,
        max_results=max_results,
        include_preview=include_preview,
    )
    with _client() as c:
        path = f"/me/mailFolders/{quote(folder, safe='')}/messages"
        if fetch_all:
            msgs = _paginated_get(c, path, params, max_results)
        else:
            r = c.get(path, params=params)
            r.raise_for_status()
            msgs = r.json().get("value", [])
    _note_senders(msgs)
    return msgs


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
    destination_folder = resolve_folder(destination_folder)
    with _client() as c:
        r = c.post(f"/me/messages/{quote(message_id, safe='')}/move", json={"destinationId": destination_folder})
        r.raise_for_status()
        return r.json()


def delete_message(message_id: str, mark_as_read: bool = True) -> None:
    """Move the message to the *visible* Deleted Items folder, marking it read
    first so deleted mail doesn't keep inflating the unread count.

    Deliberately uses the /move action, NOT the HTTP DELETE verb. Graph's DELETE
    on a message does not put it in Deleted Items — it soft-deletes into the
    hidden Recoverable Items ("dumpster") area, where the message vanishes from
    Outlook entirely and is only retrievable via "Recover items deleted from
    this folder". That is not what a user means by "delete", so we move instead.

    Order matters: a move mints a NEW message id, so the read flag must be set
    before the move, not after.
    """
    message_id = resolve_id(message_id)
    if mark_as_read:
        mark_read(message_id, True)
    move_message(message_id, "deleteditems")


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
    destination_folder = resolve_folder(destination_folder)
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


def bulk_delete_messages(message_ids: list[str], mark_as_read: bool = True) -> dict:
    """Move each message to the visible Deleted Items folder, marking them read
    first — same semantics as delete_message(), and for the same reason it uses
    /move rather than the HTTP DELETE verb (DELETE hides mail in the Recoverable
    Items dumpster).

    The read pass runs before the move because moving mints new message ids."""
    message_ids = [resolve_id(m) for m in message_ids]
    marked = bulk_mark_read(message_ids, True)["ok"] if mark_as_read else None
    res = bulk_move_messages(message_ids, "deleteditems")
    if marked is not None:
        res["marked_read"] = marked
    return res


def create_draft(
    subject: str,
    body: str,
    to: list[str],
    cc: list[str] | None = None,
    body_type: str = "Text",
) -> dict:
    """Create a brand-new draft message (not a reply). Draft-only: this never
    sends. Mail.ReadWrite (our scope) can create drafts but cannot send, so the
    human sends it from Outlook after review."""
    msg = {
        "subject": subject,
        "body": {"contentType": body_type, "content": body},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
    }
    if cc:
        msg["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc]
    with _client() as c:
        r = c.post("/me/messages", json=msg)
        r.raise_for_status()
        return r.json()


# --- Newsletter / unsubscribe -----------------------------------------------
def get_headers(message_id: str) -> dict:
    """Fetch a single message's internet (RFC 5322) headers plus from/subject."""
    message_id = resolve_id(message_id)
    with _client() as c:
        r = c.get(
            f"/me/messages/{quote(message_id, safe='')}",
            params={"$select": "internetMessageHeaders,from,subject,id"},
        )
        r.raise_for_status()
        return r.json()


def batch_get_headers(message_ids: list[str]) -> list[dict]:
    """Fetch internetMessageHeaders for many messages via $batch (chunked to
    Graph's 20/call limit). Returns the message dicts that came back OK, in
    request order; failed sub-requests are dropped."""
    message_ids = [resolve_id(m) for m in message_ids]
    requests = [
        {
            "id": str(i),
            "method": "GET",
            "url": f"/me/messages/{quote(mid, safe='')}?$select=internetMessageHeaders,from,subject,id",
        }
        for i, mid in enumerate(message_ids)
    ]
    out = []
    for resp in _batch(requests):
        if 200 <= resp.get("status", 0) < 300 and isinstance(resp.get("body"), dict):
            out.append(resp["body"])
    _note_senders(out)
    return out


def parse_unsubscribe(message: dict) -> dict:
    """Read the List-Unsubscribe / List-Unsubscribe-Post headers and classify
    the unsubscribe mechanism the sender offers.

    method is one of:
      one-click  - RFC 8058 one-click; POST to `url` unsubscribes (no browser).
      mailto     - send an email to `mailto` to unsubscribe.
      weblink    - only an https link; a human should open `url` in a browser.
      none       - no List-Unsubscribe header (probably not a bulk sender).

    `url` and `mailto` are both returned whenever the header offers them,
    regardless of which method won — a one-click POST that the sender's
    endpoint rejects (some privacy portals answer 405) can then fall back to
    the very same link a human would click, with no second header fetch.
    """
    import re

    headers = {
        (h.get("name") or "").lower(): (h.get("value") or "")
        for h in (message.get("internetMessageHeaders") or [])
    }
    lu = headers.get("list-unsubscribe")
    if not lu:
        return {"method": "none"}
    targets = re.findall(r"<([^>]+)>", lu)
    mailtos = [t for t in targets if t.lower().startswith("mailto:")]
    https = [t for t in targets if t.lower().startswith("http")]
    # RFC 8058: List-Unsubscribe-Post: List-Unsubscribe=One-Click
    one_click = "one-click" in headers.get("list-unsubscribe-post", "").lower()
    if one_click and https:
        method = "one-click"
    elif https:
        method = "weblink"
    elif mailtos:
        method = "mailto"
    else:
        return {"method": "unknown", "raw": lu}
    info = {"method": method}
    if https:
        info["url"] = https[0]
    if mailtos:
        info["mailto"] = mailtos[0]
    return info


def http_one_click_unsubscribe(url: str) -> int:
    """Perform an RFC 8058 one-click unsubscribe: POST the standard form body to
    the sender's List-Unsubscribe URL. Unauthenticated request to an external
    host by design (the URL comes from the email). Returns the HTTP status."""
    with httpx.Client(timeout=20.0, follow_redirects=True) as c:
        r = c.post(
            url,
            content="List-Unsubscribe=One-Click",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        return r.status_code
