"""A keep-list of senders that bulk operations must not touch by default.

Bulk triage is fast and irreversible-feeling; the mail that matters least in
volume (a benefits claim, a legal notice, a bill) is exactly the mail that
matters most to lose. Holding the list in the server rather than in whatever
notes the caller happens to be reading means every sweep is checked, not just
the ones where somebody remembered.
"""
import fnmatch
import json
import os
from pathlib import Path

DEFAULT_CONFIG_PATH = Path.home() / ".outlook_mcp_config.json"


def config_path() -> Path:
    return Path(os.environ.get("OUTLOOK_MCP_CONFIG") or DEFAULT_CONFIG_PATH)


def load_config() -> dict:
    path = config_path()
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_config(data: dict) -> Path:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, sort_keys=True)
    return path


def get_patterns() -> list[str]:
    raw = load_config().get("protected_senders") or []
    return [str(p).strip().lower() for p in raw if str(p).strip()]


def set_patterns(patterns: list[str]) -> list[str]:
    cleaned = sorted({str(p).strip().lower() for p in patterns if str(p).strip()})
    data = load_config()
    data["protected_senders"] = cleaned
    save_config(data)
    return cleaned


def match(address: str, patterns: list[str]) -> str | None:
    """Return the first pattern protecting `address`, or None.

    A pattern containing a glob character is matched with fnmatch against the
    whole address (`*@edfenergy.com`); one without is a plain substring test
    (`dwp`), which is what a human writing a keep-list usually means.
    """
    addr = (address or "").strip().lower()
    if not addr:
        return None
    for p in patterns:
        if any(ch in p for ch in "*?["):
            if fnmatch.fnmatch(addr, p):
                return p
        elif p in addr:
            return p
    return None


def partition(messages: list[dict], patterns: list[str], sender_of) -> tuple[list, list]:
    """Split messages into (unprotected, protected) using sender_of(msg)."""
    if not patterns:
        return list(messages), []
    keep, hit = [], []
    for m in messages:
        (hit if match(sender_of(m), patterns) else keep).append(m)
    return keep, hit
