"""
VK Community Messages Exporter
Exports all messages from a VK community via VK API.
"""

import csv
import json
import logging
import os
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv()

# ── Configuration ────────────────────────────────────────────────────────────

VK_TOKEN = os.getenv("VK_TOKEN", "")
GROUP_ID = os.getenv("GROUP_ID", "")
VK_API_VERSION = os.getenv("VK_API_VERSION", "5.131")
EXPORT_DIR = Path(os.getenv("EXPORT_DIR", "export"))

REQUEST_DELAY = float(os.getenv("REQUEST_DELAY", "0.34"))   # ~3 req/s limit
RETRY_COUNT = int(os.getenv("RETRY_COUNT", "3"))
RETRY_DELAY = float(os.getenv("RETRY_DELAY", "5.0"))
PAGE_SIZE = 200                                               # VK max per call

# ── Logging ──────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ── VK API client ─────────────────────────────────────────────────────────────


class VKAPIError(Exception):
    pass


def vk_call(method: str, params: dict, *, retry: int = RETRY_COUNT) -> dict:
    """Call a VK API method with automatic retry on transient errors."""
    url = f"https://api.vk.com/method/{method}"
    payload = {
        "access_token": VK_TOKEN,
        "v": VK_API_VERSION,
        **params,
    }

    for attempt in range(1, retry + 2):
        try:
            time.sleep(REQUEST_DELAY)
            resp = requests.post(url, data=payload, timeout=30)
            resp.raise_for_status()
            data = resp.json()

            if "error" in data:
                err = data["error"]
                code = err.get("error_code", 0)
                msg = err.get("error_msg", "unknown")
                # Rate-limit (6) or captcha (14): wait and retry
                if code in (6, 14) and attempt <= retry:
                    wait = RETRY_DELAY * attempt
                    log.warning(
                        "VK error %s (%s). Retry %d/%d in %.0fs…",
                        code, msg, attempt, retry, wait,
                    )
                    time.sleep(wait)
                    continue
                raise VKAPIError(f"VK error {code}: {msg}")

            return data.get("response", {})

        except requests.RequestException as exc:
            if attempt <= retry:
                wait = RETRY_DELAY * attempt
                log.warning(
                    "Network error: %s. Retry %d/%d in %.0fs…",
                    exc, attempt, retry, wait,
                )
                time.sleep(wait)
            else:
                raise

    raise VKAPIError(f"Failed after {retry} retries: {method}")


# ── Helpers ───────────────────────────────────────────────────────────────────


def unix_to_iso(ts: int) -> str:
    return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


def determine_sender_type(from_id: int, group_id_int: int) -> str:
    """
    from_id > 0  → regular user (client)
    from_id < 0  → community or bot
    from_id == -group_id → this community itself (manager/auto-reply)
    """
    if from_id > 0:
        return "client"
    if from_id == -group_id_int:
        return "community"
    if from_id < 0:
        return "manager"
    return "unknown"


def extract_attachments(attachments: list) -> list[dict]:
    """Extract type + best available URL from each attachment."""
    result = []
    for att in attachments:
        att_type = att.get("type", "unknown")
        obj = att.get(att_type, {})
        url = _pick_url(att_type, obj)
        result.append({"type": att_type, "url": url})
    return result


def _pick_url(att_type: str, obj: dict) -> str:
    """Return the most useful URL for a given attachment object."""
    if att_type == "photo":
        sizes = obj.get("sizes", [])
        if sizes:
            best = max(sizes, key=lambda s: s.get("width", 0) * s.get("height", 0))
            return best.get("url", "")
        return obj.get("photo_2560") or obj.get("photo_1280") or obj.get("photo_807") or ""

    if att_type == "video":
        # Direct MP4 not always available; return player page
        vid_id = obj.get("id", "")
        owner_id = obj.get("owner_id", "")
        if vid_id and owner_id:
            return f"https://vk.com/video{owner_id}_{vid_id}"
        return ""

    if att_type == "audio":
        return obj.get("url", "")

    if att_type == "doc":
        return obj.get("url", "")

    if att_type == "link":
        return obj.get("url", "")

    if att_type == "sticker":
        images = obj.get("images", [])
        if images:
            return images[-1].get("url", "")
        return ""

    if att_type == "wall":
        owner = obj.get("owner_id", "")
        post_id = obj.get("id", "")
        if owner and post_id:
            return f"https://vk.com/wall{owner}_{post_id}"
        return ""

    return obj.get("url", "")


# ── Fetchers ──────────────────────────────────────────────────────────────────


def fetch_all_conversations(group_id_int: int) -> list[dict]:
    """Return all conversations for the community (paginated)."""
    conversations = []
    offset = 0

    log.info("Fetching conversations list…")
    while True:
        resp = vk_call("messages.getConversations", {
            "group_id": group_id_int,
            "offset": offset,
            "count": PAGE_SIZE,
            "extended": 1,
            "fields": "first_name,last_name,screen_name,name,photo_50",
        })

        items = resp.get("items", [])
        if not items:
            break

        conversations.extend(items)
        total = resp.get("count", 0)
        offset += len(items)
        log.info("  Loaded %d / %d conversations", len(conversations), total)

        if offset >= total:
            break

    log.info("Total conversations found: %d", len(conversations))
    return conversations


def fetch_all_messages(peer_id: int, group_id_int: int) -> list[dict]:
    """Return every message in a conversation (paginated, oldest-first)."""
    messages = []
    start_message_id = 0   # 0 = latest; we'll walk backwards

    # We gather going backwards then reverse, so the list is chronological.
    while True:
        params = {
            "peer_id": peer_id,
            "group_id": group_id_int,
            "count": PAGE_SIZE,
            "rev": 0,        # newest first
            "extended": 0,
        }
        if start_message_id:
            params["start_message_id"] = start_message_id
            params["offset"] = 1   # skip the anchor itself

        resp = vk_call("messages.getHistory", params)
        items = resp.get("items", [])
        if not items:
            break

        messages.extend(items)

        if len(items) < PAGE_SIZE:
            break                   # reached the oldest message

        # Prepare next page: start from the oldest message we have so far
        start_message_id = items[-1]["id"]

    messages.reverse()
    return messages


# ── Message normaliser ────────────────────────────────────────────────────────


def normalise_message(
    msg: dict,
    group_id_int: int,
    conversation: dict,
) -> dict:
    """Convert a raw VK message dict into the export schema."""
    conv_settings = conversation.get("conversation", {}).get("chat_settings", {})
    conv_title = conv_settings.get("title", "")

    # For direct (non-chat) conversations derive title from profiles/groups
    # embedded in the conversation extended data if available.
    if not conv_title:
        conv_title = conversation.get("_peer_title", "")

    from_id = msg.get("from_id", 0)
    attachments_raw = msg.get("attachments", [])
    reply_raw = msg.get("reply_message")

    return {
        "conversation_id": msg.get("peer_id"),
        "message_id": msg.get("id"),
        "date": unix_to_iso(msg.get("date", 0)),
        "date_unix": msg.get("date", 0),
        "sender_id": from_id,
        "sender_type": determine_sender_type(from_id, group_id_int),
        "text": msg.get("text", ""),
        "attachments": extract_attachments(attachments_raw),
        "reply_to_message_id": reply_raw.get("id") if reply_raw else None,
        "conversation_title": conv_title,
    }


# ── Peer title resolution ──────────────────────────────────────────────────────


def resolve_peer_titles(conversations: list[dict]) -> None:
    """Inject _peer_title into each conversation item using extended profiles/groups."""
    # VK returns profiles and groups as top-level keys when extended=1
    # But they're per-page; we reconstruct from what we have.
    # For simplicity we use peer type + id.
    for item in conversations:
        conv = item.get("conversation", {})
        peer = conv.get("peer", {})
        peer_type = peer.get("type", "")
        peer_id = peer.get("id", 0)

        if peer_type == "chat":
            title = conv.get("chat_settings", {}).get("title", f"Chat {peer_id}")
        elif peer_type == "user":
            title = f"User {peer_id}"
        elif peer_type == "group":
            title = f"Group {peer_id}"
        else:
            title = f"Peer {peer_id}"

        item["_peer_title"] = title


# ── Export writers ─────────────────────────────────────────────────────────────

CSV_FIELDS = [
    "conversation_id",
    "conversation_title",
    "message_id",
    "date",
    "sender_id",
    "sender_type",
    "text",
    "attachments",
    "reply_to_message_id",
]


def save_json(all_data: list[dict], path: Path) -> None:
    path.write_text(json.dumps(all_data, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("JSON saved → %s", path)


def save_csv(all_messages: list[dict], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS, extrasaction="ignore")
        writer.writeheader()
        for msg in all_messages:
            row = dict(msg)
            # Serialise attachments list as a compact JSON string for CSV
            atts = row.get("attachments", [])
            row["attachments"] = json.dumps(atts, ensure_ascii=False) if atts else ""
            writer.writerow(row)
    log.info("CSV saved → %s", path)


# ── Main ──────────────────────────────────────────────────────────────────────


def validate_config() -> int:
    if not VK_TOKEN:
        raise SystemExit("ERROR: VK_TOKEN is not set. Copy .env.example → .env and fill it in.")
    if not GROUP_ID:
        raise SystemExit("ERROR: GROUP_ID is not set. Copy .env.example → .env and fill it in.")
    try:
        gid = int(GROUP_ID)
    except ValueError:
        raise SystemExit("ERROR: GROUP_ID must be an integer (e.g. 123456789).")
    if gid < 0:
        gid = -gid   # allow negative form
    return gid


def main() -> None:
    group_id_int = validate_config()
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)

    log.info("=== VK Community Messages Exporter ===")
    log.info("Group ID : %d", group_id_int)
    log.info("Export   : %s", EXPORT_DIR.resolve())

    # ── Step 1: conversations ─────────────────────────────────────────────
    conversations = fetch_all_conversations(group_id_int)
    resolve_peer_titles(conversations)

    if not conversations:
        log.warning("No conversations found. Check token permissions.")
        return

    # ── Step 2: messages per conversation ─────────────────────────────────
    export_data: list[dict] = []          # full structured data for JSON
    flat_messages: list[dict] = []        # flattened rows for CSV
    total_messages = 0
    processed = 0

    for idx, conv_item in enumerate(conversations, start=1):
        conv = conv_item.get("conversation", {})
        peer = conv.get("peer", {})
        peer_id = peer.get("id", 0)
        title = conv_item.get("_peer_title", str(peer_id))

        log.info(
            "[%d/%d] Conversation peer_id=%d  (%s)",
            idx, len(conversations), peer_id, title,
        )

        try:
            messages_raw = fetch_all_messages(peer_id, group_id_int)
        except VKAPIError as exc:
            log.error("  Skipping conversation %d: %s", peer_id, exc)
            continue

        messages_norm = [
            normalise_message(m, group_id_int, conv_item)
            for m in messages_raw
        ]

        export_data.append({
            "peer_id": peer_id,
            "conversation_title": title,
            "peer_type": peer.get("type", "unknown"),
            "message_count": len(messages_norm),
            "messages": messages_norm,
        })
        flat_messages.extend(messages_norm)
        total_messages += len(messages_norm)
        processed += 1

        log.info("  → %d messages loaded (total so far: %d)", len(messages_norm), total_messages)

    # ── Step 3: save ──────────────────────────────────────────────────────
    log.info("─" * 60)
    log.info("Processed: %d / %d conversations", processed, len(conversations))
    log.info("Total messages: %d", total_messages)

    save_json(export_data, EXPORT_DIR / "vk_conversations_raw.json")
    save_csv(flat_messages, EXPORT_DIR / "vk_messages_flat.csv")

    log.info("=== Done ===")


if __name__ == "__main__":
    main()
