#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Graph-only mail processor with delta sync + 24h backfill fallback.
- Uses MSAL client credentials (application permissions) for auth.
- Tracks @odata.deltaLink per folder in delta_tokens.json.
- On token 410/invalid: clears token and reseeds from now-1d.
- Processes messages: skips already read or already PAIRActioned,
  otherwise (a) calls user-defined `classify_email` + `handle_new_email` if available,
  else (b) marks read + adds PAIRActioned category.
"""

import os
import json
import time
from datetime import datetime, timedelta, timezone
from typing import Dict, Tuple, List, Optional
import requests
import msal
import re

# =====================
# Configuration
# =====================

# Azure AD app (Application permissions scenario recommended for headless server)
MS_CLIENT_ID     = os.getenv("MS_CLIENT_ID") or os.getenv("CLIENT_ID") or os.getenv("MSAL_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET") or os.getenv("CLIENT_SECRET") or os.getenv("MSAL_CLIENT_SECRET")
MS_TENANT_ID     = os.getenv("MS_TENANT_ID") or os.getenv("TENANT_ID")
MS_SCOPE         = ["https://graph.microsoft.com/.default"]

# The mailbox user principal to operate on (e.g., info@mlfa.org)
EMAIL_TO_WATCH = os.getenv("EMAIL_TO_WATCH")

# Polling & backfill
POLL_INTERVAL_SECS = int(os.getenv("POLL_INTERVAL_SECS", "10"))
BACKFILL_DAYS      = int(os.getenv("BACKFILL_DAYS", "1"))

# Human-approval mode (if you already have a Flask approval hub, set True)
HUMAN_CHECK = os.getenv("HUMAN_CHECK", "false").lower() == "true"

# Delta token store
DELTA_FILE = os.getenv("DELTA_FILE", "delta_tokens.json")

# Categories
PROCESSED_CATEGORY = "PAIRActioned"

# =====================
# Utilities
# =====================

def utc_now() -> datetime:
    return datetime.now(timezone.utc)

def to_graph_z(dt: datetime) -> str:
    dt = dt.astimezone(timezone.utc).replace(microsecond=0)
    s = dt.isoformat()
    return s.replace("+00:00", "Z")

def load_delta_tokens() -> Dict[str, str]:
    try:
        with open(DELTA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
    except FileNotFoundError:
        return {}
    except Exception as e:
        print(f"⚠️ Could not read {DELTA_FILE}: {e}")
    return {}

def save_delta_tokens(tokens: Dict[str, str]) -> None:
    tmp = DELTA_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(tokens, f, indent=2, ensure_ascii=False)
        os.replace(tmp, DELTA_FILE)
    except Exception as e:
        print(f"⚠️ Could not write {DELTA_FILE}: {e}")

# =====================
# Auth via MSAL
# =====================

_msal_app: Optional[msal.ConfidentialClientApplication] = None

def msal_app() -> msal.ConfidentialClientApplication:
    global _msal_app
    if _msal_app is None:
        if not (MS_CLIENT_ID and MS_CLIENT_SECRET and MS_TENANT_ID):
            raise RuntimeError("Missing MSAL environment vars: MS_CLIENT_ID / MS_CLIENT_SECRET / MS_TENANT_ID")
        _msal_app = msal.ConfidentialClientApplication(
            MS_CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{MS_TENANT_ID}",
            client_credential=MS_CLIENT_SECRET,
        )
    return _msal_app

def get_bearer_token() -> str:
    app = msal_app()
    result = app.acquire_token_silent(MS_SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=MS_SCOPE)
    if "access_token" in result:
        return result["access_token"]
    raise RuntimeError(f"Auth failed: {result.get('error_description') or result}")

def auth_headers() -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {get_bearer_token()}",
        "Content-Type": "application/json"
    }

# =====================
# Graph helpers
# =====================

def get_folder_ids() -> Dict[str, str]:
    """Resolve Inbox and Junk folder IDs via Graph; fallback to well-known names."""
    try:
        res = requests.get(
            "https://graph.microsoft.com/v1.0/me/mailFolders?$select=id,displayName",
            headers=auth_headers(), timeout=15
        )
        res.raise_for_status()
        inbox_id = None
        junk_id  = None
        for f in res.json().get("value", []):
            name = (f.get("displayName") or "").lower()
            if name == "inbox":
                inbox_id = f["id"]
            if name in ("junk email", "junkemail"):
                junk_id = f["id"]
        print(f"📁 Folder IDs - Inbox: {inbox_id or 'inbox'}, Junk: {junk_id or 'junkemail'}")
        return {"inbox": inbox_id or "inbox", "junk": junk_id or "junkemail"}
    except Exception as e:
        print(f"❌ Error getting folder IDs: {e}")
        return {"inbox": "inbox", "junk": "junkemail"}

def graph_delta_sync(folder_name: str, folder_id: str, *, delta_url: Optional[str]=None, since_utc_iso: Optional[str]=None) -> Tuple[List[Dict], Optional[str]]:
    """
    Perform delta sync via Graph. Returns (changes, final_delta_url).
    Each change is {'removed': bool, 'id': str} or {'removed': False, 'message': {...}}.
    """
    if not EMAIL_TO_WATCH:
        raise RuntimeError("EMAIL_TO_WATCH is not set")

    headers = auth_headers()

    if delta_url:
        url = delta_url
        print(f"🔄 Resuming delta sync for {folder_name}")
    else:
        base = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages/delta"
        if since_utc_iso:
            # Use params to ensure proper encoding
            url = f"{base}?$filter=receivedDateTime ge {since_utc_iso}"
        else:
            url = base
        print(f"🌱 Starting fresh delta sync for {folder_name} since {since_utc_iso}")

    changes: List[Dict] = []
    final_delta_url: Optional[str] = None

    while url:
        try:
            print(f"📡 GET {url[:120]}{'…' if len(url)>120 else ''}")
            r = requests.get(url, headers=headers, timeout=30)
            if r.status_code == 410:
                print(f"⚠️ 410 Gone for {folder_name} delta; need to reseed")
                return [], None
            r.raise_for_status()
            data = r.json()

            for item in data.get("value", []):
                if "@removed" in item:
                    changes.append({"removed": True, "id": item.get("id")})
                else:
                    changes.append({"removed": False, "message": item})

            next_link  = data.get("@odata.nextLink")
            delta_link = data.get("@odata.deltaLink")
            if next_link:
                url = next_link
            else:
                final_delta_url = delta_link
                url = None

        except requests.exceptions.HTTPError as e:
            print(f"❌ HTTP error in delta sync for {folder_name}: {e}")
            return [], None
        except Exception as e:
            print(f"❌ Error in delta sync for {folder_name}: {e}")
            return [], None

    print(f"✅ Delta {folder_name}: {len(changes)} changes; new token: {'Yes' if final_delta_url else 'No'}")
    return changes, final_delta_url

# =====================
# Processing
# =====================

def html_to_text(html: str) -> str:
    # very lightweight tag-stripper
    return re.sub(r"<[^>]+>", " ", html or "").replace("&nbsp;", " ").strip()

def classify_and_route(subject: str, body_text: str) -> Optional[dict]:
    """
    If user-defined classify/handle functions exist in the runtime (e.g. imported),
    use them; otherwise return None and we'll just mark processed.
    """
    try:
        # resolve from globals if present
        classify_email = globals().get("classify_email")
        handle_new_email = globals().get("handle_new_email")
        if callable(classify_email) and callable(handle_new_email):
            result = classify_email(subject, body_text)
            return {"result": result, "handler": handle_new_email}
    except Exception as e:
        print(f"⚠️ Classification pipeline failed: {e}")
    return None

def graph_patch_message(msg_id: str, payload: dict) -> None:
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}",
        headers=auth_headers(), json=payload, timeout=20
    )
    r.raise_for_status()

def graph_move_message(msg_id: str, dest_folder_id: str) -> dict:
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/move",
        headers=auth_headers(), json={"destinationId": dest_folder_id}, timeout=20
    )
    r.raise_for_status()
    return r.json()

def process_graph_message(message_data: dict, folder_name: str) -> None:
    try:
        msg_id = message_data.get("id")
        if not msg_id:
            return

        subject = message_data.get("subject") or ""
        categories = set(message_data.get("categories") or [])
        is_read = bool(message_data.get("isRead"))
        body_html = (message_data.get("body") or {}).get("content") or ""
        body_text = html_to_text(body_html)

        # Idempotency
        if any((c or "").startswith(PROCESSED_CATEGORY) for c in categories):
            return
        if is_read:
            return

        # Try user's pipeline
        pipeline = classify_and_route(subject, body_text)

        if HUMAN_CHECK:
            # If you have an approval hub, stash message here by implementing stash_for_approval
            stash = globals().get("stash_for_approval")
            if callable(stash):
                stash(msg_id, subject, body_text, pipeline["result"] if pipeline else None, message_data)
                print(f"📥 Stashed for approval: {subject}")
                return

        if pipeline:
            try:
                # Let user's handler act; we assume it can work off Graph msg id via an adapter.
                handler = pipeline["handler"]
                # If they expect an object, they should adapt to Graph-id externally.
                handler(message_data, pipeline["result"])
            except Exception as e:
                print(f"⚠️ User handler failed: {e} — falling back to mark as processed")

        # Default action: mark as read + add processed category
        new_cats = sorted(categories.union({PROCESSED_CATEGORY}))
        graph_patch_message(msg_id, {"isRead": True, "categories": new_cats})
        print(f"✅ Marked processed: {subject}")

    except Exception as e:
        print(f"❌ Error processing message: {e}")

# =====================
# Folder driver
# =====================

def process_folders_with_graph_api(delta_tokens: Dict[str, str]) -> Dict[str, str]:
    """Process inbox + junk using Graph delta; returns possibly updated delta_tokens."""
    folder_ids = get_folder_ids()
    start_time_iso = to_graph_z(utc_now() - timedelta(days=BACKFILL_DAYS))

    for folder_name in ("inbox", "junk"):
        print(f"\n🔄 Processing {folder_name.upper()}…")
        folder_id = folder_ids[folder_name]
        existing_delta_url = delta_tokens.get(folder_name)

        # Run delta
        if existing_delta_url:
            changes, new_delta_url = graph_delta_sync(folder_name, folder_id, delta_url=existing_delta_url)
        else:
            changes, new_delta_url = graph_delta_sync(folder_name, folder_id, since_utc_iso=start_time_iso)

        # Handle changes
        for change in changes:
            if change.get("removed"):
                print(f"🗑️ Removed: {change.get('id')}")
            else:
                process_graph_message(change.get("message") or {}, folder_name)

        # Finalize token once
        if new_delta_url:
            delta_tokens[folder_name] = new_delta_url
            print(f"💾 Saved delta token for {folder_name}")
        else:
            if existing_delta_url:
                # Token likely expired or not returned; clear to reseed next pass
                delta_tokens.pop(folder_name, None)
                print(f"🧹 Cleared stale delta token for {folder_name}")

    return delta_tokens

# =====================
# Main loop
# =====================

def main():
    if not EMAIL_TO_WATCH:
        raise RuntimeError("Set EMAIL_TO_WATCH to the mailbox UPN (e.g., info@mlfa.org)")

    print(f"📧 Monitoring: {EMAIL_TO_WATCH}")
    print(f"🕒 Poll interval: {POLL_INTERVAL_SECS}s; Backfill: last {BACKFILL_DAYS} day(s)")

    delta_tokens = load_delta_tokens()
    print(f"📚 Loaded tokens for: {list(delta_tokens.keys())}")

    consecutive_errors = 0
    last_ok = time.time()

    while True:
        try:
            delta_tokens = process_folders_with_graph_api(delta_tokens)
            save_delta_tokens(delta_tokens)
            consecutive_errors = 0
            last_ok = time.time()
        except Exception as e:
            consecutive_errors += 1
            print(f"❌ Loop error (#{consecutive_errors}): {e}")
            if consecutive_errors >= 3:
                # brief backoff
                time.sleep(30)
                consecutive_errors = 0
        time.sleep(POLL_INTERVAL_SECS)

if __name__ == "__main__":
    main()
