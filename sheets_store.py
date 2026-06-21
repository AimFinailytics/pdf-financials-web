"""Email-signup storage.

Signups are sent to a Google Form (published, "Anyone with the link", email
collection OFF), which writes every submission to its linked Google Sheet — that
Sheet is the mailing list. This needs NO credentials on the server: it's a plain
HTTPS POST to the form's public formResponse endpoint.

It is intentionally resilient: signups are also mirrored to a local CSV, so if the
form POST ever fails the email is still captured. Nothing here raises into the
request path.

Configure via environment variables (set on Render):
  FORM_POST_URL     -> the form's .../formResponse URL
  FORM_EMAIL_ENTRY  -> the email field id, e.g. "entry.906489358"
  FORM_NAME_ENTRY   -> the name field id,  e.g. "entry.659201863"  (optional)
"""

from __future__ import annotations

import csv
import os
import threading
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LOCAL_CSV = BASE_DIR / "signups.csv"
_HEADERS = ["timestamp_utc", "email", "name", "ip", "user_agent"]
_lock = threading.Lock()


def _post_to_form(email: str, name: str) -> bool:
    """POST one signup to the Google Form. Returns True on success. Never raises."""
    url = os.environ.get("FORM_POST_URL", "").strip()
    email_entry = os.environ.get("FORM_EMAIL_ENTRY", "").strip()
    if not url or not email_entry:
        return False
    fields = {email_entry: email}
    name_entry = os.environ.get("FORM_NAME_ENTRY", "").strip()
    if name_entry and name:
        fields[name_entry] = name
    try:
        data = urllib.parse.urlencode(fields).encode()
        req = urllib.request.Request(
            url, data=data,
            headers={"User-Agent": "Mozilla/5.0", "Content-Type": "application/x-www-form-urlencoded"},
        )
        with urllib.request.urlopen(req, timeout=15) as r:
            return r.status == 200
    except Exception as exc:  # noqa: BLE001 — never break the request path
        print(f"[sheets_store] form POST failed, using local CSV only: {exc}")
        return False


def _append_local_csv(row: list[str]) -> None:
    new = not LOCAL_CSV.exists()
    try:
        with LOCAL_CSV.open("a", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            if new:
                writer.writerow(_HEADERS)
            writer.writerow(row)
    except OSError as exc:
        print(f"[sheets_store] could not write local CSV: {exc}")


def record_signup(email: str, name: str = "", ip: str = "", user_agent: str = "") -> bool:
    """Append a signup to the Google Form (-> Sheet) and mirror to local CSV.
    Returns True if stored anywhere. Never raises."""
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    row = [ts, email, name, ip, user_agent[:300]]
    with _lock:
        _post_to_form(email, name)        # the durable mailing list
        _append_local_csv(row)            # local safety-net mirror
    return True


def storage_mode() -> str:
    """For diagnostics: 'google-form' when the form POST is configured, else 'local-csv'."""
    return "google-form" if os.environ.get("FORM_POST_URL", "").strip() else "local-csv"
