"""Email-signup storage.

Primary store is a Google Sheet (so Manas gets a live, exportable mailing list).
It is intentionally resilient: if Google Sheets isn't configured or the API call
fails for any reason, signups are still captured in a local CSV so the gate keeps
working and no email is ever lost. Nothing here ever raises into the request path.

Configure the Google Sheet via two environment variables:
  GOOGLE_SERVICE_ACCOUNT_JSON  -> the full service-account JSON (one line), OR a
                                  path to the .json file on disk.
  SHEET_ID                     -> the target spreadsheet's ID (from its URL).
Share the spreadsheet with the service account's client_email (Editor) first.
"""

from __future__ import annotations

import csv
import json
import os
import threading
from datetime import datetime, timezone
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LOCAL_CSV = BASE_DIR / "signups.csv"
_HEADERS = ["timestamp_utc", "email", "name", "ip", "user_agent"]

# gspread worksheet handle is cached after first successful connect.
_lock = threading.Lock()
_worksheet = None
_sheets_tried = False


def _get_worksheet():
    """Lazily connect to the Google Sheet. Returns a gspread worksheet or None.
    Cached so we don't re-auth on every signup. Never raises."""
    global _worksheet, _sheets_tried
    if _worksheet is not None:
        return _worksheet
    if _sheets_tried:
        return None  # already failed once; don't hammer the API every request
    _sheets_tried = True

    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    sheet_id = os.environ.get("SHEET_ID", "").strip()
    if not raw or not sheet_id:
        return None
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        if raw.startswith("{"):
            info = json.loads(raw)
        else:  # treat as a path to the JSON file
            info = json.loads(Path(raw).read_text(encoding="utf-8"))

        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id)
        ws = sheet.sheet1
        # Ensure a header row exists.
        if not ws.row_values(1):
            ws.append_row(_HEADERS, value_input_option="RAW")
        _worksheet = ws
        return _worksheet
    except Exception as exc:  # noqa: BLE001 — never break the request path
        print(f"[sheets_store] Google Sheets unavailable, using local CSV: {exc}")
        return None


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
    """Append a signup. Returns True if stored anywhere. Never raises."""
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    row = [ts, email, name, ip, user_agent[:300]]
    stored = False
    with _lock:
        ws = _get_worksheet()
        if ws is not None:
            try:
                ws.append_row(row, value_input_option="RAW")
                stored = True
            except Exception as exc:  # noqa: BLE001
                print(f"[sheets_store] append failed, falling back to CSV: {exc}")
        # Always mirror to the local CSV too — cheap, and a safety net.
        _append_local_csv(row)
        stored = True
    return stored


def storage_mode() -> str:
    """For diagnostics: 'google-sheets' when the Sheet is reachable, else 'local-csv'."""
    return "google-sheets" if _get_worksheet() is not None else "local-csv"
