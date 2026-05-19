"""
GE Finance Slack → Google Sheets daily sync.

Flow:
  1. Read Slack channel where GE Finance posts at 06:00 BRT
  2. Find the GE Finance message (file attachment or text)
  3. If file: download Excel/CSV and upload to Google Sheets "Dados Completos"
  4. If text only: log raw content for format inspection

Environment variables (see .env.example):
  SLACK_TOKEN                  xoxp-... or xoxb-... token
  SLACK_CHANNEL_ID             C0B4LAZPNS2
  GOOGLE_SERVICE_ACCOUNT_JSON  full JSON string of service account key
"""

import io
import json
import logging
import os
import time
from datetime import datetime, timedelta, timezone

import gspread
import openpyxl
import requests
from google.oauth2.service_account import Credentials

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

SPREADSHEET_ID = "1rlB0XEnwkbnyU8nNViz4u6uDwpqa0Yfc2oNtMTlgg0I"
TARGET_TAB = "Dados Completos"
SLACK_API = "https://slack.com/api"


# ---------------------------------------------------------------------------
# Slack helpers
# ---------------------------------------------------------------------------

def _slack_get(endpoint: str, token: str, params: dict) -> dict:
    resp = requests.get(
        f"{SLACK_API}/{endpoint}",
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(f"Slack API error ({endpoint}): {data.get('error')}")
    return data


def fetch_recent_messages(token: str, channel_id: str, hours_back: int = 20) -> list:
    """Return messages posted in the last `hours_back` hours."""
    oldest = (datetime.now(timezone.utc) - timedelta(hours=hours_back)).timestamp()
    data = _slack_get(
        "conversations.history",
        token,
        {"channel": channel_id, "oldest": str(oldest), "limit": 50},
    )
    messages = data.get("messages", [])
    log.info("Fetched %d messages from channel (last %dh).", len(messages), hours_back)
    return messages


def find_ge_finance_message(messages: list) -> dict | None:
    """
    Try to identify the GE Finance automated message.
    Heuristics (adjust after seeing the real message tomorrow):
      - Has a file attachment, OR
      - Bot/app post (subtype == 'bot_message'), OR
      - Text contains keywords like 'vendas', 'pedidos', 'relatório'
    """
    keywords = ["venda", "pedido", "relatório", "relatorio", "ge finance", "exportação"]
    for msg in messages:
        # Bot messages are the most reliable signal
        if msg.get("subtype") == "bot_message" or msg.get("bot_id"):
            log.info("Found bot message: %.120s", msg.get("text", "")[:120])
            return msg
        # Messages with file attachments
        if msg.get("files"):
            log.info("Found message with file attachment.")
            return msg
        # Keyword match in text
        text = (msg.get("text") or "").lower()
        if any(kw in text for kw in keywords):
            log.info("Found keyword-matched message: %.120s", text[:120])
            return msg
    return None


def download_slack_file(token: str, file_info: dict) -> tuple[bytes, str]:
    """Download a file shared in Slack. Returns (content_bytes, filename)."""
    url = file_info.get("url_private_download") or file_info.get("url_private")
    name = file_info.get("name", "download")
    if not url:
        raise ValueError(f"No download URL in file info: {file_info}")
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=120,
    )
    resp.raise_for_status()
    log.info("Downloaded Slack file '%s' (%d bytes).", name, len(resp.content))
    return resp.content, name


# ---------------------------------------------------------------------------
# Parse file content
# ---------------------------------------------------------------------------

def parse_excel(content: bytes) -> list[list]:
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.values]
    wb.close()
    return rows


def parse_csv(content: bytes) -> list[list]:
    import csv
    text = content.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    return list(reader)


def parse_file(content: bytes, filename: str) -> list[list]:
    name_lower = filename.lower()
    if name_lower.endswith((".xlsx", ".xls")):
        return parse_excel(content)
    if name_lower.endswith(".csv"):
        return parse_csv(content)
    # Try Excel first, fall back to CSV
    try:
        return parse_excel(content)
    except Exception:
        return parse_csv(content)


# ---------------------------------------------------------------------------
# Google Sheets upload
# ---------------------------------------------------------------------------

def upload_to_google_sheets(rows: list[list]) -> None:
    def _clean(v):
        if v is None:
            return ""
        if isinstance(v, (int, float)):
            return v
        return str(v)

    clean_rows = [[_clean(c) for c in row] for row in rows]
    log.info("Uploading %d rows × %d cols to Google Sheets…", len(clean_rows), len(clean_rows[0]) if clean_rows else 0)

    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    sa_info = json.loads(sa_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    gc = gspread.authorize(creds)

    sheet = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sheet.worksheet(TARGET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title=TARGET_TAB, rows=1, cols=1)
        log.info("Created new tab '%s'.", TARGET_TAB)

    ws.clear()
    ws.update(clean_rows, value_input_option="USER_ENTERED")
    log.info("Tab '%s' updated successfully.", TARGET_TAB)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    token = os.environ["SLACK_TOKEN"]
    channel_id = os.environ.get("SLACK_CHANNEL_ID", "C0B4LAZPNS2")

    messages = fetch_recent_messages(token, channel_id, hours_back=20)

    if not messages:
        log.warning("No messages found in the last 20 hours. Exiting.")
        return

    msg = find_ge_finance_message(messages)

    if not msg:
        log.warning(
            "Could not identify GE Finance message. Dumping all messages for inspection:\n%s",
            json.dumps(messages, indent=2, ensure_ascii=False),
        )
        return

    # --- Case 1: file attachment ---
    files = msg.get("files", [])
    if files:
        content, filename = download_slack_file(token, files[0])
        rows = parse_file(content, filename)
        if not rows:
            raise ValueError("Parsed file is empty.")
        upload_to_google_sheets(rows)
        log.info("Done! Uploaded %d rows from file '%s'.", len(rows), filename)
        return

    # --- Case 2: text-only message (format unknown until tomorrow) ---
    log.info(
        "Message has no file attachment. Raw text for inspection:\n%s",
        json.dumps(msg, indent=2, ensure_ascii=False),
    )
    log.warning(
        "Text-only message detected. Share the log above so we can build a parser."
    )


if __name__ == "__main__":
    main()
