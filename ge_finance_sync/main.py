"""
GE Finance → Google Sheets daily sync.

Flow:
  1. Login to GE Finance → get fresh JWT
  2. Trigger spreadsheet export (D-91 to D-1, status filters)
  3. Poll until S3 download URL is ready
  4. Download Excel file
  5. Replace data in Google Sheets tab "Dados Completos"

Required environment variables (see .env.example):
  GE_EMAIL, GE_PASSWORD
  GOOGLE_SERVICE_ACCOUNT_JSON  (full JSON string of service account key)
"""

import io
import json
import logging
import os
import time
from datetime import date, timedelta

import gspread
import openpyxl
import requests
from google.oauth2.service_account import Credentials

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

BASE_URL = "https://gateway-web.ge.finance"
SPREADSHEET_ID = "1rlB0XEnwkbnyU8nNViz4u6uDwpqa0Yfc2oNtMTlgg0I"
TARGET_TAB = "Dados Completos"
BUCKET_NAME = "prod-gecom-spreadsheet-excelexport-production"
STATUS_FILTERS = ["Enviado", "Aprovado", "Entregue", "Pronto para envio"]
POLL_INTERVAL_SEC = 5
POLL_MAX_ATTEMPTS = 60  # 5 min max


# ---------------------------------------------------------------------------
# Step 1 – Authentication
# ---------------------------------------------------------------------------

def login(email: str, password: str) -> dict:
    """POST /api/Auth/Login — returns token + account metadata."""
    url = f"{BASE_URL}/api/Auth/Login"
    payload = {"email": email, "password": password}
    resp = requests.post(url, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    # The API may nest the result under different keys depending on version.
    # Try the most common shapes:
    token = (
        data.get("token")
        or data.get("accessToken")
        or data.get("data", {}).get("token")
        or data.get("data", {}).get("accessToken")
    )
    if not token:
        raise ValueError(f"Login succeeded but no token found in response: {data}")

    customer_id = (
        data.get("customerId")
        or data.get("data", {}).get("customerId")
        or data.get("customer", {}).get("id")
    )
    plan_id = (
        data.get("customerPlanId")
        or data.get("data", {}).get("customerPlanId")
        or data.get("planId")
    )

    log.info("Login OK. customerId=%s planId=%s", customer_id, plan_id)
    return {"token": token, "customerId": customer_id, "customerPlanId": plan_id}


def _auth_headers(auth: dict) -> dict:
    return {
        "authorization": f"Bearer {auth['token']}",
        "customerid": str(auth["customerId"]),
        "customerplanid": str(auth["customerPlanId"]),
        "language": "pt-BR",
        "accept": "application/json, text/plain, */*",
    }


# ---------------------------------------------------------------------------
# Step 2 – Trigger export
# ---------------------------------------------------------------------------

def trigger_export(auth: dict, first_date: date, end_date: date) -> str:
    """
    GET /api/SpreadSheet/GetSpreadSheetSystemNew
    Returns requestId string.
    """
    url = f"{BASE_URL}/api/SpreadSheet/GetSpreadSheetSystemNew"

    # refresh_date = right now (used by the UI as a cache-busting param)
    refresh_ts = f"{date.today()}T00:00:00.000"

    params = [
        ("pageSize", "500"),
        ("firstDate", first_date.strftime("%Y-%m-%d")),
        ("endDate", end_date.strftime("%Y-%m-%d")),
        ("refreshDate", refresh_ts),
        ("sortColumn", "date"),
        ("sortType", "2"),
        ("currentPage", "1"),
        ("exportType", "product"),
        ("customerId", auth["customerId"]),
        ("customerPlanId", auth["customerPlanId"]),
        ("isTrial", "false"),
        ("SortFilters", ""),
        ("SalesChannel", ""),
        *[("StatusOrder", s) for s in STATUS_FILTERS],
        ("Brands", ""),
        ("sku", ""),
        ("productName", ""),
        ("search", ""),
    ]

    resp = requests.get(url, params=params, headers=_auth_headers(auth), timeout=60)
    resp.raise_for_status()
    data = resp.json()

    request_id = (
        data.get("requestId")
        or data.get("data", {}).get("requestId")
        or data.get("id")
        # sometimes the response IS the string itself
        or (data if isinstance(data, str) else None)
    )
    if not request_id:
        raise ValueError(f"No requestId in export trigger response: {data}")

    log.info("Export triggered. requestId=%s", request_id)
    return request_id


# ---------------------------------------------------------------------------
# Step 3 – Poll for S3 URL
# ---------------------------------------------------------------------------

def poll_for_download_url(auth: dict, request_id: str) -> str:
    """
    GET /api/SpreadSheet/GetSpreadSheetDownloadS3New
    Polls until a download URL is returned.
    """
    url = f"{BASE_URL}/api/SpreadSheet/GetSpreadSheetDownloadS3New"
    params = {"bucketName": BUCKET_NAME, "requestId": request_id}

    for attempt in range(1, POLL_MAX_ATTEMPTS + 1):
        resp = requests.get(url, params=params, headers=_auth_headers(auth), timeout=30)
        resp.raise_for_status()
        data = resp.json()

        download_url = (
            data.get("url")
            or data.get("downloadUrl")
            or data.get("data", {}).get("url")
            or data.get("data", {}).get("downloadUrl")
            or (data if isinstance(data, str) and data.startswith("http") else None)
        )

        if download_url:
            log.info("File ready after %d poll(s). URL: %s…", attempt, download_url[:60])
            return download_url

        log.info("Attempt %d/%d — not ready yet, waiting %ds…", attempt, POLL_MAX_ATTEMPTS, POLL_INTERVAL_SEC)
        time.sleep(POLL_INTERVAL_SEC)

    raise TimeoutError(f"Export not ready after {POLL_MAX_ATTEMPTS} attempts ({POLL_MAX_ATTEMPTS * POLL_INTERVAL_SEC}s)")


# ---------------------------------------------------------------------------
# Step 4 – Download Excel
# ---------------------------------------------------------------------------

def download_excel(download_url: str) -> bytes:
    resp = requests.get(download_url, timeout=120)
    resp.raise_for_status()
    log.info("Downloaded %d bytes.", len(resp.content))
    return resp.content


# ---------------------------------------------------------------------------
# Step 5 – Upload to Google Sheets
# ---------------------------------------------------------------------------

def upload_to_google_sheets(excel_bytes: bytes) -> None:
    """
    Reads the first sheet of the downloaded Excel and replaces all data
    in the "Dados Completos" tab of the target Google Sheet.
    """
    # --- Parse Excel ---
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.values]
    wb.close()

    if not rows:
        raise ValueError("Downloaded Excel file is empty.")

    # Convert all cells to plain Python types (openpyxl may return datetime, etc.)
    def _clean(v):
        if v is None:
            return ""
        if isinstance(v, (int, float)):
            return v
        return str(v)

    clean_rows = [[_clean(c) for c in row] for row in rows]
    log.info("Excel parsed: %d rows × %d cols", len(clean_rows), len(clean_rows[0]))

    # --- Authenticate with Google ---
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
        worksheet = sheet.worksheet(TARGET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title=TARGET_TAB, rows=1, cols=1)
        log.info("Created new tab '%s'.", TARGET_TAB)

    # --- Clear and write ---
    worksheet.clear()
    worksheet.update(clean_rows, value_input_option="USER_ENTERED")
    log.info("Google Sheets tab '%s' updated with %d rows.", TARGET_TAB, len(clean_rows))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    email = os.environ["GE_EMAIL"]
    password = os.environ["GE_PASSWORD"]

    today = date.today()
    end_date = today - timedelta(days=1)      # D-1  (yesterday)
    first_date = today - timedelta(days=91)   # D-91

    log.info("Date range: %s → %s", first_date, end_date)

    auth = login(email, password)
    request_id = trigger_export(auth, first_date, end_date)
    download_url = poll_for_download_url(auth, request_id)
    excel_bytes = download_excel(download_url)
    upload_to_google_sheets(excel_bytes)

    log.info("Done!")


if __name__ == "__main__":
    main()
