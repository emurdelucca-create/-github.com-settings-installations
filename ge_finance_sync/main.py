"""
GE Finance → Google Sheets daily sync.
Flow:
  1. Login na API do GE Finance com email/senha
  2. Solicita exportação dos últimos 90 dias
  3. Aguarda o arquivo ficar pronto (polling S3)
  4. Baixa o Excel
  5. Faz upload na aba "Dados Completos" do Google Sheets
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

GE_BASE        = "https://gateway-web.ge.finance"
GE_APP         = "https://app.ge.finance"
SPREADSHEET_ID = "1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo"
TARGET_TAB     = "Dados Completos"
EXPORT_DAYS    = 90
BUCKET_NAME    = "prod-gecom-spreadsheet-excelexport-production"
STATUS_ORDERS  = ["Entregue", "Enviado", "Aprovado", "Pronto para envio"]


# ---------------------------------------------------------------------------
# 1. Login
# ---------------------------------------------------------------------------
def ge_login(email: str, password: str) -> dict:
    resp = requests.post(
        f"{GE_BASE}/api/Auth/login",
        headers={
            "accept": "application/json, text/plain, */*",
            "content-type": "application/json",
            "language": "pt-BR",
            "origin": GE_APP,
            "referer": f"{GE_APP}/",
        },
        json={"username": email, "password": password},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    log.info("Login OK para %s", email)
    log.debug("Login response: %s", json.dumps(data, ensure_ascii=False)[:500])
    return data


def _extrair_auth(data: dict) -> tuple[str, str, str]:
    """Extrai token, customerId e customerPlanId da resposta de login."""
    log.info("Login response keys: %s", list(data.keys()))
    inner = data.get("data") or data
    if isinstance(inner, dict):
        log.info("Inner keys: %s", list(inner.keys()))

    token = (
        inner.get("token")
        or inner.get("accessToken")
        or inner.get("access_token")
        or data.get("token")
    )
    customer_id = (
        inner.get("customerId")
        or inner.get("customer_id")
        or data.get("customerId")
    )

    # customerPlanIds pode vir como lista, string JSON ou inteiro
    raw_plans = (
        inner.get("customerPlanIds")
        or inner.get("customerPlanId")
        or data.get("customerPlanIds")
        or data.get("customerPlanId")
    )
    log.info("customerPlanIds raw: %s (type=%s)", raw_plans, type(raw_plans).__name__)

    if isinstance(raw_plans, list) and raw_plans:
        plan_id = str(raw_plans[0])
    elif isinstance(raw_plans, str):
        # Pode ser string JSON "[1099,1274,...]" ou só "1099"
        try:
            parsed = json.loads(raw_plans)
            plan_id = str(parsed[0]) if isinstance(parsed, list) else str(parsed)
        except Exception:
            plan_id = raw_plans.strip()
    elif raw_plans is not None:
        plan_id = str(raw_plans)
    else:
        # Fallback: extrair do token JWT (campo customerPlanIds)
        try:
            import base64
            payload = token.split(".")[1]
            payload += "=" * (4 - len(payload) % 4)
            jwt_data = json.loads(base64.b64decode(payload))
            raw_jwt = jwt_data.get("customerPlanIds", "")
            parsed_jwt = json.loads(raw_jwt) if isinstance(raw_jwt, str) else raw_jwt
            plan_id = str(parsed_jwt[0]) if isinstance(parsed_jwt, list) else str(parsed_jwt)
            log.info("customerPlanId extraído do JWT: %s", plan_id)
        except Exception as e:
            log.warning("Falha ao extrair planId do JWT: %s", e)
            plan_id = "1099"  # fallback conhecido

    if not token:
        raise ValueError(f"Não foi possível extrair o token da resposta: {data}")
    log.info("customerId=%s | customerPlanId=%s", customer_id, plan_id)
    return str(token), str(customer_id), str(plan_id)


# ---------------------------------------------------------------------------
# 2. Solicitar exportação
# ---------------------------------------------------------------------------
def ge_solicitar_exportacao(token: str, customer_id: str, plan_id: str) -> str:
    hoje      = datetime.now(timezone.utc).date()
    end_date  = hoje.strftime("%Y-%m-%d")
    start_date = (hoje - timedelta(days=EXPORT_DAYS)).strftime("%Y-%m-%d")
    refresh   = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000")

    params = [
        ("pageSize",       500),
        ("firstDate",      start_date),
        ("endDate",        end_date),
        ("refreshDate",    refresh),
        ("sortColumn",     "date"),
        ("sortType",       2),
        ("currentPage",    1),
        ("exportType",     "product"),
        ("customerId",     customer_id),
        ("customerPlanId", plan_id),
        ("isTrial",        "false"),
        ("SortFilters",    ""),
        ("SalesChannel",   ""),
        ("Brands",         ""),
        ("sku",            ""),
        ("productName",    ""),
        ("search",         ""),
    ]
    for s in STATUS_ORDERS:
        params.append(("StatusOrder", s))

    resp = requests.get(
        f"{GE_BASE}/api/SpreadSheet/GetSpreadSheetSystemNew",
        headers=_headers(token, customer_id, plan_id),
        params=params,
        timeout=60,
    )
    resp.raise_for_status()
    data = resp.json()
    log.debug("Resposta exportação: %s", json.dumps(data, ensure_ascii=False)[:500])

    request_id = (
        data.get("requestId")
        or data.get("id")
        or (data.get("data") if isinstance(data.get("data"), str) else None)
        or (data.get("data", {}).get("requestId") if isinstance(data.get("data"), dict) else None)
    )
    if not request_id:
        raise ValueError(f"requestId não encontrado na resposta: {data}")

    log.info("Exportação solicitada. requestId=%s | período=%s → %s", request_id, start_date, end_date)
    return request_id


# ---------------------------------------------------------------------------
# 3. Aguardar + baixar
# ---------------------------------------------------------------------------
def ge_aguardar_e_baixar(token: str, customer_id: str, plan_id: str,
                          request_id: str, max_espera: int = 300, intervalo: int = 5) -> bytes:
    params = {"bucketName": BUCKET_NAME, "requestId": request_id}
    aguardado = 0

    while aguardado < max_espera:
        resp = requests.get(
            f"{GE_BASE}/api/SpreadSheet/GetSpreadSheetDownloadS3New",
            headers=_headers(token, customer_id, plan_id),
            params=params,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        log.debug("Poll resposta: %s", str(data)[:200])

        url = (
            data.get("url")
            or data.get("downloadUrl")
            or data.get("fileUrl")
            or (data.get("data") if isinstance(data.get("data"), str) and data.get("data", "").startswith("http") else None)
            or (data.get("data", {}).get("url") if isinstance(data.get("data"), dict) else None)
        )
        if url:
            log.info("Arquivo pronto! Baixando de: %s", url[:80])
            file_resp = requests.get(url, timeout=120)
            file_resp.raise_for_status()
            log.info("Baixados %d bytes.", len(file_resp.content))
            return file_resp.content

        log.info("Aguardando arquivo... (%ds / %ds)", aguardado, max_espera)
        time.sleep(intervalo)
        aguardado += intervalo

    raise TimeoutError(f"Arquivo não ficou pronto em {max_espera}s.")


# ---------------------------------------------------------------------------
# 4. Parsear Excel
# ---------------------------------------------------------------------------
def parsear_excel(content: bytes) -> list[list]:
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.values]
    wb.close()
    log.info("Excel parseado: %d linhas × %d colunas.", len(rows), len(rows[0]) if rows else 0)
    return rows


# ---------------------------------------------------------------------------
# 5. Upload Google Sheets
# ---------------------------------------------------------------------------
def upload_google_sheets(rows: list[list]) -> None:
    def _limpar(v):
        if v is None:
            return ""
        if isinstance(v, (int, float)):
            return v
        return str(v)

    linhas = [[_limpar(c) for c in row] for row in rows]
    log.info("Enviando %d linhas para Google Sheets…", len(linhas))

    sa_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds   = Credentials.from_service_account_info(sa_info, scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    gc    = gspread.authorize(creds)
    sheet = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sheet.worksheet(TARGET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title=TARGET_TAB, rows=1, cols=1)
        log.info("Aba '%s' criada.", TARGET_TAB)

    ws.clear()
    ws.update(linhas, value_input_option="USER_ENTERED")
    log.info("Aba '%s' atualizada com %d linhas.", TARGET_TAB, len(linhas))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _headers(token: str, customer_id: str, plan_id: str) -> dict:
    return {
        "accept":         "application/json, text/plain, */*",
        "accept-language":"pt-BR,pt;q=0.9",
        "authorization":  f"Bearer {token}",
        "customerid":     customer_id,
        "customerplanid": plan_id,
        "language":       "pt-BR",
        "origin":         GE_APP,
        "referer":        f"{GE_APP}/",
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    email    = os.environ["GE_EMAIL"]
    password = os.environ["GE_PASSWORD"]

    auth_data              = ge_login(email, password)
    token, cid, plan_id   = _extrair_auth(auth_data)
    request_id             = ge_solicitar_exportacao(token, cid, plan_id)
    content                = ge_aguardar_e_baixar(token, cid, plan_id, request_id)
    rows                   = parsear_excel(content)

    if not rows:
        raise ValueError("Arquivo Excel vazio.")

    upload_google_sheets(rows)
    log.info("✅ Sync concluído!")


if __name__ == "__main__":
    main()
