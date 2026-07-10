"""
BaseLinker Inventory CSV → Google Sheets (SNAPSHOT_BL)

Flow por armazém (Padrão + Armazenamento):
  1. Cria inventário rascunho via UI do BaseLinker (Playwright)
  2. Exporta CSV "Exportar itens de inventário (CSV)"
  3. Exclui o rascunho
  4. Parseia CSV → {sku: {picking, armPad/arm, locF/locG}}

Dados de CHEGOU e dimensões via API BaseLinker.

Escreve na aba SNAPSHOT_BL da planilha:
  linha 1: timestamp + contagem
  linha 2: cabeçalho SKU|picking|armPad|arm|chg|locF|locG|peso|vol
  linhas 3+: dados (SKU em formato texto para preservar zeros à esquerda)
"""

import asyncio
import csv
import io
import json
import logging
import os
import re
import time
from datetime import datetime, timezone
from pathlib import Path

import gspread
import requests
from google.oauth2.service_account import Credentials
from playwright.async_api import async_playwright

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

BL_PANEL         = "https://panel-u.baselinker.com"
BL_LOGIN         = "https://login.baselinker.com"
BL_API_URL       = "https://api.baselinker.com/connector.php"
INVENTORY_ID     = 39947
WH_PADRAO        = "bl_44285"
WH_ARMAZENAMENTO = "bl_50394"
WH_CHEGOU        = "bl_51442"

# Planilha onde o AppsScript lê o SNAPSHOT_BL
SPREADSHEET_ID = "1aWRJItE7pSFU8Mic3b4Vesl09UFGbF-Iqi47X8G4mC0"
SNAPSHOT_TAB   = "SNAPSHOT_BL"

# Screenshots de debug (carregados como artefato do GitHub Actions)
DEBUG_DIR = Path("debug_screenshots")

BL_EMAIL    = os.environ["BASELINKER_EMAIL"]
BL_PASSWORD = os.environ["BASELINKER_PASSWORD"]
BL_API_KEY  = os.environ["BASELINKER_API_KEY"]
SA_JSON     = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]


# ─── BaseLinker API ───────────────────────────────────────────────────────────

def bl_call(method, params):
    resp = requests.post(
        BL_API_URL,
        data={
            "token":      BL_API_KEY,
            "method":     method,
            "parameters": json.dumps(params),
        },
        timeout=60,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("status") != "SUCCESS":
        raise RuntimeError(f"BaseLinker API [{method}]: {data.get('error_message', data)}")
    return data


def fetch_chegou_and_dimensions():
    """
    Retorna {sku: {chg, peso, vol}} via getInventoryProductsData.
    Usado para CHEGOU (armazém bl_51442, uma só localização) e dimensões.
    """
    log.info("[API] Buscando estoque CHEGOU + dimensões...")

    pid_to_sku = {}
    page_num = 1
    while True:
        r = bl_call("getInventoryProductsList", {"inventory_id": INVENTORY_ID, "page": page_num})
        entries = list((r.get("products") or {}).items())
        if not entries:
            break
        for pid, info in entries:
            sku = str(info.get("sku") or "").strip()
            if sku:
                pid_to_sku[pid] = sku
        if len(entries) < 1000:
            break
        page_num += 1
        time.sleep(0.2)

    log.info(f"[API] {len(pid_to_sku)} produtos listados")

    result = {}
    pids = list(pid_to_sku.keys())
    LOTE = 1000
    for i in range(0, len(pids), LOTE):
        lote = [int(p) for p in pids[i : i + LOTE]]
        r = bl_call("getInventoryProductsData", {"inventory_id": INVENTORY_ID, "products": lote})
        for pid, p in (r.get("products") or {}).items():
            sku = pid_to_sku.get(str(pid))
            if not sku:
                continue
            s    = p.get("stock") or {}
            chg  = int(s.get(WH_CHEGOU, 0) or 0)
            h    = float(p.get("height") or 0)
            w    = float(p.get("width")  or 0)
            c    = float(p.get("length") or 0)
            peso = float(p.get("weight") or 0)
            vol  = round(h * w * c, 2) if h and w and c else 0
            if sku not in result:
                result[sku] = {"chg": 0, "peso": 0, "vol": 0}
            result[sku]["chg"] += chg
            if peso > 0 and not result[sku]["peso"]:
                result[sku]["peso"] = peso
                result[sku]["vol"]  = vol
        if i + LOTE < len(pids):
            time.sleep(0.3)

    log.info(f"[API] {len(result)} SKUs com CHEGOU/dimensões")
    return result


# ─── Playwright helpers ───────────────────────────────────────────────────────

async def screenshot(page, name):
    DEBUG_DIR.mkdir(exist_ok=True)
    path = str(DEBUG_DIR / f"{name}.png")
    try:
        await page.screenshot(path=path, full_page=True)
        log.info(f"[SCREENSHOT] {path}")
    except Exception as e:
        log.warning(f"[SCREENSHOT] Falhou em {name}: {e}")


async def try_click(page, locators, label):
    """Tenta clicar em uma lista de localizadores — para no primeiro visível."""
    for loc in locators:
        try:
            el = loc.first if hasattr(loc, "first") else loc
            if await el.is_visible(timeout=3000):
                await el.click()
                log.info(f"[UI] Clicou: {label}")
                return True
        except Exception:
            pass
    log.warning(f"[UI] Não encontrou: {label}")
    return False


async def bl_login(page):
    log.info("[LOGIN] Acessando login.baselinker.com...")
    await page.goto(BL_LOGIN, wait_until="networkidle")
    await screenshot(page, "01_login_page")

    # Aguarda qualquer input aparecer (página pode ser JS-heavy)
    try:
        await page.wait_for_selector("input", timeout=15_000)
    except Exception:
        pass

    # Loga todos os inputs para debug em caso de falha futura
    inputs = await page.locator("input").all()
    for i, inp in enumerate(inputs):
        try:
            t = await inp.get_attribute("type")
            n = await inp.get_attribute("name")
            iid = await inp.get_attribute("id")
            ph  = await inp.get_attribute("placeholder")
            log.info(f"[LOGIN] input[{i}] type={t} name={n} id={iid} placeholder={ph}")
        except Exception:
            pass

    # ── Preenche email ────────────────────────────────────────────────────────
    email_filled = False
    for sel in [
        "input[type='email']",
        "input[name='email']",
        "input[name='login']",
        "input[id='email']",
        "input[id='login']",
        "input[autocomplete='email']",
        "input[autocomplete='username']",
        "input[placeholder*='mail']",
        "input[placeholder*='ogin']",
        "input[placeholder*='ser']",
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.is_visible(timeout=2000):
                await loc.fill(BL_EMAIL)
                email_filled = True
                log.info(f"[LOGIN] Email via: {sel}")
                break
        except Exception:
            pass

    if not email_filled:
        # Fallback: primeiro input visível (que não seja password/hidden)
        for nth in range(min(len(inputs), 5)):
            try:
                inp = page.locator("input:visible").nth(nth)
                t   = await inp.get_attribute("type") or ""
                if t.lower() in ("password", "hidden", "submit", "checkbox", "radio"):
                    continue
                await inp.fill(BL_EMAIL)
                email_filled = True
                log.info(f"[LOGIN] Email via input visível #{nth}")
                break
            except Exception:
                pass

    if not email_filled:
        await screenshot(page, "ERR_login_no_email_field")
        raise RuntimeError("[LOGIN] Campo de email não encontrado")

    # ── Preenche senha ────────────────────────────────────────────────────────
    pwd_filled = False
    for sel in [
        "input[type='password']",
        "input[name='password']",
        "input[id='password']",
        "input[autocomplete='current-password']",
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.is_visible(timeout=2000):
                await loc.fill(BL_PASSWORD)
                pwd_filled = True
                log.info(f"[LOGIN] Senha via: {sel}")
                break
        except Exception:
            pass

    if not pwd_filled:
        await screenshot(page, "ERR_login_no_pwd_field")
        raise RuntimeError("[LOGIN] Campo de senha não encontrado")

    await screenshot(page, "02_login_filled")

    # ── Submit ────────────────────────────────────────────────────────────────
    submitted = False
    for sel in [
        "button[type='submit']",
        "input[type='submit']",
        "button:has-text('Entrar')",
        "button:has-text('Login')",
        "button:has-text('Sign in')",
        "button:has-text('Loguj')",
        "button:has-text('Zaloguj')",
        "form button",
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.is_visible(timeout=2000):
                await loc.click()
                submitted = True
                log.info(f"[LOGIN] Submit via: {sel}")
                break
        except Exception:
            pass

    if not submitted:
        # Pressiona Enter como último recurso
        await page.keyboard.press("Enter")
        log.info("[LOGIN] Submit via Enter")

    await page.wait_for_url(f"{BL_PANEL}/**", timeout=30_000)
    await page.wait_for_load_state("networkidle")
    log.info(f"[LOGIN] OK — {page.url}")
    await screenshot(page, "03_panel")


async def export_warehouse_csv(page, wh_key, wh_display):
    """
    Cria inventário rascunho para o armazém, exporta CSV, exclui rascunho.
    Retorna o conteúdo do CSV como string.
    """
    log.info(f"[{wh_display}] Abrindo lista de inventários...")
    await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
    await page.wait_for_timeout(1500)
    await screenshot(page, f"04_{wh_key}_list")

    # ── 1. Criar novo inventário ─────────────────────────────────────────────
    opened = await try_click(page, [
        page.get_by_role("button", name=re.compile(r"(Adicionar|Novo|Add|New|\+)", re.I)),
        page.get_by_role("link",   name=re.compile(r"(Adicionar|Novo|Add|New)", re.I)),
        page.locator("a.btn-success, button.btn-success, a.btn-primary, button.btn-primary").first,
        page.locator("[data-action='add'], [id*='add_btn'], [class*='add']").first,
    ], f"Adicionar inventário ({wh_display})")

    if not opened:
        await screenshot(page, f"ERR_{wh_key}_no_add_btn")
        raise RuntimeError(f"[{wh_display}] Botão de criar inventário não encontrado")

    await page.wait_for_timeout(1500)
    await screenshot(page, f"05_{wh_key}_modal")

    # ── 2. Selecionar armazém no formulário ──────────────────────────────────
    # Tenta <select> com a opção pelo nome do armazém
    select_with_wh = page.locator("select").filter(
        has=page.locator(f"option:has-text('{wh_display}')")
    )
    try:
        if await select_with_wh.first.is_visible(timeout=3000):
            await select_with_wh.first.select_option(label=wh_display)
            log.info(f"[{wh_display}] Armazém selecionado via <select>")
    except Exception:
        # Tenta clicar no texto do armazém (radio, checkbox, ou item de lista)
        wh_option = page.locator(
            f"label:has-text('{wh_display}'), li:has-text('{wh_display}'), "
            f"div:has-text('{wh_display}'), span:has-text('{wh_display}')"
        ).first
        try:
            if await wh_option.is_visible(timeout=3000):
                await wh_option.click()
                log.info(f"[{wh_display}] Armazém selecionado via texto")
        except Exception:
            log.warning(f"[{wh_display}] Seleção de armazém não confirmada — continuando...")

    await screenshot(page, f"06_{wh_key}_wh_selected")

    # ── 3. Confirmar criação ─────────────────────────────────────────────────
    await try_click(page, [
        page.get_by_role("button", name=re.compile(r"(Salvar|Criar|Confirmar|OK|Zapisz|Dodaj)", re.I)),
        page.locator("button[type='submit']").first,
        page.locator(".modal-footer .btn-primary, .modal-footer .btn-success").first,
        page.locator("form .btn-primary").first,
    ], f"Confirmar criação ({wh_display})")

    # Aguarda o sistema gerar o inventário (usuário relatou ~5 segundos)
    log.info(f"[{wh_display}] Aguardando inventário ser gerado (até 15s)...")
    await page.wait_for_timeout(15_000)
    await page.wait_for_load_state("networkidle")
    await screenshot(page, f"07_{wh_key}_after_create")

    # ── 4. Selecionar o inventário recém-criado (primeiro da lista) ──────────
    # Se redirecionou para a página do inventário, volta à lista
    if "/inventory_stocktakes" not in page.url:
        await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
        await page.wait_for_timeout(1000)

    selected = await try_click(page, [
        page.locator("table tbody tr:first-child input[type='checkbox']"),
        page.locator(".list-item:first-child input[type='checkbox']"),
        page.locator("tr:first-child input[type='checkbox']").first,
        page.locator("input[type='checkbox']").first,
    ], f"Checkbox do inventário ({wh_display})")

    if not selected:
        await screenshot(page, f"ERR_{wh_key}_no_checkbox")
        raise RuntimeError(f"[{wh_display}] Checkbox do inventário não encontrado")

    await screenshot(page, f"08_{wh_key}_checked")

    # ── 5. Clicar em IMPRIMIR ────────────────────────────────────────────────
    clicked_print = await try_click(page, [
        page.get_by_role("button", name=re.compile(r"imprimir|print", re.I)),
        page.locator("button:has-text('IMPRIMIR'), a:has-text('IMPRIMIR')").first,
        page.locator("[data-action='print'], .btn-print").first,
    ], f"IMPRIMIR ({wh_display})")

    if not clicked_print:
        await screenshot(page, f"ERR_{wh_key}_no_print_btn")
        raise RuntimeError(f"[{wh_display}] Botão IMPRIMIR não encontrado")

    await page.wait_for_timeout(700)
    await screenshot(page, f"09_{wh_key}_print_menu")

    # ── 6. Exportar itens de inventário (CSV) ────────────────────────────────
    async with page.expect_download(timeout=30_000) as dl_info:
        clicked_csv = await try_click(page, [
            page.get_by_text(re.compile(r"exportar.*csv|csv.*export|itens.*inventário", re.I)),
            page.locator("li:has-text('CSV'), a:has-text('CSV')").first,
        ], f"Exportar CSV ({wh_display})")

        if not clicked_csv:
            await screenshot(page, f"ERR_{wh_key}_no_csv_option")
            raise RuntimeError(f"[{wh_display}] Opção de exportar CSV não encontrada")

    download = await dl_info.value
    log.info(f"[{wh_display}] Download: {download.suggested_filename}")

    tmp_path = await download.path()
    csv_content = Path(tmp_path).read_text(encoding="utf-8-sig", errors="replace")
    log.info(f"[{wh_display}] CSV: {len(csv_content)} chars, {csv_content.count(chr(10))} linhas")
    await screenshot(page, f"10_{wh_key}_downloaded")

    # ── 7. Excluir rascunho ──────────────────────────────────────────────────
    deleted = await try_click(page, [
        page.get_by_role("button", name=re.compile(r"(Excluir|Deletar|Remover|Delete|Usuń)", re.I)),
        page.locator("button.btn-danger, a.btn-danger").first,
        page.locator("[data-action='delete']").first,
    ], f"Excluir rascunho ({wh_display})")

    if deleted:
        await page.wait_for_timeout(500)
        # Confirma diálogo de confirmação se aparecer
        await try_click(page, [
            page.get_by_role("button", name=re.compile(r"(Sim|Yes|Confirmar|OK|Tak)", re.I)),
            page.locator(".modal-footer .btn-danger, .bootbox .btn-danger").first,
        ], f"Confirmar exclusão ({wh_display})")
        log.info(f"[{wh_display}] Rascunho excluído")
    else:
        log.warning(f"[{wh_display}] Não conseguiu excluir rascunho — remova manualmente")

    return csv_content


# ─── CSV Parsing ─────────────────────────────────────────────────────────────

def detect_delimiter(sample):
    return ";" if sample.count(";") >= sample.count(",") else ","


def find_col(headers, *candidates):
    """Retorna índice da primeira coluna cujo header contenha algum candidato (case-insensitive)."""
    hl = [h.lower().strip() for h in headers]
    for cand in candidates:
        cand_l = cand.lower()
        for i, h in enumerate(hl):
            if cand_l in h:
                return i
    return None


def parse_csv(csv_content, is_padrao):
    """
    Parseia o CSV de inventário do BaseLinker.
    Para Padrão:       A8 locations → picking; A9 → armPad; locF/locG
    Para Armazenamento: tudo → arm; locG
    Retorna {sku: {picking, armPad, arm, locF, locG}}
    """
    result = {}
    delim = detect_delimiter(csv_content[:3000])
    reader = csv.reader(io.StringIO(csv_content), delimiter=delim)

    headers = None
    sku_col = loc_col = qty_col = None

    for row in reader:
        if not any(c.strip() for c in row):
            continue

        if headers is None:
            headers = row
            sku_col = find_col(headers, "sku", "cod", "codigo", "code", "ean", "referencia")
            loc_col = find_col(headers, "locali", "location", "posicao", "posição", "prateleira", "shelf")
            qty_col = find_col(headers, "esperado", "expected", "quantidade", "quantity", "estoque", "qty")
            log.info(f"[CSV {'Padrão' if is_padrao else 'Arm'}] headers={headers}")
            log.info(f"[CSV] cols → sku:{sku_col} loc:{loc_col} qty:{qty_col}")
            if sku_col is None:
                sku_col = 0
                log.warning("[CSV] Coluna SKU não identificada — usando col 0")
            if qty_col is None:
                qty_col = len(headers) - 1
                log.warning("[CSV] Coluna quantidade não identificada — usando última col")
            continue

        max_col = max(c for c in [sku_col, loc_col, qty_col] if c is not None)
        if len(row) <= max_col:
            continue

        sku = str(row[sku_col]).strip()
        if not sku or sku.lower() in ("sku", "cod", "codigo", "ean", "referencia"):
            continue

        loc = str(row[loc_col]).strip() if loc_col is not None else ""
        try:
            qty = int(float(str(row[qty_col]).replace(",", ".").strip() or "0"))
        except (ValueError, IndexError):
            qty = 0

        if qty <= 0:
            continue

        loc_up = loc.upper()
        is_a8 = loc_up.startswith("A8")
        is_a9 = loc_up.startswith("A9")

        if sku not in result:
            result[sku] = {"picking": 0, "armPad": 0, "arm": 0, "locF": set(), "locG": set()}

        d = result[sku]

        if is_padrao:
            if is_a8:
                d["picking"] += qty
                if loc:
                    d["locF"].add(loc)
            elif is_a9:
                d["armPad"] += qty
                if loc:
                    d["locG"].add(loc)
            else:
                d["picking"] += qty   # fallback sem prefixo A8/A9
        else:
            d["arm"] += qty
            if loc:
                d["locG"].add(loc)

    wh = "Padrão" if is_padrao else "Armazenamento"
    log.info(f"[CSV {wh}] {len(result)} SKUs parseados")
    return result


# ─── Merge + Google Sheets ────────────────────────────────────────────────────

def merge_data(padrao, arm, api):
    """Combina dados de CSV (picking/armPad/arm/loc) e API (chg/peso/vol)."""
    all_skus = sorted(set(padrao) | set(arm) | set(api))
    rows = []
    for sku in all_skus:
        p = padrao.get(sku, {})
        a = arm.get(sku, {})
        q = api.get(sku, {})

        picking = p.get("picking", 0) or 0
        arm_pad = p.get("armPad",  0) or 0
        arm_qty = a.get("arm",     0) or 0
        chg     = q.get("chg",     0) or 0

        locF = " / ".join(sorted(p.get("locF") or set())) or ""
        locG_set = (p.get("locG") or set()) | (a.get("locG") or set())
        locG = " / ".join(sorted(locG_set)) or ""

        peso = q.get("peso", 0) or 0
        vol  = q.get("vol",  0) or 0

        rows.append([sku, picking, arm_pad, arm_qty, chg, locF, locG, peso, vol])

    return rows


def write_to_sheets(rows):
    log.info(f"[SHEETS] Escrevendo {len(rows)} SKUs em {SNAPSHOT_TAB}...")

    creds = Credentials.from_service_account_info(
        json.loads(SA_JSON),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SNAPSHOT_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SNAPSHOT_TAB, rows=len(rows) + 10, cols=10)

    ws.clear()

    ts  = datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M:%S UTC")
    hdr = ["SKU", "picking", "armPad", "arm", "chg", "locF", "locG", "peso", "vol"]

    # Linha 1: info, Linha 2: cabeçalho, Linhas 3+: dados
    all_data = [
        [f"Atualizado: {ts}  |  SKUs: {len(rows)}"] + [""] * 8,
        hdr,
    ] + rows

    # RAW preserva strings com zeros à esquerda (ex: "00118" não vira 118)
    ws.update("A1", all_data, value_input_option="RAW")

    # Garante que coluna A (SKU) seja tratada como texto no Sheets
    ws.format("A3:A", {"numberFormat": {"type": "TEXT"}})

    log.info(f"[SHEETS] OK — {len(rows)} SKUs gravados")


# ─── Main ─────────────────────────────────────────────────────────────────────

async def run():
    DEBUG_DIR.mkdir(exist_ok=True)

    # 1. Dados de CHEGOU + dimensões via API (sem browser)
    api_data = fetch_chegou_and_dimensions()

    # 2. Exportar CSVs via Playwright
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        ctx = await browser.new_context(
            accept_downloads=True,
            locale="pt-BR",
            viewport={"width": 1280, "height": 900},
        )
        page = await ctx.new_page()

        await bl_login(page)

        padrao_csv = await export_warehouse_csv(page, "padrao", "Padrão")
        arm_csv    = await export_warehouse_csv(page, "arm",    "Armazenamento")

        await browser.close()

    # Salva CSVs brutos para debug
    (DEBUG_DIR / "padrao.csv").write_bytes(padrao_csv.encode("utf-8"))
    (DEBUG_DIR / "armazenamento.csv").write_bytes(arm_csv.encode("utf-8"))

    # 3. Parseia
    padrao_data = parse_csv(padrao_csv, is_padrao=True)
    arm_data    = parse_csv(arm_csv,    is_padrao=False)

    # 4. Merge
    rows = merge_data(padrao_data, arm_data, api_data)

    # 5. Grava no Sheets
    write_to_sheets(rows)

    log.info(f"[DONE] {len(rows)} SKUs no SNAPSHOT_BL")


if __name__ == "__main__":
    asyncio.run(run())
