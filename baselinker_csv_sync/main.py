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
from playwright_stealth import stealth_async

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


async def bl_inject_session(ctx):
    """
    Alternativa ao login via formulário: injeta cookies de sessão diretamente.
    Usado quando BASELINKER_SESSION_COOKIE está definido no ambiente.
    Formato da variável: JSON array de objetos cookie do DevTools.
    Ex: [{"name":"PHPSESSID","value":"abc123","domain":".baselinker.com","path":"/"}]
    """
    raw = os.environ.get("BASELINKER_SESSION_COOKIE", "")
    if not raw:
        return False
    try:
        cookies = json.loads(raw)
        # Garante que domain e path estejam presentes
        for c in cookies:
            c.setdefault("domain", "panel-u.baselinker.com")
            c.setdefault("path", "/")
        await ctx.add_cookies(cookies)
        log.info(f"[LOGIN] {len(cookies)} cookie(s) de sessão injetado(s)")
        return True
    except Exception as e:
        log.warning(f"[LOGIN] Falha ao injetar cookies: {e}")
        return False


async def bl_login(page):
    log.info("[LOGIN] Acessando login.baselinker.com...")
    await page.goto(BL_LOGIN, wait_until="networkidle")
    await screenshot(page, "01_login_page")

    # Aguarda formulário aparecer (Cloudflare pode atrasar)
    try:
        await page.wait_for_selector("input:not([type='hidden'])", timeout=20_000)
    except Exception:
        pass

    # Loga todos os inputs encontrados para diagnóstico
    inputs = await page.locator("input").all()
    log.info(f"[LOGIN] {len(inputs)} inputs encontrados na página:")
    for i, inp in enumerate(inputs):
        try:
            t   = await inp.get_attribute("type")
            n   = await inp.get_attribute("name")
            iid = await inp.get_attribute("id")
            ph  = await inp.get_attribute("placeholder")
            log.info(f"  input[{i}] type={t} name={n} id={iid} placeholder={ph}")
        except Exception:
            pass

    # ── Preenche campo LOGIN ──────────────────────────────────────────────────
    # A página usa label "LOGIN" — get_by_label é a forma mais robusta
    login_filled = False
    for attempt in [
        lambda: page.get_by_label("LOGIN", exact=True),
        lambda: page.get_by_label("Login", exact=False),
        lambda: page.locator("input[name='login']"),
        lambda: page.locator("input[name='email']"),
        lambda: page.locator("input[id='login']"),
        lambda: page.locator("input[id='email']"),
        lambda: page.locator("input[type='text']").first,
    ]:
        try:
            loc = attempt()
            if await loc.first.is_visible(timeout=3000):
                await loc.first.fill(BL_EMAIL)
                login_filled = True
                log.info("[LOGIN] Campo LOGIN preenchido")
                break
        except Exception:
            pass

    if not login_filled:
        await screenshot(page, "ERR_login_no_login_field")
        raise RuntimeError("[LOGIN] Campo LOGIN não encontrado — verifique screenshot ERR_login_no_login_field.png")

    # ── Preenche SENHA ────────────────────────────────────────────────────────
    pwd_filled = False
    for attempt in [
        lambda: page.get_by_label("SENHA", exact=True),
        lambda: page.get_by_label("Senha", exact=False),
        lambda: page.locator("input[type='password']").first,
        lambda: page.locator("input[name='password']").first,
    ]:
        try:
            loc = attempt()
            if await loc.first.is_visible(timeout=3000):
                await loc.first.fill(BL_PASSWORD)
                pwd_filled = True
                log.info("[LOGIN] Campo SENHA preenchido")
                break
        except Exception:
            pass

    if not pwd_filled:
        await screenshot(page, "ERR_login_no_pwd_field")
        raise RuntimeError("[LOGIN] Campo SENHA não encontrado")

    await screenshot(page, "02_login_filled")

    # ── Clica em "Faça o login" ───────────────────────────────────────────────
    submitted = False
    for attempt in [
        lambda: page.get_by_role("button", name=re.compile(r"fa.a o login|entrar|login", re.I)),
        lambda: page.locator("button[type='submit']").first,
        lambda: page.locator("input[type='submit']").first,
        lambda: page.locator("form button").first,
    ]:
        try:
            loc = attempt()
            if await loc.first.is_visible(timeout=3000):
                await loc.first.click()
                submitted = True
                log.info("[LOGIN] Botão de login clicado")
                break
        except Exception:
            pass

    if not submitted:
        await page.keyboard.press("Enter")
        log.info("[LOGIN] Submit via Enter")

    await page.wait_for_url(f"{BL_PANEL}/**", timeout=30_000)
    await page.wait_for_load_state("networkidle")
    log.info(f"[LOGIN] OK — {page.url}")
    await screenshot(page, "03_panel")


async def export_inventory_csv(page):
    """
    Cria um inventário rascunho (sem seleção de armazém — modal é simples),
    exporta o CSV completo, exclui o rascunho.
    Retorna o conteúdo CSV como string.

    FIX: o botão "Criar inventário" existe tanto na PÁGINA (abre o modal)
    quanto DENTRO do modal (confirma a criação). O `.last` garante que
    clicamos no do modal, não no da página.
    """
    log.info("[EXPORT] Abrindo inventory_stocktakes...")
    await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
    await page.wait_for_timeout(2000)
    await screenshot(page, "04_list")

    # ── 1. Abre o modal clicando no botão da PÁGINA ──────────────────────────
    # O botão pode ser <a> ou <button> com classe btn-success ou btn-primary
    opened = await try_click(page, [
        page.locator("a.btn-success, button.btn-success, a.btn-primary, button.btn-primary").first,
        page.locator("[data-action='add'], [id*='add_btn'], [class*='create']").first,
        page.get_by_role("button", name=re.compile(r"criar|add|novo|new", re.I)).first,
        page.get_by_role("link",   name=re.compile(r"criar|add|novo|new", re.I)).first,
    ], "Abrir modal 'Criar inventário'")

    if not opened:
        await screenshot(page, "ERR_no_open_modal_btn")
        raise RuntimeError("Botão para abrir modal de criação não encontrado")

    # Aguarda modal aparecer
    try:
        await page.wait_for_selector(
            ".modal.in, .modal.show, [role='dialog'][aria-modal='true']",
            state="visible", timeout=10_000
        )
    except Exception:
        pass
    await page.wait_for_timeout(2000)
    await screenshot(page, "05_modal")

    # ── 2. Clica "Criar inventário" DENTRO do modal ──────────────────────────
    # CRÍTICO: usar .last para pegar o botão do modal, não o que abriu o modal
    confirm_clicked = False
    modal = page.locator(".modal.in, .modal.show, [role='dialog']").first

    for btn_loc in [
        # Escopo do modal — mais confiável
        modal.get_by_role("button", name=re.compile(r"criar inventário", re.I)),
        modal.locator("button.btn-primary").last,
        modal.locator("button.btn-success").last,
        modal.locator(".modal-footer button").last,
        # Fallback global: ÚLTIMO botão com esse texto (o do modal)
        page.get_by_role("button", name=re.compile(r"criar inventário", re.I)).last,
    ]:
        try:
            if await btn_loc.is_visible(timeout=2000):
                await btn_loc.click()
                confirm_clicked = True
                log.info("[EXPORT] Botão confirmar (modal) clicado")
                break
        except Exception:
            pass

    if not confirm_clicked:
        await screenshot(page, "ERR_no_confirm_btn")
        raise RuntimeError("Botão 'Criar inventário' no modal não encontrado")

    # Aguarda modal fechar
    log.info("[EXPORT] Aguardando modal fechar e inventário ser gerado (~15s)...")
    try:
        await page.wait_for_selector(
            ".modal.in, .modal.show, [role='dialog'][aria-modal='true']",
            state="hidden", timeout=30_000
        )
    except Exception:
        pass
    await page.wait_for_timeout(8_000)
    await page.wait_for_load_state("networkidle")
    await screenshot(page, "06_after_create")

    # ── 3. Volta à lista e seleciona o primeiro inventário (mais recente) ────
    if "/inventory_stocktakes" not in page.url:
        await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
        await page.wait_for_timeout(2000)
    await screenshot(page, "07_list_after_create")

    first_cb = page.locator("input[type='checkbox']").first
    try:
        await first_cb.wait_for(state="attached", timeout=10_000)
        # dispatch_event ignora viewport (checkbox class="px" é customizado via CSS)
        await first_cb.dispatch_event("click")
        log.info("[EXPORT] Primeiro inventário selecionado via dispatch_event")
    except Exception as e:
        # Fallback: clica via JavaScript diretamente no DOM
        try:
            await page.evaluate("document.querySelector('input[type=\"checkbox\"]').click()")
            log.info("[EXPORT] Primeiro inventário selecionado via JS eval")
        except Exception as e2:
            await screenshot(page, "ERR_no_checkbox")
            raise RuntimeError(f"Checkbox não clicável: dispatch={e} | js={e2}")

    await screenshot(page, "08_checked")

    # ── 4. IMPRIMIR → Exportar CSV ───────────────────────────────────────────
    print_clicked = await try_click(page, [
        page.get_by_role("button", name=re.compile(r"imprimir|print", re.I)),
        page.locator("button:has-text('IMPRIMIR')").first,
        page.locator("[data-action='print']").first,
    ], "IMPRIMIR")

    if not print_clicked:
        await screenshot(page, "ERR_no_print_btn")
        raise RuntimeError("Botão IMPRIMIR não encontrado")

    await page.wait_for_timeout(700)
    await screenshot(page, "09_print_menu")

    async with page.expect_download(timeout=30_000) as dl_info:
        csv_clicked = await try_click(page, [
            page.get_by_text(re.compile(r"exportar.*csv|itens.*inventário", re.I)),
            page.locator("li:has-text('CSV'), a:has-text('CSV')").first,
        ], "Exportar CSV")

        if not csv_clicked:
            await screenshot(page, "ERR_no_csv_option")
            raise RuntimeError("Opção exportar CSV não encontrada")

    download = await dl_info.value
    tmp_path = await download.path()
    csv_content = Path(tmp_path).read_text(encoding="utf-8-sig", errors="replace")
    log.info(f"[EXPORT] CSV: {len(csv_content)} chars, {csv_content.count(chr(10))} linhas")
    await screenshot(page, "10_downloaded")

    # ── 5. Exclui o rascunho ─────────────────────────────────────────────────
    deleted = await try_click(page, [
        page.get_by_role("button", name=re.compile(r"excluir|deletar|delete|remover", re.I)),
        page.locator("button.btn-danger").first,
        page.locator("[data-action='delete']").first,
    ], "Excluir inventário")

    if deleted:
        await page.wait_for_timeout(500)
        await try_click(page, [
            page.get_by_role("button", name=re.compile(r"sim|yes|confirmar|ok", re.I)),
            page.locator(".modal-footer .btn-danger, .bootbox-accept").first,
        ], "Confirmar exclusão")
        log.info("[EXPORT] Inventário excluído")
    else:
        log.warning("[EXPORT] Não conseguiu excluir — remova manualmente")

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


def parse_csv(csv_content):
    """
    Parseia o CSV único de inventário do BaseLinker.

    Lógica de zona por localização:
      A8... → picking   (locF)
      A9... → armPad    (locG)   [Padrão — zona de armazenagem pequena]
      Sem A8/A9 → picking (fallback)

    Se o CSV tiver coluna de armazém, produtos do armazém "Armazenamento"
    com A9 vão para arm em vez de armPad (detectado pelo nome da coluna).

    Retorna {sku: {picking, armPad, arm, locF: set, locG: set}}

    NOTA: depois de ver o CSV real os headers serão logados —
    ajuste as colunas se necessário.
    """
    result = {}
    delim = detect_delimiter(csv_content[:3000])
    reader = csv.reader(io.StringIO(csv_content), delimiter=delim)

    headers = None
    sku_col = loc_col = qty_col = wh_col = None

    for row in reader:
        if not any(c.strip() for c in row):
            continue

        if headers is None:
            headers = row
            sku_col = find_col(headers, "sku", "cod", "codigo", "code", "ean", "referencia")
            loc_col = find_col(headers, "locali", "location", "posicao", "posição", "prateleira", "shelf")
            qty_col = find_col(headers, "esperado", "expected", "quantidade", "quantity", "estoque", "qty")
            wh_col  = find_col(headers, "armazem", "armazém", "warehouse", "deposito", "depósito")
            log.info(f"[CSV] headers={headers}")
            log.info(f"[CSV] cols → sku:{sku_col} loc:{loc_col} qty:{qty_col} wh:{wh_col}")
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
        wh  = str(row[wh_col]).strip()  if wh_col  is not None else ""
        try:
            qty = int(float(str(row[qty_col]).replace(",", ".").strip() or "0"))
        except (ValueError, IndexError):
            qty = 0

        if qty <= 0:
            continue

        loc_up = loc.upper()
        is_a8 = loc_up.startswith("A8")
        is_a9 = loc_up.startswith("A9")
        is_arm_wh = "armazenamento" in wh.lower() if wh else False

        if sku not in result:
            result[sku] = {"picking": 0, "armPad": 0, "arm": 0, "locF": set(), "locG": set()}

        d = result[sku]

        if is_a8:
            d["picking"] += qty
            if loc:
                d["locF"].add(loc)
        elif is_a9:
            if is_arm_wh:
                # A9 no armazém Armazenamento → arm
                d["arm"] += qty
            else:
                # A9 no armazém Padrão (ou sem info de armazém) → armPad
                d["armPad"] += qty
            if loc:
                d["locG"].add(loc)
        else:
            d["picking"] += qty  # fallback

    log.info(f"[CSV] {len(result)} SKUs parseados")
    return result


# ─── Merge + Google Sheets ────────────────────────────────────────────────────

def merge_data(csv_data, api):
    """Combina dados do CSV único (picking/armPad/arm/loc) com API (chg/peso/vol)."""
    all_skus = sorted(set(csv_data) | set(api))
    rows = []
    for sku in all_skus:
        c = csv_data.get(sku, {})
        q = api.get(sku, {})

        picking = c.get("picking", 0) or 0
        arm_pad = c.get("armPad",  0) or 0
        arm_qty = c.get("arm",     0) or 0
        chg     = q.get("chg",     0) or 0

        locF = " / ".join(sorted(c.get("locF") or set())) or ""
        locG = " / ".join(sorted(c.get("locG") or set())) or ""

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
        # headless=False + Xvfb (no workflow) para passar pelo Cloudflare Turnstile
        browser = await pw.chromium.launch(
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
            ],
        )
        ctx = await browser.new_context(
            accept_downloads=True,
            locale="pt-BR",
            viewport={"width": 1280, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/126.0.0.0 Safari/537.36"
            ),
        )
        page = await ctx.new_page()
        await stealth_async(page)  # mascara sinais de automação

        # Tenta injetar cookie de sessão salvo (bypass do Cloudflare)
        session_injected = await bl_inject_session(ctx)
        if session_injected:
            # Verifica se a sessão ainda é válida indo direto ao painel
            await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
            if BL_LOGIN in page.url or "login" in page.url.lower():
                log.warning("[LOGIN] Cookie expirado — fazendo login via formulário")
                await bl_login(page)
            else:
                log.info(f"[LOGIN] Sessão via cookie OK — {page.url}")
                await screenshot(page, "03_panel_via_cookie")
        else:
            await bl_login(page)

        # Uma única exportação — o inventário cobre todos os armazéns
        inv_csv = await export_inventory_csv(page)

        await browser.close()

    # Salva CSV bruto para debug
    (DEBUG_DIR / "inventory.csv").write_bytes(inv_csv.encode("utf-8"))

    # 3. Parseia: A8→picking, A9 Padrão→armPad, A9 Arm→arm
    # parse_csv com is_padrao=True extrai picking+armPad (A8/A9)
    # O merge_data combina com api_data (CHEGOU) para arm separado quando necessário
    csv_data = parse_csv(inv_csv)

    # 4. Merge
    rows = merge_data(csv_data, api_data)

    # 5. Grava no Sheets
    write_to_sheets(rows)

    log.info(f"[DONE] {len(rows)} SKUs no SNAPSHOT_BL")


if __name__ == "__main__":
    asyncio.run(run())
