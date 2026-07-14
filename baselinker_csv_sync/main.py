"""
BaseLinker Inventory CSV → Google Sheets (SNAPSHOT_BL)

Flow por armazém (Padrão + Armazenamento):
  1. Cria inventário rascunho via UI do BaseLinker (Playwright)
  2. Recarrega a página (rascunho vem sem produtos; após reload lista completa)
  3. Exporta CSV "Exportar itens de inventário (CSV)"
  4. Exclui o rascunho
  5. Parseia CSV → {sku: {picking, armPad/arm, locF/locG}}

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
    login_filled = False
    for attempt in [
        lambda: page.locator("input[id='loginField']"),          # id exato da BaseLinker
        lambda: page.locator("input[name='login'][type='text']"), # name='login' mas só type=text
        lambda: page.get_by_label("LOGIN", exact=True),
        lambda: page.get_by_label("Login", exact=False),
        lambda: page.locator("input[name='login']:not([type='hidden'])"),
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
        lambda: page.locator("input[id='passwordField']"),           # id exato da BaseLinker
        lambda: page.locator("input[name='password'][id='passwordField']"),
        lambda: page.get_by_label("SENHA", exact=True),
        lambda: page.get_by_label("Senha", exact=False),
        lambda: page.locator("input[type='password']#passwordField"),
        lambda: page.locator("input[type='password']:not([id*='change']):not([id*='pin'])").first,
        lambda: page.locator("input[name='password']").last,
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


async def _find_visible_modal(page):
    """Retorna info do modal visível (getBoundingClientRect().width > 0)."""
    return await page.evaluate("""
        () => {
            const modal = [...document.querySelectorAll('.modal, [role="dialog"], .bootbox')]
                .find(m => {
                    const r = m.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                });
            if (!modal) return null;

            // Selects nativos
            const selects = [...modal.querySelectorAll('select')].map((sel, i) => ({
                i, id: sel.id, name: sel.name, value: sel.value,
                options: [...sel.options].map(o => o.text.trim()),
                rect: (() => { const r = sel.getBoundingClientRect();
                    return [Math.round(r.left), Math.round(r.top),
                            Math.round(r.width), Math.round(r.height)]; })()
            }));

            // Todos os elementos de formulário (para detectar select2/custom)
            const formEls = [...modal.querySelectorAll(
                'input:not([type=hidden]), select, textarea, ' +
                '[role="combobox"], [role="listbox"], [role="option"], ' +
                '.select2-container, .select2-choice, ' +
                '[data-toggle="dropdown"], ' +
                'div[class*="select"], span[class*="select"]'
            )].map(el => ({
                tag: el.tagName,
                type: el.getAttribute('type') || '',
                id: el.id || '',
                cls: el.className.substring(0, 80),
                role: el.getAttribute('role') || '',
                text: el.textContent.trim().substring(0, 60),
                rect: (() => { const r = el.getBoundingClientRect();
                    return [Math.round(r.left), Math.round(r.top),
                            Math.round(r.width), Math.round(r.height)]; })()
            }));

            const buttons = [...modal.querySelectorAll('button, input[type="submit"]')].map(btn => ({
                text: btn.textContent.trim(),
                classes: btn.className,
                dismiss: btn.getAttribute('data-dismiss'),
                rect: (() => { const r = btn.getBoundingClientRect();
                    return [Math.round(r.left), Math.round(r.top),
                            Math.round(r.width), Math.round(r.height)]; })()
            }));

            return {selects, formEls, buttons};
        }
    """)


async def export_inventory_csv(page, warehouse_name, tag=""):
    """
    Cria inventário rascunho para o armazém indicado, recarrega a página para
    garantir que todos os produtos/locações apareçam, exporta o CSV e exclui
    o rascunho.

    warehouse_name: nome exato do armazém como aparece no select do modal
                    (ex: "Padrão" ou "Armazenamento")
    tag:            prefixo curto para nomear screenshots (ex: "padrao" ou "arm")

    IMPORTANTE: após criar o rascunho, recarregar a página é necessário porque
    na primeira carga o inventário aparece sem produtos; após reload lista completa.
    """
    pfx = f"{tag}_" if tag else ""
    log.info(f"[EXPORT:{warehouse_name}] Abrindo inventory_stocktakes...")
    await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
    await page.wait_for_timeout(2000)
    await screenshot(page, f"{pfx}04_list")

    # ── 1. Abre modal "Criar inventário" ─────────────────────────────────────
    opened = await try_click(page, [
        page.locator("a.btn-success, button.btn-success, a.btn-primary, button.btn-primary").first,
        page.locator("[data-action='add'], [id*='add_btn'], [class*='create']").first,
        page.get_by_role("button", name=re.compile(r"criar|add|novo|new", re.I)).first,
        page.get_by_role("link",   name=re.compile(r"criar|add|novo|new", re.I)).first,
    ], f"Abrir modal 'Criar inventário' [{warehouse_name}]")

    if not opened:
        await screenshot(page, f"ERR_{pfx}no_open_modal_btn")
        raise RuntimeError(f"Botão para abrir modal de criação não encontrado [{warehouse_name}]")

    # Aguarda modal visível aparecer
    try:
        await page.wait_for_selector(
            ".modal.in, .modal.show, [role='dialog'][aria-modal='true']",
            state="visible", timeout=10_000
        )
    except Exception:
        pass
    await page.wait_for_timeout(1500)
    await screenshot(page, f"{pfx}05_modal")

    # ── 2. Diagnostica modal visível ─────────────────────────────────────────
    modal_info = await _find_visible_modal(page)
    log.info(f"[EXPORT:{warehouse_name}] Modal info: {json.dumps(modal_info)}")

    # ── 3. Define armazém e escopo no modal ──────────────────────────────────
    # Tenta selects nativos primeiro; caso não haja, tenta select2 / custom.
    set_result = await page.evaluate("""
        (warehouseName) => {
            const modal = [...document.querySelectorAll('.modal, [role="dialog"], .bootbox')]
                .find(m => {
                    const r = m.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                });
            if (!modal) return {error: 'no visible modal'};

            const selects = [...modal.querySelectorAll('select')];
            let whSet = false, escopoSet = false;

            // 1) Tenta native <select>
            for (const sel of selects) {
                if (!whSet) {
                    const whOpt = [...sel.options].find(o =>
                        o.text.trim().toLowerCase().includes(warehouseName.toLowerCase())
                    );
                    if (whOpt) {
                        sel.value = whOpt.value;
                        sel.dispatchEvent(new Event('change', {bubbles: true}));
                        whSet = true;
                        continue;
                    }
                }
                if (!escopoSet) {
                    const todosOpt = [...sel.options].find(o =>
                        o.text.trim().toLowerCase().includes('todos')
                    );
                    if (todosOpt) {
                        sel.value = todosOpt.value;
                        sel.dispatchEvent(new Event('change', {bubbles: true}));
                        escopoSet = true;
                    }
                }
            }

            // 2) Tenta Select2 / custom dropdown (data-value / aria-selected)
            if (!whSet) {
                const allOpts = [...modal.querySelectorAll(
                    '[role="option"], li[data-value], li[data-id], .select2-result'
                )];
                const whOpt = allOpts.find(el =>
                    el.textContent.trim().toLowerCase().includes(warehouseName.toLowerCase())
                );
                if (whOpt) {
                    whOpt.click();
                    whSet = true;
                }
            }
            if (!escopoSet) {
                const allOpts = [...modal.querySelectorAll(
                    '[role="option"], li[data-value], li[data-id], .select2-result'
                )];
                const todosOpt = allOpts.find(el =>
                    el.textContent.trim().toLowerCase().includes('todos')
                );
                if (todosOpt) {
                    todosOpt.click();
                    escopoSet = true;
                }
            }

            return {whSet, escopoSet, totalSelects: selects.length};
        }
    """, warehouse_name)
    log.info(f"[EXPORT:{warehouse_name}] Modal set result: {set_result}")

    await page.wait_for_timeout(500)
    await screenshot(page, f"{pfx}05b_modal_filled")

    # ── 4. Clica botão "Criar inventário" no modal via coordenadas reais ─────
    submit_pos = await page.evaluate("""
        () => {
            const modal = [...document.querySelectorAll('.modal, [role="dialog"], .bootbox')]
                .find(m => {
                    const r = m.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                });
            if (!modal) return null;

            const candidates = [...modal.querySelectorAll('button, input[type="submit"], a.btn')];
            for (const btn of candidates) {
                const r = btn.getBoundingClientRect();
                if (r.width <= 0 || r.height <= 0) continue;

                const text = btn.textContent.trim().toLowerCase();
                const isCancel = btn.getAttribute('data-dismiss') === 'modal'
                    || text.includes('cancelar') || text.includes('cancel')
                    || btn.classList.contains('close')
                    || text === 'x' || text === '×';

                if (!isCancel) {
                    return {
                        text: btn.textContent.trim(),
                        x: Math.round(r.left + r.width / 2),
                        y: Math.round(r.top + r.height / 2)
                    };
                }
            }
            return null;
        }
    """)

    if submit_pos:
        log.info(f"[EXPORT:{warehouse_name}] Clicando submit '{submit_pos['text']}' em ({submit_pos['x']}, {submit_pos['y']})")
        await page.mouse.click(submit_pos["x"], submit_pos["y"])
    else:
        log.warning(f"[EXPORT:{warehouse_name}] Botão submit não encontrado — tentando Enter")
        await page.keyboard.press("Enter")

    # Aguarda o modal FECHAR — indica que o AJAX de criação completou.
    # 5s era insuficiente: o servidor leva até ~20s para criar o inventário.
    log.info(f"[EXPORT:{warehouse_name}] Aguardando modal fechar (até 30s)...")
    try:
        await page.wait_for_function(
            """() => {
                return ![...document.querySelectorAll('.modal, [role="dialog"], .bootbox')]
                    .some(m => {
                        const r = m.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    });
            }""",
            timeout=30_000,
        )
        log.info(f"[EXPORT:{warehouse_name}] Modal fechou — inventário criado")
    except Exception:
        log.warning(f"[EXPORT:{warehouse_name}] Modal não fechou em 30s — forçando fechamento")

    await screenshot(page, f"{pfx}06_after_confirm")

    # Fecha modal caso ainda esteja aberto
    await page.evaluate("""
        () => {
            const x = document.querySelector('.modal .close, .modal [data-dismiss="modal"]');
            if (x) { x.click(); return; }
            if (window.$) {
                const m = $('.modal:visible');
                if (m.length) m.modal('hide');
            }
        }
    """)
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(1000)

    # ── 5. Navega à lista e RECARREGA ────────────────────────────────────────
    # O rascunho é criado primeiro sem produtos; após reload todos aparecem.
    log.info(f"[EXPORT:{warehouse_name}] Navegando à lista de inventários...")
    await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
    await page.wait_for_timeout(2000)
    await screenshot(page, f"{pfx}07_list_after_create")

    log.info(f"[EXPORT:{warehouse_name}] Recarregando para garantir produtos completos...")
    await page.reload(wait_until="networkidle")
    await page.wait_for_timeout(5000)  # 5s: produtos carregam via AJAX após o reload
    await screenshot(page, f"{pfx}07b_list_after_reload")

    # Pega o ID do inventário mais recente (maior ID = recém-criado).
    # Usa seletor amplo: tr[id^="row_"] cobre também novas linhas sem a classe stocktake_row.
    inv_info = await page.evaluate("""
        () => {
            const rows = [...document.querySelectorAll('tr[id^="row_"], tr.stocktake_row')];
            const ids = rows.map(r => {
                const m = r.id.match(/row_(\\d+)/);
                return m ? parseInt(m[1]) : 0;
            }).filter(Boolean);
            const unique = [...new Set(ids)];
            if (!unique.length) return null;
            const maxId = Math.max(...unique);
            return {first: String(unique[0]), max: String(maxId), all: unique.map(String)};
        }
    """)
    log.info(f"[EXPORT:{warehouse_name}] Inventários na lista: {json.dumps(inv_info)}")
    inv_id = inv_info["max"] if inv_info else None
    log.info(f"[EXPORT:{warehouse_name}] Usando inventário ID: {inv_id}")

    # ── 6. Seleciona checkbox via JS; captura URL e clica Imprimir ───────────
    await page.evaluate("""
        () => {
            window.__capturedUrls = [];
            const _open = window.open.bind(window);
            window.open = function(url) {
                if (url) window.__capturedUrls.push(String(url));
                return _open.apply(window, arguments);
            };
        }
    """)

    export_requests = []
    def _on_req(req):
        url = req.url
        if any(k in url.lower() for k in ["export", "csv", "stocktake", "download", "print"]):
            export_requests.append(url)
            log.info(f"[NET:{warehouse_name}] {url}")
    page.on("request", _on_req)

    # Seleciona o checkbox do inventário com o ID mais recente
    cb_ok = await page.evaluate("""
        (invId) => {
            // Tenta selecionar o checkbox do inventário específico primeiro
            let row = invId ? document.querySelector(`tr#row_${invId}`) : null;
            if (!row) row = document.querySelector('tr.stocktake_row');
            if (!row) return false;
            const cb = row.querySelector('input[type="checkbox"]');
            if (!cb) return false;
            cb.checked = true;
            cb.dispatchEvent(new Event('change', {bubbles: true}));
            return row.id || true;
        }
    """, inv_id)
    log.info(f"[EXPORT:{warehouse_name}] Checkbox JS: {cb_ok}")
    await page.wait_for_timeout(800)

    # Acha coordenadas do botão Imprimir
    btn_pos = await page.evaluate("""
        () => {
            const btn = [...document.querySelectorAll('button')]
                .find(b => b.textContent.trim() === 'Imprimir' && b.classList.contains('btn-reset'));
            if (!btn) return null;
            const r = btn.getBoundingClientRect();
            return {x: r.left + r.width / 2, y: r.top + r.height / 2, cls: btn.className};
        }
    """)
    log.info(f"[EXPORT:{warehouse_name}] Botão Imprimir: {btn_pos}")

    if not btn_pos:
        await screenshot(page, f"ERR_{pfx}no_print_btn")
        raise RuntimeError(f"Botão Imprimir não encontrado [{warehouse_name}]")

    # Clica com mouse real — o botão tem classe 'collapse-btn dropdown-toggle',
    # então abre um painel inline (Bootstrap collapse), não um dropdown flutuante.
    await page.mouse.click(btn_pos["x"], btn_pos["y"])
    log.info(f"[EXPORT:{warehouse_name}] mouse.click(Imprimir) em ({btn_pos['x']:.0f}, {btn_pos['y']:.0f})")
    await page.wait_for_timeout(1200)

    # Força abertura de dropdown pai E de qualquer collapse associado
    await page.evaluate("""
        (pos) => {
            // Bootstrap dropdown
            const el = document.elementFromPoint(pos.x, pos.y);
            if (el) {
                const parent = el.closest('.dropdown, .btn-group');
                if (parent) parent.classList.add('open');
            }
            // Bootstrap collapse: qualquer .collapse adjacente ao botão
            const btn = [...document.querySelectorAll('button.collapse-btn')]
                .find(b => b.textContent.trim() === 'Imprimir');
            if (btn) {
                const target = btn.getAttribute('data-target') || btn.getAttribute('href');
                if (target) {
                    const el2 = document.querySelector(target);
                    if (el2) { el2.classList.add('in'); el2.classList.add('show'); }
                }
                // Também tenta abrir o próximo .collapse irmão
                let sib = btn.parentElement;
                while (sib) {
                    const c = sib.querySelector('.collapse');
                    if (c) { c.classList.add('in'); c.classList.add('show'); break; }
                    sib = sib.parentElement;
                    if (!sib || sib === document.body) break;
                }
            }
        }
    """, {"x": btn_pos["x"], "y": btn_pos["y"]})
    await page.wait_for_timeout(600)
    await screenshot(page, f"{pfx}09_dropdown_open")

    # ── 7. Procura e clica no link de exportação CSV ──────────────────────────
    # O painel é um Bootstrap collapse (clipped by table overflow), por isso
    # getBoundingClientRect() retorna 0 para itens dentro da tabela.
    # Usamos JS .click() diretamente no elemento, que funciona independente de overflow.
    #
    # Log diagnóstico: mostra TODOS os itens dentro de qualquer dropdown-menu / collapse.
    diag = await page.evaluate("""
        () => {
            const containers = [
                ...document.querySelectorAll(
                    '.collapse.in, .collapse.show, ' +
                    '.dropdown-menu, .open .dropdown-menu, ' +
                    '.btn-group.open .dropdown-menu, ' +
                    '[aria-expanded="true"] ~ *, [aria-expanded="true"] + *'
                )
            ];
            return containers.flatMap(c =>
                [...c.querySelectorAll('a, button, li')].map(el => ({
                    tag: el.tagName,
                    text: el.textContent.trim().substring(0, 80),
                    href: el.getAttribute('href') || ''
                }))
            ).slice(0, 50);
        }
    """)
    log.info(f"[DEBUG:{warehouse_name}] Itens em colapso/dropdown: {json.dumps(diag)}")

    # Clique direto via JS — não depende de coordenadas nem de visibilidade CSS
    js_click = await page.evaluate("""
        () => {
            const KEYWORDS = ['csv', 'exportar item', 'export item', 'exportar itens',
                              'export itens', 'baixar', 'download'];
            // Escopo prioritário: collapse aberto + dropdown-menu
            const primary = [
                ...document.querySelectorAll(
                    '.collapse.in a, .collapse.show a, ' +
                    '.dropdown-menu a, .open .dropdown-menu a, ' +
                    '.btn-group.open .dropdown-menu a'
                )
            ];
            for (const el of primary) {
                const t = el.textContent.toLowerCase();
                if (KEYWORDS.some(k => t.includes(k))) {
                    el.click();
                    return {clicked: el.textContent.trim(), from: 'primary'};
                }
            }

            // Busca ampla: qualquer <a> com CSV no texto (exceto sidebar)
            for (const el of document.querySelectorAll('a')) {
                const t = el.textContent.toLowerCase();
                if (!t.includes('csv')) continue;
                const href = el.getAttribute('href') || '';
                if (href.includes('inventory_import')) continue;
                const r = el.getBoundingClientRect();
                // Aceita mesmo com rect zerado (overflow clipped) se não estiver na sidebar
                if (r.left > 0 && r.left < 250) continue;
                el.click();
                return {clicked: el.textContent.trim(), from: 'broad',
                        rect: [Math.round(r.left), Math.round(r.top)]};
            }

            // Log de todos os .dropdown-menu no DOM para diagnóstico
            const allMenuText = [...document.querySelectorAll('.dropdown-menu a, .collapse a')]
                .map(a => a.textContent.trim().substring(0, 60));
            return {not_found: true, allMenuText};
        }
    """)
    log.info(f"[EXPORT:{warehouse_name}] JS click result: {json.dumps(js_click)}")

    if not js_click or not js_click.get("clicked"):
        # Tenta Playwright force-click como último recurso
        for pl_loc in [
            page.locator('.collapse a').filter(has_text=re.compile(r'csv', re.I)),
            page.locator('.dropdown-menu a').filter(has_text=re.compile(r'csv', re.I)),
            page.get_by_role("menuitem", name=re.compile(r'csv', re.I)),
            page.get_by_text(re.compile(r'exportar.*csv|csv', re.I)).first,
        ]:
            try:
                cnt = await pl_loc.count()
                if cnt > 0:
                    await pl_loc.first.click(force=True, timeout=2000)
                    log.info(f"[EXPORT:{warehouse_name}] CSV clicado via Playwright force ({cnt})")
                    js_click = {"clicked": "playwright-force", "from": "playwright"}
                    break
            except Exception as pe:
                log.warning(f"[EXPORT:{warehouse_name}] Playwright force: {pe}")

    if not js_click or not js_click.get("clicked"):
        await screenshot(page, f"ERR_{pfx}no_csv_item")
        raise RuntimeError(
            f"Item CSV não encontrado no dropdown/collapse [{warehouse_name}]. "
            f"diag={json.dumps(diag)}"
        )

    log.info(f"[EXPORT:{warehouse_name}] CSV clicado: '{js_click.get('clicked')}' via {js_click.get('from')}")

    # ── 8. Captura o download disparado pelo clique JS ────────────────────────
    # O clique JS acima já foi executado. Agora aguardamos o download.
    csv_content = None

    try:
        # O download pode já estar em andamento; expect_download com timeout curto
        async with page.expect_download(timeout=12_000) as dl_info:
            # Re-dispara caso o download anterior não tenha iniciado
            await page.evaluate("""
                () => {
                    const KEYWORDS = ['csv', 'exportar item', 'exportar itens'];
                    for (const scope of ['.collapse.in', '.collapse.show', '.dropdown-menu', '.open .dropdown-menu']) {
                        for (const a of document.querySelectorAll(scope + ' a')) {
                            const t = a.textContent.toLowerCase();
                            if (KEYWORDS.some(k => t.includes(k))) { a.click(); return; }
                        }
                    }
                }
            """)
            log.info(f"[EXPORT:{warehouse_name}] Aguardando download...")
        dl = await dl_info.value
        dl_path = await dl.path()
        csv_content = Path(dl_path).read_text(encoding="utf-8-sig", errors="replace")
        log.info(f"[EXPORT:{warehouse_name}] Download OK: {dl.suggested_filename}")
    except Exception as e:
        log.warning(f"[EXPORT:{warehouse_name}] expect_download falhou: {e} — verificando URL capturada")
        await page.wait_for_timeout(3000)

    if not csv_content:
        opened_urls = await page.evaluate("() => window.__capturedUrls || []")
        page.remove_listener("request", _on_req)

        # Exclui JS/CSS/imagens — apenas URLs com csv/export/stocktake/download
        BAD_EXTS = (".js", ".css", ".png", ".ico", ".gif", ".woff", ".woff2", ".ttf", ".svg")
        target_url = None
        for u in list(opened_urls) + list(export_requests):
            lower_u = u.lower()
            if (any(k in lower_u for k in ["csv", "export", "stocktake", "download"])
                    and not any(lower_u.endswith(ext) for ext in BAD_EXTS)):
                target_url = u
                break

        if target_url:
            log.info(f"[EXPORT:{warehouse_name}] Baixando via URL: {target_url}")
            try:
                cookies_js = await page.context.cookies()
                sess = requests.Session()
                for c in cookies_js:
                    sess.cookies.set(c["name"], c["value"], domain=c.get("domain", ""))
                resp = sess.get(target_url, timeout=60)
                resp.raise_for_status()
                csv_content = resp.text
                log.info(f"[EXPORT:{warehouse_name}] URL download OK: {len(csv_content)} chars")
            except Exception as e2:
                log.error(f"[EXPORT:{warehouse_name}] Falha ao baixar URL: {e2}")
        else:
            log.error(
                f"[EXPORT:{warehouse_name}] Nenhuma URL de download capturada. "
                f"opened={opened_urls} net={export_requests}"
            )
            await screenshot(page, f"ERR_{pfx}no_download_url")

    if not csv_content:
        raise RuntimeError(f"Conteúdo CSV vazio [{warehouse_name}]")

    log.info(f"[EXPORT:{warehouse_name}] CSV: {len(csv_content)} chars, {csv_content.count(chr(10))} linhas")
    await screenshot(page, f"{pfx}10_downloaded")

    # ── 7. Exclui o rascunho ─────────────────────────────────────────────────
    deleted = await try_click(page, [
        page.get_by_role("button", name=re.compile(r"excluir|deletar|delete|remover", re.I)),
        page.locator("button.btn-danger").first,
        page.locator("[data-action='delete']").first,
    ], f"Excluir inventário [{warehouse_name}]")

    if deleted:
        await page.wait_for_timeout(500)
        await try_click(page, [
            page.get_by_role("button", name=re.compile(r"sim|yes|confirmar|ok", re.I)),
            page.locator(".modal-footer .btn-danger, .bootbox-accept").first,
        ], "Confirmar exclusão")
        log.info(f"[EXPORT:{warehouse_name}] Inventário excluído")
    else:
        log.warning(f"[EXPORT:{warehouse_name}] Não conseguiu excluir — remova manualmente")

    # Aguarda exclusão processar antes de criar o próximo
    await page.wait_for_timeout(2000)

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


def parse_csv(csv_content, warehouse_name="Padrão"):
    """
    Parseia CSV de inventário do BaseLinker para um armazém específico.

    Headers reais: #;Nome do produto;EAN;SKU;Localização;Número de unidades;...
      col 3 → SKU
      col 4 → Localização
      col 5 → Número de unidades (estoque atual)

    Lógica de zona:
      Padrão:        A8... → picking (locF)
                     A9... → armPad (locG)
                     else  → picking (fallback)
      Armazenamento: qualquer localização → arm (locG)

    Retorna {sku: {picking, armPad, arm, locF: set, locG: set}}
    """
    result = {}
    delim = detect_delimiter(csv_content[:3000])
    reader = csv.reader(io.StringIO(csv_content), delimiter=delim)

    headers = None
    sku_col = loc_col = qty_col = None
    is_armazenamento = "armazenamento" in warehouse_name.lower()

    for row in reader:
        if not any(c.strip() for c in row):
            continue

        if headers is None:
            headers = row
            sku_col = find_col(headers, "sku", "cod", "codigo", "code", "referencia")
            loc_col = find_col(headers, "locali", "location", "posicao", "posição", "prateleira", "shelf")
            # "Número de unidades" → candidate "unidades" (col 5 nos CSVs reais)
            qty_col = find_col(
                headers,
                "unidades",       # "Número de unidades" — CSVs reais da BaseLinker
                "esperado", "expected", "quantidade", "quantity", "estoque", "qty",
            )
            log.info(f"[CSV:{warehouse_name}] headers={headers}")
            log.info(f"[CSV:{warehouse_name}] cols → sku:{sku_col} loc:{loc_col} qty:{qty_col}")
            if sku_col is None:
                sku_col = 3  # col 3 no formato real (#;Nome;EAN;SKU;...)
                log.warning(f"[CSV:{warehouse_name}] SKU col não identificada — usando col 3")
            if loc_col is None:
                loc_col = 4  # col 4 no formato real
                log.warning(f"[CSV:{warehouse_name}] Loc col não identificada — usando col 4")
            if qty_col is None:
                qty_col = 5  # col 5 = "Número de unidades"
                log.warning(f"[CSV:{warehouse_name}] Qty col não identificada — usando col 5")
            continue

        max_col = max(c for c in [sku_col, loc_col, qty_col] if c is not None)
        if len(row) <= max_col:
            continue

        sku = str(row[sku_col]).strip()
        if not sku or sku.lower() in ("#", "sku", "cod", "codigo", "ean", "referencia"):
            continue

        loc = str(row[loc_col]).strip() if loc_col is not None else ""
        try:
            qty = int(float(str(row[qty_col]).replace(",", ".").strip() or "0"))
        except (ValueError, IndexError):
            qty = 0

        if qty <= 0:
            continue

        if sku not in result:
            result[sku] = {"picking": 0, "armPad": 0, "arm": 0, "locF": set(), "locG": set()}

        d = result[sku]
        loc_up = loc.upper()

        if is_armazenamento:
            # Armazenamento: todas as locações → arm
            d["arm"] += qty
            if loc:
                d["locG"].add(loc)
        else:
            # Padrão: A8→picking, A9→armPad
            if loc_up.startswith("A8"):
                d["picking"] += qty
                if loc:
                    d["locF"].add(loc)
            elif loc_up.startswith("A9"):
                d["armPad"] += qty
                if loc:
                    d["locG"].add(loc)
            else:
                d["picking"] += qty  # fallback para locações desconhecidas

    log.info(f"[CSV:{warehouse_name}] {len(result)} SKUs parseados")
    return result


# ─── Merge + Google Sheets ────────────────────────────────────────────────────

def merge_data(padrao_data, armazenamento_data, api):
    """
    Combina:
      padrao_data        → picking (A8) e armPad (A9) do armazém Padrão
      armazenamento_data → arm de todas as locações do armazém Armazenamento
      api                → chg (CHEGOU), peso e vol via API BaseLinker
    """
    all_skus = sorted(set(padrao_data) | set(armazenamento_data) | set(api))
    rows = []
    for sku in all_skus:
        p = padrao_data.get(sku, {})
        a = armazenamento_data.get(sku, {})
        q = api.get(sku, {})

        picking = p.get("picking", 0) or 0
        arm_pad = p.get("armPad",  0) or 0
        arm_qty = a.get("arm",     0) or 0
        chg     = q.get("chg",     0) or 0

        # locF: locações A8 do Padrão
        locF = " / ".join(sorted(p.get("locF") or set())) or ""
        # locG: locações A9 do Padrão + locações do Armazenamento
        locG_parts = (p.get("locG") or set()) | (a.get("locG") or set())
        locG = " / ".join(sorted(locG_parts)) or ""

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

    # 2. Exportar CSVs via Playwright (dois armazéns separados)
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
        await stealth_async(page)

        # Tenta injetar cookie de sessão salvo (bypass do Cloudflare)
        session_injected = await bl_inject_session(ctx)
        if session_injected:
            await page.goto(f"{BL_PANEL}/inventory_stocktakes", wait_until="networkidle")
            if BL_LOGIN in page.url or "login" in page.url.lower():
                log.warning("[LOGIN] Cookie expirado — fazendo login via formulário")
                await bl_login(page)
            else:
                log.info(f"[LOGIN] Sessão via cookie OK — {page.url}")
                await screenshot(page, "03_panel_via_cookie")
        else:
            await bl_login(page)

        # Exporta armazém Padrão (picking A8 + armPad A9)
        log.info("[RUN] Exportando armazém Padrão...")
        padrao_csv = await export_inventory_csv(page, "Padrão", tag="padrao")

        # Exporta armazém Armazenamento (arm)
        log.info("[RUN] Exportando armazém Armazenamento...")
        armazenamento_csv = await export_inventory_csv(page, "Armazenamento", tag="arm")

        await browser.close()

    # Salva CSVs brutos para debug
    (DEBUG_DIR / "padrao.csv").write_bytes(padrao_csv.encode("utf-8"))
    (DEBUG_DIR / "armazenamento.csv").write_bytes(armazenamento_csv.encode("utf-8"))

    # 3. Parseia cada CSV com contexto de armazém
    padrao_data       = parse_csv(padrao_csv,       warehouse_name="Padrão")
    armazenamento_data = parse_csv(armazenamento_csv, warehouse_name="Armazenamento")

    # 4. Merge: Padrão (picking/armPad) + Armazenamento (arm) + API (chg/peso/vol)
    rows = merge_data(padrao_data, armazenamento_data, api_data)

    # 5. Grava no Sheets
    write_to_sheets(rows)

    log.info(f"[DONE] {len(rows)} SKUs no SNAPSHOT_BL")


if __name__ == "__main__":
    asyncio.run(run())
