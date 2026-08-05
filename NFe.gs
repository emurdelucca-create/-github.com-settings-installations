// ============================================================
// NF-e IMPORTADOR
// Importa XMLs de NF-e a partir de ZIPs via diálogo HTML.
//
// Fluxo:
//   1. Menu → "📤 Importar XMLs de ZIPs"
//   2. Diálogo carrega apenas fornecedores (rápido, ~2s)
//   3. Usuário sobe ZIPs → browser extrai e parseia XMLs
//   4. Apps Script grava na aba "Dados NF" (col O = Num NF)
//   5. Menu → "🔍 Buscar Canal de Venda"  (ou botão no diálogo)
//      → consulta BaseLinker por número de NF, preenche col M
//
// Pré-requisito: BASELINKER_API_KEY em Extensões → Apps Script
//                → ⚙️ Propriedades do script
// ============================================================

const NFE_CFG = {
  ABA_DADOS: 'Dados NF',
  ID_FORN:   '1VSraBQz0pnXwcCV0QjQmU8vtcIy9UFbDUh4H3ZNYc68',
};

const NFE_BL_URL = 'https://api.baselinker.com/connector.php';

const NFE_CANAL_MAP = {
  'humble':                 'Shopee Humble',
  'humble shopee':          'Shopee Humble',
  'najumi_pecas':           'Shopee Najumi',
  'najumi shopee':          'Shopee Najumi',
  'sky motoparts':          'Shopee Sky',
  'shopee sky':             'Shopee Sky',
  'edmur de lucca junior':  'ML Humble',
  'humble ml':              'ML Humble',
  'mirian aparecida':       'ML Najumi',
  'najumi ml':              'ML Najumi',
  'edmotos ltda':           'ML Edmotos',
  'edmotos ml':             'ML Edmotos',
  'moto vibe':              'ML Moto Vibe',
  'moto vibe ml':           'ML Moto Vibe',
  'sky motoparts ml':       'ML Sky',
};

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📄 NF-e')
    .addItem('📤 Importar XMLs de ZIPs',      'nfe_abrirDialog')
    .addItem('🔍 Buscar Canal de Venda',       'nfe_abrirDialogCanal')
    .addSeparator()
    .addItem('🔬 Diagnóstico BL — ver pedido', 'nfe_diagnosticarPedido')
    .addToUi();
}

function nfe_abrirDialog() {
  const html = HtmlService.createHtmlOutputFromFile('NFeUpload')
    .setWidth(620)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '📄 Importar NF-e de ZIPs');
}

function nfe_abrirDialogCanal() {
  const html = HtmlService.createHtmlOutputFromFile('NFeUpload')
    .setWidth(620)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Buscar Canal de Venda');
}

// ── Pré-carrega apenas fornecedores (rápido) ──────────────────
function nfe_precarregar() {
  try {
    const fornMap = _nfe_loadForn();
    return JSON.stringify({ ok: true, fornMap: fornMap });
  } catch(e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ── Busca canal para todas as NFs já gravadas na planilha ─────
// Lê coluna O (Num NF), consulta BaseLinker, preenche coluna M.
function nfe_buscarCanal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(NFE_CFG.ABA_DADOS);
    if (!ws) return JSON.stringify({ ok: false, error: 'Aba "' + NFE_CFG.ABA_DADOS + '" não encontrada.' });

    const lastRow = ws.getLastRow();
    if (lastRow < 2) return JSON.stringify({ ok: true, filled: 0, total: 0 });

    // Coluna O (índice 15, base-1) = Num NF
    const numNFs = ws.getRange(2, 15, lastRow - 1, 1).getValues();

    const canalMap = _nfe_buildCanalMap();

    var filled = 0;
    var novosCanais = numNFs.map(function(row) {
      const numNf = String(row[0] || '').trim();
      if (!numNf) return [''];
      const parts  = numNf.split('/');
      const canal  = canalMap[numNf] ||
                     (parts.length === 2
                       ? (canalMap[parts[1] + '/' + parts[0]] || canalMap[parts[0]])
                       : '') || '';
      if (canal) filled++;
      return [canal];
    });

    // Coluna M (índice 13, base-1) = Canal
    ws.getRange(2, 13, novosCanais.length, 1).setValues(novosCanais);
    SpreadsheetApp.flush();
    return JSON.stringify({ ok: true, filled: filled, total: lastRow - 1 });
  } catch(e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ── Constrói mapa: invoice_number → canal (via getOrders) ─────
function _nfe_buildCanalMap() {
  const map = {};
  // 2026-05-01T00:00:00Z
  const dateFrom = 1746057600;

  let page = 1;
  while (true) {
    const r = _nfe_bl('getOrders', {
      date_from:       dateFrom,
      get_unconfirmed: true,
      page:            page,
    });
    const orders = r.orders || [];
    if (!orders.length) break;

    orders.forEach(function(o) {
      const inv   = String(o.invoice_fullnumber || o.invoice_number || '').trim();
      const login = String(o.order_source_login || '').trim().toLowerCase();
      if (!inv || !login) return;

      const canal = _nfe_detectCanal(login);
      if (!canal) return;

      map[inv] = canal;
      const parts = inv.split('/');
      if (parts.length === 2) {
        map[parts[1] + '/' + parts[0]] = canal;
        map[parts[0]] = canal;
      }
    });

    if (orders.length < 1000) break;
    page++;
    Utilities.sleep(300);
  }
  return map;
}

function _nfe_detectCanal(login) {
  if (NFE_CANAL_MAP[login]) return NFE_CANAL_MAP[login];
  for (var k in NFE_CANAL_MAP) {
    if (login.startsWith(k) || k.startsWith(login)) return NFE_CANAL_MAP[k];
  }
  return '';
}

// ── Carrega mapa SKU → Fornecedor ─────────────────────────────
function _nfe_loadForn() {
  const ss   = SpreadsheetApp.openById(NFE_CFG.ID_FORN);
  const aba  = ss.getSheets()[0];
  const data = aba.getDataRange().getValues();
  const map  = {};
  for (var i = 1; i < data.length; i++) {
    const sku  = String(data[i][0] || '').trim();
    const forn = String(data[i][2] || '').trim() || 'Sem Fornecedor';
    if (!sku) continue;
    map[sku] = forn;
    const norm = sku.replace(/^0+/, '') || '0';
    if (norm !== sku) map[norm] = forn;
  }
  return map;
}

// ── Recebe lote de linhas do browser e grava na planilha ──────
// rows: 15 colunas (A-O); O = Num NF; M = Canal (vazio aqui)
function nfe_processarLote(rows, isFirst) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws   = ss.getSheetByName(NFE_CFG.ABA_DADOS);

    if (isFirst) {
      if (!ws) ws = ss.insertSheet(NFE_CFG.ABA_DADOS);
      ws.clearContents();
      const hdr = [
        'Data', 'CNPJ/CPF', 'Cliente', 'Cons. Final', 'UF',
        'CFOP', 'NCM', 'Descrição', 'Origem',
        'Qtd', 'Vlr Unit', 'Vlr Total',
        'Canal', 'Fornecedor', 'Num NF',
      ];
      ws.getRange(1, 1, 1, hdr.length).setValues([hdr]);
      ws.getRange(1, 1, 1, hdr.length).setFontWeight('bold');
    } else {
      if (!ws) throw new Error('Aba "' + NFE_CFG.ABA_DADOS + '" não encontrada.');
    }

    if (rows.length > 0) {
      const lastRow = ws.getLastRow();
      ws.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Diagnóstico: mostra campos de um pedido por order_id ──────
function nfe_diagnosticarPedido() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('🔬 Diagnóstico Pedido BL',
    'Digite o order_id do pedido:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const oid = parseInt(res.getResponseText().trim(), 10);
  if (!oid) { ui.alert('ID inválido.'); return; }

  try {
    const r = _nfe_bl('getOrders', { order_id: oid });
    const o = (r.orders || [])[0];
    if (!o) { ui.alert('Pedido não encontrado.'); return; }

    let msg = '══ Pedido ' + oid + ' ══\n';
    msg += 'order_source_login= ' + o.order_source_login + '\n';
    msg += 'invoice_fullnumber= ' + o.invoice_fullnumber + '\n';
    msg += 'invoice_number    = ' + o.invoice_number + '\n';
    msg += 'date_add          = ' + o.date_add + '\n';
    msg += '\nTodos os campos:\n' + Object.keys(o).join(', ');
    ui.alert(msg.substring(0, 1500));
  } catch(e) {
    ui.alert('❌ Erro:\n' + e.message);
  }
}

// ── BaseLinker API ────────────────────────────────────────────
function _nfe_bl(method, params) {
  const key = PropertiesService.getScriptProperties().getProperty('BASELINKER_API_KEY');
  if (!key) throw new Error('BASELINKER_API_KEY não configurada nas propriedades do script.');
  const resp = UrlFetchApp.fetch(NFE_BL_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'token=' + key +
             '&method=' + method +
             '&parameters=' + JSON.stringify(params || {}),
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (data.status !== 'SUCCESS') {
    throw new Error('[BL:' + method + '] ' + (data.error_message || JSON.stringify(data)));
  }
  return data;
}
