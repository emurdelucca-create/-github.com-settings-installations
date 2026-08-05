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
  BL_INV:    39947,        // ID do inventário BaseLinker
  BL_ORIG:   '',           // ID do campo "Código da Origem" (ex: 'extra_field_73315')
                           // Deixe vazio até rodar o diagnóstico abaixo.
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
    .addItem('📤 Importar XMLs de ZIPs',         'nfe_abrirDialog')
    .addItem('🔍 Buscar Canal de Venda',          'nfe_buscarCanal')
    .addSeparator()
    .addItem('🔬 Diagnóstico BL — ver pedido',    'nfe_diagnosticarPedido')
    .addItem('🔬 Diagnóstico BL — campo Origem',  'nfe_diagnosticarOrigem')
    .addItem('🔬 Diagnóstico BL — primeiros pedidos', 'nfe_debugOrders')
    .addToUi();
}

function nfe_abrirDialog() {
  const html = HtmlService.createHtmlOutputFromFile('NFeUpload')
    .setWidth(620)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '📄 Importar NF-e de ZIPs');
}

// ── Pré-carrega fornecedores + mapa de origem ─────────────────
function nfe_precarregar() {
  try {
    const fornMap = _nfe_loadForn();
    const origMap = NFE_CFG.BL_ORIG ? _nfe_loadOrigMap() : {};
    return JSON.stringify({ ok: true, fornMap: fornMap, origMap: origMap });
  } catch(e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// ── Carrega mapa SKU → Código de Origem (via BL) ──────────────
function _nfe_loadOrigMap() {
  const map  = {};
  const inv  = NFE_CFG.BL_INV;
  const fOrig = NFE_CFG.BL_ORIG;

  function parseTextField(tf) {
    if (!tf) return {};
    if (typeof tf === 'object') return tf;
    try { return JSON.parse(tf); } catch(e) { return {}; }
  }

  let page = 1;
  while (true) {
    const r1 = _nfe_bl('getInventoryProductsList', {
      inventory_id: inv, page: page, filter_limit: 1000,
    });
    const prods = r1.products || {};
    const ids   = Object.keys(prods);
    if (!ids.length) break;

    // busca em lotes de 1000
    for (var i = 0; i < ids.length; i += 1000) {
      const batch = ids.slice(i, i + 1000);
      const r2 = _nfe_bl('getInventoryProductsData', {
        inventory_id: inv, products: batch,
      });
      Object.values(r2.products || {}).forEach(function(p) {
        const tf   = parseTextField(p.text_fields);
        const orig = String(tf[fOrig] || '').trim();
        const sku  = String(p.sku || '').trim();
        if (!sku || orig === '') return;
        map[sku] = orig;
        const norm = sku.replace(/^0+/, '') || '0';
        if (norm !== sku) map[norm] = orig;
      });
    }

    if (ids.length < 1000) break;
    page++;
    Utilities.sleep(200);
  }
  return map;
}

// ── Diagnóstico: descobre o campo de Origem num produto por SKU ─
function nfe_diagnosticarOrigem() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('🔬 Diagnóstico Origem BL',
    'Digite o SKU do produto (ex: 04426-S):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const sku = res.getResponseText().trim();
  if (!sku) return;

  try {
    const r1 = _nfe_bl('getInventoryProductsList', {
      inventory_id: NFE_CFG.BL_INV, filter_sku: sku,
    });
    const ids = Object.keys(r1.products || {});
    if (!ids.length) { ui.alert('Produto não encontrado.'); return; }

    const r2 = _nfe_bl('getInventoryProductsData', {
      inventory_id: NFE_CFG.BL_INV, products: [ids[0]],
    });
    const p  = Object.values(r2.products || {})[0];
    if (!p) { ui.alert('Dados não encontrados.'); return; }

    let tf;
    try { tf = typeof p.text_fields === 'string' ? JSON.parse(p.text_fields) : (p.text_fields || {}); }
    catch(e) { tf = {}; }

    let msg = '══ ' + sku + ' (id=' + ids[0] + ') ══\n';
    msg += 'sku = ' + p.sku + '\n\ntext_fields:\n';
    Object.keys(tf).forEach(function(k) {
      if (tf[k]) msg += '  ' + k + ' = ' + tf[k] + '\n';
    });
    ui.alert(msg.substring(0, 1500));
  } catch(e) {
    ui.alert('❌ Erro:\n' + e.message);
  }
}

// ── Busca canal para todas as NFs já gravadas na planilha ─────
// Menu → "🔍 Buscar Canal de Venda"
// Lê coluna O (Num NF), consulta BaseLinker por todos os pedidos,
// cruza pelo número da NF e preenche coluna M (Canal).
function nfe_buscarCanal() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(NFE_CFG.ABA_DADOS);

  if (!ws) {
    ui.alert('❌ Aba "' + NFE_CFG.ABA_DADOS + '" não encontrada.\nImporte os XMLs primeiro.');
    return;
  }

  const lastRow = ws.getLastRow();
  if (lastRow < 2) { ui.alert('Nenhuma linha encontrada na aba "' + NFE_CFG.ABA_DADOS + '".'); return; }

  ss.toast('Consultando BaseLinker… pode levar 30–90 s.', '🔍 Buscar Canal', 120);

  try {
    // Coluna O (15, base-1) = Num NF
    const numNFs = ws.getRange(2, 15, lastRow - 1, 1).getValues();

    const canalMap = _nfe_buildCanalMap();
    const total    = Object.keys(canalMap).length;

    var filled = 0;
    var novosCanais = numNFs.map(function(row) {
      const numNf = String(row[0] || '').trim();
      if (!numNf) return [''];
      const parts = numNf.split('/');
      const canal = canalMap[numNf] ||
                    (parts.length === 2
                      ? (canalMap[parts[1] + '/' + parts[0]] || canalMap[parts[0]])
                      : '') || '';
      if (canal) filled++;
      return [canal];
    });

    // Coluna M (13, base-1) = Canal
    ws.getRange(2, 13, novosCanais.length, 1).setValues(novosCanais);
    SpreadsheetApp.flush();

    ui.alert(
      '✅ Canal preenchido!\n\n' +
      'NFs encontradas na BaseLinker: ' + total + '\n' +
      'Linhas preenchidas: ' + filled + ' de ' + (lastRow - 1) + '\n' +
      (filled < lastRow - 1
        ? '\n⚠ ' + (lastRow - 1 - filled) + ' linhas sem correspondência.\n' +
          'Use "Diagnóstico BL" para verificar se o nº da NF está armazenado\n' +
          'em extra_field_1 ou extra_field_2 do pedido.'
        : '')
    );
  } catch(e) {
    ui.alert('❌ Erro:\n' + e.message);
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
    // A API pode retornar orders como array ou como objeto {id: {...}}
    const raw    = r.orders || [];
    const orders = Array.isArray(raw) ? raw : Object.values(raw);
    if (!orders.length) break;

    orders.forEach(function(o) {
      // Tenta order_source_login e order_source_info como fonte do canal
      const login = (
        String(o.order_source_login || '').trim() ||
        String(o.order_source_info  || '').trim()
      ).toLowerCase();
      if (!login) return;
      const canal = _nfe_detectCanal(login);
      if (!canal) return;

      // Tenta todos os campos candidatos que podem conter o nº da NF
      const candidatos = [
        o.invoice_fullnumber,
        o.invoice_number,
        o.extra_field_1,
        o.extra_field_2,
      ];

      candidatos.forEach(function(raw) {
        const inv = String(raw || '').trim();
        if (!inv) return;
        map[inv] = canal;
        const parts = inv.split('/');
        if (parts.length === 2) {
          map[parts[1] + '/' + parts[0]] = canal;
          map[parts[0]] = canal;
        }
      });
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
// rows: 17 colunas (A-Q)
//   M = Canal (Shopee Humble, ML Najumi…)
//   O = Num NF  |  P = ID Canal (idCadIntTran)  |  Q = Emitente
function nfe_processarLote(rows, isFirst) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws   = ss.getSheetByName(NFE_CFG.ABA_DADOS);

    const hdr = [
      'Data', 'CNPJ/CPF', 'Cliente', 'Cons. Final', 'UF',
      'CFOP', 'NCM', 'Descrição', 'Origem',
      'Qtd', 'Vlr Unit', 'Vlr Total',
      'Canal', 'Fornecedor', 'Num NF', 'ID Canal', 'Emitente', 'SKU',
    ];

    if (isFirst) {
      // Regulares: recria a aba do zero
      if (!ws) ws = ss.insertSheet(NFE_CFG.ABA_DADOS);
      ws.clearContents();
      ws.getRange(1, 1, 1, hdr.length).setValues([hdr]);
      ws.getRange(1, 1, 1, hdr.length).setFontWeight('bold');
      // Coluna R (18) = SKU: força formato texto para preservar zeros à esquerda
      ws.getRange(1, 18, ws.getMaxRows(), 1).setNumberFormat('@');
    } else {
      // Full (append): cria aba com header se não existir ainda
      if (!ws) {
        ws = ss.insertSheet(NFE_CFG.ABA_DADOS);
        ws.getRange(1, 1, 1, hdr.length).setValues([hdr]);
        ws.getRange(1, 1, 1, hdr.length).setFontWeight('bold');
        ws.getRange(1, 18, ws.getMaxRows(), 1).setNumberFormat('@');
      }
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

// ── Diagnóstico: mostra estrutura dos primeiros pedidos ───────
// Ajuda a descobrir onde está o nº da NF e o order_source_login.
function nfe_debugOrders() {
  const ui = SpreadsheetApp.getUi();
  try {
    const dateFrom = 1746057600; // 2026-05-01
    const r   = _nfe_bl('getOrders', { date_from: dateFrom, get_unconfirmed: true, page: 1 });
    const raw = r.orders || {};
    const arr = Array.isArray(raw) ? raw : Object.values(raw);

    let msg = 'Total pedidos 1ª pág: ' + arr.length + '\n';
    msg += 'Tipo de r.orders: ' + (Array.isArray(raw) ? 'array' : typeof raw) + '\n\n';

    const amostra = arr.slice(0, 3);
    amostra.forEach(function(o, i) {
      msg += '── Pedido ' + (i+1) + ' (id=' + o.order_id + ') ──\n';
      msg += 'order_source_login = ' + o.order_source_login + '\n';
      msg += 'invoice_fullnumber = ' + o.invoice_fullnumber + '\n';
      msg += 'invoice_number     = ' + o.invoice_number + '\n';
      msg += 'extra_field_1      = ' + o.extra_field_1 + '\n';
      msg += 'extra_field_2      = ' + o.extra_field_2 + '\n';
      msg += 'date_add           = ' + o.date_add + '\n\n';
    });

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
