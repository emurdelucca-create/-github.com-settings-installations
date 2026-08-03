// ============================================================
// ANÁLISE NAJUMI
// Gera 3 abas: SKU | NCM | Fornecedor
//
// Layout de colunas (aba SKU):
//   A         → SKU
//   B         → NCM (Campos adicionais do BaseLinker)
//   C         → Fornecedor (planilha de fornecedores)
//   D-I (6)   → Vendas Todas as Contas: Qtd/Fat por período
//   J-L (3)   → Curva ABC por período (0-30 / 30-60 / 60-90)
//   M-R (6)   → Shopee Najumi (sem FULL)
//   S-X (6)   → Shopee Najumi FULL
//   Y-AD (6)  → Shopee Najumi (sem FULL + FULL juntos)
//
// Abas NCM e Fornecedor: mesma lógica, agregado por grupo.
//
// Pré-requisito: BASELINKER_API_KEY em Extensões → Apps Script
//                → ⚙️ Propriedades do script
// ============================================================

const NJ = {
  ID_GE:   '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE:  'Dados Completos',

  ID_FORN: '1VSraBQz0pnXwcCV0QjQmU8vtcIy9UFbDUh4H3ZNYc68',

  BL_INV:   39947,
  BL_NCM:   'extra_field_73312',  // NCM em text_fields (identificado via diagnóstico)

  GE_SKU:   1,
  GE_QTD:   2,
  GE_DATA:  3,
  GE_CANAL: 7,
  GE_FAT:   21,

  ABC_A: 0.70,
  ABC_B: 0.90,

  CANAIS_TODOS: [
    'ML Edmotos',    'ML Edmotos FULL',
    'ML Humble',     'ML Humble FULL',
    'ML Najumi',     'ML Najumi FULL',
    'ML Sky',        'ML Sky FULL',
    'ML Moto Vibe',  'ML Moto Vibe FULL',
    'Shopee Humble',
    'Shopee Najumi', 'Shopee Najumi FULL',
    'Shopee Sky',
  ],
  CANAIS_NJ:      ['Shopee Najumi'],
  CANAIS_NJ_FULL: ['Shopee Najumi FULL'],
  CANAIS_NJ_ALL:  ['Shopee Najumi', 'Shopee Najumi FULL'],

  ABA_SKU:  'SKU',
  ABA_NCM:  'NCM',
  ABA_FORN: 'Fornecedor',

  SEM_FORN: 'Sem Fornecedor',
};

const NJ_BL_URL = 'https://api.baselinker.com/connector.php';

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Análise Najumi')
    .addItem('🔄 Gerar Análise Completa', 'nj_gerarAnalise')
    .addSeparator()
    .addItem('🔍 Buscar Campos Extras por SKU', 'nj_diagnosticarCamposExtras')
    .addToUi();
}

// ── Ponto de entrada ──────────────────────────────────────────
function nj_gerarAnalise() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando Análise Najumi... Aguarde 2-3 minutos.');

  try {
    const blData  = _nj_fetchBL();
    const fornMap = _nj_loadForn();
    const vendas  = _nj_loadVendas();

    _nj_calcABC(vendas);

    _nj_writeSkuTab(vendas, blData, fornMap);
    _nj_writeGroupTab(NJ.ABA_NCM,  vendas, blData, fornMap, 'ncm');
    _nj_writeGroupTab(NJ.ABA_FORN, vendas, blData, fornMap, 'forn');

    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    ui.alert('✅ Análise gerada!\n' + Object.keys(vendas).length + ' SKUs — ' + ts);
  } catch(e) {
    ui.alert('❌ Erro:\n' + e.message + '\n\n' + e.stack);
  }
}

// ── BaseLinker API ────────────────────────────────────────────
function _nj_bl(method, params) {
  const key = PropertiesService.getScriptProperties().getProperty('BASELINKER_API_KEY');
  if (!key) throw new Error('BASELINKER_API_KEY não configurada nas propriedades do script.');
  const resp = UrlFetchApp.fetch(NJ_BL_URL, {
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

// ── Busca NCM de todos os produtos via API ───────────────────
function _nj_fetchBL() {
  // 1. Lista todos os produtos (paginado)
  const pidToSku = {};
  let page = 1;
  while (true) {
    const r = _nj_bl('getInventoryProductsList', { inventory_id: NJ.BL_INV, page });
    const entries = Object.entries(r.products || {});
    if (!entries.length) break;
    for (const [pid, info] of entries) {
      const sku = String(info.sku || '').trim();
      if (sku) pidToSku[pid] = sku;
    }
    if (entries.length < 1000) break;
    page++;
    Utilities.sleep(250);
  }

  const pids = Object.keys(pidToSku);

  // 2. Chave do NCM em text_fields (hardcoded; identificado via diagnóstico)
  const ncmKey = NJ.BL_NCM;

  // 3. Dados detalhados: text_fields[ncmKey] → NCM
  const result = {};
  for (let i = 0; i < pids.length; i += 1000) {
    const batch = pids.slice(i, i + 1000).map(Number);
    const r = _nj_bl('getInventoryProductsData', { inventory_id: NJ.BL_INV, products: batch });
    for (const [pid, p] of Object.entries(r.products || {})) {
      const sku = pidToSku[String(pid)];
      if (!sku || result[sku]) continue;
      let ncm = '';
      if (ncmKey) {
        const tf = _nj_parseTextField(p.text_fields);
        ncm = String(tf[ncmKey] || '').trim();
      }
      result[sku] = { ncm };
    }
    if (i + 1000 < pids.length) Utilities.sleep(300);
  }

  return result;
}

// text_fields pode chegar como string JSON ou como objeto
function _nj_parseTextField(tf) {
  if (!tf) return {};
  if (typeof tf === 'object') return tf;
  try { return JSON.parse(tf); } catch(_) { return {}; }
}

// ── Diagnóstico: dump completo de um produto por ID ──────────
// Usa product_id diretamente para ver TODOS os campos retornados pela API.
function nj_diagnosticarCamposExtras() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('🔍 Dump Completo do Produto',
    'Digite o product_id (número) do produto com NCM preenchido.\n\nEx: 83984421', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pid = parseInt(res.getResponseText().trim(), 10);
  if (!pid) { ui.alert('ID inválido.'); return; }

  try {
    const r = _nj_bl('getInventoryProductsData', { inventory_id: NJ.BL_INV, products: [pid] });
    const p = (r.products || {})[String(pid)];
    if (!p) { ui.alert('Produto ' + pid + ' não encontrado.'); return; }

    // Mostra todos os campos de topo e o extra_fields
    let msg = '══ PRODUTO id=' + pid + ' sku="' + p.sku + '" ══\n\n';
    msg += 'Campos de topo:\n';
    for (const [k, v] of Object.entries(p)) {
      if (k === 'extra_fields' || k === 'stock' || k === 'locations' || k === 'images') continue;
      const s = typeof v === 'object' ? JSON.stringify(v) : String(v);
      msg += '  ' + k + ' = ' + s.substring(0, 80) + '\n';
    }
    msg += '\ntext_fields (campos adicionais):\n';
    const tf = _nj_parseTextField(p.text_fields);
    const tfKeys = Object.keys(tf).filter(k => k.startsWith('extra_field_') || k === 'name');
    if (!tfKeys.length) {
      msg += '  (nenhum extra_field encontrado)\n';
    } else {
      for (const k of tfKeys) {
        msg += '  ' + k + ' = "' + String(tf[k] || '').substring(0, 60) + '"\n';
      }
    }
    msg += '\nextra_fields (outro objeto):\n';
    const ef = p.extra_fields || {};
    msg += !Object.keys(ef).length ? '  (vazio)\n'
      : Object.entries(ef).map(([k,v]) => '  ' + k + '="' + String(v).substring(0,40) + '"').join('\n') + '\n';
    msg += '\nstock: ' + JSON.stringify(p.stock || {}).substring(0, 200) + '\n';

    // Grava resultado na aba oculta para análise completa
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName('_nj_debug');
    if (!ws) { ws = ss.insertSheet('_nj_debug'); ws.hideSheet(); }
    ws.clearContents();
    ws.getRange(1, 1).setValue(JSON.stringify(p, null, 2));

    ui.alert(msg.substring(0, 1500) +
      (msg.length > 1500 ? '\n...(ver aba "_nj_debug" para JSON completo)' : '') +
      '\n\nJSON completo gravado em aba "_nj_debug".');
  } catch(e) {
    ui.alert('❌ Erro:\n' + e.message);
  }
}

// ── Fornecedor (planilha externa) ─────────────────────────────
function _nj_loadForn() {
  const ss   = SpreadsheetApp.openById(NJ.ID_FORN);
  const aba  = ss.getSheets()[0];
  const data = aba.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const sku  = String(data[i][0] || '').trim();
    const forn = String(data[i][2] || '').trim() || NJ.SEM_FORN;
    if (!sku) continue;
    if (!map[sku]) map[sku] = forn;
    const norm = sku.replace(/^0+/, '') || '0';
    if (norm !== sku && !map[norm]) map[norm] = forn;
  }
  return map;
}

function _nj_lookupForn(fornMap, sku) {
  if (fornMap[sku]) return fornMap[sku];
  const norm = sku.replace(/^0+/, '') || '0';
  return fornMap[norm] || NJ.SEM_FORN;
}

// ── Vendas do GE Finance ──────────────────────────────────────
function _nj_loadVendas() {
  const geSS  = SpreadsheetApp.openById(NJ.ID_GE);
  const geAba = geSS.getSheetByName(NJ.ABA_GE);
  if (!geAba) throw new Error('Aba "' + NJ.ABA_GE + '" não encontrada em GE Finance.');
  const rows = geAba.getDataRange().getValues();

  // Acha data mais recente → base dos períodos
  let dataMax = new Date(0);
  for (let i = 1; i < rows.length; i++) {
    const d = _nj_parseDate(rows[i][NJ.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(23, 59, 59, 999);

  const todosSet  = new Set(NJ.CANAIS_TODOS.map(c => c.toLowerCase()));
  const njSet     = new Set(NJ.CANAIS_NJ.map(c => c.toLowerCase()));
  const njFulSet  = new Set(NJ.CANAIS_NJ_FULL.map(c => c.toLowerCase()));
  const njAllSet  = new Set(NJ.CANAIS_NJ_ALL.map(c => c.toLowerCase()));

  const vendas = {};

  for (let i = 1; i < rows.length; i++) {
    const sku   = String(rows[i][NJ.GE_SKU]   || '').trim();
    const canal = String(rows[i][NJ.GE_CANAL] || '').trim().toLowerCase();
    const qtd   = Number(rows[i][NJ.GE_QTD]   || 0);
    const fat   = Number(rows[i][NJ.GE_FAT]   || 0);
    if (!sku || !todosSet.has(canal)) continue;

    const d = _nj_parseDate(rows[i][NJ.GE_DATA]);
    if (!d) continue;
    const dias = Math.floor((dataMax - d) / 86400000);
    if (dias < 0 || dias >= 90) continue;

    const suf = dias < 30 ? '030' : dias < 60 ? '3060' : '6090';

    if (!vendas[sku]) vendas[sku] = _nj_novaSku();
    const v = vendas[sku];

    v.todos['q' + suf] += qtd;
    v.todos['f' + suf] += fat;
    if (njSet.has(canal))    { v.nj['q'     + suf] += qtd; v.nj['f'     + suf] += fat; }
    if (njFulSet.has(canal)) { v.njFull['q' + suf] += qtd; v.njFull['f' + suf] += fat; }
    if (njAllSet.has(canal)) { v.njAll['q'  + suf] += qtd; v.njAll['f'  + suf] += fat; }
  }

  return vendas;
}

function _nj_novaSku() {
  const blk = () => ({ q030:0, f030:0, q3060:0, f3060:0, q6090:0, f6090:0 });
  return { todos:blk(), nj:blk(), njFull:blk(), njAll:blk(),
           abc030:'', abc3060:'', abc6090:'' };
}

// ── Curva ABC (Fat de todas as contas, por período) ───────────
function _nj_calcABC(vendas) {
  const skus = Object.keys(vendas);
  [['030','f030'], ['3060','f3060'], ['6090','f6090']].forEach(([per, fKey]) => {
    const sorted = [...skus].sort((a, b) => vendas[b].todos[fKey] - vendas[a].todos[fKey]);
    const total  = sorted.reduce((s, sk) => s + vendas[sk].todos[fKey], 0);
    let cum = 0;
    sorted.forEach(sk => {
      cum += vendas[sk].todos[fKey];
      const pct = total > 0 ? cum / total : 1;
      vendas[sk]['abc' + per] = pct <= NJ.ABC_A ? 'A' : pct <= NJ.ABC_B ? 'B' : 'C';
    });
  });
}

// ── Helpers de linha ──────────────────────────────────────────
function _nj_buildRow(prefix, v) {
  return [
    ...prefix,
    // Todas as contas (D-I para SKU / B-G para grupos)
    v.todos.q030, v.todos.f030, v.todos.q3060, v.todos.f3060, v.todos.q6090, v.todos.f6090,
    // ABC (J-L para SKU / H-J para grupos)
    v.abc030, v.abc3060, v.abc6090,
    // Shopee Najumi sem FULL (M-R para SKU / K-P para grupos)
    v.nj.q030, v.nj.f030, v.nj.q3060, v.nj.f3060, v.nj.q6090, v.nj.f6090,
    // Shopee Najumi FULL (S-X para SKU / Q-V para grupos)
    v.njFull.q030, v.njFull.f030, v.njFull.q3060, v.njFull.f3060, v.njFull.q6090, v.njFull.f6090,
    // Shopee Najumi All (Y-AD para SKU / W-AB para grupos)
    v.njAll.q030, v.njAll.f030, v.njAll.q3060, v.njAll.f3060, v.njAll.q6090, v.njAll.f6090,
  ];
}

// ── Aba SKU ───────────────────────────────────────────────────
const NJ_HDR_SKU = [
  'SKU', 'NCM', 'Fornecedor',
  'Qtd 0-30',  'Fat 0-30',  'Qtd 30-60', 'Fat 30-60', 'Qtd 60-90', 'Fat 60-90',
  'ABC 0-30',  'ABC 30-60', 'ABC 60-90',
  'NJ Qtd 0-30',     'NJ Fat 0-30',     'NJ Qtd 30-60',     'NJ Fat 30-60',     'NJ Qtd 60-90',     'NJ Fat 60-90',
  'NJ Full Qtd 0-30','NJ Full Fat 0-30','NJ Full Qtd 30-60','NJ Full Fat 30-60','NJ Full Qtd 60-90','NJ Full Fat 60-90',
  'NJ All Qtd 0-30', 'NJ All Fat 0-30', 'NJ All Qtd 30-60', 'NJ All Fat 30-60', 'NJ All Qtd 60-90', 'NJ All Fat 60-90',
];

function _nj_writeSkuTab(vendas, blData, fornMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(NJ.ABA_SKU);
  if (!ws) ws = ss.insertSheet(NJ.ABA_SKU);
  ws.clearContents();

  const rows = Object.keys(vendas).sort().map(sku => {
    const ncm  = (blData[sku] && blData[sku].ncm) || '';
    const forn = _nj_lookupForn(fornMap, sku);
    return _nj_buildRow([sku, ncm, forn], vendas[sku]);
  });

  const all = [NJ_HDR_SKU, ...rows];
  ws.getRange(1, 1, all.length, NJ_HDR_SKU.length).setValues(all);
  if (rows.length) ws.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  SpreadsheetApp.flush();
}

// ── Abas NCM e Fornecedor ─────────────────────────────────────
function _nj_writeGroupTab(abaName, vendas, blData, fornMap, groupBy) {
  // Agrega por grupo
  const groups = {};
  for (const sku of Object.keys(vendas)) {
    const v   = vendas[sku];
    const key = groupBy === 'ncm'
      ? ((blData[sku] && blData[sku].ncm) || 'Sem NCM')
      : _nj_lookupForn(fornMap, sku);

    if (!groups[key]) groups[key] = _nj_novaSku();
    const g = groups[key];
    for (const blk of ['todos','nj','njFull','njAll']) {
      for (const m of ['q030','f030','q3060','f3060','q6090','f6090']) {
        g[blk][m] += v[blk][m];
      }
    }
  }

  _nj_calcABC(groups);

  const label = groupBy === 'ncm' ? 'NCM' : 'Fornecedor';
  const hdr   = [
    label,
    'Qtd 0-30',  'Fat 0-30',  'Qtd 30-60', 'Fat 30-60', 'Qtd 60-90', 'Fat 60-90',
    'ABC 0-30',  'ABC 30-60', 'ABC 60-90',
    'NJ Qtd 0-30',     'NJ Fat 0-30',     'NJ Qtd 30-60',     'NJ Fat 30-60',     'NJ Qtd 60-90',     'NJ Fat 60-90',
    'NJ Full Qtd 0-30','NJ Full Fat 0-30','NJ Full Qtd 30-60','NJ Full Fat 30-60','NJ Full Qtd 60-90','NJ Full Fat 60-90',
    'NJ All Qtd 0-30', 'NJ All Fat 0-30', 'NJ All Qtd 30-60', 'NJ All Fat 30-60', 'NJ All Qtd 60-90', 'NJ All Fat 60-90',
  ];

  const rows = Object.keys(groups).sort().map(key => _nj_buildRow([key], groups[key]));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(abaName);
  if (!ws) ws = ss.insertSheet(abaName);
  ws.clearContents();

  const all = [hdr, ...rows];
  ws.getRange(1, 1, all.length, hdr.length).setValues(all);
  SpreadsheetApp.flush();
}

// ── Parse DD/MM/YYYY (mesmo padrão do BaseCompras) ───────────
function _nj_parseDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
