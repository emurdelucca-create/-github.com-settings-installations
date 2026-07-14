// ============================================================
// SNAPSHOT_BL — Atualização de estoque via upload manual de CSV
//
// Fluxo:
//   1. Menu BaseLinker → "📦 Atualizar SNAPSHOT_BL..."
//   2. Usuário faz upload dos CSVs exportados manualmente do BaseLinker:
//      • Armazém Padrão  (picking A8, armPad A9)
//      • Armazém Armazenamento  (arm — todas as locações)
//   3. Script busca estoque CHEGOU + dimensões via API BaseLinker
//   4. Faz merge e grava na aba SNAPSHOT_BL
//
// Pré-requisito: BASELINKER_API_KEY em Extensões → Apps Script
//                → ⚙️ Propriedades do script
//
// SKU matching: o CSV exportado pelo BaseLinker pode perder zeros
// à esquerda (e.g. "00118" vira "118" quando aberto no Excel).
// O script normaliza ambos os lados (remove zeros à esquerda)
// para fazer o match e usa o SKU canônico da API na planilha.
// ============================================================

const SNAPSHOT_TAB   = 'SNAPSHOT_BL';
const BL_INVENTORY   = 39947;
const BL_WH_CHEGOU   = 'bl_51442';

// ── Dialog (chamado por me_atualizarSnapshotEstoque em MovimentacaoEstoque.gs) ──

function showSnapshotUpload() {
  const html = HtmlService.createHtmlOutputFromFile('SnapshotBLUpload')
    .setWidth(520)
    .setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, '📦 Atualizar Estoque — SNAPSHOT_BL');
}

// ── Ponto de entrada chamado pelo dialog ─────────────────────

function snapshotProcessCsvs(padraoContent, armContent) {
  try {
    // 1. Busca CHEGOU + dimensões via API
    const apiData = _snapshotFetchApi();

    // 2. Mapa de normalização: sku-sem-zeros → sku-canônico (formato BaseLinker)
    const skuMap = _buildSkuMap(apiData);

    // 3. Parseia os dois CSVs
    const padraoData = _parseCsv(padraoContent, false, skuMap);
    const armData    = _parseCsv(armContent,    true,  skuMap);

    // 4. Merge
    const rows = _mergeData(padraoData, armData, apiData);

    // 5. Grava
    _writeSheet(rows);

    return { ok: true, count: rows.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── BaseLinker API ───────────────────────────────────────────

function _snapshotFetchApi() {
  // Mapeia productId → SKU canônico
  const pidToSku = {};
  let page = 1;
  while (true) {
    const r = _bl_call('getInventoryProductsList', { inventory_id: BL_INVENTORY, page: page });
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

  // Busca estoque + dimensões em lotes de 1000
  const result = {};
  const pids = Object.keys(pidToSku);
  for (let i = 0; i < pids.length; i += 1000) {
    const batch = pids.slice(i, i + 1000).map(p => parseInt(p, 10));
    const r = _bl_call('getInventoryProductsData', { inventory_id: BL_INVENTORY, products: batch });
    for (const [pid, p] of Object.entries(r.products || {})) {
      const sku = pidToSku[String(pid)];
      if (!sku) continue;
      const chg  = parseInt((p.stock || {})[BL_WH_CHEGOU] || 0, 10) || 0;
      const h    = parseFloat(p.height || 0);
      const w    = parseFloat(p.width  || 0);
      const c    = parseFloat(p.length || 0);
      const peso = parseFloat(p.weight || 0);
      const vol  = (h && w && c) ? Math.round(h * w * c * 100) / 100 : 0;
      if (!result[sku]) result[sku] = { chg: 0, peso: 0, vol: 0 };
      result[sku].chg += chg;
      if (peso > 0 && !result[sku].peso) { result[sku].peso = peso; result[sku].vol = vol; }
    }
    if (i + 1000 < pids.length) Utilities.sleep(350);
  }

  return result;
}

// ── SKU matching ─────────────────────────────────────────────

function _normSku(sku) {
  // Remove zeros à esquerda para comparação: "00118" → "118"
  const s = String(sku).trim().replace(/^0+/, '');
  return s || '0';
}

function _buildSkuMap(apiData) {
  // { "118": "00118", ... }
  const map = {};
  for (const sku of Object.keys(apiData)) {
    const n = _normSku(sku);
    if (!map[n]) map[n] = sku;
  }
  return map;
}

// ── CSV Parsing ───────────────────────────────────────────────

function _detectDelim(sample) {
  return sample.split(';').length >= sample.split(',').length ? ';' : ',';
}

function _parseLine(line, delim) {
  const cells = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { inQ = !inQ; }
    else if (ch === delim && !inQ) { cells.push(cur.trim()); cur = ''; }
    else { cur += ch; }
  }
  cells.push(cur.trim());
  return cells;
}

function _findCol(headers, candidates) {
  const hl = headers.map(h => h.toLowerCase().trim());
  for (const c of candidates) {
    const idx = hl.findIndex(h => h.includes(c));
    if (idx >= 0) return idx;
  }
  return null;
}

function _parseCsv(content, isArmazenamento, skuMap) {
  // Remove BOM se presente
  const text = content.replace(/^﻿/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = text.split('\n');
  const delim = _detectDelim(lines.slice(0, 5).join('\n'));

  const result = {};
  let headers = null, skuCol = null, locCol = null, qtyCol = null;

  for (const line of lines) {
    if (!line.trim()) continue;
    const row = _parseLine(line, delim);

    if (headers === null) {
      headers = row;
      skuCol = _findCol(headers, ['sku', 'cod', 'codigo', 'code', 'referencia']);
      locCol = _findCol(headers, ['locali', 'location', 'posicao', 'posição', 'prateleira', 'shelf']);
      qtyCol = _findCol(headers, ['unidades', 'esperado', 'expected', 'quantidade', 'quantity', 'estoque', 'qty']);
      if (skuCol === null) skuCol = 3;
      if (locCol === null) locCol = 4;
      if (qtyCol === null) qtyCol = 5;
      continue;
    }

    const maxCol = Math.max(skuCol, locCol, qtyCol);
    if (row.length <= maxCol) continue;

    const rawSku = String(row[skuCol] || '').trim();
    if (!rawSku || ['#', 'sku', 'cod', 'ean'].includes(rawSku.toLowerCase())) continue;

    const loc = String(row[locCol] || '').trim();
    const qty = parseInt(parseFloat(String(row[qtyCol] || '').replace(',', '.')) || 0, 10);
    if (qty <= 0) continue;

    // Resolve para SKU canônico via mapa de normalização
    const canonical = skuMap[_normSku(rawSku)] || rawSku;

    if (!result[canonical]) {
      result[canonical] = { picking: 0, armPad: 0, arm: 0, locF: [], locG: [] };
    }
    const d = result[canonical];
    const locUp = loc.toUpperCase();

    if (isArmazenamento) {
      d.arm += qty;
      if (loc && !d.locG.includes(loc)) d.locG.push(loc);
    } else {
      if (locUp.startsWith('A8')) {
        d.picking += qty;
        if (loc && !d.locF.includes(loc)) d.locF.push(loc);
      } else if (locUp.startsWith('A9')) {
        d.armPad += qty;
        if (loc && !d.locG.includes(loc)) d.locG.push(loc);
      } else {
        d.picking += qty;
      }
    }
  }

  return result;
}

// ── Merge ────────────────────────────────────────────────────

function _mergeData(padraoData, armData, apiData) {
  const skus = [...new Set([
    ...Object.keys(padraoData),
    ...Object.keys(armData),
    ...Object.keys(apiData),
  ])].sort();

  return skus.map(sku => {
    const p = padraoData[sku] || {};
    const a = armData[sku]    || {};
    const q = apiData[sku]    || {};

    const locFArr = (p.locF || []).sort();
    const locGArr = [...new Set([...(p.locG || []), ...(a.locG || [])])].sort();

    return [
      sku,
      p.picking || 0,
      p.armPad  || 0,
      a.arm     || 0,
      q.chg     || 0,
      locFArr.join(' / '),
      locGArr.join(' / '),
      q.peso    || 0,
      q.vol     || 0,
    ];
  });
}

// ── Write to sheet ───────────────────────────────────────────

function _writeSheet(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(SNAPSHOT_TAB);
  if (!ws) ws = ss.insertSheet(SNAPSHOT_TAB);

  ws.clearContents();
  ws.clearFormats();

  const ts  = Utilities.formatDate(new Date(), 'UTC', 'dd/MM/yyyy HH:mm:ss') + ' UTC';
  const hdr = ['SKU', 'picking', 'armPad', 'arm', 'chg', 'locF', 'locG', 'peso', 'vol'];
  const all = [
    ['Atualizado: ' + ts + '  |  SKUs: ' + rows.length, '', '', '', '', '', '', '', ''],
    hdr,
    ...rows,
  ];

  ws.getRange(1, 1, all.length, 9).setValues(all);

  // Coluna A como texto puro para preservar zeros à esquerda
  if (rows.length > 0) {
    const skuRange = ws.getRange(3, 1, rows.length, 1);
    skuRange.setNumberFormat('@STRING@');
    skuRange.setValues(rows.map(r => [String(r[0])]));
  }
}
