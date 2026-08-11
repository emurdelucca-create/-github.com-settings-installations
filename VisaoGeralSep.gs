// ============================================================
// VISÃO GERAL SEPARAÇÃO
// Preenche abas "Datas" e "SKUs" com pedidos dos 8 status
// configurados em VG_STATUS_ALVO, usando date_confirmed.
//
// Pré-requisito: BASELINKER_API_KEY em
//   Extensões → Apps Script → ⚙ Propriedades do script
// ============================================================

const VG_BL_URL = 'https://api.baselinker.com/connector.php';

const VG_STATUS_ALVO = [
  'NF Emitida',
  'Erro NF',
  'Erro Etiqueta',
  'Movimentação',
  '[SEP] ML Flex',
  '[SEP] ML Agência',
  '[SEP] Shopee Direta',
  '[SEP] Shopee Xpress',
];

// Paleta dark dashboard
const VG_T = {
  bgBase:    '#0d1117',
  bgCard:    '#161b22',
  bgHeader:  '#1c2333',
  bgAlt:     '#131920',
  accent:    '#1f6feb',
  accentSub: '#388bfd',
  green:     '#3fb950',
  amber:     '#d29922',
  red:       '#f85149',
  cyan:      '#39d353',
  textMain:  '#e6edf3',
  textSub:   '#8b949e',
  textDim:   '#484f58',
  border:    '#30363d',
};

// ── Web App ───────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('VisaoGeralDashboard')
    .setTitle('Visão Geral — Separação')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Lê as abas já preenchidas e retorna JSON para o dashboard HTML
function vg_getDados() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const tz   = ss.getSpreadsheetTimeZone();
  const hoje = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  const dados = { hoje, timestamp: '', datas: [], status: [], skus: [] };

  const abaDatas = ss.getSheetByName('Datas');
  if (abaDatas) {
    const ts = String(abaDatas.getRange('A1').getValue());
    dados.timestamp = ts.replace(/.*atualização:\s*/i, '').trim();

    const lastRow = abaDatas.getLastRow();
    if (lastRow >= 3) {
      abaDatas.getRange(3, 1, lastRow - 2, 2).getValues().forEach(([d, p]) => {
        if (!d) return;
        dados.datas.push({ data: Utilities.formatDate(new Date(d), tz, 'dd/MM/yyyy'), pedidos: Number(p) || 0 });
      });
    }

    abaDatas.getRange(3, 4, VG_STATUS_ALVO.length, 2).getValues().forEach(([nome, dataAntiga]) => {
      if (!nome) return;
      dados.status.push({
        nome: String(nome),
        dataAntiga: dataAntiga ? Utilities.formatDate(new Date(dataAntiga), tz, 'dd/MM/yyyy') : '',
      });
    });
  }

  const abaSKUs = ss.getSheetByName('SKUs');
  if (abaSKUs) {
    const STEP    = 3;
    const lastCol = abaSKUs.getLastColumn();
    const lastRowS = abaSKUs.getLastRow();
    if (lastCol >= 1 && lastRowS >= 2) {
      const row2    = abaSKUs.getRange(2, 1, 1, lastCol).getValues()[0];
      const nRows   = Math.min(12, Math.max(0, lastRowS - 3));
      const skuData = nRows > 0 ? abaSKUs.getRange(4, 1, nRows, lastCol).getValues() : [];
      const blocks  = [];
      for (let c = 0; c < lastCol; c += STEP) {
        if (!row2[c]) continue;
        const dataStr = Utilities.formatDate(new Date(row2[c]), tz, 'dd/MM/yyyy');
        const itens   = skuData.map(row => ({ sku: String(row[c] || ''), qty: Number(row[c + 1]) || 0 }))
                                .filter(r => r.sku);
        blocks.push({ data: dataStr, itens });
      }
      dados.skus = blocks.slice(-4).reverse(); // 4 datas mais recentes, mais nova primeiro
    }
  }

  return dados;
}

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Visão Geral')
    .addItem('🔄 Atualizar Datas e SKUs', 'vg_atualizar')
    .addSeparator()
    .addItem('⏱ Configurar atualização automática (5 min)', 'vg_configurarGatilho')
    .addItem('🗑 Remover atualização automática', 'vg_removerGatilho')
    .addToUi();
}

// ── Chamada à API BaseLinker ──────────────────────────────────
function vg_bl(method, params) {
  const token = PropertiesService.getScriptProperties().getProperty('BASELINKER_API_KEY');
  if (!token) throw new Error(
    'Configure BASELINKER_API_KEY em Extensões → Apps Script → ⚙ Propriedades do script.'
  );
  const resp = UrlFetchApp.fetch(VG_BL_URL, {
    method: 'post',
    payload: { token, method, parameters: JSON.stringify(params || {}) },
    muteHttpExceptions: true,
  });
  const d = JSON.parse(resp.getContentText());
  if (d.status !== 'SUCCESS') {
    throw new Error('[' + method + '] ' + (d.error_message || JSON.stringify(d)));
  }
  return d;
}

// ── Gatilho automático ─────────────────────────────────────────
function vg_atualizar_automatico() {
  vg_atualizar(true);
}

function vg_configurarGatilho() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'vg_atualizar_automatico') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('vg_atualizar_automatico').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('✅ Atualização automática configurada a cada 5 minutos.');
}

function vg_removerGatilho() {
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'vg_atualizar_automatico') { ScriptApp.deleteTrigger(t); removidos++; }
  });
  SpreadsheetApp.getUi().alert(
    removidos > 0 ? '✅ Atualização automática removida.' : 'Nenhum gatilho automático encontrado.'
  );
}

// ── Função principal ──────────────────────────────────────────
function vg_atualizar(silencioso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const ui = silencioso ? null : SpreadsheetApp.getUi();

  try {
    const statusResp = vg_bl('getOrderStatusList', {});
    const nameToId   = {};
    (statusResp.statuses || []).forEach(s => { nameToId[s.name] = s.id; });

    const encontrados    = VG_STATUS_ALVO.filter(n =>  nameToId[n]);
    const naoEncontrados = VG_STATUS_ALVO.filter(n => !nameToId[n]);

    if (!encontrados.length) {
      if (ui) ui.alert('❌ Nenhum dos status configurados foi encontrado no BaseLinker.');
      return;
    }
    if (naoEncontrados.length) {
      ss.toast('Aviso — status não encontrados: ' + naoEncontrados.join(', '), '⚠', 10);
    }

    ss.toast('Buscando pedidos em ' + encontrados.length + ' status…', '📊 Visão Geral', 600);

    const dateMap      = {};
    const statusOldest = {};

    for (const nome of encontrados) {
      const sid  = nameToId[nome];
      let idFrom = 0;
      while (true) {
        const r     = vg_bl('getOrders', { status_id: sid, id_from: idFrom });
        const batch = r.orders || [];
        if (!batch.length) break;
        for (const pedido of batch) {
          const ts = pedido.date_confirmed;
          if (!ts) continue;
          if (!statusOldest[nome] || ts < statusOldest[nome]) statusOldest[nome] = ts;
          const ds = Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd');
          if (!dateMap[ds]) dateMap[ds] = { count: 0, skus: {} };
          dateMap[ds].count++;
          for (const prod of (pedido.products || [])) {
            const sku = String(prod.sku || '').trim();
            if (!sku) continue;
            const qty = Number(prod.quantity) || 0;
            dateMap[ds].skus[sku] = (dateMap[ds].skus[sku] || 0) + qty;
          }
        }
        if (batch.length < 100) break;
        idFrom = batch[batch.length - 1].order_id;
      }
    }

    const sortedDates = Object.keys(dateMap).sort();
    if (!sortedDates.length) {
      if (ui) ui.alert('Nenhum pedido com date_confirmed encontrado nos status selecionados.');
      return;
    }

    const agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
    _vg_escreverDatas(ss, tz, sortedDates, dateMap, statusOldest, agora);
    _vg_escreverSKUs(ss, tz, sortedDates, dateMap, agora);

    SpreadsheetApp.flush();
    if (ui) ui.alert('✅ Concluído!\n' + sortedDates.length + ' datas encontradas\n' + encontrados.length + ' status analisados');

  } catch (e) {
    if (ui) ui.alert('❌ Erro: ' + e.message);
    else ss.toast('❌ Erro na atualização automática: ' + e.message, '📊 Visão Geral', 30);
  }
}

// ── Escrever aba "Datas" ──────────────────────────────────────
function _vg_escreverDatas(ss, tz, sortedDates, dateMap, statusOldest, agora) {
  let aba = ss.getSheetByName('Datas');
  if (!aba) aba = ss.insertSheet('Datas');
  aba.clearContents();
  aba.clearFormats();
  aba.setTabColor(VG_T.accent);

  const nDatas  = sortedDates.length;
  const nStatus = VG_STATUS_ALVO.length;
  const totalR  = Math.max(nDatas, nStatus) + 6;

  // ── Fundo geral ──
  aba.getRange(1, 1, totalR, 6)
    .setBackground(VG_T.bgBase)
    .setFontColor(VG_T.textMain)
    .setFontFamily('Arial')
    .setFontSize(11);

  // ── Linha 1: banner de atualização ──
  aba.getRange(1, 1, 1, 6).setBackground(VG_T.bgCard);
  aba.getRange('A1:E1').merge()
    .setValue('  ⏱  Última atualização: ' + agora)
    .setFontColor(VG_T.textSub)
    .setFontStyle('italic')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  aba.setRowHeight(1, 28);

  // ── Linha 2: cabeçalhos ──
  aba.getRange(2, 1, 1, 6).setBackground(VG_T.bgHeader);
  aba.setRowHeight(2, 36);

  const hdrDados = [['DATA', 'PEDIDOS', '', 'STATUS', 'DATA MAIS ANTIGA', '']];
  aba.getRange(2, 1, 1, 6)
    .setValues(hdrDados)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(VG_T.accentSub)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // ── Dados: Datas + Pedidos (cols A-B) ──
  const linhasDatas = sortedDates.map(ds => {
    const [y, m, d] = ds.split('-').map(Number);
    return [new Date(y, m - 1, d), dateMap[ds].count];
  });

  if (linhasDatas.length) {
    // Fundo alternado
    const bgsDatas = linhasDatas.map((_, i) => [
      i % 2 === 0 ? VG_T.bgBase : VG_T.bgAlt,
      i % 2 === 0 ? VG_T.bgBase : VG_T.bgAlt,
    ]);
    aba.getRange(3, 1, linhasDatas.length, 2).setBackgrounds(bgsDatas);

    aba.getRange(3, 1, linhasDatas.length, 2).setValues(linhasDatas).setFontSize(12).setVerticalAlignment('middle');
    aba.getRange(3, 1, linhasDatas.length, 1)
      .setNumberFormat('dd/mm/yyyy')
      .setFontColor(VG_T.accentSub)
      .setHorizontalAlignment('center');
    aba.getRange(3, 2, linhasDatas.length, 1)
      .setFontColor(VG_T.green)
      .setFontWeight('bold')
      .setFontSize(13)
      .setHorizontalAlignment('center');
  }

  // ── Dados: Status + Data mais antiga (cols D-E) ──
  const linhasStatus = VG_STATUS_ALVO.map(nome => {
    const ts = statusOldest[nome];
    if (!ts) return [nome, ''];
    const ds = Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd');
    const [y, m, d] = ds.split('-').map(Number);
    return [nome, new Date(y, m - 1, d)];
  });

  const bgsStatus = linhasStatus.map((_, i) => [
    i % 2 === 0 ? VG_T.bgBase : VG_T.bgAlt,
    i % 2 === 0 ? VG_T.bgBase : VG_T.bgAlt,
  ]);
  aba.getRange(3, 4, linhasStatus.length, 2).setBackgrounds(bgsStatus);
  aba.getRange(3, 4, linhasStatus.length, 2)
    .setValues(linhasStatus)
    .setFontSize(11)
    .setVerticalAlignment('middle');
  aba.getRange(3, 4, linhasStatus.length, 1)
    .setFontColor(VG_T.textMain)
    .setHorizontalAlignment('left');
  aba.getRange(3, 5, linhasStatus.length, 1)
    .setNumberFormat('dd/mm/yyyy')
    .setFontColor(VG_T.amber)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  // ── Larguras de coluna ──
  aba.setColumnWidth(1, 120); // Data
  aba.setColumnWidth(2, 100); // Pedidos
  aba.setColumnWidth(3, 24);  // gap
  aba.setColumnWidth(4, 200); // Status
  aba.setColumnWidth(5, 130); // Data mais antiga
  aba.setColumnWidth(6, 20);  // margem

  // ── Alturas de linha ──
  if (nDatas > 0) {
    for (let r = 3; r < 3 + Math.max(nDatas, nStatus); r++) {
      aba.setRowHeight(r, 30);
    }
  }

  aba.setFrozenRows(2);
}

// ── Escrever aba "SKUs" ───────────────────────────────────────
function _vg_escreverSKUs(ss, tz, sortedDates, dateMap, agora) {
  let aba = ss.getSheetByName('SKUs');
  if (!aba) aba = ss.insertSheet('SKUs');
  aba.clearContents();
  aba.clearFormats();
  aba.setTabColor(VG_T.accentSub);

  const STEP    = 3;  // 2 colunas dados + 1 gap
  const nCols   = sortedDates.length * STEP + 1;

  // ── Fundo geral ──
  // Limita a uma área razoável para não ser lento
  const maxSkuRows = Math.max(...sortedDates.map(ds => Object.keys(dateMap[ds].skus).length));
  const totalR = maxSkuRows + 6;
  aba.getRange(1, 1, totalR, nCols)
    .setBackground(VG_T.bgBase)
    .setFontColor(VG_T.textMain)
    .setFontFamily('Arial')
    .setFontSize(11);

  // ── Linha 1: banner de atualização (só col A) ──
  aba.getRange(1, 1, 1, nCols).setBackground(VG_T.bgCard);
  aba.getRange(1, 1, 1, Math.min(nCols, 10)).merge()
    .setValue('  ⏱  Última atualização: ' + agora)
    .setFontColor(VG_T.textSub)
    .setFontStyle('italic')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  aba.setRowHeight(1, 28);
  aba.setRowHeight(2, 36);
  aba.setRowHeight(3, 30);

  // ── Um bloco por data ──
  sortedDates.forEach((ds, idx) => {
    const col     = 1 + idx * STEP;
    const [y, m, d] = ds.split('-').map(Number);

    // Linha 2: data — fundo accent, fonte grande
    aba.getRange(2, col, 1, 2)
      .setBackground(VG_T.bgHeader)
      .setFontColor(VG_T.accentSub)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    aba.getRange(2, col)
      .setValue(new Date(y, m - 1, d))
      .setNumberFormat('dd/mm/yyyy');
    aba.getRange(2, col + 1).setValue(''); // merge visual apenas via cor

    // Linha 3: cabeçalhos SKU / Qtd.
    aba.getRange(3, col, 1, 2)
      .setBackground(VG_T.bgCard)
      .setFontColor(VG_T.textSub)
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    aba.getRange(3, col    ).setValue('SKU');
    aba.getRange(3, col + 1).setValue('QTD');

    // Linha 4+: dados
    const skuRows = Object.entries(dateMap[ds].skus)
      .sort((a, b) => b[1] - a[1])
      .map(([sku, qty]) => [sku, qty]);

    if (skuRows.length) {
      // Fundo alternado
      const bgs = skuRows.map((_, i) => [
        i % 2 === 0 ? VG_T.bgBase : VG_T.bgAlt,
        i % 2 === 0 ? VG_T.bgBase : VG_T.bgAlt,
      ]);
      aba.getRange(4, col, skuRows.length, 2).setBackgrounds(bgs);
      aba.getRange(4, col, skuRows.length, 2).setValues(skuRows).setVerticalAlignment('middle');
      aba.getRange(4, col, skuRows.length, 1)
        .setNumberFormat('@')
        .setFontColor(VG_T.textMain)
        .setFontSize(11)
        .setHorizontalAlignment('left');
      aba.getRange(4, col + 1, skuRows.length, 1)
        .setFontColor(VG_T.green)
        .setFontWeight('bold')
        .setFontSize(12)
        .setHorizontalAlignment('center');
    }

    // Larguras das colunas do bloco
    aba.setColumnWidth(col,     130); // SKU
    aba.setColumnWidth(col + 1,  70); // Qtd
    if (col + 2 <= nCols) aba.setColumnWidth(col + 2, 16); // gap
  });

  aba.setFrozenRows(3);
}
