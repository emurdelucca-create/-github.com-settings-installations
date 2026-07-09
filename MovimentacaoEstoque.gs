// ============================================================
// MOVIMENTAÇÃO DE ESTOQUE  v2 — Corretiva / Planejada / Excesso
// 35 colunas (A–AI), 5 linhas de cabeçalho, dados a partir da linha 6
// Arquivo autossuficiente — não depende de outros .gs
// ============================================================

// ── Chamada à API BaseLinker (lê chave de Script Properties) ──
function _me_bl_call(method, params) {
  const key = PropertiesService.getScriptProperties().getProperty('BASELINKER_API_KEY');
  if (!key) throw new Error('BASELINKER_API_KEY não configurada nas propriedades do script.');
  const resp = UrlFetchApp.fetch('https://api.baselinker.com/connector.php', {
    method:      'post',
    contentType: 'application/x-www-form-urlencoded',
    payload:     'token=' + key + '&method=' + method + '&parameters=' + JSON.stringify(params || {}),
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (data.status !== 'SUCCESS') {
    throw new Error('[BL:' + method + '] ' + (data.error_message || JSON.stringify(data)));
  }
  return data;
}

const ME = {
  ID_GE:  '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE: 'Dados Completos',
  ID_BL:  '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL: 'estoque',

  GE_SKU: 1, GE_QTD: 2, GE_DATA: 3, GE_CANAL: 7,
  BL_SKU: 1, BL_COD: 2, BL_QTDE: 3, BL_TIPO:  7,

  INVENTORY_ID:     39947,
  WH_PADRAO:        'bl_44285',
  WH_ARMAZENAMENTO: 'bl_50394',
  WH_CHEGOU:        'bl_51442',

  STATUS_IDS: [
    252551, 268797, 268798, 268799,   // Novos, NF Emitida, Erro NF, Erro Etiqueta
    268813, 268814, 268806, 268805,   // [SEP]
    268800, 268801, 268803, 268804,   // [EMB]
    252552, 268792, 268794, 268795,   // [EXP]
  ],

  ABC_A: 0.70, ABC_B: 0.90,
  HEADER_ROWS: 5,
  NCOLS: 35,
};

// Peso AABBCC para sort da Corretiva (menor = mais urgente)
const ME_AABBCC_W = {
  AA:0.88, AB:0.93, AC:0.98,
  BA:0.91, BB:0.96, BC:1.02,
  CA:0.94, CB:0.99, CC:1.06,
};

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Mov. Estoque')
    .addItem('🔴 Corretiva',       'meGerarCorretiva')
    .addItem('🟡 Planejada',       'meGerarPlanejada')
    .addItem('🟢 Excesso',         'meGerarExcesso')
    .addSeparator()
    .addItem('🔄 Atualizar Todas', 'meGerarTodas')
    .addToUi();
}

function meGerarCorretiva() { _me_gerarAba('Corretiva'); }
function meGerarPlanejada()  { _me_gerarAba('Planejada'); }
function meGerarExcesso()    { _me_gerarAba('Excesso'); }

// Carrega dados uma vez → gera as 3 abas
function meGerarTodas() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando todas as abas... Aguarde 2-3 minutos.');
  try {
    const itens = _me_carregarTodos();
    ['Corretiva','Planejada','Excesso'].forEach(t => _me_escreverAba(t, itens));
    ui.alert('✅ Todas as abas atualizadas! ' + itens.length + ' SKUs.');
  } catch(e) { ui.alert('❌ Erro: ' + e.message + '\n\n' + e.stack); }
}

// Carrega dados e gera uma aba
function _me_gerarAba(tipoAba) {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando "' + tipoAba + '"... Aguarde 1-2 minutos.');
  try {
    const itens = _me_carregarTodos();
    _me_escreverAba(tipoAba, itens);
    ui.alert('✅ "' + tipoAba + '" atualizada! ' + itens.length + ' SKUs.');
  } catch(e) { ui.alert('❌ Erro: ' + e.message + '\n\n' + e.stack); }
}

// ============================================================
// PIPELINE COMPLETO
// ============================================================
function _me_carregarTodos() {
  Logger.log('[ME] Estrutura BL...');
  const { mapaCompostos } = _me_carregarBL();
  Logger.log('[ME] Vendas GE...');
  const vendas = _me_carregarVendas(mapaCompostos);
  Logger.log('[ME] Dados API BL...');
  const stockLocDim = _me_carregarDadosAPI();
  Logger.log('[ME] Pedidos abertos...');
  const pedidos = _me_carregarPedidos();
  Logger.log('[ME] Montando itens...');
  const itens = _me_montarItens(vendas, stockLocDim, pedidos, mapaCompostos);
  Logger.log('[ME] Curvas ABC...');
  _me_calcularCurvas(itens);
  return itens;
}

// ============================================================
// 1. ESTRUTURA BL — Compostos → lista de componentes
// ============================================================
function _me_carregarBL() {
  const ss  = SpreadsheetApp.openById(ME.ID_BL);
  const aba = ss.getSheetByName(ME.ABA_BL);
  if (!aba) throw new Error('Aba BL "' + ME.ABA_BL + '" não encontrada.');
  const rows = aba.getDataRange().getValues();
  const mapaCompostos = {};

  for (let i = 1; i < rows.length; i++) {
    const r    = rows[i];
    const sku  = String(r[ME.BL_SKU]  || '').trim();
    const tipo = String(r[ME.BL_TIPO] || '').trim();
    if (!sku || tipo !== 'Composto') continue;
    const cod  = String(r[ME.BL_COD]  || '').trim();
    const qtde = Number(r[ME.BL_QTDE] || 0);
    if (!cod) continue;
    if (!mapaCompostos[sku]) mapaCompostos[sku] = [];
    mapaCompostos[sku].push({ codEst: cod, qtde });
  }
  return { mapaCompostos };
}

// ============================================================
// 2. VENDAS GE (0-5 / 0-7 / 0-10 / 0-15 dias, sem FULL)
// ============================================================
function _me_carregarVendas(mapaCompostos) {
  const ss  = SpreadsheetApp.openById(ME.ID_GE);
  const aba = ss.getSheetByName(ME.ABA_GE);
  if (!aba) throw new Error('Aba GE "' + ME.ABA_GE + '" não encontrada.');
  const rows = aba.getDataRange().getValues();

  let dataMax = new Date(0);
  for (let i = 1; i < rows.length; i++) {
    const d = _me_parseData(rows[i][ME.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(0, 0, 0, 0);

  const vendas = {};

  function acum(sku, dias, qty) {
    if (!vendas[sku]) vendas[sku] = { v5:0, v7:0, v10:0, v15:0 };
    if (dias <  5) vendas[sku].v5  += qty;
    if (dias <  7) vendas[sku].v7  += qty;
    if (dias < 10) vendas[sku].v10 += qty;
    if (dias < 15) vendas[sku].v15 += qty;
  }

  for (let i = 1; i < rows.length; i++) {
    const r     = rows[i];
    const canal = String(r[ME.GE_CANAL] || '').toUpperCase();
    if (canal.includes('FULL')) continue;
    const d = _me_parseData(r[ME.GE_DATA]);
    if (!d) continue;
    const dNorm = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const dias  = Math.floor((dataMax - dNorm) / 86400000);
    if (dias < 0 || dias >= 15) continue;
    const sku = String(r[ME.GE_SKU] || '').trim();
    const qty = Number(r[ME.GE_QTD] || 0);
    if (!sku || !qty) continue;

    if (mapaCompostos[sku]) {
      mapaCompostos[sku].forEach(c => acum(c.codEst, dias, qty * c.qtde));
    } else {
      acum(sku, dias, qty);
    }
  }
  return vendas;
}

// ============================================================
// 3. DADOS API — stock / localizações / dimensões
//    Lógica A8/A9:
//    • Warehouses secundários (≠ bl_44285/50394/51442) → mapeados
//      por prefixo de localização para zona A8 ou A9.
//    • bl_44285 (Padrão master) → stock distribuído pelo prefixo
//      da location string (A8 → picking, A9 → armPad).
// ============================================================
function _me_carregarDadosAPI() {
  const KNOWN = new Set([ME.WH_PADRAO, ME.WH_ARMAZENAMENTO, ME.WH_CHEGOU]);

  // Passo 1: lista de produtos (paginada)
  const pidToSku = {};
  let page = 1;
  while (true) {
    const r       = _me_bl_call('getInventoryProductsList', { inventory_id: ME.INVENTORY_ID, page });
    const entries = Object.entries(r.products || {});
    if (!entries.length) break;
    entries.forEach(([pid, info]) => { pidToSku[pid] = String(info.sku || '').trim(); });
    if (entries.length < 1000) break;
    page++;
    Utilities.sleep(200);
  }

  // Passo 2: dados completos em lotes de 1000
  const rawData = {};
  const allPids = Object.keys(pidToSku);
  const LOTE    = 1000;
  for (let i = 0; i < allPids.length; i += LOTE) {
    const lote = allPids.slice(i, i + LOTE).map(Number);
    const r    = _me_bl_call('getInventoryProductsData', { inventory_id: ME.INVENTORY_ID, products: lote });
    Object.entries(r.products || {}).forEach(([pid, p]) => { rawData[pid] = p; });
    if (i + LOTE < allPids.length) Utilities.sleep(300);
  }

  // Passo 3a: detectar zona de sub-warehouses a partir de suas próprias location strings
  const subWhZone = {}; // wid → 'A8' | 'A9'
  Object.values(rawData).forEach(p => {
    Object.entries(p.locations || {}).forEach(([wid, locStr]) => {
      if (KNOWN.has(wid) || subWhZone[wid] || !locStr) return;
      const locs = _me_parseLocs(locStr);
      const lA8  = locs.some(x => x.toUpperCase().startsWith('A8'));
      const lA9  = locs.some(x => x.toUpperCase().startsWith('A9'));
      if (lA8 && !lA9)       subWhZone[wid] = 'A8';
      else if (lA9 && !lA8)  subWhZone[wid] = 'A9';
    });
  });

  // Passo 3b: inferir zona de sub-warehouses com location string vazia
  // a partir de produtos onde o armazém Padrão tem SOMENTE A8 ou SOMENTE A9
  Object.values(rawData).forEach(p => {
    const padLocs = _me_parseLocs((p.locations || {})[ME.WH_PADRAO] || '');
    const hasA8   = padLocs.some(x => x.toUpperCase().startsWith('A8'));
    const hasA9   = padLocs.some(x => x.toUpperCase().startsWith('A9'));
    if (hasA8 === hasA9) return; // misto ou sem localização → não serve para inferência
    const zone = hasA8 ? 'A8' : 'A9';
    Object.keys(p.stock || {}).forEach(wid => {
      if (!KNOWN.has(wid) && !subWhZone[wid]) subWhZone[wid] = zone;
    });
  });

  const zoneA8 = new Set(Object.keys(subWhZone).filter(w => subWhZone[w] === 'A8'));
  const zoneA9 = new Set(Object.keys(subWhZone).filter(w => subWhZone[w] === 'A9'));

  // Passo 4: consolida por SKU
  const stockLocDim = {};

  Object.entries(rawData).forEach(([pid, p]) => {
    const sku = pidToSku[pid];
    if (!sku) return;

    const s = p.stock     || {};
    const l = p.locations || {};

    // Soma stock dos sub-warehouses de cada zona (apenas os desta produto)
    let subA8 = 0, subA9 = 0;
    const a8Locs = [], a9Locs = [];

    Object.keys(s).forEach(wid => {
      if (KNOWN.has(wid)) return;
      const qty = Number(s[wid] || 0);
      if (zoneA8.has(wid)) {
        subA8 += qty;
        _me_parseLocs(l[wid] || '').filter(x => x.toUpperCase().startsWith('A8')).forEach(x => a8Locs.push(x));
      } else if (zoneA9.has(wid)) {
        subA9 += qty;
        _me_parseLocs(l[wid] || '').filter(x => x.toUpperCase().startsWith('A9')).forEach(x => a9Locs.push(x));
      }
    });

    // Locações da string do armazém Padrão (bl_44285) — usadas para colunas F/G
    const padraoLocs = _me_parseLocs(l[ME.WH_PADRAO] || '');
    padraoLocs.filter(x => x.toUpperCase().startsWith('A8')).forEach(x => a8Locs.push(x));
    padraoLocs.filter(x => x.toUpperCase().startsWith('A9')).forEach(x => a9Locs.push(x));

    let picking, armPad;
    if (subA8 + subA9 > 0) {
      // Sub-warehouses identificados — usa stock deles diretamente
      picking = subA8;
      armPad  = subA9;
    } else {
      // Sem sub-warehouses: usa stock total do bl_44285 com prefixo da localização
      const padraoStock = Number(s[ME.WH_PADRAO] || 0);
      const pA8 = a8Locs.length > 0;
      const pA9 = a9Locs.length > 0;
      if (pA8 && !pA9)       { picking = padraoStock; armPad = 0; }
      else if (pA9 && !pA8)  { picking = 0; armPad = padraoStock; }
      else                   { picking = padraoStock; armPad = 0; } // fallback
    }

    const arm = Number(s[ME.WH_ARMAZENAMENTO] || 0);
    const chg = Number(s[ME.WH_CHEGOU]        || 0);

    const locF   = [...new Set(a8Locs)].join(' / ');
    const locA9s = [...new Set(a9Locs)].join(' / ');
    const locArm = _me_parseLocs(l[ME.WH_ARMAZENAMENTO] || '').join(' / ');
    const locG   = locA9s || locArm;

    const h    = Number(p.height || 0);
    const w    = Number(p.width  || 0);
    const c    = Number(p.length || 0);
    const peso = Number(p.weight || 0);
    const vol  = h && w && c ? h * w * c : 0;

    if (!stockLocDim[sku]) {
      stockLocDim[sku] = { picking:0, armPad:0, arm:0, chg:0, locF:'', locG:'', h:0, w:0, c:0, peso:0, vol:0 };
    }
    const d = stockLocDim[sku];
    d.picking += picking;
    d.armPad  += armPad;
    d.arm     += arm;
    d.chg     += chg;
    if (locF && !d.locF) d.locF = locF;
    if (locG && !d.locG) d.locG = locG;
    if (peso > 0 && !d.peso) { d.h = h; d.w = w; d.c = c; d.peso = peso; d.vol = vol; }
  });

  return stockLocDim;
}

// ============================================================
// 4. PEDIDOS ABERTOS (payment_done > 0)
//    Agrupa por método de entrega
// ============================================================
function _me_carregarPedidos() {
  const pedMap = {};

  function add(sku, metodo, qty) {
    if (!pedMap[sku]) pedMap[sku] = { total:0, entrDireta:0, me2:0, retirada:0, xpress:0 };
    const m = String(metodo).toLowerCase();
    pedMap[sku].total += qty;
    if      (m.includes('entrega direta') || m.includes('flex'))        pedMap[sku].entrDireta += qty;
    else if (m.includes('me2') || m.includes('mercado envios') ||
             m.includes('agên') || m.includes('agen'))                  pedMap[sku].me2        += qty;
    else if (m.includes('retirada'))                                     pedMap[sku].retirada   += qty;
    else if (m.includes('xpress'))                                       pedMap[sku].xpress     += qty;
  }

  for (const sid of ME.STATUS_IDS) {
    let idFrom = 0;
    for (let iter = 0; iter < 200; iter++) {
      const r     = _me_bl_call('getOrders', { status_id: sid, id_from: idFrom });
      const batch = r.orders || [];
      if (!batch.length) break;
      batch.forEach(o => {
        if (!(Number(o.payment_done) > 0)) return;
        const metodo = String(o.delivery_method || '');
        (o.products || []).forEach(prod => {
          const sku = String(prod.sku || '').trim();
          const qty = Number(prod.quantity) || 0;
          if (sku && qty) add(sku, metodo, qty);
        });
      });
      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }
    Utilities.sleep(100);
  }
  return pedMap;
}

// ============================================================
// 5. MONTAR ITENS
// ============================================================
function _me_montarItens(vendas, stockLocDim, pedMap, mapaCompostos) {
  const allSkus = new Set([...Object.keys(vendas), ...Object.keys(stockLocDim)]);
  Object.keys(mapaCompostos).forEach(pai => allSkus.delete(pai)); // exclui PAIs/Compostos

  return [...allSkus].map(sku => {
    const v  = vendas[sku]      || { v5:0, v7:0, v10:0, v15:0 };
    const s  = stockLocDim[sku] || { picking:0, armPad:0, arm:0, chg:0, locF:'', locG:'', h:0, w:0, c:0, peso:0, vol:0 };
    const p  = pedMap[sku]      || { total:0, entrDireta:0, me2:0, retirada:0, xpress:0 };

    const totalStock = s.picking + s.armPad + s.arm + s.chg;

    function dif(vd) { return s.picking - vd; }
    function mov(vd) {
      const raw = vd - s.picking;
      // raw > 0: déficit (precisa mover pra picking) — capado pelo totalStock, mínimo 0
      // raw ≤ 0: excesso em picking — retorna negativo (usado pelo filtro da aba Excesso)
      return raw > 0 ? Math.max(0, Math.min(raw, totalStock)) : raw;
    }

    return {
      sku,
      picking: s.picking, armPad: s.armPad, arm: s.arm, chg: s.chg,
      locF: s.locF, locG: s.locG,
      v5: v.v5, v7: v.v7, v10: v.v10, v15: v.v15,
      dif5:  dif(v.v5),  dif7:  dif(v.v7),  dif10: dif(v.v10), dif15: dif(v.v15),
      mov5:  mov(v.v5),  mov7:  mov(v.v7),   mov10: mov(v.v10), mov15: mov(v.v15),
      h: s.h, w: s.w, c: s.c, peso: s.peso, vol: s.vol,
      pedTotal: p.total, entrDireta: p.entrDireta, me2: p.me2,
      retirada: p.retirada, xpress: p.xpress,
      abc: '', pctAbc: 0, aabbcc: '', pctAabbcc: 0,
    };
  });
}

// ============================================================
// 6. CURVAS ABC + AABBCC (baseado em vendas 0-15)
// ============================================================
function _me_calcularCurvas(itens) {
  const sorted = [...itens].sort((a, b) => b.v15 - a.v15);
  const total  = sorted.reduce((s, x) => s + x.v15, 0);

  let cum = 0;
  sorted.forEach(x => {
    cum += x.v15;
    const pct = total > 0 ? cum / total : 1;
    x.abc    = pct <= ME.ABC_A ? 'A' : pct <= ME.ABC_B ? 'B' : 'C';
    x.pctAbc = pct;
  });

  ['A', 'B', 'C'].forEach(grp => {
    const gi   = sorted.filter(x => x.abc === grp);
    const gTot = gi.reduce((s, x) => s + x.v15, 0);
    let gCum = 0;
    gi.forEach(x => {
      gCum += x.v15;
      const sp = gTot > 0 ? gCum / gTot : 1;
      const sl = sp <= ME.ABC_A ? 'A' : sp <= ME.ABC_B ? 'B' : 'C';
      x.aabbcc    = grp + sl;
      x.pctAabbcc = x.pctAbc;
    });
  });
}

// ============================================================
// 7. ESCREVER ABA
// ============================================================
function _me_escreverAba(tipoAba, todosItens) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let   aba  = ss.getSheetByName(tipoAba);
  if (!aba) {
    aba = ss.insertSheet(tipoAba);
  } else {
    try { aba.getRange(1, 1, ME.HEADER_ROWS + 1, ME.NCOLS + 1).breakApart(); } catch(_) {}
    aba.clearContents();
    aba.clearFormats();
  }

  const PRIMA = ME.HEADER_ROWS + 1; // linha 6

  // ── Filtro e ordenação ──
  // Apenas produtos com vendas nos últimos 15 dias em todas as abas
  const comVendas = todosItens.filter(x => x.v15 > 0);

  let itens;
  if (tipoAba === 'Corretiva') {
    itens = [...comVendas].sort((a, b) => {
      const wa = ME_AABBCC_W[a.aabbcc] || 1.0;
      const wb = ME_AABBCC_W[b.aabbcc] || 1.0;
      const sa = a.picking * wa;
      const sb = b.picking * wb;
      if (Math.abs(sa - sb) > 0.0001) return sa - sb;
      return wa - wb;
    });
  } else if (tipoAba === 'Planejada') {
    // Coluna Q = Movimentar 0-7, do maior para o menor
    itens = [...comVendas].sort((a, b) => b.mov7 - a.mov7);
  } else {
    // Excesso: somente mov10 < 0, do menor (mais excesso) para o maior
    itens = comVendas.filter(x => x.mov10 < 0).sort((a, b) => a.mov10 - b.mov10);
  }

  // ── Cabeçalho ──
  _me_escreverCabecalho(aba);
  if (!itens.length) { SpreadsheetApp.flush(); return; }

  // Formata SKU como texto ANTES de escrever
  aba.getRange(PRIMA, 1, itens.length, 1).setNumberFormat('@');

  // Excesso exibe valores negativos reais; Corretiva/Planejada exibe 0 para negativos
  const floorMov = tipoAba !== 'Excesso';
  const m = (v) => floorMov ? Math.max(0, v) : v;

  const linhas = itens.map(x => [
    String(x.sku),                                                   // A
    x.picking, x.armPad, x.arm, x.chg,                              // B-E
    x.locF, x.locG,                                                  // F-G
    x.v5, x.v7, x.v10, x.v15,                                       // H-K
    x.dif5, x.dif7, x.dif10, x.dif15,                               // L-O
    m(x.mov5), m(x.mov7), m(x.mov10), m(x.mov15),                  // P-S
    x.abc,    x.pctAbc,                                              // T-U
    x.aabbcc, x.pctAabbcc,                                           // V-W
    x.h, x.w, x.c, x.peso, x.vol,                                   // X-AB
    '', '',                                                           // AC-AD (em branco)
    x.pedTotal, x.entrDireta, x.me2, x.retirada, x.xpress,          // AE-AI
  ]);

  aba.getRange(PRIMA, 1, linhas.length, ME.NCOLS).setValues(linhas);

  // Percentuais (U=21, W=23)
  aba.getRange(PRIMA, 21, linhas.length, 1).setNumberFormat('0.0%');
  aba.getRange(PRIMA, 23, linhas.length, 1).setNumberFormat('0.0%');

  // Colorir Diferença negativa (L-O = cols 12-15)
  _me_colorirCols(aba, itens, PRIMA,
    [12,13,14,15], (x,c) => [x.dif5,x.dif7,x.dif10,x.dif15][c-12] < 0, '#ffcdd2');

  // Colorir Movimentar positivo (P-S = cols 16-19)
  _me_colorirCols(aba, itens, PRIMA,
    [16,17,18,19], (x,c) => [x.mov5,x.mov7,x.mov10,x.mov15][c-16] > 0, '#fff9c4');

  // Colorir curvas
  _me_colorirABC(aba, itens, PRIMA);
  _me_colorirAABBCC(aba, itens, PRIMA);

  // Timestamp
  aba.getRange(1, ME.NCOLS + 1).setValue('Atualização');
  aba.getRange(2, ME.NCOLS + 1).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  );

  SpreadsheetApp.flush();
  Logger.log('[ME] ' + tipoAba + ': ' + itens.length + ' linhas escritas.');
}

// ============================================================
// 8. CABEÇALHO 5 LINHAS
// ============================================================
function _me_escreverCabecalho(aba) {
  const N = ME.NCOLS;
  const H = Array.from({ length: 5 }, () => new Array(N).fill(''));

  // Linha 1 — grupos (índices 0-based = coluna - 1)
  H[0][0]  = 'SKU';
  H[0][1]  = 'Qntd. Estoque';
  H[0][5]  = 'Endereço';
  H[0][7]  = 'Qntd. Vend.';
  H[0][11] = 'Diferença';
  H[0][15] = 'Movimentar';
  H[0][19] = 'Curva ABC';
  H[0][21] = 'Curva AABBCC';
  H[0][23] = 'Medidas Produto';
  H[0][28] = 'Classificação Medida';
  H[0][29] = 'Classificação Peso';
  H[0][30] = 'Qntd. / Método Entrega';

  // Linha 2 — Armazém
  H[1][1] = 'Armazém';

  // Linha 3 — sub-armazéns
  H[2][1] = 'Padrão';
  H[2][3] = 'Armazenamento';
  H[2][4] = 'CHEGOU';

  // Linha 4 — zonas
  H[3][1] = 'A8';
  H[3][2] = 'A9';

  // Linha 5 — nomes das colunas (35 elementos)
  H[4] = [
    '',                                                         // A
    'Picking','Armazenamento','Armazenamento','CHEGOU',         // B-E
    'Picking','Armazenamento',                                  // F-G
    '0-5','0-7','0-10','0-15',                                  // H-K
    '0-5','0-7','0-10','0-15',                                  // L-O
    '0-5','0-7','0-10','0-15',                                  // P-S
    'Curva','%',                                                // T-U
    'Curva','%',                                                // V-W
    'Altura','Largura','Comprimento','Peso','cm³',              // X-AB
    '','',                                                      // AC-AD
    'Geral (Todos os Métodos)',                                  // AE
    'Entrega Direta',                                           // AF
    'ME2 - Mercado Envios Places',                              // AG
    'Retirada pelo Comprador',                                  // AH
    'Shopee Xpress',                                            // AI
  ];

  aba.getRange(1, 1, 5, N).setValues(H);

  // ── Merges linha 1 (row=1-indexed, col=1-indexed, span) ──
  [
    [1, 2, 4],   // B-E  Qntd. Estoque
    [1, 6, 2],   // F-G  Endereço
    [1, 8, 4],   // H-K  Qntd. Vend.
    [1, 12, 4],  // L-O  Diferença
    [1, 16, 4],  // P-S  Movimentar
    [1, 20, 2],  // T-U  Curva ABC
    [1, 22, 2],  // V-W  Curva AABBCC
    [1, 24, 5],  // X-AB Medidas Produto
    [1, 31, 5],  // AE-AI Qntd./Método
  ].forEach(([row, col, span]) => aba.getRange(row, col, 1, span).merge());

  // ── Merges linha 2 ──
  aba.getRange(2, 2, 1, 4).merge(); // B-E "Armazém"

  // ── Merges linha 3 ──
  aba.getRange(3, 2, 1, 2).merge(); // B-C "Padrão"

  _me_estilizarCabecalho(aba);
}

// ============================================================
// 9. ESTILIZAR CABEÇALHO
// ============================================================
function _me_estilizarCabecalho(aba) {
  // Cores por grupo de colunas
  [
    { de:1,  ate:1,  top:'#455a64', bot:'#263238' },  // SKU
    { de:2,  ate:5,  top:'#1565c0', bot:'#0d47a1' },  // Estoque
    { de:6,  ate:7,  top:'#6a1b9a', bot:'#4a148c' },  // Endereço
    { de:8,  ate:11, top:'#00838f', bot:'#006064' },  // Vendas
    { de:12, ate:15, top:'#c62828', bot:'#b71c1c' },  // Diferença
    { de:16, ate:19, top:'#e65100', bot:'#bf360c' },  // Movimentar
    { de:20, ate:21, top:'#558b2f', bot:'#33691e' },  // ABC
    { de:22, ate:23, top:'#2e7d32', bot:'#1b5e20' },  // AABBCC
    { de:24, ate:28, top:'#4e342e', bot:'#3e2723' },  // Medidas
    { de:29, ate:30, top:'#37474f', bot:'#263238' },  // Classificação
    { de:31, ate:35, top:'#4527a0', bot:'#311b92' },  // Pedidos
  ].forEach(g => {
    const w = g.ate - g.de + 1;
    aba.getRange(1, g.de, 4, w)
       .setBackground(g.top).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
    aba.getRange(5, g.de, 1, w)
       .setBackground(g.bot).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center')
       .setVerticalAlignment('middle').setWrap(true);
  });

  // Alturas
  for (let r = 1; r <= 4; r++) aba.setRowHeight(r, 22);
  aba.setRowHeight(5, 45);

  // Larguras de destaque
  aba.setColumnWidth(1, 115); // SKU
  aba.setColumnWidth(6, 150); // Loc A8
  aba.setColumnWidth(7, 150); // Loc A9

  // Congelar 5 linhas e coluna A
  aba.setFrozenRows(ME.HEADER_ROWS);
  aba.setFrozenColumns(1);
}

// ============================================================
// 10. COLORAÇÃO
// ============================================================
function _me_colorirCols(aba, itens, prima, cols, testFn, bg) {
  cols.forEach(col => {
    const letter = _me_colLetter(col);
    const cells  = [];
    itens.forEach((x, i) => { if (testFn(x, col)) cells.push(letter + (prima + i)); });
    if (cells.length) aba.getRangeList(cells).setBackground(bg);
  });
}

function _me_colorirABC(aba, itens, prima) {
  const BG = { A:'#c8e6c9', B:'#fff9c4', C:'#ffcdd2' };
  const FG = { A:'#1b5e20', B:'#f57f17', C:'#b71c1c' };
  const grps = {};
  itens.forEach((x, i) => {
    const g = x.abc;
    if (!g || !BG[g]) return;
    if (!grps[g]) grps[g] = [];
    grps[g].push('T' + (prima + i));
  });
  Object.keys(grps).forEach(g => aba.getRangeList(grps[g]).setBackground(BG[g]).setFontColor(FG[g]));
}

function _me_colorirAABBCC(aba, itens, prima) {
  const BG = {
    AA:'#1b5e20', AB:'#388e3c', AC:'#a5d6a7',
    BA:'#f57f17', BB:'#fbc02d', BC:'#fff176',
    CA:'#b71c1c', CB:'#e57373', CC:'#ffcdd2',
  };
  const FG = {
    AA:'#fff', AB:'#fff', AC:'#000',
    BA:'#fff', BB:'#000', BC:'#000',
    CA:'#fff', CB:'#000', CC:'#000',
  };
  const grps = {};
  itens.forEach((x, i) => {
    const g = x.aabbcc;
    if (!g || !BG[g]) return;
    if (!grps[g]) grps[g] = [];
    grps[g].push('V' + (prima + i));
  });
  Object.keys(grps).forEach(g => aba.getRangeList(grps[g]).setBackground(BG[g]).setFontColor(FG[g]));
}

// ============================================================
// 11. HELPERS
// ============================================================
function _me_parseLocs(str) {
  if (!str) return [];
  return str.split(/[,;\n|]/).map(s => s.trim()).filter(Boolean);
}

function _me_parseData(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function _me_colLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
