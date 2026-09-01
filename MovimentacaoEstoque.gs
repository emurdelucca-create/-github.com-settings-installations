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
  NCOLS: 35, // A–AI (35 colunas)

  SNAPSHOT_SHEET: 'SNAPSHOT_BL',
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
    .addItem('🔴 Corretiva',   'meGerarCorretiva')
    .addItem('🟡 Planejada',   'meGerarPlanejada')
    .addItem('🟢 Excesso',     'meGerarExcesso')
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
  const stockLocDim = _me_carregarDadosAPI(mapaCompostos);
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
//    Usa quantidade direta por armazém (sem análise de prefixo):
//    • WH_PADRAO (bl_44285)        → picking
//    • WH_ARMAZENAMENTO (bl_50394) → estoque caixa fechada (arm)
//    • WH_CHEGOU (bl_51442)        → chegou
// ============================================================
function _me_carregarDadosAPI(mapaCompostos) {
  const compostos = new Set(Object.keys(mapaCompostos || {}));

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

  // Passo 3: consolida por SKU usando IDs de armazém diretamente
  const stockLocDim = {};

  Object.entries(rawData).forEach(([pid, p]) => {
    const sku = pidToSku[pid];
    if (!sku) return;
    if (compostos.has(sku)) return;

    const s = p.stock     || {};
    const l = p.locations || {};

    const picking = Number(s[ME.WH_PADRAO]        || 0);
    const arm     = Number(s[ME.WH_ARMAZENAMENTO] || 0);
    const chg     = Number(s[ME.WH_CHEGOU]        || 0);

    const locF = String(l[ME.WH_PADRAO]        || '');
    const locG = String(l[ME.WH_ARMAZENAMENTO] || '');

    const h    = Number(p.height || 0);
    const w    = Number(p.width  || 0);
    const c    = Number(p.length || 0);
    const peso = Number(p.weight || 0);
    const vol  = h && w && c ? h * w * c : 0;

    if (!stockLocDim[sku]) {
      stockLocDim[sku] = { picking:0, arm:0, chg:0, locF:'', locG:'', h:0, w:0, c:0, peso:0, vol:0 };
    }
    const d = stockLocDim[sku];
    d.picking += picking;
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
//    Agrupa por método de entrega + status Movimentação separado
// ============================================================
function _me_carregarPedidos() {
  const pedMap = {};
  const MOV_STATUS = 269416;

  function addSep(sku, metodo, qty) {
    if (!pedMap[sku]) pedMap[sku] = { total:0, entrDireta:0, me2:0, retirada:0, xpress:0, mov:0 };
    const m = String(metodo).toLowerCase();
    pedMap[sku].total += qty;
    if      (m.includes('entrega direta') || m.includes('flex'))       pedMap[sku].entrDireta += qty;
    else if (m.includes('me2') || m.includes('mercado envios') ||
             m.includes('agên') || m.includes('agen'))                 pedMap[sku].me2        += qty;
    else if (m.includes('retirada'))                                    pedMap[sku].retirada   += qty;
    else if (m.includes('xpress'))                                      pedMap[sku].xpress     += qty;
  }

  function addMov(sku, qty) {
    if (!pedMap[sku]) pedMap[sku] = { total:0, entrDireta:0, me2:0, retirada:0, xpress:0, mov:0 };
    pedMap[sku].mov += qty;
  }

  // Separação / Embalagem / Expedição
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
          if (sku && qty) addSep(sku, metodo, qty);
        });
      });
      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }
    Utilities.sleep(100);
  }

  // Movimentação (status separado — sem distinção de método)
  let idFrom = 0;
  for (let iter = 0; iter < 200; iter++) {
    const r     = _me_bl_call('getOrders', { status_id: MOV_STATUS, id_from: idFrom });
    const batch = r.orders || [];
    if (!batch.length) break;
    batch.forEach(o => {
      if (!(Number(o.payment_done) > 0)) return;
      (o.products || []).forEach(prod => {
        const sku = String(prod.sku || '').trim();
        const qty = Number(prod.quantity) || 0;
        if (sku && qty) addMov(sku, qty);
      });
    });
    if (batch.length < 100) break;
    idFrom = batch[batch.length - 1].order_id;
  }

  return pedMap;
}

// ============================================================
// 5. MONTAR ITENS
// ============================================================
function _me_montarItens(vendas, stockLocDim, pedMap, mapaCompostos) {
  const allSkus = new Set([...Object.keys(vendas), ...Object.keys(stockLocDim)]);
  Object.keys(mapaCompostos).forEach(pai => allSkus.delete(pai));

  return [...allSkus].map(sku => {
    const v = vendas[sku]      || { v5:0, v7:0, v10:0, v15:0 };
    const s = stockLocDim[sku] || { picking:0, arm:0, chg:0, locF:'', locG:'', h:0, w:0, c:0, peso:0, vol:0 };
    const p = pedMap[sku]      || { total:0, entrDireta:0, me2:0, retirada:0, xpress:0, mov:0 };

    const totalStock = s.picking + s.arm + s.chg;

    function dif(vd) { return s.picking - vd; }
    function mov(vd) {
      const raw = vd - s.picking;
      return raw > 0 ? Math.max(0, Math.min(raw, totalStock)) : raw;
    }

    return {
      sku,
      picking: s.picking, arm: s.arm, chg: s.chg,
      locF: s.locF, locG: s.locG,
      v5: v.v5, v7: v.v7, v10: v.v10, v15: v.v15,
      dif5:  dif(v.v5),  dif7:  dif(v.v7),  dif10: dif(v.v10), dif15: dif(v.v15),
      mov5:  mov(v.v5),  mov7:  mov(v.v7),   mov10: mov(v.v10), mov15: mov(v.v15),
      h: s.h, w: s.w, c: s.c, peso: s.peso, vol: s.vol,
      pedTotal: p.total, entrDireta: p.entrDireta, me2: p.me2,
      retirada: p.retirada, xpress: p.xpress, pedMov: p.mov,
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
    x.picking, x.arm, x.chg,                                        // B-D
    x.locF, x.locG,                                                  // E-F
    x.v5, x.v7, x.v10, x.v15,                                       // G-J
    x.dif5, x.dif7, x.dif10, x.dif15,                               // K-N
    m(x.mov5), m(x.mov7), m(x.mov10), m(x.mov15),                  // O-R
    x.abc,    x.pctAbc,                                              // S-T
    x.aabbcc, x.pctAabbcc,                                           // U-V
    x.h, x.w, x.c, x.peso, x.vol,                                   // W-AA
    '', '',                                                           // AB-AC (em branco)
    x.pedTotal, x.entrDireta, x.me2, x.retirada, x.xpress, x.pedMov, // AD-AI
  ]);

  aba.getRange(PRIMA, 1, linhas.length, ME.NCOLS).setValues(linhas);

  // Percentuais (T=20, V=22)
  aba.getRange(PRIMA, 20, linhas.length, 1).setNumberFormat('0.0%');
  aba.getRange(PRIMA, 22, linhas.length, 1).setNumberFormat('0.0%');

  // Colorir Diferença negativa (K-N = cols 11-14)
  _me_colorirCols(aba, itens, PRIMA,
    [11,12,13,14], (x,c) => [x.dif5,x.dif7,x.dif10,x.dif15][c-11] < 0, '#ffcdd2');

  // Colorir Movimentar positivo (O-R = cols 15-18)
  _me_colorirCols(aba, itens, PRIMA,
    [15,16,17,18], (x,c) => [x.mov5,x.mov7,x.mov10,x.mov15][c-15] > 0, '#fff9c4');

  // Colorir curvas (S=19, U=21)
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
  H[0][4]  = 'Endereço';
  H[0][6]  = 'Qntd. Vend.';
  H[0][10] = 'Diferença';
  H[0][14] = 'Movimentar';
  H[0][18] = 'Curva ABC';
  H[0][20] = 'Curva AABBCC';
  H[0][22] = 'Medidas Produto';
  H[0][29] = 'Qntd. / Método Entrega';

  // Linha 2 — Armazém
  H[1][1] = 'Armazém';

  // Linha 3 — sub-armazéns
  H[2][1] = 'Padrão (Picking)';
  H[2][2] = 'Armazenamento';
  H[2][3] = 'CHEGOU';

  // Linha 5 — nomes das colunas (35 elementos)
  H[4] = [
    '',                                                         // A
    'Picking','Armazenamento','CHEGOU',                         // B-D
    'Picking','Armazenamento',                                  // E-F
    '0-5','0-7','0-10','0-15',                                  // G-J
    '0-5','0-7','0-10','0-15',                                  // K-N
    '0-5','0-7','0-10','0-15',                                  // O-R
    'Curva','%',                                                // S-T
    'Curva','%',                                                // U-V
    'Altura','Largura','Comprimento','Peso','cm³',              // W-AA
    '','',                                                      // AB-AC
    'Geral (Todos os Métodos)',                                  // AD
    'Entrega Direta',                                           // AE
    'ME2 - Mercado Envios Places',                              // AF
    'Retirada pelo Comprador',                                  // AG
    'Shopee Xpress',                                            // AH
    'Movimentação',                                             // AI
  ];

  aba.getRange(1, 1, 5, N).setValues(H);

  // ── Merges linha 1 ──
  [
    [1, 2, 3],   // B-D  Qntd. Estoque
    [1, 5, 2],   // E-F  Endereço
    [1, 7, 4],   // G-J  Qntd. Vend.
    [1, 11, 4],  // K-N  Diferença
    [1, 15, 4],  // O-R  Movimentar
    [1, 19, 2],  // S-T  Curva ABC
    [1, 21, 2],  // U-V  Curva AABBCC
    [1, 23, 5],  // W-AA Medidas Produto
    [1, 30, 6],  // AD-AI Qntd./Método
  ].forEach(([row, col, span]) => aba.getRange(row, col, 1, span).merge());

  // ── Merges linha 2 ──
  aba.getRange(2, 2, 1, 3).merge(); // B-D "Armazém"

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
    grps[g].push('S' + (prima + i));
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
    grps[g].push('U' + (prima + i));
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

// ============================================================
// DIAGNÓSTICO — dumpa dados brutos da API para um SKU específico
// Cria (ou limpa) a aba DEBUG_SKU com todos os campos retornados
// pela API: stock por warehouse, locations por warehouse, dimensões.
// Modifique SKU_ALVO antes de executar.
// ============================================================
function meDiagnosticarSKU() {
  const SKU_ALVO = '06045-S'; // ← altere aqui

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   aba = ss.getSheetByName('DEBUG_SKU');
  if (!aba) aba = ss.insertSheet('DEBUG_SKU');
  aba.clearContents();

  // 1. Encontra o product_id
  let pid = null;
  let page = 1;
  while (true) {
    const r       = _me_bl_call('getInventoryProductsList', { inventory_id: ME.INVENTORY_ID, page });
    const entries = Object.entries(r.products || {});
    if (!entries.length) break;
    for (const [id, info] of entries) {
      if (String(info.sku || '').trim() === SKU_ALVO) { pid = id; break; }
    }
    if (pid || entries.length < 1000) break;
    page++;
    Utilities.sleep(200);
  }

  if (!pid) {
    aba.getRange(1, 1).setValue('SKU não encontrado: ' + SKU_ALVO);
    SpreadsheetApp.flush();
    return;
  }

  // 2. Busca dados completos
  const r = _me_bl_call('getInventoryProductsData', {
    inventory_id: ME.INVENTORY_ID,
    products: [Number(pid)],
  });
  const p = (r.products || {})[pid];

  if (!p) {
    aba.getRange(1, 1).setValue('Produto não retornado por getInventoryProductsData');
    SpreadsheetApp.flush();
    return;
  }

  const rows = [];
  rows.push(['SKU', SKU_ALVO]);
  rows.push(['product_id', pid]);
  rows.push(['name', p.name || '']);
  rows.push(['height (cm)', p.height || 0]);
  rows.push(['width (cm)',  p.width  || 0]);
  rows.push(['length (cm)', p.length || 0]);
  rows.push(['weight (kg)', p.weight || 0]);
  rows.push(['', '']);
  rows.push(['=== STOCK (wid → qty) ===', '']);
  Object.entries(p.stock || {}).forEach(([wid, qty]) => rows.push([wid, qty]));
  rows.push(['', '']);
  rows.push(['=== LOCATIONS (wid → string) ===', '']);
  Object.entries(p.locations || {}).forEach(([wid, loc]) => rows.push([wid, loc || '(vazio)']));
  rows.push(['', '']);
  rows.push(['=== STOCK bruto JSON ===', JSON.stringify(p.stock)]);
  rows.push(['=== LOCATIONS bruto JSON ===', JSON.stringify(p.locations)]);

  aba.getRange(1, 1, rows.length, 2).setValues(rows);
  aba.autoResizeColumns(1, 2);
  SpreadsheetApp.flush();

  Logger.log('Diagnóstico de "' + SKU_ALVO + '" (pid=' + pid + ') escrito na aba DEBUG_SKU');
}

// ============================================================
// COMPARAÇÃO — confronta o CSV manual de 10/07/2026 contra o
// SNAPSHOT_BL atual (se existir) ou a API de produtos.
// Use para validar que o snapshot está correto após configurar
// o trigger horário. Veja a aba DEBUG_Comparacao.
// ============================================================
function meCompararComCSV() {
  // Snapshot do CSV exportado do BaseLinker em 10/07/2026 12:18
  // Colunas: a8 = soma picking (prefixo A8), a9 = soma armazenamento (prefixo A9)
  const CSV_EST = {"04426-S":{"a8":0,"a9":3},"04425-S":{"a8":0,"a9":6},"04422-S":{"a8":0,"a9":3},"04421-S":{"a8":0,"a9":23},"04420-S":{"a8":0,"a9":57},"00083-S":{"a8":0,"a9":65},"TP021":{"a8":17,"a9":0},"TC910":{"a8":8,"a9":0},"00084-S":{"a8":0,"a9":261},"00082-S":{"a8":0,"a9":1},"02170-S":{"a8":0,"a9":1},"02168-S":{"a8":0,"a9":26},"02167-S":{"a8":0,"a9":7},"02166-S":{"a8":0,"a9":21},"02164-S":{"a8":0,"a9":74},"02163-S":{"a8":0,"a9":3},"02161-S":{"a8":0,"a9":2},"24077-S":{"a8":0,"a9":76},"04413-S":{"a8":0,"a9":1},"04412-S":{"a8":0,"a9":4},"04411-S":{"a8":0,"a9":75},"04410-S":{"a8":0,"a9":251},"04409-S":{"a8":0,"a9":375},"04408-S":{"a8":0,"a9":9},"04407-S":{"a8":0,"a9":27},"25811":{"a8":0,"a9":106},"FA160":{"a8":179,"a9":160},"LBJ9002":{"a8":692,"a9":81},"0054-E":{"a8":0,"a9":868},"TM990":{"a8":0,"a9":623},"25180":{"a8":0,"a9":55},"TM0211":{"a8":2,"a9":1124},"TM0153":{"a8":0,"a9":400},"TP033":{"a8":3,"a9":380},"TS015":{"a8":19,"a9":40},"6791":{"a8":2,"a9":136},"6790":{"a8":414,"a9":529},"13170-E":{"a8":0,"a9":1250},"TM861":{"a8":0,"a9":98},"8604":{"a8":0,"a9":113},"10035-E":{"a8":27,"a9":115},"9392":{"a8":0,"a9":85},"24076-S":{"a8":0,"a9":2466},"04122-S":{"a8":2,"a9":546},"04123-S":{"a8":1,"a9":905},"15014-E":{"a8":0,"a9":10882},"05900-S":{"a8":0,"a9":1418},"05901-S":{"a8":0,"a9":1418},"04208-S":{"a8":0,"a9":504},"04209-S":{"a8":0,"a9":354},"4700-E":{"a8":0,"a9":283},"5090-E":{"a8":0,"a9":2649},"06045-S":{"a8":0,"a9":1091},"05854-S":{"a8":0,"a9":292},"05853-S":{"a8":0,"a9":291},"TP005":{"a8":0,"a9":1300},"24071-S":{"a8":0,"a9":4574},"20087-S":{"a8":0,"a9":1401},"TM863":{"a8":1,"a9":0},"14998-E":{"a8":0,"a9":83},"PA0066":{"a8":0,"a9":20},"SM0027":{"a8":0,"a9":475},"825-0001":{"a8":38,"a9":0},"00061-S":{"a8":0,"a9":5},"04094-S":{"a8":43,"a9":0},"04095-S":{"a8":1,"a9":0},"04614-S":{"a8":0,"a9":37},"04232-S":{"a8":0,"a9":962},"04233-S":{"a8":0,"a9":2605},"04009-S":{"a8":0,"a9":536},"05902-S":{"a8":0,"a9":25836},"ADES-BRANCO":{"a8":0,"a9":469},"ADES-PRETO":{"a8":0,"a9":1717},"750-0041":{"a8":0,"a9":190},"750-0040":{"a8":0,"a9":245},"10684-E":{"a8":0,"a9":360},"LOGO-HONDA":{"a8":0,"a9":194},"9708":{"a8":92,"a9":16},"TP1110":{"a8":0,"a9":15},"12506":{"a8":0,"a9":149},"03422-S":{"a8":0,"a9":1612},"03420-S":{"a8":0,"a9":17},"13955":{"a8":353,"a9":50},"2420210":{"a8":4,"a9":0},"10082":{"a8":0,"a9":571},"LBJ3009":{"a8":0,"a9":404},"TP044":{"a8":0,"a9":3},"TM971":{"a8":0,"a9":322},"TR310":{"a8":0,"a9":162},"10281":{"a8":0,"a9":154},"33170":{"a8":0,"a9":148},"29480":{"a8":0,"a9":118},"33093":{"a8":0,"a9":76},"860-0001":{"a8":182,"a9":0},"829-0004":{"a8":50,"a9":37},"829-0003":{"a8":10,"a9":0},"829-0002":{"a8":35,"a9":0},"829-0001":{"a8":92,"a9":0},"ML012":{"a8":1456,"a9":0},"TM853":{"a8":59,"a9":0},"TM445":{"a8":0,"a9":18},"TM401":{"a8":0,"a9":658},"TM402":{"a8":1,"a9":681},"TM451":{"a8":0,"a9":182},"TM466":{"a8":195,"a9":1000},"979-0001":{"a8":9,"a9":192},"29440":{"a8":14,"a9":0},"33021":{"a8":0,"a9":1011},"F120105":{"a8":0,"a9":236},"F120101":{"a8":0,"a9":29},"F100105":{"a8":0,"a9":20},"F120103":{"a8":0,"a9":399},"F100101":{"a8":0,"a9":11},"F100103":{"a8":0,"a9":89},"F120102":{"a8":0,"a9":291},"F120104":{"a8":0,"a9":360},"F120106":{"a8":0,"a9":106},"F100102":{"a8":0,"a9":56},"F100104":{"a8":0,"a9":30},"F100106":{"a8":0,"a9":72},"20046-S":{"a8":0,"a9":888},"TP011":{"a8":0,"a9":47},"22204-S":{"a8":0,"a9":278},"949-0004":{"a8":31,"a9":25},"TR305":{"a8":0,"a9":394},"05861-S":{"a8":0,"a9":133},"01115-S":{"a8":27,"a9":0},"04004-S":{"a8":4,"a9":0},"05862-S":{"a8":0,"a9":261},"01114-S":{"a8":11,"a9":0},"20089-S":{"a8":0,"a9":95},"20066-S":{"a8":0,"a9":11},"06110-S":{"a8":0,"a9":52},"03400-S":{"a8":0,"a9":155},"02404-S":{"a8":0,"a9":97},"02231-S":{"a8":0,"a9":147},"20073-S":{"a8":0,"a9":5},"LBJ403":{"a8":0,"a9":669},"00079-S":{"a8":0,"a9":464},"00078-S":{"a8":0,"a9":160},"TM932":{"a8":18,"a9":190},"AMM008":{"a8":14,"a9":48},"220281":{"a8":1,"a9":0},"220264":{"a8":0,"a9":11},"15675-E":{"a8":0,"a9":323},"10485-E":{"a8":0,"a9":32},"03467-S":{"a8":0,"a9":1},"03466-S":{"a8":1,"a9":0},"03464-S":{"a8":0,"a9":10},"03463-S":{"a8":0,"a9":1},"03446-S":{"a8":0,"a9":50},"25041-S":{"a8":0,"a9":267},"25040-S":{"a8":82,"a9":1302},"25038-S":{"a8":17,"a9":0},"25018-S":{"a8":0,"a9":108},"25017-S":{"a8":0,"a9":136},"25016-S":{"a8":0,"a9":127},"25010-S":{"a8":0,"a9":100},"01267-S":{"a8":4,"a9":0},"04202-S":{"a8":0,"a9":400},"05863-S":{"a8":0,"a9":1},"04428-S":{"a8":0,"a9":60},"04027-S":{"a8":0,"a9":9},"03481-S":{"a8":0,"a9":8},"03479-S":{"a8":0,"a9":19},"03476-S":{"a8":0,"a9":4},"03475-S":{"a8":0,"a9":17},"03474-S":{"a8":0,"a9":14},"03473-S":{"a8":0,"a9":17},"03472-S":{"a8":0,"a9":18},"03461-S":{"a8":0,"a9":6},"03437-S":{"a8":0,"a9":7},"03435-S":{"a8":0,"a9":16},"03434-S":{"a8":0,"a9":10},"03433-S":{"a8":0,"a9":10},"03421-S":{"a8":0,"a9":1},"00129-S":{"a8":0,"a9":15},"00111-S":{"a8":0,"a9":10},"00110-S":{"a8":0,"a9":16},"00109-S":{"a8":0,"a9":7},"00108-S":{"a8":0,"a9":3},"00126-S":{"a8":0,"a9":26},"04028-S":{"a8":0,"a9":9},"00112-S":{"a8":4,"a9":0},"FA152":{"a8":0,"a9":17},"TM450":{"a8":0,"a9":234},"KTC0127":{"a8":0,"a9":140},"KTC0160":{"a8":0,"a9":218},"KTC9125":{"a8":0,"a9":21},"RETM015":{"a8":0,"a9":260},"RETM042":{"a8":0,"a9":80},"RAM010":{"a8":50,"a9":0},"PSIP026":{"a8":33,"a9":0},"CAR013":{"a8":0,"a9":33},"10791-E":{"a8":0,"a9":7},"13976-E":{"a8":0,"a9":236},"9422":{"a8":0,"a9":172},"12408-E":{"a8":0,"a9":62},"15782-E":{"a8":0,"a9":730},"15483-E":{"a8":5,"a9":0},"11353-E":{"a8":0,"a9":100},"11486-E":{"a8":0,"a9":17},"16108-E":{"a8":0,"a9":100},"16102-E":{"a8":0,"a9":236},"11551-E":{"a8":0,"a9":102},"15967-E":{"a8":0,"a9":7},"14478-E":{"a8":0,"a9":50},"14724-E":{"a8":0,"a9":235},"8691":{"a8":0,"a9":236},"8573":{"a8":0,"a9":31},"13439-E":{"a8":57,"a9":0},"15881-E":{"a8":0,"a9":16},"13162-E":{"a8":8,"a9":15},"9096-E":{"a8":0,"a9":65},"9095-E":{"a8":0,"a9":62},"9116-E":{"a8":0,"a9":7},"9115-E":{"a8":0,"a9":7},"10947-E":{"a8":0,"a9":236},"04430-S":{"a8":0,"a9":10},"03482-S":{"a8":0,"a9":9},"03477-S":{"a8":0,"a9":9},"PT006":{"a8":0,"a9":60},"CBRD1006":{"a8":0,"a9":56},"PA0077":{"a8":0,"a9":22},"PA0069":{"a8":0,"a9":18},"EMS003":{"a8":0,"a9":7},"BFE0002":{"a8":0,"a9":42},"12310-E":{"a8":0,"a9":4},"11681-E":{"a8":0,"a9":102},"15884-E":{"a8":0,"a9":58},"15882-E":{"a8":0,"a9":39},"TR322":{"a8":0,"a9":485},"TR322A":{"a8":0,"a9":511},"TC451":{"a8":0,"a9":78},"TC1176":{"a8":0,"a9":40},"ATR466PTO":{"a8":0,"a9":41},"FA210":{"a8":0,"a9":198},"TM093":{"a8":0,"a9":300},"TM874":{"a8":0,"a9":272},"TM0139":{"a8":0,"a9":300},"TC3245":{"a8":0,"a9":71},"LBJ202":{"a8":0,"a9":353},"TM0509":{"a8":0,"a9":150},"04429-S":{"a8":0,"a9":49},"03376-S":{"a8":0,"a9":369},"CW0019":{"a8":0,"a9":176},"CW504":{"a8":0,"a9":300},"949-0001":{"a8":40,"a9":0},"03471-S":{"a8":0,"a9":1},"25020-S":{"a8":6,"a9":0},"TM0131":{"a8":1,"a9":0},"04010-S":{"a8":3,"a9":0},"06132-S":{"a8":1,"a9":0},"ATRT133":{"a8":0,"a9":1},"03426-S":{"a8":0,"a9":2},"03427-S":{"a8":0,"a9":1},"00125-S":{"a8":0,"a9":1},"00127-S":{"a8":0,"a9":1},"20060-S":{"a8":1,"a9":0},"20095-S":{"a8":1,"a9":0}};

  Logger.log('Carregando estrutura BL (compostos)...');
  const { mapaCompostos } = _me_carregarBL();
  Logger.log('Carregando dados da API...');
  const stockLocDim = _me_carregarDadosAPI(mapaCompostos);

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   aba = ss.getSheetByName('DEBUG_Comparacao');
  if (!aba) aba = ss.insertSheet('DEBUG_Comparacao');
  aba.clearContents();
  aba.clearFormats();

  const header = [
    'SKU',
    'CSV A8 (Picking)', 'API A8 (Picking)', 'Dif A8',
    'CSV A9 (ArmPad)', 'API A9 (ArmPad)', 'Dif A9',
    'Status',
  ];
  const rows = [header];

  const allSkus = new Set([...Object.keys(CSV_EST), ...Object.keys(stockLocDim)]);

  allSkus.forEach(sku => {
    const csv    = CSV_EST[sku]     || { a8: 0, a9: 0 };
    const api    = stockLocDim[sku] || { picking: 0, armPad: 0, arm: 0 };
    const apiA9  = (api.armPad || 0) + (api.arm || 0); // A9 = zona padrao + WH_ARMAZENAMENTO
    const difA8  = api.picking - csv.a8;
    const difA9  = apiA9 - csv.a9;
    const ok = (Math.abs(difA8) <= 5 && Math.abs(difA9) <= 5); // tolerância de 5 unidades
    rows.push([sku, csv.a8, api.picking, difA8, csv.a9, apiA9, difA9, ok ? 'OK' : 'DIVERGENTE']);
  });

  // Ordena: DIVERGENTE primeiro, depois por SKU
  rows.sort((a, b) => {
    if (a === header) return -1;
    if (b === header) return 1;
    if (a[7] !== b[7]) return a[7] === 'DIVERGENTE' ? -1 : 1;
    return String(a[0]).localeCompare(String(b[0]));
  });

  aba.getRange(1, 1, rows.length, header.length).setValues(rows);

  // Cabeçalho
  aba.getRange(1, 1, 1, header.length)
     .setBackground('#37474f').setFontColor('#fff').setFontWeight('bold');

  // Colorir divergentes
  rows.slice(1).forEach((r, i) => {
    if (r[7] === 'DIVERGENTE') {
      aba.getRange(i + 2, 1, 1, header.length).setBackground('#ffcdd2');
    }
  });

  aba.autoResizeColumns(1, header.length);
  SpreadsheetApp.flush();

  const divergentes = rows.slice(1).filter(r => r[7] === 'DIVERGENTE').length;
  Logger.log('Comparação concluída. Divergentes: ' + divergentes + ' / ' + (rows.length - 1));
  SpreadsheetApp.getUi().alert('Comparação concluída!\n\nDivergentes: ' + divergentes + ' de ' + (rows.length - 1) + ' SKUs\nVeja a aba DEBUG_Comparacao.');
}
