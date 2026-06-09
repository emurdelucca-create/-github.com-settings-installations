// ============================================================
// GIRO DE ESTOQUE — v1
// 3 abas: Menor Giro / Maior Giro / Maior Estoque
// Produtos sempre desmembrados (SKUs unitários)
// ============================================================

const GCFG = {
  ID_GE:      '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE:     'Dados Completos',
  ID_BL:      '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL:     'estoque',
  ID_CUSTOS:  '1C9Z4vT9SaamEGJ_hCWXdOm8vWOMhxw_vAb049mvtrP0',
  ABA_CUSTOS: 'Soma Composições',
  ID_PEND:    '1bYe0kxiHCUzPPHszQiK8JarKqzj5lxijAy7M3zpy1mg',

  // "Dados Completos" — índices de coluna (base 0)
  GE_SKU:  1, GE_QTD: 2, GE_DATA: 3, GE_FAT: 21,

  // BaseLinker
  BL_SKU: 1, BL_COD_EST: 2, BL_QTDE_EST: 3,
  BL_SALDO_PAD: 4, BL_SALDO_ARM: 5, BL_SALDO_CHG: 6,
  BL_TIPO: 7,

  CUST_SKU: 0, CUST_VALOR: 1,
  PEND_SKU: 1, PEND_QTD:   2,

  ABC_A: 0.70,
  ABC_B: 0.90,
  HEADER_ROWS: 4,
  NCOLS: 29,
};

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Giro de Estoque')
    .addItem('🔄 Atualizar Dados', 'gerarGiroEstoque')
    .addToUi();
}

// ============================================================
// ENTRADA PRINCIPAL
// ============================================================
function gerarGiroEstoque() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Atualizando Giro de Estoque... Aguarde.');

  try {
    const mapaCustos                         = _ge_carregarCustos();
    const { mapaCompostos, proporcoes, estoque } = _ge_carregarBL(mapaCustos);
    const pivotPend                          = _ge_carregarPendentes();
    const vendas                             = _ge_carregarVendas(mapaCompostos, proporcoes);

    const itens = _ge_montarItens(vendas, estoque, mapaCustos, pivotPend);
    _ge_calcularCurvas(itens);

    _ge_escreverAba('Giro — Menor Giro',    itens, (a, b) => a.dijSum  - b.dijSum);
    _ge_escreverAba('Giro — Maior Giro',    itens, (a, b) => b.dijSum  - a.dijSum);
    _ge_escreverAba('Giro — Maior Estoque', itens, (a, b) => b.estPend - a.estPend);

    ui.alert('✅ Giro de Estoque atualizado! ' + itens.length + ' SKUs.');
  } catch (e) {
    ui.alert('❌ Erro: ' + e.message + '\n\nStack:\n' + e.stack);
  }
}

// ============================================================
// PARSE DD/MM/YYYY  (idêntico ao BaseCompras)
// ============================================================
function _ge_parseData(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

// ============================================================
// HELPER: aba com tolerância a acentos corrompidos
// ============================================================
function _ge_getAba(ss, nomeDesejado) {
  let aba = ss.getSheetByName(nomeDesejado);
  if (aba) return aba;
  const norm = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const target = norm(nomeDesejado);
  for (const a of ss.getSheets()) {
    if (norm(a.getName()) === target) return a;
  }
  return null;
}

// ============================================================
// CARREGAR CUSTOS (aba "Soma Composições")
// ============================================================
function _ge_carregarCustos() {
  const ss  = SpreadsheetApp.openById(GCFG.ID_CUSTOS);
  const aba = _ge_getAba(ss, GCFG.ABA_CUSTOS);
  if (!aba) throw new Error('Aba Custos "' + GCFG.ABA_CUSTOS + '" não encontrada.');
  const dados = aba.getDataRange().getValues();
  const mapa  = {};
  for (let i = 1; i < dados.length; i++) {
    const sku   = String(dados[i][GCFG.CUST_SKU]   || '').trim();
    const custo = parseFloat(String(dados[i][GCFG.CUST_VALOR] || '0').replace(',', '.')) || 0;
    if (sku && !(sku in mapa)) mapa[sku] = custo;
  }
  return mapa;
}

// ============================================================
// CARREGAR BASELINKER — estoque + estrutura de kits
// ============================================================
function _ge_carregarBL(mapaCustos) {
  const ss  = SpreadsheetApp.openById(GCFG.ID_BL);
  const aba = _ge_getAba(ss, GCFG.ABA_BL);
  if (!aba) throw new Error('Aba BaseLinker "' + GCFG.ABA_BL + '" não encontrada.');
  const dados = aba.getDataRange().getValues();

  const estoque       = {};
  const mapaCompostos = {};

  for (let i = 1; i < dados.length; i++) {
    const row    = dados[i];
    const sku    = String(row[GCFG.BL_SKU]      || '').trim();
    const tipo   = String(row[GCFG.BL_TIPO]     || '').trim();
    const codEst = String(row[GCFG.BL_COD_EST]  || '').trim();
    const qtde   = Number(row[GCFG.BL_QTDE_EST] || 0);
    const saldo  = Number(row[GCFG.BL_SALDO_PAD] || 0)
                 + Number(row[GCFG.BL_SALDO_ARM] || 0)
                 + Number(row[GCFG.BL_SALDO_CHG] || 0);

    if (!sku) continue;

    if (tipo === 'Composto') {
      if (!codEst) continue;
      if (!mapaCompostos[sku]) mapaCompostos[sku] = [];
      mapaCompostos[sku].push({ codEst, qtde });
    } else {
      // Simples e PAI: estoque unitário direto
      estoque[sku] = (estoque[sku] || 0) + saldo;
    }
  }

  // Proporções de custo para desmembrar faturamento dos kits
  const proporcoes = {};
  for (const pai in mapaCompostos) {
    const comps = mapaCompostos[pai];
    let total = 0;
    comps.forEach(c => { total += (mapaCustos[c.codEst] || 0) * c.qtde; });
    proporcoes[pai] = {};
    comps.forEach(c => {
      proporcoes[pai][c.codEst] = total > 0
        ? ((mapaCustos[c.codEst] || 0) * c.qtde) / total
        : 1 / comps.length;
    });
  }

  return { mapaCompostos, proporcoes, estoque };
}

// ============================================================
// CARREGAR COMPRAS PENDENTES
// ============================================================
function _ge_carregarPendentes() {
  const ss    = SpreadsheetApp.openById(GCFG.ID_PEND);
  const aba   = ss.getSheets()[0];
  const dados = aba.getDataRange().getValues();
  const pivot = {};
  for (let i = 1; i < dados.length; i++) {
    const sku = String(dados[i][GCFG.PEND_SKU] || '').trim();
    const qtd = Number(dados[i][GCFG.PEND_QTD] || 0);
    if (!sku || !qtd) continue;
    pivot[sku] = (pivot[sku] || 0) + qtd;
  }
  return pivot;
}

// ============================================================
// CARREGAR VENDAS + DESMEMBRAMENTO
// ============================================================
function _ge_carregarVendas(mapaCompostos, proporcoes) {
  const geSS  = SpreadsheetApp.openById(GCFG.ID_GE);
  const geAba = _ge_getAba(geSS, GCFG.ABA_GE);
  if (!geAba) throw new Error('Aba GE Finance "' + GCFG.ABA_GE + '" não encontrada.');
  const dados = geAba.getDataRange().getValues();

  // dataMax = data mais recente na planilha → base para calcular "dias atrás"
  let dataMax = new Date(0);
  for (let i = 1; i < dados.length; i++) {
    const d = _ge_parseData(dados[i][GCFG.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(0, 0, 0, 0);

  const vendas = {};

  function acum(sku, dataVenda, dias, qtd, fat) {
    if (!vendas[sku]) {
      vendas[sku] = {
        qtd030: 0, qtd060: 0, qtd090: 0,
        fat030: 0, fat3060: 0, fat6090: 0,
        ultimaVenda: null,
      };
    }
    const v = vendas[sku];

    // Qtd acumulada (0-30 ⊂ 0-60 ⊂ 0-90)
    if      (dias < 30) { v.qtd030 += qtd; v.qtd060 += qtd; v.qtd090 += qtd; }
    else if (dias < 60) {                  v.qtd060 += qtd; v.qtd090 += qtd; }
    else                {                                    v.qtd090 += qtd; }

    // Fat por período exclusivo (para curvas ABC)
    if      (dias < 30) v.fat030  += fat;
    else if (dias < 60) v.fat3060 += fat;
    else                v.fat6090 += fat;

    if (!v.ultimaVenda || dataVenda > v.ultimaVenda) v.ultimaVenda = dataVenda;
  }

  for (let i = 1; i < dados.length; i++) {
    const d = _ge_parseData(dados[i][GCFG.GE_DATA]);
    if (!d) continue;
    const dNorm = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const dias  = Math.floor((dataMax - dNorm) / 86400000);
    if (dias < 0 || dias >= 90) continue;

    const skuGE = String(dados[i][GCFG.GE_SKU] || '').trim();
    const qtd   = Number(dados[i][GCFG.GE_QTD] || 0);
    const fat   = Number(dados[i][GCFG.GE_FAT] || 0);
    if (!skuGE) continue;

    if (mapaCompostos[skuGE]) {
      mapaCompostos[skuGE].forEach(comp => {
        const prop = (proporcoes[skuGE] && proporcoes[skuGE][comp.codEst]) || 0;
        acum(comp.codEst, dNorm, dias, qtd * comp.qtde, fat * prop);
      });
    } else {
      acum(skuGE, dNorm, dias, qtd, fat);
    }
  }

  return vendas;
}

// ============================================================
// MONTAR ITENS (1 objeto por SKU)
// ============================================================
function _ge_montarItens(vendas, estoque, mapaCustos, pivotPend) {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const allSkus = new Set([...Object.keys(vendas), ...Object.keys(estoque)]);
  const itens   = [];

  allSkus.forEach(sku => {
    const v    = vendas[sku] || { qtd030: 0, qtd060: 0, qtd090: 0, fat030: 0, fat3060: 0, fat6090: 0, ultimaVenda: null };
    const est  = estoque[sku] !== undefined ? estoque[sku] : 0;
    const custo = mapaCustos[sku] || 0;
    const pend  = pivotPend[sku]  || 0;

    const fat090       = v.fat030 + v.fat3060 + v.fat6090;
    const ultimaVenda  = v.ultimaVenda;
    const diasUltVenda = ultimaVenda ? Math.floor((hoje - ultimaVenda) / 86400000) : null;

    const i_030 = (v.qtd030 - est) * custo;
    const i_060 = (v.qtd060 - est) * custo;
    const i_090 = (v.qtd090 - est) * custo;

    itens.push({
      sku, est, ultimaVenda, diasUltVenda, custo,
      qtd030:  v.qtd030,
      qtd060:  v.qtd060,
      qtd090:  v.qtd090,
      i_030, i_060, i_090,
      custoEst: est > 0 ? est * custo : 0,
      fat030:  v.fat030,
      fat3060: v.fat3060,
      fat6090: v.fat6090,
      fat090,
      pend,
      dijSum:  i_030 + i_060 + i_090,
      estPend: est + pend,
    });
  });

  return itens;
}

// ============================================================
// CURVAS ABC + AABBCC (4 períodos × 2 tipos)
// ============================================================
function _ge_calcularCurvas(itens) {
  const periodos = [
    { fatKey: 'fat030',  abcKey: 'abc030',  aabbccKey: 'aabbcc030'  },
    { fatKey: 'fat3060', abcKey: 'abc3060', aabbccKey: 'aabbcc3060' },
    { fatKey: 'fat6090', abcKey: 'abc6090', aabbccKey: 'aabbcc6090' },
    { fatKey: 'fat090',  abcKey: 'abc090',  aabbccKey: 'aabbcc090'  },
  ];

  periodos.forEach(({ fatKey, abcKey, aabbccKey }) => {
    const sorted = [...itens].sort((a, b) => b[fatKey] - a[fatKey]);
    const total  = sorted.reduce((s, x) => s + x[fatKey], 0);

    // Curva ABC primária
    let cum = 0;
    sorted.forEach(item => {
      cum += item[fatKey];
      const pct   = total > 0 ? cum / total : 1;
      const letra = pct <= GCFG.ABC_A ? 'A' : pct <= GCFG.ABC_B ? 'B' : 'C';
      item[abcKey] = { letra, pct };
    });

    // AABBCC: sub-classificação ABC dentro de cada grupo primário
    ['A', 'B', 'C'].forEach(grp => {
      const grpItems = sorted.filter(x => x[abcKey].letra === grp);
      const grpTotal = grpItems.reduce((s, x) => s + x[fatKey], 0);
      let grpCum = 0;
      grpItems.forEach(item => {
        grpCum += item[fatKey];
        const subPct   = grpTotal > 0 ? grpCum / grpTotal : 1;
        const subLetra = subPct <= GCFG.ABC_A ? 'A' : subPct <= GCFG.ABC_B ? 'B' : 'C';
        item[aabbccKey] = { letra: grp + subLetra, pct: item[abcKey].pct };
      });
    });
  });
}

// ============================================================
// ESCREVER ABA
// Layout (29 colunas):
//  A(1)   SKU
//  B(2)   Qntd. Estoque
//  C(3)   Data Última Venda
//  D(4)   Qntd. Dias Última Venda
//  E(5)   Custo Unitário  ← vermelho se ≤ 0
//  F(6)   Qtd Vendida 0-30
//  G(7)   Qtd Vendida 0-60
//  H(8)   Qtd Vendida 0-90
//  I(9)   Diferença R$ 0-30  = (F-B)×E
//  J(10)  Diferença R$ 0-60
//  K(11)  Diferença R$ 0-90
//  L(12)  Custo Estoque = B×E (0 se estoque ≤ 0)
//  M(13)  Curva ABC 0-30  letra
//  N(14)  Curva ABC 0-30  %
//  O(15)  Curva ABC 30-60 letra
//  P(16)  Curva ABC 30-60 %
//  Q(17)  Curva ABC 60-90 letra
//  R(18)  Curva ABC 60-90 %
//  S(19)  Curva ABC 0-90  letra
//  T(20)  Curva ABC 0-90  %
//  U(21)  Curva AABBCC 0-30  letra
//  V(22)  Curva AABBCC 0-30  %
//  W(23)  Curva AABBCC 30-60 letra
//  X(24)  Curva AABBCC 30-60 %
//  Y(25)  Curva AABBCC 60-90 letra
//  Z(26)  Curva AABBCC 60-90 %
//  AA(27) Curva AABBCC 0-90  letra
//  AB(28) Curva AABBCC 0-90  %
//  AC(29) Compras Pendentes Qtd.
// ============================================================
function _ge_escreverAba(nomeAba, itens, sortFn) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);

  aba.clearContents();
  aba.clearFormats();

  const PRIMA = GCFG.HEADER_ROWS + 1;
  const NCOLS = GCFG.NCOLS;

  // ── Cabeçalho 4 linhas ──
  const h = [
    Array(NCOLS).fill(''),
    Array(NCOLS).fill(''),
    Array(NCOLS).fill(''),
    ['SKU','Qntd. Estoque','Data Última Venda','Qntd. Dias Última Venda','Custo Unitário',
     '0-30','0-60','0-90','0-30','0-60','0-90','Custo Estoque',
     'Curva','%','Curva','%','Curva','%','Curva','%',
     'Curva','%','Curva','%','Curva','%','Curva','%',
     'Qntd.'],
  ];
  // Linha 1: grupos de alto nível
  h[0][12] = 'Curva ABC';
  h[0][28] = 'Compras Pendentes';
  // Linha 2: períodos (ABC cols 13–20, AABBCC cols 21–28, base-0 = 12–27)
  h[1][12] = '0-30';  h[1][14] = '30-60'; h[1][16] = '60-90'; h[1][18] = '0-90';
  h[1][20] = '0-30';  h[1][22] = '30-60'; h[1][24] = '60-90'; h[1][26] = '0-90';
  // Linha 3: sub-grupos
  h[2][5]  = 'Qntd. Vendida';
  h[2][8]  = 'Diferença em R$';
  h[2][20] = 'Curva AABBCC';

  aba.getRange(1, 1, 4, NCOLS).setValues(h);

  // ── Dados ──
  const sorted = [...itens].sort(sortFn);
  // Data/hora de atualização em AE1:AE2
  aba.getRange(1, 31).setValue('Data Atualização');
  aba.getRange(2, 31).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));

  if (sorted.length === 0) { SpreadsheetApp.flush(); return; }

  const tz   = Session.getScriptTimeZone();
  const rows = sorted.map(item => [
    item.sku,                                                                // A
    item.est,                                                                // B
    item.ultimaVenda
      ? Utilities.formatDate(item.ultimaVenda, tz, 'dd/MM/yyyy')
      : '',                                                                  // C
    item.diasUltVenda !== null ? item.diasUltVenda : '',                     // D
    item.custo,                                                              // E
    item.qtd030,                                                             // F
    item.qtd060,                                                             // G
    item.qtd090,                                                             // H
    item.i_030,                                                              // I
    item.i_060,                                                              // J
    item.i_090,                                                              // K
    item.custoEst,                                                           // L
    item.abc030   ? item.abc030.letra   : '',                                // M
    item.abc030   ? item.abc030.pct     : 0,                                 // N
    item.abc3060  ? item.abc3060.letra  : '',                                // O
    item.abc3060  ? item.abc3060.pct    : 0,                                 // P
    item.abc6090  ? item.abc6090.letra  : '',                                // Q
    item.abc6090  ? item.abc6090.pct    : 0,                                 // R
    item.abc090   ? item.abc090.letra   : '',                                // S
    item.abc090   ? item.abc090.pct     : 0,                                 // T
    item.aabbcc030  ? item.aabbcc030.letra  : '',                            // U
    item.aabbcc030  ? item.aabbcc030.pct    : 0,                             // V
    item.aabbcc3060 ? item.aabbcc3060.letra : '',                            // W
    item.aabbcc3060 ? item.aabbcc3060.pct   : 0,                             // X
    item.aabbcc6090 ? item.aabbcc6090.letra : '',                            // Y
    item.aabbcc6090 ? item.aabbcc6090.pct   : 0,                             // Z
    item.aabbcc090  ? item.aabbcc090.letra  : '',                            // AA
    item.aabbcc090  ? item.aabbcc090.pct    : 0,                             // AB
    item.pend,                                                               // AC
  ]);

  aba.getRange(PRIMA, 1, rows.length, NCOLS).setValues(rows);

  // SKU como texto (evita conversão numérica)
  aba.getRange(PRIMA, 1, rows.length, 1).setNumberFormat('@');

  // Percentuais: N(14), P(16), R(18), T(20), V(22), X(24), Z(26), AB(28)
  [14, 16, 18, 20, 22, 24, 26, 28].forEach(col => {
    aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('0.0%');
  });

  // Fundo vermelho na col E onde custo ≤ 0
  const redA1 = [];
  sorted.forEach((item, i) => {
    if (item.custo <= 0) redA1.push(aba.getRange(PRIMA + i, 5).getA1Notation());
  });
  if (redA1.length) aba.getRangeList(redA1).setBackground('#ffcdd2');

  // Colorir letras ABC (M=13, O=15, Q=17, S=19)
  _ge_colorirABC(aba, sorted, PRIMA,
    [13, 15, 17, 19],
    ['abc030', 'abc3060', 'abc6090', 'abc090']);

  // Colorir letras AABBCC (U=21, W=23, Y=25, AA=27)
  _ge_colorirAABBCC(aba, sorted, PRIMA,
    [21, 23, 25, 27],
    ['aabbcc030', 'aabbcc3060', 'aabbcc6090', 'aabbcc090']);

  SpreadsheetApp.flush();
}

// ============================================================
// COLORIR ABC
// ============================================================
function _ge_colorirABC(aba, itens, prima, cols, keys) {
  const BG = { 'A': '#c8e6c9', 'B': '#fff9c4', 'C': '#ffcdd2' };
  const FG = { 'A': '#1b5e20', 'B': '#f57f17', 'C': '#b71c1c' };

  cols.forEach((col, idx) => {
    const key    = keys[idx];
    const grupos = { 'A': [], 'B': [], 'C': [] };
    itens.forEach((item, i) => {
      const l = item[key] && item[key].letra;
      if (l) grupos[l].push(aba.getRange(prima + i, col).getA1Notation());
    });
    Object.keys(grupos).forEach(l => {
      if (!grupos[l].length) return;
      aba.getRangeList(grupos[l]).setBackground(BG[l]).setFontColor(FG[l]);
    });
  });
}

// ============================================================
// COLORIR AABBCC
// ============================================================
function _ge_colorirAABBCC(aba, itens, prima, cols, keys) {
  const BG = {
    'AA': '#1b5e20', 'AB': '#388e3c', 'AC': '#a5d6a7',
    'BA': '#f57f17', 'BB': '#fbc02d', 'BC': '#fff176',
    'CA': '#b71c1c', 'CB': '#e57373', 'CC': '#ffcdd2',
  };
  const FG = {
    'AA': '#ffffff', 'AB': '#ffffff', 'AC': '#000000',
    'BA': '#ffffff', 'BB': '#000000', 'BC': '#000000',
    'CA': '#ffffff', 'CB': '#000000', 'CC': '#000000',
  };

  cols.forEach((col, idx) => {
    const key    = keys[idx];
    const grupos = {};
    itens.forEach((item, i) => {
      const l = item[key] && item[key].letra;
      if (!l || !BG[l]) return;
      if (!grupos[l]) grupos[l] = [];
      grupos[l].push(aba.getRange(prima + i, col).getA1Notation());
    });
    Object.keys(grupos).forEach(l => {
      aba.getRangeList(grupos[l]).setBackground(BG[l]).setFontColor(FG[l]);
    });
  });
}
