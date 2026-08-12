// ============================================================
// ANÁLISE NAJUMI vs DEMAIS — v2
// 62 colunas | 5 linhas de cabeçalho | dados a partir da linha 6
//
// Fontes:
//   GE Finance  → vendas (qtd, fat, caixa, margem)
//   BaseLinker  → estoque (PAD + ARM + CHEGOU)
//   Fornecedores → col A = SKU, col C = Fornecedor
//
// Ordenação: por AP (Fat. Total Geral 0-90) decrescente
// ============================================================

const NJCFG = {
  ID_GE:      '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE:     'Dados Completos',

  ID_BL:      '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL:     'estoque',

  ID_FORNEC:  '1VSraBQz0pnXwcCV0QjQmU8vtcIy9UFbDUh4H3ZNYc68',

  // Índices GE Finance
  GE_SKU:    1,
  GE_QTD:    2,
  GE_DATA:   3,
  GE_CANAL:  7,
  GE_FAT:    21,
  GE_CAIXA:  44,
  GE_MARGEM: 45,

  // Índices BaseLinker
  BL_SKU:       1,
  BL_SALDO_PAD: 4,
  BL_SALDO_ARM: 5,
  BL_SALDO_CHG: 6,
  BL_TIPO:      7,

  // Índices Fornecedor
  FORNEC_SKU:  0,  // col A
  FORNEC_NOME: 2,  // col C

  ABC_A: 0.70,
  ABC_B: 0.90,

  ABA_SAIDA: 'Análise Najumi',
};

// Canais reconhecidos (case-insensitive match na leitura)
const NJ_TODOS_CANAIS = [
  'ML Edmotos',     'ML Edmotos FULL',
  'ML Humble',      'ML Humble FULL',
  'ML Najumi',      'ML Najumi FULL',
  'ML Sky',         'ML Sky FULL',
  'ML Moto Vibe',   'ML Moto Vibe FULL',
  'Shopee Humble',
  'Shopee Najumi',  'Shopee Najumi FULL',
  'Shopee Sky',
];

const NJ_CANAIS_NAJUMI       = new Set(['ML Najumi', 'ML Najumi FULL', 'Shopee Najumi', 'Shopee Najumi FULL']);
const NJ_CANAIS_ML_NAJ       = new Set(['ML Najumi', 'ML Najumi FULL']);
const NJ_CANAIS_SHOP_TOTAL   = new Set(['Shopee Najumi', 'Shopee Najumi FULL']);
const NJ_CANAIS_SHOP_FULL    = new Set(['Shopee Najumi FULL']);
const NJ_CANAIS_SHOP_SEM     = new Set(['Shopee Najumi']);

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Análise Najumi')
    .addItem('🚀 Gerar Análise', 'gerarAnaliseNajumi')
    .addToUi();
}

// ── Ponto de entrada ──────────────────────────────────────────
function gerarAnaliseNajumi() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando análise Najumi... Aguarde 1-2 minutos.');
  try {
    const fornecMap  = _njCarregarFornecedores();
    const estoqueMap = _njCarregarEstoque();
    const dados      = _njAcumularVendas();
    const itens      = _njCalcularABC(dados);
    _njEscreverPlanilha(itens, fornecMap, estoqueMap);
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    ui.alert('✅ Análise gerada! ' + itens.length + ' SKUs — ' + ts);
  } catch(e) {
    ui.alert('❌ Erro:\n' + e.message + '\n\n' + e.stack);
  }
}

// ── Fornecedores ──────────────────────────────────────────────
function _njCarregarFornecedores() {
  const ss   = SpreadsheetApp.openById(NJCFG.ID_FORNEC);
  const aba  = ss.getSheets()[0];
  const data = aba.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const sku  = String(data[i][NJCFG.FORNEC_SKU]  || '').trim();
    const forn = String(data[i][NJCFG.FORNEC_NOME] || '').trim();
    if (sku && !map[sku]) map[sku] = forn;
  }
  return map;
}

// ── Estoque BaseLinker ─────────────────────────────────────────
function _njCarregarEstoque() {
  const ss  = SpreadsheetApp.openById(NJCFG.ID_BL);
  const aba = ss.getSheetByName(NJCFG.ABA_BL);
  if (!aba) throw new Error('Aba BaseLinker "' + NJCFG.ABA_BL + '" não encontrada.');
  const data = aba.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const sku  = String(data[i][NJCFG.BL_SKU]  || '').trim();
    const tipo = String(data[i][NJCFG.BL_TIPO] || '').trim();
    if (!sku || tipo === 'Composto') continue;
    const saldo = Number(data[i][NJCFG.BL_SALDO_PAD] || 0)
                + Number(data[i][NJCFG.BL_SALDO_ARM] || 0)
                + Number(data[i][NJCFG.BL_SALDO_CHG] || 0);
    if (map[sku] === undefined) map[sku] = saldo;
  }
  return map;
}

// ── Parse de data ─────────────────────────────────────────────
function _njParseData(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

// ── Estrutura de acumulador por grupo de canais ───────────────
function _njNovoPeriodo() {
  return { qty: 0, fat: 0, caixa: 0, margem: 0 };
}
function _njNovoGrupo() {
  return { '0-30': _njNovoPeriodo(), '30-60': _njNovoPeriodo(), '60-90': _njNovoPeriodo() };
}
function _njNovaEntry() {
  return {
    najumi:    _njNovoGrupo(),
    mlNaj:     _njNovoGrupo(),
    shopTotal: _njNovoGrupo(),
    shopFull:  _njNovoGrupo(),
    shopSem:   _njNovoGrupo(),
    geral:     _njNovoGrupo(),
  };
}

// ── Acumular vendas do GE Finance ─────────────────────────────
function _njAcumularVendas() {
  const ss  = SpreadsheetApp.openById(NJCFG.ID_GE);
  const aba = ss.getSheetByName(NJCFG.ABA_GE);
  if (!aba) throw new Error('Aba GE Finance "' + NJCFG.ABA_GE + '" não encontrada.');
  const rows = aba.getDataRange().getValues();

  // Data mais recente como referência dos períodos
  let dataMax = new Date(0);
  for (let i = 1; i < rows.length; i++) {
    const d = _njParseData(rows[i][NJCFG.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(23, 59, 59, 999);

  // Lookup rápido de canal (lower → normalizado)
  const canalIndex = {};
  NJ_TODOS_CANAIS.forEach(c => { canalIndex[c.toLowerCase()] = c; });

  const dados = {};

  for (let i = 1; i < rows.length; i++) {
    const row    = rows[i];
    const sku    = String(row[NJCFG.GE_SKU]    || '').trim();
    const canal  = String(row[NJCFG.GE_CANAL]  || '').trim().toLowerCase();
    const qtd    = Number(row[NJCFG.GE_QTD]    || 0);
    const fat    = Number(row[NJCFG.GE_FAT]    || 0);
    const caixa  = Number(row[NJCFG.GE_CAIXA]  || 0);
    const margem = Number(row[NJCFG.GE_MARGEM] || 0);

    if (!sku) continue;
    const canalNorm = canalIndex[canal];
    if (!canalNorm) continue;

    const d = _njParseData(row[NJCFG.GE_DATA]);
    if (!d) continue;
    const dias = Math.floor((dataMax - d) / 86400000);
    if (dias < 0 || dias >= 90) continue;

    const per = dias < 30 ? '0-30' : dias < 60 ? '30-60' : '60-90';

    if (!dados[sku]) dados[sku] = _njNovaEntry();
    const item = dados[sku];

    function add(grupo) {
      grupo[per].qty    += qtd;
      grupo[per].fat    += fat;
      grupo[per].caixa  += caixa;
      grupo[per].margem += margem;
    }

    if (NJ_CANAIS_NAJUMI.has(canalNorm)) {
      add(item.najumi);
      if (NJ_CANAIS_ML_NAJ.has(canalNorm))     add(item.mlNaj);
      if (NJ_CANAIS_SHOP_TOTAL.has(canalNorm)) add(item.shopTotal);
      if (NJ_CANAIS_SHOP_FULL.has(canalNorm))  add(item.shopFull);
      if (NJ_CANAIS_SHOP_SEM.has(canalNorm))   add(item.shopSem);
    } else {
      add(item.geral);
    }
  }

  return dados;
}

// ── Soma 0-90 de um grupo ─────────────────────────────────────
function _njSoma090(grupo) {
  return {
    qty:    grupo['0-30'].qty    + grupo['30-60'].qty    + grupo['60-90'].qty,
    fat:    grupo['0-30'].fat    + grupo['30-60'].fat    + grupo['60-90'].fat,
    caixa:  grupo['0-30'].caixa  + grupo['30-60'].caixa  + grupo['60-90'].caixa,
    margem: grupo['0-30'].margem + grupo['30-60'].margem + grupo['60-90'].margem,
    pctFat: 0, pctAcum: 0, curva: '',
  };
}

// ── ABC: recebe array de {fat}, atribui pctFat/pctAcum/curva ─
function _njABC(arr, getFat) {
  const total = arr.reduce((s, x) => s + getFat(x), 0);
  const sorted = [...arr].sort((a, b) => getFat(b) - getFat(a));
  let cum = 0;
  sorted.forEach(x => {
    cum += getFat(x);
    x.pctFat  = total > 0 ? getFat(x) / total : 0;
    x.pctAcum = total > 0 ? cum / total : 0;
    x.curva   = x.pctAcum <= NJCFG.ABC_A ? 'A' : x.pctAcum <= NJCFG.ABC_B ? 'B' : 'C';
  });
}

// ── Calcular ABC e ordenar ─────────────────────────────────────
function _njCalcularABC(dados) {
  const itens = Object.keys(dados).map(sku => {
    const d = dados[sku];
    return {
      sku,
      najumi:    d.najumi,
      mlNaj:     d.mlNaj,
      shopTotal: d.shopTotal,
      shopFull:  d.shopFull,
      shopSem:   d.shopSem,
      geral:     d.geral,
      naj090:    _njSoma090(d.najumi),
      ger090:    _njSoma090(d.geral),
      ger030:    Object.assign({}, d.geral['0-30'],  { pctFat:0, pctAcum:0, curva:'' }),
      ger3060:   Object.assign({}, d.geral['30-60'], { pctFat:0, pctAcum:0, curva:'' }),
      ger6090:   Object.assign({}, d.geral['60-90'], { pctFat:0, pctAcum:0, curva:'' }),
    };
  });

  // Ordenação principal: Fat Geral 0-90 decrescente (coluna AP)
  itens.sort((a, b) => b.ger090.fat - a.ger090.fat);

  // ABC Najumi 0-90
  _njABC(itens.map(x => x.naj090), x => x.fat);

  // ABC Geral 0-90 (já ordenado por fat → pctAcum monotônico)
  _njABC(itens.map(x => x.ger090), x => x.fat);

  // ABC Geral por período (cada um com sua própria ordenação interna)
  ['ger030', 'ger3060', 'ger6090'].forEach(campo => {
    _njABC(itens.map(x => x[campo]), x => x.fat);
  });

  return itens;
}

// ── Escrever planilha ─────────────────────────────────────────
function _njEscreverPlanilha(itens, fornecMap, estoqueMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba  = ss.getSheetByName(NJCFG.ABA_SAIDA);
  if (!aba) aba = ss.insertSheet(NJCFG.ABA_SAIDA);

  const NCOLS  = 62;
  const NHDR   = 5;   // linhas de cabeçalho
  const PRIMA  = NHDR + 1;

  // Limpa tudo
  aba.clearContents();
  aba.clearFormats();

  // ── Cabeçalho linha 1: grandes grupos ────────────────────────
  const h1 = _njArr(NCOLS);
  h1[0]  = 'SKU';
  h1[1]  = 'Fornecedor';
  h1[2]  = 'Estoque';
  h1[3]  = 'Dados de venda Najumi';
  h1[40] = 'Geral (Todas as demais contas, tirando Najumi)';
  aba.getRange(1, 1, 1, NCOLS).setValues([h1]);

  // ── Cabeçalho linha 2: subgrupos ─────────────────────────────
  const h2 = _njArr(NCOLS);
  h2[3]  = 'Najumi Total';
  h2[16] = 'ML Najumi';
  h2[22] = 'Shopee Najumi (Total)';
  h2[28] = 'Shopee Najumi (Só Full)';
  h2[34] = 'Shopee Najumi (Sem Full)';
  aba.getRange(2, 1, 1, NCOLS).setValues([h2]);

  // ── Cabeçalho linha 3: sub-subgrupos ─────────────────────────
  const h3 = _njArr(NCOLS);
  h3[3]  = 'Vendas por período';
  h3[9]  = 'Curva ABC';
  h3[40] = 'Curva ABC';
  aba.getRange(3, 1, 1, NCOLS).setValues([h3]);

  // ── Cabeçalho linha 4: períodos ──────────────────────────────
  const h4 = _njArr(NCOLS);
  // Najumi Total períodos
  h4[3]=  '0-30';  h4[5]= '30-60'; h4[7]= '60-90';
  // Najumi ABC
  h4[9]=  '0-90';
  // ML Najumi
  h4[16]= '0-30'; h4[18]= '30-60'; h4[20]= '60-90';
  // Shopee Total
  h4[22]= '0-30'; h4[24]= '30-60'; h4[26]= '60-90';
  // Shopee Full
  h4[28]= '0-30'; h4[30]= '30-60'; h4[32]= '60-90';
  // Shopee Sem Full
  h4[34]= '0-30'; h4[36]= '30-60'; h4[38]= '60-90';
  // Geral ABC
  h4[40]= '0-90';
  // Geral períodos
  h4[47]= '0-30'; h4[52]= '30-60'; h4[57]= '60-90';
  aba.getRange(4, 1, 1, NCOLS).setValues([h4]);

  // ── Cabeçalho linha 5: nomes de colunas ──────────────────────
  const h5 = [
    'SKU', 'Fornecedor', 'Estoque',
    // D-I  Najumi Total períodos
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    // J-P  Najumi ABC 0-90
    'Qntd. Total Vend.', 'Fat. Total', 'Caixa Total', 'Margem Total',
    '% Fat', '% Acumulado', 'Curva',
    // Q-V  ML Najumi
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    // W-AB Shopee Total
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    // AC-AH Shopee Full
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    // AI-AN Shopee Sem Full
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    // AO-AU Geral ABC 0-90
    'Qntd. Total Vend.', 'Fat. Total', 'Caixa Total', 'Margem Total',
    '% Fat', '% Acumulado', 'Curva',
    // AV-AZ Geral 0-30
    'Qntd. Total Vend.', 'Fat. Total', '% Fat', '% Acumulado', 'Curva',
    // BA-BE Geral 30-60
    'Qntd. Total Vend.', 'Fat. Total', '% Fat', '% Acumulado', 'Curva',
    // BF-BJ Geral 60-90
    'Qntd. Total Vend.', 'Fat. Total', '% Fat', '% Acumulado', 'Curva',
  ];
  aba.getRange(5, 1, 1, NCOLS).setValues([h5]);

  // ── Dados ─────────────────────────────────────────────────────
  const rows = itens.map(item => {
    const n   = item.najumi;
    const ml  = item.mlNaj;
    const st  = item.shopTotal;
    const sf  = item.shopFull;
    const ss2 = item.shopSem;
    const n9  = item.naj090;
    const g9  = item.ger090;
    const g3  = item.ger030;
    const g36 = item.ger3060;
    const g69 = item.ger6090;

    return [
      item.sku,                       // A
      fornecMap[item.sku]  || '',     // B
      estoqueMap[item.sku] || 0,      // C
      // D-I Najumi Total por período
      n['0-30'].qty,  n['0-30'].fat,
      n['30-60'].qty, n['30-60'].fat,
      n['60-90'].qty, n['60-90'].fat,
      // J-P Najumi ABC 0-90
      n9.qty, n9.fat, n9.caixa, n9.margem,
      n9.pctFat, n9.pctAcum, n9.curva,
      // Q-V ML Najumi
      ml['0-30'].qty,  ml['0-30'].fat,
      ml['30-60'].qty, ml['30-60'].fat,
      ml['60-90'].qty, ml['60-90'].fat,
      // W-AB Shopee Total
      st['0-30'].qty,  st['0-30'].fat,
      st['30-60'].qty, st['30-60'].fat,
      st['60-90'].qty, st['60-90'].fat,
      // AC-AH Shopee Full
      sf['0-30'].qty,  sf['0-30'].fat,
      sf['30-60'].qty, sf['30-60'].fat,
      sf['60-90'].qty, sf['60-90'].fat,
      // AI-AN Shopee Sem Full
      ss2['0-30'].qty,  ss2['0-30'].fat,
      ss2['30-60'].qty, ss2['30-60'].fat,
      ss2['60-90'].qty, ss2['60-90'].fat,
      // AO-AU Geral ABC 0-90
      g9.qty, g9.fat, g9.caixa, g9.margem,
      g9.pctFat, g9.pctAcum, g9.curva,
      // AV-AZ Geral 0-30
      g3.qty, g3.fat, g3.pctFat, g3.pctAcum, g3.curva,
      // BA-BE Geral 30-60
      g36.qty, g36.fat, g36.pctFat, g36.pctAcum, g36.curva,
      // BF-BJ Geral 60-90
      g69.qty, g69.fat, g69.pctFat, g69.pctAcum, g69.curva,
    ];
  });

  if (rows.length > 0) {
    aba.getRange(PRIMA, 1, rows.length, NCOLS).setValues(rows);

    // SKU como texto
    aba.getRange(PRIMA, 1, rows.length, 1).setNumberFormat('@');

    // Percentuais: N(14), O(15), AS(45), AT(46), AX(50), AY(51), BC(55), BD(56), BH(60), BI(61)
    [14, 15, 45, 46, 50, 51, 55, 56, 60, 61].forEach(col => {
      aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('0.0%');
    });

    // Moeda: colunas de Fat e Caixa e Margem
    // E(5),G(7),I(9),K(11),L(12),M(13) — Najumi
    // R(18),T(20),V(22) — ML
    // X(24),Z(26),AB(28) — Shopee Total
    // AD(30),AF(32),AH(34) — Shopee Full
    // AJ(36),AL(38),AN(40) — Shopee Sem
    // AP(42),AQ(43),AR(44) — Geral 0-90
    // AW(49),BB(54),BG(59) — Geral períodos fat
    [5,7,9,11,12,13,18,20,22,24,26,28,30,32,34,36,38,40,42,43,44,49,54,59].forEach(col => {
      aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('R$ #,##0.00');
    });
  }

  // ── Mesclar cabeçalhos ────────────────────────────────────────
  // Linha 1
  aba.getRange(1, 1, 5, 1).merge();   // A1:A5 SKU
  aba.getRange(1, 2, 5, 1).merge();   // B1:B5 Fornecedor
  aba.getRange(1, 3, 5, 1).merge();   // C1:C5 Estoque
  aba.getRange(1, 4, 1, 37).merge();  // D1:AN1 Dados Najumi
  aba.getRange(1, 41, 1, 22).merge(); // AO1:BJ1 Geral
  // Linha 2
  aba.getRange(2, 4,  1, 13).merge(); // D2:P2  Najumi Total
  aba.getRange(2, 17, 1, 6).merge();  // Q2:V2  ML Najumi
  aba.getRange(2, 23, 1, 6).merge();  // W2:AB2 Shopee Total
  aba.getRange(2, 29, 1, 6).merge();  // AC2:AH2 Shopee Full
  aba.getRange(2, 35, 1, 6).merge();  // AI2:AN2 Shopee Sem Full
  aba.getRange(2, 41, 1, 22).merge(); // AO2:BJ2
  // Linha 3
  aba.getRange(3, 4,  1, 6).merge();  // D3:I3  Vendas por período
  aba.getRange(3, 10, 1, 7).merge();  // J3:P3  Curva ABC
  aba.getRange(3, 17, 1, 6).merge();  // Q3:V3
  aba.getRange(3, 23, 1, 6).merge();  // W3:AB3
  aba.getRange(3, 29, 1, 6).merge();  // AC3:AH3
  aba.getRange(3, 35, 1, 6).merge();  // AI3:AN3
  aba.getRange(3, 41, 1, 7).merge();  // AO3:AU3 Curva ABC Geral
  aba.getRange(3, 48, 1, 22).merge(); // AV3:BJ3
  // Linha 4
  aba.getRange(4, 4,  1, 2).merge();  // D4:E4   0-30
  aba.getRange(4, 6,  1, 2).merge();  // F4:G4   30-60
  aba.getRange(4, 8,  1, 2).merge();  // H4:I4   60-90
  aba.getRange(4, 10, 1, 7).merge();  // J4:P4   0-90
  aba.getRange(4, 17, 1, 2).merge();  // Q4:R4   0-30
  aba.getRange(4, 19, 1, 2).merge();  // S4:T4   30-60
  aba.getRange(4, 21, 1, 2).merge();  // U4:V4   60-90
  aba.getRange(4, 23, 1, 2).merge();  // W4:X4
  aba.getRange(4, 25, 1, 2).merge();
  aba.getRange(4, 27, 1, 2).merge();
  aba.getRange(4, 29, 1, 2).merge();
  aba.getRange(4, 31, 1, 2).merge();
  aba.getRange(4, 33, 1, 2).merge();
  aba.getRange(4, 35, 1, 2).merge();
  aba.getRange(4, 37, 1, 2).merge();
  aba.getRange(4, 39, 1, 2).merge();
  aba.getRange(4, 41, 1, 7).merge();  // AO4:AU4 0-90
  aba.getRange(4, 48, 1, 5).merge();  // AV4:AZ4 0-30
  aba.getRange(4, 53, 1, 5).merge();  // BA4:BE4 30-60
  aba.getRange(4, 58, 1, 5).merge();  // BF4:BJ4 60-90

  // ── Estilo cabeçalho ─────────────────────────────────────────
  const HDR_BG   = '#1c2c4a';
  const HDR_FG   = '#3a8dff';
  const HDR_BOLD = true;
  aba.getRange(1, 1, NHDR, NCOLS).setBackground(HDR_BG)
     .setFontColor(HDR_FG)
     .setFontWeight(HDR_BOLD ? 'bold' : 'normal')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');

  // ── Fixar linhas e colunas ───────────────────────────────────
  aba.setFrozenRows(NHDR);
  aba.setFrozenColumns(3);

  SpreadsheetApp.flush();
}

// ── Util ──────────────────────────────────────────────────────
function _njArr(n) { return Array(n).fill(''); }
