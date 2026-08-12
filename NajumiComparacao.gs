// ============================================================
// NAJUMI COMPARAÇÃO — v1
// Planilha: "Najumi comparação"
// 62 colunas | 5 linhas de cabeçalho | dados a partir da linha 6
//
// Fontes:
//   GE Finance   → vendas (qtd, fat, caixa, margem)
//   BaseLinker   → estoque (PAD + ARM + CHEGOU)
//   Fornecedores → col A = SKU, col C = Fornecedor
//
// Ordenação: por AP (Fat. Total Geral 0-90) decrescente
// ============================================================

const NC = {
  ID_GE:      '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE:     'Dados Completos',

  ID_BL:      '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL:     'estoque',

  ID_FORNEC:  '1VSraBQz0pnXwcCV0QjQmU8vtcIy9UFbDUh4H3ZNYc68',

  GE_SKU:    1,
  GE_QTD:    2,
  GE_DATA:   3,
  GE_CANAL:  7,
  GE_FAT:    21,
  GE_CAIXA:  44,
  GE_MARGEM: 45,

  BL_SKU:       1,
  BL_SALDO_PAD: 4,
  BL_SALDO_ARM: 5,
  BL_SALDO_CHG: 6,
  BL_TIPO:      7,

  FORNEC_SKU:  0,
  FORNEC_NOME: 2,

  ABC_A: 0.70,
  ABC_B: 0.90,

  ABA_SAIDA: 'Análise Najumi',
};

const NC_TODOS_CANAIS = [
  'ML Edmotos',     'ML Edmotos FULL',
  'ML Humble',      'ML Humble FULL',
  'ML Najumi',      'ML Najumi FULL',
  'ML Sky',         'ML Sky FULL',
  'ML Moto Vibe',   'ML Moto Vibe FULL',
  'Shopee Humble',
  'Shopee Najumi',  'Shopee Najumi FULL',
  'Shopee Sky',
];

const NC_NAJUMI      = new Set(['ML Najumi', 'ML Najumi FULL', 'Shopee Najumi', 'Shopee Najumi FULL']);
const NC_ML_NAJ      = new Set(['ML Najumi', 'ML Najumi FULL']);
const NC_SHOP_TOTAL  = new Set(['Shopee Najumi', 'Shopee Najumi FULL']);
const NC_SHOP_FULL   = new Set(['Shopee Najumi FULL']);
const NC_SHOP_SEM    = new Set(['Shopee Najumi']);

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Najumi Comparação')
    .addItem('🚀 Gerar Análise', 'ncGerarAnalise')
    .addToUi();
}

// ── Ponto de entrada ──────────────────────────────────────────
function ncGerarAnalise() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando análise... Aguarde 1-2 minutos.');
  try {
    const fornecMap  = _ncFornecedores();
    const estoqueMap = _ncEstoque();
    const dados      = _ncAcumular();
    const itens      = _ncABC(dados);
    _ncEscrever(itens, fornecMap, estoqueMap);
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    ui.alert('✅ Concluído! ' + itens.length + ' SKUs — ' + ts);
  } catch(e) {
    ui.alert('❌ Erro:\n' + e.message + '\n\n' + e.stack);
  }
}

// ── Fornecedores ──────────────────────────────────────────────
function _ncFornecedores() {
  const ss   = SpreadsheetApp.openById(NC.ID_FORNEC);
  const data = ss.getSheets()[0].getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const sku  = String(data[i][NC.FORNEC_SKU]  || '').trim();
    const forn = String(data[i][NC.FORNEC_NOME] || '').trim();
    if (sku && !map[sku]) map[sku] = forn;
  }
  return map;
}

// ── Estoque BaseLinker ─────────────────────────────────────────
function _ncEstoque() {
  const ss  = SpreadsheetApp.openById(NC.ID_BL);
  const aba = ss.getSheetByName(NC.ABA_BL);
  if (!aba) throw new Error('Aba BaseLinker "' + NC.ABA_BL + '" não encontrada.');
  const data = aba.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const sku  = String(data[i][NC.BL_SKU]  || '').trim();
    const tipo = String(data[i][NC.BL_TIPO] || '').trim();
    if (!sku || tipo === 'Composto') continue;
    const saldo = Number(data[i][NC.BL_SALDO_PAD] || 0)
                + Number(data[i][NC.BL_SALDO_ARM] || 0)
                + Number(data[i][NC.BL_SALDO_CHG] || 0);
    if (map[sku] === undefined) map[sku] = saldo;
  }
  return map;
}

// ── Parse de data DD/MM/YYYY ──────────────────────────────────
function _ncData(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

// ── Estruturas de acumulação ──────────────────────────────────
function _ncPer()   { return { qty: 0, fat: 0, caixa: 0, margem: 0 }; }
function _ncGrupo() { return { '0-30': _ncPer(), '30-60': _ncPer(), '60-90': _ncPer() }; }
function _ncEntry() {
  return {
    najumi: _ncGrupo(), mlNaj: _ncGrupo(),
    shopTotal: _ncGrupo(), shopFull: _ncGrupo(), shopSem: _ncGrupo(),
    geral: _ncGrupo(),
  };
}

// ── Acumular vendas ───────────────────────────────────────────
function _ncAcumular() {
  const ss  = SpreadsheetApp.openById(NC.ID_GE);
  const aba = ss.getSheetByName(NC.ABA_GE);
  if (!aba) throw new Error('Aba GE Finance "' + NC.ABA_GE + '" não encontrada.');
  const rows = aba.getDataRange().getValues();

  let dataMax = new Date(0);
  for (let i = 1; i < rows.length; i++) {
    const d = _ncData(rows[i][NC.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(23, 59, 59, 999);

  const idx = {};
  NC_TODOS_CANAIS.forEach(c => { idx[c.toLowerCase()] = c; });

  const dados = {};

  for (let i = 1; i < rows.length; i++) {
    const row    = rows[i];
    const sku    = String(row[NC.GE_SKU]    || '').trim();
    const canal  = String(row[NC.GE_CANAL]  || '').trim().toLowerCase();
    const qtd    = Number(row[NC.GE_QTD]    || 0);
    const fat    = Number(row[NC.GE_FAT]    || 0);
    const caixa  = Number(row[NC.GE_CAIXA]  || 0);
    const margem = Number(row[NC.GE_MARGEM] || 0);

    if (!sku) continue;
    const cn = idx[canal];
    if (!cn) continue;

    const d = _ncData(row[NC.GE_DATA]);
    if (!d) continue;
    const dias = Math.floor((dataMax - d) / 86400000);
    if (dias < 0 || dias >= 90) continue;

    const per = dias < 30 ? '0-30' : dias < 60 ? '30-60' : '60-90';
    if (!dados[sku]) dados[sku] = _ncEntry();
    const item = dados[sku];

    function add(g) {
      g[per].qty    += qtd;
      g[per].fat    += fat;
      g[per].caixa  += caixa;
      g[per].margem += margem;
    }

    if (NC_NAJUMI.has(cn)) {
      add(item.najumi);
      if (NC_ML_NAJ.has(cn))     add(item.mlNaj);
      if (NC_SHOP_TOTAL.has(cn)) add(item.shopTotal);
      if (NC_SHOP_FULL.has(cn))  add(item.shopFull);
      if (NC_SHOP_SEM.has(cn))   add(item.shopSem);
    } else {
      add(item.geral);
    }
  }

  return dados;
}

// ── Soma 0-90 ─────────────────────────────────────────────────
function _ncSoma(g) {
  return {
    qty:    g['0-30'].qty    + g['30-60'].qty    + g['60-90'].qty,
    fat:    g['0-30'].fat    + g['30-60'].fat    + g['60-90'].fat,
    caixa:  g['0-30'].caixa  + g['30-60'].caixa  + g['60-90'].caixa,
    margem: g['0-30'].margem + g['30-60'].margem + g['60-90'].margem,
    pctFat: 0, pctAcum: 0, curva: '',
  };
}

// ── ABC genérico ──────────────────────────────────────────────
function _ncCalcABC(arr, getFat) {
  const total  = arr.reduce((s, x) => s + getFat(x), 0);
  const sorted = [...arr].sort((a, b) => getFat(b) - getFat(a));
  let cum = 0;
  sorted.forEach(x => {
    cum += getFat(x);
    x.pctFat  = total > 0 ? getFat(x) / total : 0;
    x.pctAcum = total > 0 ? cum / total : 0;
    x.curva   = x.pctAcum <= NC.ABC_A ? 'A' : x.pctAcum <= NC.ABC_B ? 'B' : 'C';
  });
}

// ── Calcular todos os ABCs e ordenar ─────────────────────────
function _ncABC(dados) {
  const itens = Object.keys(dados).map(sku => {
    const d = dados[sku];
    return {
      sku,
      najumi: d.najumi, mlNaj: d.mlNaj,
      shopTotal: d.shopTotal, shopFull: d.shopFull, shopSem: d.shopSem,
      geral: d.geral,
      n90:  _ncSoma(d.najumi),
      g90:  _ncSoma(d.geral),
      g030: Object.assign({}, d.geral['0-30'],  { pctFat:0, pctAcum:0, curva:'' }),
      g36:  Object.assign({}, d.geral['30-60'], { pctFat:0, pctAcum:0, curva:'' }),
      g69:  Object.assign({}, d.geral['60-90'], { pctFat:0, pctAcum:0, curva:'' }),
    };
  });

  // Ordenação principal: Fat Geral 0-90 (AP) decrescente
  itens.sort((a, b) => b.g90.fat - a.g90.fat);

  _ncCalcABC(itens.map(x => x.n90),  x => x.fat);
  _ncCalcABC(itens.map(x => x.g90),  x => x.fat);
  _ncCalcABC(itens.map(x => x.g030), x => x.fat);
  _ncCalcABC(itens.map(x => x.g36),  x => x.fat);
  _ncCalcABC(itens.map(x => x.g69),  x => x.fat);

  return itens;
}

// ── Escrever planilha ─────────────────────────────────────────
function _ncEscrever(itens, fornecMap, estoqueMap) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let aba   = ss.getSheetByName(NC.ABA_SAIDA);
  if (!aba) aba = ss.insertSheet(NC.ABA_SAIDA);

  const NCOLS = 62;
  const NHDR  = 5;
  const PRIMA = NHDR + 1;

  aba.clearContents();
  aba.clearFormats();

  // ── Linha 1 ───────────────────────────────────────────────────
  const h1 = _ncArr(NCOLS);
  h1[0] = 'SKU'; h1[1] = 'Fornecedor'; h1[2] = 'Estoque';
  h1[3] = 'Dados de venda Najumi';
  h1[40] = 'Geral (Todas as demais contas, tirando Najumi)';
  aba.getRange(1, 1, 1, NCOLS).setValues([h1]);

  // ── Linha 2 ───────────────────────────────────────────────────
  const h2 = _ncArr(NCOLS);
  h2[3]  = 'Najumi Total';
  h2[16] = 'ML Najumi';
  h2[22] = 'Shopee Najumi (Total)';
  h2[28] = 'Shopee Najumi (Só Full)';
  h2[34] = 'Shopee Najumi (Sem Full)';
  aba.getRange(2, 1, 1, NCOLS).setValues([h2]);

  // ── Linha 3 ───────────────────────────────────────────────────
  const h3 = _ncArr(NCOLS);
  h3[3]  = 'Vendas por período';
  h3[9]  = 'Curva ABC';
  h3[40] = 'Curva ABC';
  aba.getRange(3, 1, 1, NCOLS).setValues([h3]);

  // ── Linha 4 ───────────────────────────────────────────────────
  const h4 = _ncArr(NCOLS);
  h4[3]=  '0-30'; h4[5]= '30-60'; h4[7]= '60-90'; h4[9]= '0-90';
  h4[16]= '0-30'; h4[18]= '30-60'; h4[20]= '60-90';
  h4[22]= '0-30'; h4[24]= '30-60'; h4[26]= '60-90';
  h4[28]= '0-30'; h4[30]= '30-60'; h4[32]= '60-90';
  h4[34]= '0-30'; h4[36]= '30-60'; h4[38]= '60-90';
  h4[40]= '0-90';
  h4[47]= '0-30'; h4[52]= '30-60'; h4[57]= '60-90';
  aba.getRange(4, 1, 1, NCOLS).setValues([h4]);

  // ── Linha 5 ───────────────────────────────────────────────────
  const h5 = [
    'SKU', 'Fornecedor', 'Estoque',
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    'Qntd. Total Vend.', 'Fat. Total', 'Caixa Total', 'Margem Total', '% Fat', '% Acumulado', 'Curva',
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)', 'Qntd.', 'Fat (R$)',
    'Qntd. Total Vend.', 'Fat. Total', 'Caixa Total', 'Margem Total', '% Fat', '% Acumulado', 'Curva',
    'Qntd. Total Vend.', 'Fat. Total', '% Fat', '% Acumulado', 'Curva',
    'Qntd. Total Vend.', 'Fat. Total', '% Fat', '% Acumulado', 'Curva',
    'Qntd. Total Vend.', 'Fat. Total', '% Fat', '% Acumulado', 'Curva',
  ];
  aba.getRange(5, 1, 1, NCOLS).setValues([h5]);

  // ── Dados ─────────────────────────────────────────────────────
  const rows = itens.map(item => {
    const n  = item.najumi;
    const ml = item.mlNaj;
    const st = item.shopTotal;
    const sf = item.shopFull;
    const ss2 = item.shopSem;
    const n9  = item.n90;
    const g9  = item.g90;
    const g3  = item.g030;
    const g36 = item.g36;
    const g69 = item.g69;

    return [
      item.sku,                       // A
      fornecMap[item.sku]  || '',     // B
      estoqueMap[item.sku] || 0,      // C
      n['0-30'].qty,  n['0-30'].fat,  // D E
      n['30-60'].qty, n['30-60'].fat, // F G
      n['60-90'].qty, n['60-90'].fat, // H I
      n9.qty, n9.fat, n9.caixa, n9.margem, n9.pctFat, n9.pctAcum, n9.curva, // J-P
      ml['0-30'].qty, ml['0-30'].fat, ml['30-60'].qty, ml['30-60'].fat, ml['60-90'].qty, ml['60-90'].fat, // Q-V
      st['0-30'].qty, st['0-30'].fat, st['30-60'].qty, st['30-60'].fat, st['60-90'].qty, st['60-90'].fat, // W-AB
      sf['0-30'].qty, sf['0-30'].fat, sf['30-60'].qty, sf['30-60'].fat, sf['60-90'].qty, sf['60-90'].fat, // AC-AH
      ss2['0-30'].qty, ss2['0-30'].fat, ss2['30-60'].qty, ss2['30-60'].fat, ss2['60-90'].qty, ss2['60-90'].fat, // AI-AN
      g9.qty, g9.fat, g9.caixa, g9.margem, g9.pctFat, g9.pctAcum, g9.curva, // AO-AU
      g3.qty,  g3.fat,  g3.pctFat,  g3.pctAcum,  g3.curva,  // AV-AZ
      g36.qty, g36.fat, g36.pctFat, g36.pctAcum, g36.curva, // BA-BE
      g69.qty, g69.fat, g69.pctFat, g69.pctAcum, g69.curva, // BF-BJ
    ];
  });

  if (rows.length > 0) {
    aba.getRange(PRIMA, 1, rows.length, NCOLS).setValues(rows);
    aba.getRange(PRIMA, 1, rows.length, 1).setNumberFormat('@');
    [14,15,45,46,50,51,55,56,60,61].forEach(c =>
      aba.getRange(PRIMA, c, rows.length, 1).setNumberFormat('0.0%'));
    [5,7,9,11,12,13,18,20,22,24,26,28,30,32,34,36,38,40,42,43,44,49,54,59].forEach(c =>
      aba.getRange(PRIMA, c, rows.length, 1).setNumberFormat('R$ #,##0.00'));
  }

  // ── Mesclar cabeçalhos ────────────────────────────────────────
  aba.getRange(1,1,5,1).merge(); aba.getRange(1,2,5,1).merge(); aba.getRange(1,3,5,1).merge();
  aba.getRange(1,4,1,37).merge(); aba.getRange(1,41,1,22).merge();
  aba.getRange(2,4,1,13).merge(); aba.getRange(2,17,1,6).merge();
  aba.getRange(2,23,1,6).merge(); aba.getRange(2,29,1,6).merge();
  aba.getRange(2,35,1,6).merge(); aba.getRange(2,41,1,22).merge();
  aba.getRange(3,4,1,6).merge();  aba.getRange(3,10,1,7).merge();
  aba.getRange(3,17,1,6).merge(); aba.getRange(3,23,1,6).merge();
  aba.getRange(3,29,1,6).merge(); aba.getRange(3,35,1,6).merge();
  aba.getRange(3,41,1,7).merge(); aba.getRange(3,48,1,22).merge();
  aba.getRange(4,4,1,2).merge();  aba.getRange(4,6,1,2).merge();
  aba.getRange(4,8,1,2).merge();  aba.getRange(4,10,1,7).merge();
  aba.getRange(4,17,1,2).merge(); aba.getRange(4,19,1,2).merge(); aba.getRange(4,21,1,2).merge();
  aba.getRange(4,23,1,2).merge(); aba.getRange(4,25,1,2).merge(); aba.getRange(4,27,1,2).merge();
  aba.getRange(4,29,1,2).merge(); aba.getRange(4,31,1,2).merge(); aba.getRange(4,33,1,2).merge();
  aba.getRange(4,35,1,2).merge(); aba.getRange(4,37,1,2).merge(); aba.getRange(4,39,1,2).merge();
  aba.getRange(4,41,1,7).merge();
  aba.getRange(4,48,1,5).merge(); aba.getRange(4,53,1,5).merge(); aba.getRange(4,58,1,5).merge();

  // ── Estilo ────────────────────────────────────────────────────
  aba.getRange(1,1,NHDR,NCOLS)
     .setBackground('#1c2c4a').setFontColor('#3a8dff')
     .setFontWeight('bold').setHorizontalAlignment('center')
     .setVerticalAlignment('middle');

  aba.setFrozenRows(NHDR);
  aba.setFrozenColumns(3);

  SpreadsheetApp.flush();
}

function _ncArr(n) { return Array(n).fill(''); }
