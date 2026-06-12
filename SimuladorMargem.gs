// ============================================================
// SIMULADOR DE MARGEM — v1
// Calcula impacto de aumento de custo na margem por SKU/período
// Entradas: col A = SKU (texto), col B = custo novo
// ============================================================

const FMCFG = {
  ID_GE:  '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE: 'Dados Completos',

  // "Dados Completos" — índices de coluna (base 0)
  // SKUs na col B estão armazenados como TEXTO — comparação sempre via String()
  GE_SKU:    1,   // B — texto
  GE_QTD:    2,   // C
  GE_DATA:   3,   // D
  GE_CANAL:  7,   // H
  GE_CUSTO:  11,  // L
  GE_VENDA:  25,  // Z
  GE_LIQ:    44,  // AS (valor líquido, já com imposto deduzido)
  GE_MARGEM: 45,  // AT

  HEADER_ROWS:    3,
  NCOLS:          26,  // A–Z
  DATA_START_COL: 3,   // C (A e B são entrada do usuário)
};

const FM_CANAIS_ML = [
  'ML Edmotos', 'ML Edmotos FULL', 'ML Humble', 'ML Humble FULL',
  'ML Najumi',  'ML Najumi FULL',  'ML Sky',    'ML Sky FULL',
  'ML Moto Vibe', 'ML Moto Vibe FULL',
];
const FM_CANAIS_SHOPEE = [
  'Shopee Humble',
  'Shopee Najumi', 'Shopee Najumi FULL',
  'Shopee Sky',
];

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Simulador de Margem')
    .addItem('📊 Calcular — Geral',  'calcularGeral')
    .addItem('🟠 Calcular — Shopee', 'calcularShopee')
    .addItem('🟡 Calcular — ML',     'calcularML')
    .addToUi();
}

function calcularGeral()  { _fm_calcular('Geral',  null);             }
function calcularShopee() { _fm_calcular('Shopee', FM_CANAIS_SHOPEE); }
function calcularML()     { _fm_calcular('ML',     FM_CANAIS_ML);     }

// ============================================================
// LÓGICA PRINCIPAL
// ============================================================
function _fm_calcular(nomeAba, canaisFiltro) {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Calculando "' + nomeAba + '"... Aguarde.');

  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) { ui.alert('❌ Aba "' + nomeAba + '" não encontrada.'); return; }

    const PRIMA = FMCFG.HEADER_ROWS + 1;
    const lastRow = aba.getLastRow();

    if (lastRow < PRIMA) {
      ui.alert('Nenhum SKU encontrado. Digite os SKUs na coluna A a partir da linha ' + PRIMA + '.');
      return;
    }

    const nRows  = lastRow - PRIMA + 1;
    const inputs = aba.getRange(PRIMA, 1, nRows, 2).getValues();

    // Carregar e indexar dados do GE Finance
    const { dados, diasPorLinha } = _fm_carregarGE();

    // Recria cabeçalho e limpa colunas C–Z dos dados
    _fm_escreverCabecalho(aba, nomeAba);
    aba.getRange(PRIMA, FMCFG.DATA_START_COL, nRows, FMCFG.NCOLS - 2).clearContent();

    // Calcular linha a linha
    const resultRows = [];
    let calculados = 0;

    inputs.forEach(([skuRaw, custoNovoRaw]) => {
      const sku      = String(skuRaw     || '').trim();
      const custoNovo = Number(custoNovoRaw) || 0;
      if (!sku) { resultRows.push(Array(FMCFG.NCOLS - 2).fill('')); return; }
      resultRows.push(_fm_calcularSku(dados, diasPorLinha, sku, custoNovo, canaisFiltro));
      calculados++;
    });

    if (resultRows.length > 0) {
      aba.getRange(PRIMA, FMCFG.DATA_START_COL, resultRows.length, FMCFG.NCOLS - 2)
         .setValues(resultRows);

      // Percentuais: H(8) J(10) P(16) R(18) X(24) Z(26)
      [8, 10, 16, 18, 24, 26].forEach(col => {
        aba.getRange(PRIMA, col, resultRows.length, 1).setNumberFormat('0.00%');
      });

      // Moeda: custo médio, preço venda, fat, margem R$
      [3,4,6,7,9, 11,12,14,15,17, 19,20,22,23,25].forEach(col => {
        aba.getRange(PRIMA, col, resultRows.length, 1).setNumberFormat('R$ #,##0.00');
      });

      // Colorir Margem C/ Aumento %
      _fm_colorirMargemNova(aba, resultRows, PRIMA);
    }

    SpreadsheetApp.flush();
    ui.alert('✅ "' + nomeAba + '" calculada! ' + calculados + ' SKUs processados.');
  } catch (e) {
    ui.alert('❌ Erro: ' + e.message + '\n\nStack:\n' + e.stack);
  }
}

// ============================================================
// CARREGAR GE FINANCE
// ============================================================
function _fm_carregarGE() {
  const geSS  = SpreadsheetApp.openById(FMCFG.ID_GE);
  const geAba = geSS.getSheetByName(FMCFG.ABA_GE);
  if (!geAba) throw new Error('Aba "' + FMCFG.ABA_GE + '" não encontrada.');
  const dados = geAba.getDataRange().getValues();

  let dataMax = new Date(0);
  for (let i = 1; i < dados.length; i++) {
    const d = _fm_parseData(dados[i][FMCFG.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(0, 0, 0, 0);

  const diasPorLinha = dados.map((row, i) => {
    if (i === 0) return -1;
    const d = _fm_parseData(row[FMCFG.GE_DATA]);
    if (!d) return -1;
    const dNorm = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return Math.floor((dataMax - dNorm) / 86400000);
  });

  return { dados, diasPorLinha };
}

// ============================================================
// CALCULAR MÉTRICAS PARA UM SKU
// Retorna array de 24 valores (8 por período × 3 períodos)
// ============================================================
function _fm_calcularSku(dados, diasPorLinha, skuBusca, custoNovo, canaisFiltro) {
  // Acumuladores por período: 0-30 / 30-60 / 60-90
  const acc = [
    { qtd:0, venda:0, custo:0, liq:0, margem:0 },
    { qtd:0, venda:0, custo:0, liq:0, margem:0 },
    { qtd:0, venda:0, custo:0, liq:0, margem:0 },
  ];

  for (let i = 1; i < dados.length; i++) {
    // SKU na planilha é TEXTO — comparação String × String
    const skuGE = String(dados[i][FMCFG.GE_SKU] || '').trim();
    if (skuGE !== skuBusca) continue;

    const dias = diasPorLinha[i];
    if (dias < 0 || dias >= 90) continue;

    if (canaisFiltro) {
      const canal = String(dados[i][FMCFG.GE_CANAL] || '').trim();
      if (!canaisFiltro.some(c => c.toLowerCase() === canal.toLowerCase())) continue;
    }

    const qtd    = Number(dados[i][FMCFG.GE_QTD]    || 0);
    const venda  = Number(dados[i][FMCFG.GE_VENDA]  || 0);
    const custo  = Number(dados[i][FMCFG.GE_CUSTO]  || 0);
    const liq    = Number(dados[i][FMCFG.GE_LIQ]    || 0);
    const margem = Number(dados[i][FMCFG.GE_MARGEM] || 0);

    const p = dias < 30 ? 0 : dias < 60 ? 1 : 2;
    acc[p].qtd    += qtd;
    acc[p].venda  += venda;
    acc[p].custo  += custo;   // sum(col L) / qtd = custo médio
    acc[p].liq    += liq;
    acc[p].margem += margem;
  }

  const cols = [];
  acc.forEach(a => {
    const custoMedio    = a.qtd > 0 ? a.custo  / a.qtd : 0;
    const vendaMedio    = a.qtd > 0 ? a.venda  / a.qtd : 0;
    const margemPct     = a.venda > 0 ? a.margem / a.venda : 0;
    // Margem c/ aumento: valor líquido (já sem imposto) menos custo novo × qtd
    const margemNova    = a.liq - (a.qtd * custoNovo);
    const margemNovaPct = a.venda > 0 ? margemNova / a.venda : 0;

    cols.push(
      custoMedio,     // Custo Médio
      vendaMedio,     // Preço Venda Médio
      a.qtd,          // Qntd. Vend.
      a.venda,        // Fat. R$
      a.margem,       // Margem R$ atual
      margemPct,      // Margem % atual
      margemNova,     // Margem C/ Aumento R$
      margemNovaPct   // Margem C/ Aumento %
    );
  });

  return cols; // 24 valores
}

// ============================================================
// ESCREVER + ESTILIZAR CABEÇALHO
// Layout A-Z (26 colunas):
//  A  = SKU          B  = Custo Novo
//  C-J = 0-30        K-R = 30-60       S-Z = 60-90
//  Dentro de cada período:
//    +0 Custo Médio  +1 Preço Venda    +2 Qntd.    +3 Fat. R$
//    +4 Margem R$    +5 Margem %       +6 Marg.Aum R$  +7 Marg.Aum %
// ============================================================
function _fm_escreverCabecalho(aba, nomeAba) {
  const NCOLS  = FMCFG.NCOLS;
  const HEADER = FMCFG.HEADER_ROWS;

  aba.getRange(1, 1, HEADER, NCOLS).clearContent().clearFormat();

  const h1 = Array(NCOLS).fill('');
  const h2 = Array(NCOLS).fill('');
  const h3 = Array(NCOLS).fill('');

  // Linha 1: título da aba
  h1[0] = 'Simulador de Margem — ' + nomeAba;

  // Linha 2: grupos (base 0 → índice = col − 1)
  h2[0]  = 'Entradas';
  h2[2]  = '0-30 — Histórico';
  h2[8]  = '0-30 — C/ Aumento';
  h2[10] = '30-60 — Histórico';
  h2[16] = '30-60 — C/ Aumento';
  h2[18] = '60-90 — Histórico';
  h2[24] = '60-90 — C/ Aumento';

  // Linha 3: nomes das colunas
  h3[0]  = 'SKU';
  h3[1]  = 'Custo Novo';
  // 0-30
  h3[2]  = 'Custo Médio';      h3[3]  = 'Preço Venda Médio';
  h3[4]  = 'Qntd. Vend.';     h3[5]  = 'Fat. R$';
  h3[6]  = 'Margem R$';       h3[7]  = 'Margem %';
  h3[8]  = 'Margem R$ (novo custo)'; h3[9]  = 'Margem % (novo custo)';
  // 30-60
  h3[10] = 'Custo Médio';     h3[11] = 'Preço Venda Médio';
  h3[12] = 'Qntd. Vend.';    h3[13] = 'Fat. R$';
  h3[14] = 'Margem R$';      h3[15] = 'Margem %';
  h3[16] = 'Margem R$ (novo custo)'; h3[17] = 'Margem % (novo custo)';
  // 60-90
  h3[18] = 'Custo Médio';     h3[19] = 'Preço Venda Médio';
  h3[20] = 'Qntd. Vend.';    h3[21] = 'Fat. R$';
  h3[22] = 'Margem R$';      h3[23] = 'Margem %';
  h3[24] = 'Margem R$ (novo custo)'; h3[25] = 'Margem % (novo custo)';

  aba.getRange(1, 1, 3, NCOLS).setValues([h1, h2, h3]);

  // ── Cores por grupo ──
  const grupos = [
    { de:1,  ate:2,  bgTop:'#546e7a', bgBot:'#263238' }, // Entradas
    { de:3,  ate:8,  bgTop:'#1976d2', bgBot:'#0d47a1' }, // 0-30 histórico
    { de:9,  ate:10, bgTop:'#e53935', bgBot:'#b71c1c' }, // 0-30 c/ aumento
    { de:11, ate:16, bgTop:'#00897b', bgBot:'#004d40' }, // 30-60 histórico
    { de:17, ate:18, bgTop:'#e53935', bgBot:'#b71c1c' }, // 30-60 c/ aumento
    { de:19, ate:24, bgTop:'#8e24aa', bgBot:'#4a148c' }, // 60-90 histórico
    { de:25, ate:26, bgTop:'#e53935', bgBot:'#b71c1c' }, // 60-90 c/ aumento
  ];

  grupos.forEach(g => {
    const w = g.ate - g.de + 1;
    aba.getRange(1, g.de, 2, w)
       .setBackground(g.bgTop).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
    aba.getRange(3, g.de, 1, w)
       .setBackground(g.bgBot).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center')
       .setVerticalAlignment('middle').setWrap(true);
  });

  // ── Merges ──
  aba.getRange(1, 1, 1, NCOLS).merge();  // título
  aba.getRange(2, 1,  1, 2).merge();     // Entradas
  aba.getRange(2, 3,  1, 6).merge();     // 0-30 histórico (C-H)
  aba.getRange(2, 9,  1, 2).merge();     // 0-30 c/ aumento (I-J)
  aba.getRange(2, 11, 1, 6).merge();     // 30-60 histórico (K-P)
  aba.getRange(2, 17, 1, 2).merge();     // 30-60 c/ aumento (Q-R)
  aba.getRange(2, 19, 1, 6).merge();     // 60-90 histórico (S-X)
  aba.getRange(2, 25, 1, 2).merge();     // 60-90 c/ aumento (Y-Z)

  // ── Dimensões ──
  aba.setRowHeight(1, 30);
  aba.setRowHeight(2, 25);
  aba.setRowHeight(3, 50);
  aba.setColumnWidth(1, 120); // SKU
  aba.setColumnWidth(2, 90);  // Custo Novo
  [3,11,19].forEach(c => aba.setColumnWidth(c, 90));  // Custo Médio
  [4,12,20].forEach(c => aba.setColumnWidth(c, 110)); // Preço Venda
  [5,13,21].forEach(c => aba.setColumnWidth(c, 70));  // Qntd
  [9,10,17,18,25,26].forEach(c => aba.setColumnWidth(c, 110)); // Margem nova

  aba.setFrozenRows(FMCFG.HEADER_ROWS);
}

// ============================================================
// COLORIR MARGEM C/ AUMENTO %
// verde ≥ 15% | amarelo 10-15% | vermelho < 10%
// ============================================================
function _fm_colorirMargemNova(aba, rows, prima) {
  // Colunas de Margem C/ Aumento %: J=10, R=18, Z=26
  [10, 18, 26].forEach(col => {
    const verde = [], amarelo = [], vermelho = [];
    const idxRow = col - FMCFG.DATA_START_COL; // índice no array resultRows
    const letter = _fm_colLetter(col);

    rows.forEach((row, i) => {
      const val = row[idxRow];
      if (val === '' || val === null || val === undefined) return;
      const a1 = letter + (prima + i);
      if      (val >= 0.15) verde.push(a1);
      else if (val >= 0.10) amarelo.push(a1);
      else                  vermelho.push(a1);
    });

    if (verde.length)    aba.getRangeList(verde).setBackground('#c8e6c9').setFontColor('#1b5e20');
    if (amarelo.length)  aba.getRangeList(amarelo).setBackground('#fff9c4').setFontColor('#f57f17');
    if (vermelho.length) aba.getRangeList(vermelho).setBackground('#ffcdd2').setFontColor('#b71c1c');
  });
}

// ============================================================
// HELPERS
// ============================================================
function _fm_parseData(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function _fm_colLetter(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
