// ============================================================
// GIRO DE ESTOQUE — v3
// 3 abas: Menor Giro / Maior Giro / Maior Estoque
// Apenas SKUs Simples/PAI desmembrados (Composto excluído)
// ============================================================

const GCFG = {
  ID_GE:      '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE:     'Dados Completos',
  ID_BL:      '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL:     'estoque',
  ID_CUSTOS:  '1C9Z4vT9SaamEGJ_hCWXdOm8vWOMhxw_vAb049mvtrP0',
  ABA_CUSTOS: 'Soma Composições',
  ID_PEND:    '1bYe0kxiHCUzPPHszQiK8JarKqzj5lxijAy7M3zpy1mg',

  GE_SKU: 1, GE_QTD: 2, GE_DATA: 3, GE_FAT: 21,
  BL_SKU: 1, BL_COD_EST: 2, BL_QTDE_EST: 3,
  BL_SALDO_PAD: 4, BL_SALDO_ARM: 5, BL_SALDO_CHG: 6,
  BL_TIPO: 7,
  CUST_SKU: 0, CUST_VALOR: 1,
  PEND_SKU: 1, PEND_QTD: 2,

  ABC_A: 0.70,
  ABC_B: 0.90,
  HEADER_ROWS: 3,
  NCOLS: 29,
};

// Paleta de cores por grupo de colunas
const GE_CORES = [
  { de:1,  ate:5,  bgTop:'#455a64', bgBot:'#263238', fg:'#ffffff' }, // Info base
  { de:6,  ate:8,  bgTop:'#1976d2', bgBot:'#0d47a1', fg:'#ffffff' }, // Qtd. Vendida
  { de:9,  ate:11, bgTop:'#e53935', bgBot:'#b71c1c', fg:'#ffffff' }, // Diferença R$
  { de:12, ate:12, bgTop:'#8e24aa', bgBot:'#4a148c', fg:'#ffffff' }, // Custo Estoque
  { de:13, ate:20, bgTop:'#43a047', bgBot:'#1b5e20', fg:'#ffffff' }, // Curva ABC
  { de:21, ate:28, bgTop:'#00897b', bgBot:'#004d40', fg:'#ffffff' }, // Curva AABBCC
  { de:29, ate:29, bgTop:'#6d4c41', bgBot:'#3e2723', fg:'#ffffff' }, // Compras Pendentes
];

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Giro de Estoque')
    .addItem('📉 Menor Giro',    'gerarMenorGiro')
    .addItem('📈 Maior Giro',    'gerarMaiorGiro')
    .addItem('📦 Maior Estoque', 'gerarMaiorEstoque')
    .addSeparator()
    .addItem('🔄 Todas as Abas', 'gerarGiroEstoque')
    .addSeparator()
    .addItem('🔁 Atualizar Estoque BaseLinker', 'bl_atualizarEstoqueNaPlanilha')
    .addToUi();
}

// ============================================================
// FUNÇÕES PÚBLICAS — uma por aba para não estourar o tempo
// ============================================================
function gerarMenorGiro()    { _ge_gerarUmaAba('Giro — Menor Giro',    (a, b) => a.dijSum  - b.dijSum);  }
function gerarMaiorGiro()    { _ge_gerarUmaAba('Giro — Maior Giro',    (a, b) => b.dijSum  - a.dijSum);  }
function gerarMaiorEstoque() { _ge_gerarUmaAba('Giro — Maior Estoque', (a, b) => b.estPend - a.estPend); }

function gerarGiroEstoque() {
  gerarMenorGiro();
  gerarMaiorGiro();
  gerarMaiorEstoque();
}

function _ge_gerarUmaAba(nomeAba, sortFn) {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando "' + nomeAba + '"... Aguarde.');
  try {
    const mapaCustos                             = _ge_carregarCustos();
    const { mapaCompostos, proporcoes, estoque } = _ge_carregarBL(mapaCustos);
    const pivotPend                              = _ge_carregarPendentes();
    const vendas                                 = _ge_carregarVendas(mapaCompostos, proporcoes);
    const itens                                  = _ge_montarItens(vendas, estoque, mapaCustos, pivotPend, mapaCompostos);
    _ge_calcularCurvas(itens);
    _ge_escreverAba(nomeAba, itens, sortFn);
    ui.alert('✅ "' + nomeAba + '" atualizada! ' + itens.length + ' SKUs.');
  } catch (e) {
    ui.alert('❌ Erro: ' + e.message + '\n\nStack:\n' + e.stack);
  }
}

// ============================================================
// PARSE DD/MM/YYYY
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
// HELPER: converte número de coluna em letra (A, B, ... AA, AB...)
// ============================================================
function _ge_colLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
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
// CARREGAR CUSTOS
// ============================================================
function _ge_carregarCustos() {
  const ss  = SpreadsheetApp.openById(GCFG.ID_CUSTOS);
  const aba = _ge_getAba(ss, GCFG.ABA_CUSTOS);
  if (!aba) throw new Error('Aba "' + GCFG.ABA_CUSTOS + '" não encontrada.');
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
  const dados         = aba.getDataRange().getValues();
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
      estoque[sku] = (estoque[sku] || 0) + saldo;
    }
  }

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
// CARREGAR VENDAS + DESMEMBRAMENTO DE KITS
// ============================================================
function _ge_carregarVendas(mapaCompostos, proporcoes) {
  const geSS  = SpreadsheetApp.openById(GCFG.ID_GE);
  const geAba = _ge_getAba(geSS, GCFG.ABA_GE);
  if (!geAba) throw new Error('Aba GE Finance "' + GCFG.ABA_GE + '" não encontrada.');
  const dados = geAba.getDataRange().getValues();

  let dataMax = new Date(0);
  for (let i = 1; i < dados.length; i++) {
    const d = _ge_parseData(dados[i][GCFG.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(0, 0, 0, 0);

  const vendas = {};

  function acum(sku, dataVenda, dias, qtd, fat) {
    if (!vendas[sku]) {
      vendas[sku] = { qtd030: 0, qtd060: 0, qtd090: 0, fat030: 0, fat3060: 0, fat6090: 0, ultimaVenda: null };
    }
    const v = vendas[sku];
    if      (dias < 30) { v.qtd030 += qtd; v.qtd060 += qtd; v.qtd090 += qtd; }
    else if (dias < 60) {                  v.qtd060 += qtd; v.qtd090 += qtd; }
    else                {                                    v.qtd090 += qtd; }
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
// MONTAR ITENS — exclui chaves do mapaCompostos (SKUs de kit)
// ============================================================
function _ge_montarItens(vendas, estoque, mapaCustos, pivotPend, mapaCompostos) {
  const hoje    = new Date();
  hoje.setHours(0, 0, 0, 0);
  const allSkus = new Set([...Object.keys(vendas), ...Object.keys(estoque)]);
  const itens   = [];

  allSkus.forEach(sku => {
    if (mapaCompostos[sku]) return; // Composto: excluir

    const v     = vendas[sku] || { qtd030: 0, qtd060: 0, qtd090: 0, fat030: 0, fat3060: 0, fat6090: 0, ultimaVenda: null };
    const est   = estoque[sku] !== undefined ? estoque[sku] : 0;
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
      qtd030: v.qtd030, qtd060: v.qtd060, qtd090: v.qtd090,
      i_030, i_060, i_090,
      custoEst: est > 0 ? est * custo : 0,
      fat030: v.fat030, fat3060: v.fat3060, fat6090: v.fat6090, fat090,
      pend,
      dijSum:  i_030 + i_060 + i_090,
      estPend: est + pend,
    });
  });

  return itens;
}

// ============================================================
// CURVAS ABC + AABBCC
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

    let cum = 0;
    sorted.forEach(item => {
      cum += item[fatKey];
      const pct   = total > 0 ? cum / total : 1;
      const letra = pct <= GCFG.ABC_A ? 'A' : pct <= GCFG.ABC_B ? 'B' : 'C';
      item[abcKey] = { letra, pct };
    });

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
// ============================================================
function _ge_escreverAba(nomeAba, itens, sortFn) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);

  const PRIMA = GCFG.HEADER_ROWS + 1;
  const NCOLS = GCFG.NCOLS;

  // Cabeçalho: limpa totalmente (é 100% do script)
  aba.getRange(1, 1, GCFG.HEADER_ROWS, NCOLS + 2).clearContent().clearFormat();

  // Dados: limpa só o conteúdo (preserva formatação manual do usuário)
  // Reseta o formato apenas nas colunas que o script colore
  const lastRow = aba.getLastRow();
  if (lastRow >= PRIMA) {
    const nRows = lastRow - PRIMA + 1;
    aba.getRange(PRIMA, 1, nRows, NCOLS).clearContent();
    // E(5) = custo vermelho | M(13) O(15) Q(17) S(19) = ABC | U(21) W(23) Y(25) AA(27) = AABBCC
    [5, 13, 15, 17, 19, 21, 23, 25, 27].forEach(col => {
      aba.getRange(PRIMA, col, nRows, 1).setBackground(null).setFontColor(null);
    });
  }

  // ── Cabeçalho 3 linhas ──
  const h1 = Array(NCOLS).fill('');
  const h2 = Array(NCOLS).fill('');
  const h3 = Array(NCOLS).fill('');

  // Linha 1: grupos
  h1[0]  = 'Informações do Produto';
  h1[5]  = 'Qtd. Vendida';
  h1[8]  = 'Diferença em R$';
  h1[11] = 'Custo Estoque';
  h1[12] = 'Curva ABC';
  h1[20] = 'Curva AABBCC';
  h1[28] = 'Compras Pendentes';

  // Linha 2: períodos
  h2[5]  = '0-30';  h2[6]  = '0-60';  h2[7]  = '0-90';
  h2[8]  = '0-30';  h2[9]  = '0-60';  h2[10] = '0-90';
  h2[12] = '0-30';  h2[14] = '30-60'; h2[16] = '60-90'; h2[18] = '0-90';
  h2[20] = '0-30';  h2[22] = '30-60'; h2[24] = '60-90'; h2[26] = '0-90';

  // Linha 3: nomes das colunas
  h3[0]  = 'SKU';           h3[1]  = 'Estoque';
  h3[2]  = 'Dt. Últ. Venda'; h3[3]  = 'Dias';      h3[4]  = 'Custo Unit.';
  h3[5]  = 'Qtd.';          h3[6]  = 'Qtd.';       h3[7]  = 'Qtd.';
  h3[8]  = 'Dif. R$';       h3[9]  = 'Dif. R$';    h3[10] = 'Dif. R$';
  h3[11] = 'Custo Est.';
  h3[12] = 'Curva'; h3[13] = '%'; h3[14] = 'Curva'; h3[15] = '%';
  h3[16] = 'Curva'; h3[17] = '%'; h3[18] = 'Curva'; h3[19] = '%';
  h3[20] = 'Curva'; h3[21] = '%'; h3[22] = 'Curva'; h3[23] = '%';
  h3[24] = 'Curva'; h3[25] = '%'; h3[26] = 'Curva'; h3[27] = '%';
  h3[28] = 'Qtd.';

  aba.getRange(1, 1, 3, NCOLS).setValues([h1, h2, h3]);

  // Data/hora de atualização (coluna AE = 31)
  aba.getRange(1, 31).setValue('Data Atualização');
  aba.getRange(2, 31).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  );

  _ge_estilizarCabecalho(aba);

  // ── Dados ──
  const sorted = [...itens].sort(sortFn);
  if (sorted.length === 0) { SpreadsheetApp.flush(); return; }

  // Formata coluna A como texto ANTES de escrever (evita conversão numérica de SKUs)
  aba.getRange(PRIMA, 1, sorted.length, 1).setNumberFormat('@');

  const tz   = Session.getScriptTimeZone();
  const rows = sorted.map(item => [
    String(item.sku),                                                        // A
    item.est,                                                                // B
    item.ultimaVenda
      ? Utilities.formatDate(item.ultimaVenda, tz, 'dd/MM/yyyy') : '',      // C
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

  // Percentuais: N(14), P(16), R(18), T(20), V(22), X(24), Z(26), AB(28)
  [14, 16, 18, 20, 22, 24, 26, 28].forEach(col => {
    aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('0.0%');
  });

  // Vermelho col E onde custo ≤ 0
  const redA1 = [];
  sorted.forEach((item, i) => {
    if (item.custo <= 0) redA1.push('E' + (PRIMA + i));
  });
  if (redA1.length) aba.getRangeList(redA1).setBackground('#ffcdd2');

  _ge_colorirABC(aba, sorted, PRIMA,
    [13, 15, 17, 19], ['abc030', 'abc3060', 'abc6090', 'abc090']);
  _ge_colorirAABBCC(aba, sorted, PRIMA,
    [21, 23, 25, 27], ['aabbcc030', 'aabbcc3060', 'aabbcc6090', 'aabbcc090']);

  SpreadsheetApp.flush();
}

// ============================================================
// ESTILIZAR CABEÇALHO (cores, merges, freeze)
// ============================================================
function _ge_estilizarCabecalho(aba) {
  const HEADER = GCFG.HEADER_ROWS;

  // Cores por grupo
  GE_CORES.forEach(g => {
    const w = g.ate - g.de + 1;
    // Linhas 1-2: tom mais claro
    aba.getRange(1, g.de, 2, w)
       .setBackground(g.bgTop).setFontColor(g.fg)
       .setFontWeight('bold').setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
    // Linha 3 (nomes): tom mais escuro
    aba.getRange(3, g.de, 1, w)
       .setBackground(g.bgBot).setFontColor(g.fg)
       .setFontWeight('bold').setHorizontalAlignment('center')
       .setVerticalAlignment('middle').setWrap(true);
  });

  // Coluna AE (31) — Data Atualização
  aba.getRange(1, 31, 3, 1)
     .setBackground('#263238').setFontColor('#ffffff')
     .setFontWeight('bold').setHorizontalAlignment('center')
     .setVerticalAlignment('middle');

  // Merges linha 1 — grupos
  aba.getRange(1, 1,  1, 5).merge();  // Informações do Produto
  aba.getRange(1, 6,  1, 3).merge();  // Qtd. Vendida
  aba.getRange(1, 9,  1, 3).merge();  // Diferença em R$
  aba.getRange(1, 13, 1, 8).merge();  // Curva ABC
  aba.getRange(1, 21, 1, 8).merge();  // Curva AABBCC

  // Merges linha 2 — pares de período (Curva/%) para ABC e AABBCC
  [[13,2],[15,2],[17,2],[19,2],   // ABC: 0-30 | 30-60 | 60-90 | 0-90
   [21,2],[23,2],[25,2],[27,2],   // AABBCC: idem
  ].forEach(([c, w]) => aba.getRange(2, c, 1, w).merge());

  // Merges linha 2 — info base (A-E, L) e compras (AC)
  aba.getRange(2, 1, 1, 5).merge();
  aba.getRange(2, 6, 1, 1);  // sem merge — já tem período
  aba.getRange(2, 12, 1, 1); // Custo Estoque — sem merge
  aba.getRange(2, 29, 1, 1); // Compras Pendentes — sem merge

  // Alturas de linha
  aba.setRowHeight(1, 28);
  aba.setRowHeight(2, 28);
  aba.setRowHeight(3, 40);

  // Larguras de coluna
  aba.setColumnWidth(1, 110); // SKU
  aba.setColumnWidth(3, 105); // Dt. Última Venda
  aba.setColumnWidth(5, 85);  // Custo Unit.
  aba.setColumnWidth(12, 90); // Custo Est.

  // Congelar cabeçalho
  aba.setFrozenRows(HEADER);

  // Borda inferior no cabeçalho
  aba.getRange(HEADER, 1, 1, GCFG.NCOLS)
     .setBorder(null, null, true, null, null, null, '#000000',
                SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// ============================================================
// COLORIR ABC — constrói A1 notation sem chamar getRange por célula
// ============================================================
function _ge_colorirABC(aba, itens, prima, cols, keys) {
  const BG = { 'A': '#c8e6c9', 'B': '#fff9c4', 'C': '#ffcdd2' };
  const FG = { 'A': '#1b5e20', 'B': '#f57f17', 'C': '#b71c1c' };

  cols.forEach((col, idx) => {
    const key    = keys[idx];
    const letter = _ge_colLetter(col);
    const grupos = { 'A': [], 'B': [], 'C': [] };
    itens.forEach((item, i) => {
      const l = item[key] && item[key].letra;
      if (l) grupos[l].push(letter + (prima + i));
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
    const letter = _ge_colLetter(col);
    const grupos = {};
    itens.forEach((item, i) => {
      const l = item[key] && item[key].letra;
      if (!l || !BG[l]) return;
      if (!grupos[l]) grupos[l] = [];
      grupos[l].push(letter + (prima + i));
    });
    Object.keys(grupos).forEach(l => {
      aba.getRangeList(grupos[l]).setBackground(BG[l]).setFontColor(FG[l]);
    });
  });
}
