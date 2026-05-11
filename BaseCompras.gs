// ============================================================
// BASE COMPRAS 2026 — v3
// Gera relatório de reposição com curvas ABC e urgência
// Aba "Sem Desmembramento": SKUs como no GE Finance (kits = PAI)
// Aba "Com Desmembramento": kits decompostos em simples
// Dados = GERAL (todas as contas somadas)
// ============================================================

const RCFG = {
  ABA_SEM: 'Sem Desmembramento',
  ABA_COM: 'Com Desmembramento',
  HEADER_ROWS: 5,   // dados começam na linha 6

  ID_GE:      '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE:     'Dados Completos',

  ID_BL:      '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL:     'estoque',

  ID_CUSTOS:  '1C9Z4vT9SaamEGJ_hCWXdOm8vWOMhxw_vAb049mvtrP0',
  ABA_CUSTOS: 'Soma Composições',

  // Fontes externas para estoque composto
  ID_PEND_MAP: '1xUNgkVTl74eMhj3W5XrOD6B0OxVJpnVXR_H_pgYRoh8',
  ID_FULL:     '1vjZX4PYVQy4G_K_03BJuAo_Div9B5NQ-iA2d0rK1dic',
  ID_COMP_PEN: '1bYe0kxiHCUzPPHszQiK8JarKqzj5lxijAy7M3zpy1mg',

  // Colunas das fontes externas (0-based): col B = SKU, col C = Qtd
  EXT_SKU: 1,
  EXT_QTD: 2,

  // GE Finance colunas (0-based)
  GE_PRODUTO: 0,
  GE_SKU:     1,
  GE_QTD:     2,
  GE_DATA:    3,
  GE_CANAL:   7,
  GE_FAT:    21,
  GE_CAIXA:  44,
  GE_MARGEM: 45,

  // BaseLinker colunas (0-based)
  BL_SKU:       1,
  BL_COD_EST:   2,
  BL_QTDE_EST:  3,
  BL_SALDO_PAD: 4,
  BL_SALDO_ARM: 5,
  BL_SALDO_CHG: 6,
  BL_TIPO:      7,
  BL_PRODUTO:   8,

  // Custos colunas (0-based)
  CUST_SKU:   0,
  CUST_VALOR: 1,

  MARGEM_MIN: 0.115,   // 11,5%
  ABC_A: 0.70,
  ABC_B: 0.90,
};

const CANAIS_ML     = ['ML Edmotos', 'ML Humble', 'ML Najumi', 'ML Sky', 'ML Moto Vibe'];
const CANAIS_SHOPEE = ['Shopee Humble', 'Shopee Najumi', 'Shopee Sky'];
const CANAIS_TODOS  = [...CANAIS_ML, ...CANAIS_SHOPEE];

const PER_SEM = [
  {de:0,ate:7}, {de:7,ate:15}, {de:15,ate:30},
  {de:30,ate:45}, {de:45,ate:60}, {de:60,ate:75}, {de:75,ate:90},
];
const PER_CON = [
  {nome:'0-30',  de:0,  ate:30},
  {nome:'30-60', de:30, ate:60},
  {nome:'60-90', de:60, ate:90},
];

const C2_PESO = {
  'AA':9,'AB':8,'AC':7,'BA':6,'BB':5,'BC':4,'CA':3,'CB':2,'CC':1,
};

// ============================================================
// HELPER: busca aba com tolerância a acentos corrompidos
// ============================================================
function _getAba(ss, nomeDesejado) {
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
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Base Compras')
    .addItem('📋 Gerar Sem Desmembramento', 'gerarSemDesmembramento')
    .addItem('🔀 Gerar Com Desmembramento', 'gerarComDesmembramento')
    .addSeparator()
    .addItem('🚀 Gerar Ambas as Abas',      'gerarAmbas')
    .addSeparator()
    .addItem('🔍 Diagnóstico de Abas',      'diagnosticarAbas')
    .addToUi();
}

// ============================================================
// DIAGNÓSTICO — lista todas as abas de cada planilha fonte
// ============================================================
function diagnosticarAbas() {
  const ui = SpreadsheetApp.getUi();
  let msg = '';

  const fontes = [
    { label: 'GE Finance',          id: RCFG.ID_GE,       abaNome: RCFG.ABA_GE     },
    { label: 'BaseLinker',          id: RCFG.ID_BL,       abaNome: RCFG.ABA_BL     },
    { label: 'Custos',              id: RCFG.ID_CUSTOS,   abaNome: RCFG.ABA_CUSTOS },
    { label: 'Pendente Mapeamento', id: RCFG.ID_PEND_MAP, abaNome: null            },
    { label: 'Full',                id: RCFG.ID_FULL,     abaNome: null            },
    { label: 'Compras Pendentes',   id: RCFG.ID_COMP_PEN, abaNome: null            },
  ];

  fontes.forEach(f => {
    try {
      const ss   = SpreadsheetApp.openById(f.id);
      const abas = ss.getSheets().map(s => '"' + s.getName() + '"').join(', ');
      if (f.abaNome) {
        const ok = _getAba(ss, f.abaNome) ? '✅' : '❌ NAO ENCONTRADA';
        msg += f.label + ' → aba: "' + f.abaNome + '" ' + ok + '\n';
      } else {
        msg += f.label + ' → usando 1ª aba: "' + ss.getSheets()[0].getName() + '" ✅\n';
      }
      msg += '  Abas: ' + abas + '\n\n';
    } catch(e) {
      msg += f.label + ' → ERRO ao abrir (ID errado?): ' + e.message + '\n\n';
    }
  });

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const own = ss.getSheets().map(s => '"' + s.getName() + '"').join(', ');
  msg += 'Esta planilha → abas: ' + own;

  ui.alert('🔍 Diagnóstico', msg, ui.ButtonSet.OK);
}

function gerarAmbas() {
  gerarSemDesmembramento();
  gerarComDesmembramento();
  SpreadsheetApp.getUi().alert('✅ Ambas as abas geradas com sucesso!');
}

// ============================================================
// CARREGAR PIVOT DE PLANILHA EXTERNA
// Lê col B (SKU) e col C (Qtd) da 1ª aba, soma por SKU.
// Se desmembrar=true e o SKU for kit, distribui aos componentes.
// ============================================================
function _carregarPivotExterno(id, desmembrar, mapaCompostos) {
  const ss    = SpreadsheetApp.openById(id);
  const aba   = ss.getSheets()[0];
  const dados = aba.getDataRange().getValues();
  const pivot = {};

  for (let i = 1; i < dados.length; i++) {
    const sku = String(dados[i][RCFG.EXT_SKU] || '').trim();
    const qtd = Number(dados[i][RCFG.EXT_QTD] || 0);
    if (!sku || !qtd) continue;

    if (desmembrar && mapaCompostos[sku]) {
      // Kit → distribui qtd × quantidade do componente no kit
      mapaCompostos[sku].forEach(comp => {
        pivot[comp.codEst] = (pivot[comp.codEst] || 0) + qtd * comp.qtde;
      });
    } else {
      // Simples ou kit sem desmembramento → acumula direto
      pivot[sku] = (pivot[sku] || 0) + qtd;
    }
  }

  return pivot;
}

// ============================================================
// GERAR SEM DESMEMBRAMENTO
// ============================================================
function gerarSemDesmembramento() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando Sem Desmembramento... Aguarde 1-2 minutos.');

  const blInfo  = _carregarBL();
  const estoque = _carregarEstoque(blInfo.blDados, false);
  const { vendas, prodInfo } = _carregarVendas(false, {});

  const pivotPend = _carregarPivotExterno(RCFG.ID_PEND_MAP, false, blInfo.mapaCompostos);
  const pivotFull = _carregarPivotExterno(RCFG.ID_FULL,     false, blInfo.mapaCompostos);
  const pivotComp = _carregarPivotExterno(RCFG.ID_COMP_PEN, false, blInfo.mapaCompostos);

  const itens    = _construirItens(vendas, prodInfo, estoque, blInfo.mapaBL_prod);
  const ordenado = _calcularTudo(itens);

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RCFG.ABA_SEM);
  if (!aba) { ui.alert('❌ Aba "' + RCFG.ABA_SEM + '" não encontrada.'); return; }
  _escreverAba(aba, ordenado, pivotPend, pivotFull, pivotComp);

  ui.alert('✅ Sem Desmembramento gerado! ' + ordenado.length + ' SKUs.');
}

// ============================================================
// GERAR COM DESMEMBRAMENTO
// ============================================================
function gerarComDesmembramento() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando Com Desmembramento... Aguarde 1-2 minutos.');

  const blInfo  = _carregarBL();
  const estoque = _carregarEstoque(blInfo.blDados, true);
  const { vendas, prodInfo } = _carregarVendas(true, blInfo);

  const pivotPend = _carregarPivotExterno(RCFG.ID_PEND_MAP, true, blInfo.mapaCompostos);
  const pivotFull = _carregarPivotExterno(RCFG.ID_FULL,     true, blInfo.mapaCompostos);
  const pivotComp = _carregarPivotExterno(RCFG.ID_COMP_PEN, true, blInfo.mapaCompostos);

  const itens    = _construirItens(vendas, prodInfo, estoque, blInfo.mapaBL_prod);
  const ordenado = _calcularTudo(itens);

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RCFG.ABA_COM);
  if (!aba) { ui.alert('❌ Aba "' + RCFG.ABA_COM + '" não encontrada.'); return; }
  _escreverAba(aba, ordenado, pivotPend, pivotFull, pivotComp);

  ui.alert('✅ Com Desmembramento gerado! ' + ordenado.length + ' SKUs.');
}

// ============================================================
// CARREGAR BASELINKER
// ============================================================
function _carregarBL() {
  const blSS  = SpreadsheetApp.openById(RCFG.ID_BL);
  const blAba = _getAba(blSS, RCFG.ABA_BL);
  if (!blAba) {
    const disponiveis = blSS.getSheets().map(s => '"' + s.getName() + '"').join(', ');
    throw new Error('Aba BaseLinker "' + RCFG.ABA_BL + '" nao encontrada.\nAbas disponiveis: ' + disponiveis);
  }
  const blDados = blAba.getDataRange().getValues();

  const mapaBL_prod   = {};
  const mapaCompostos = {};

  // Pass 1: Simples
  for (let i = 1; i < blDados.length; i++) {
    const row  = blDados[i];
    const sku  = String(row[RCFG.BL_SKU]     || '').trim();
    const tipo = String(row[RCFG.BL_TIPO]    || '').trim();
    const prod = String(row[RCFG.BL_PRODUTO] || '').trim();
    const saldo = (Number(row[RCFG.BL_SALDO_PAD] || 0) +
                   Number(row[RCFG.BL_SALDO_ARM] || 0) +
                   Number(row[RCFG.BL_SALDO_CHG] || 0));
    if (!sku || tipo !== 'Simples') continue;
    if (!mapaBL_prod[sku]) mapaBL_prod[sku] = { produto: prod, saldo };
  }

  // Pass 2: PAI (kit como unidade — usado em Sem Desmembramento)
  for (let i = 1; i < blDados.length; i++) {
    const row  = blDados[i];
    const sku  = String(row[RCFG.BL_SKU]     || '').trim();
    const tipo = String(row[RCFG.BL_TIPO]    || '').trim();
    const prod = String(row[RCFG.BL_PRODUTO] || '').trim();
    const saldo = (Number(row[RCFG.BL_SALDO_PAD] || 0) +
                   Number(row[RCFG.BL_SALDO_ARM] || 0) +
                   Number(row[RCFG.BL_SALDO_CHG] || 0));
    if (!sku || tipo !== 'PAI') continue;
    if (!mapaBL_prod[sku]) mapaBL_prod[sku] = { produto: prod, saldo };
  }

  // Pass 3: Composto → estrutura dos kits
  for (let i = 1; i < blDados.length; i++) {
    const row    = blDados[i];
    const sku    = String(row[RCFG.BL_SKU]      || '').trim();
    const codEst = String(row[RCFG.BL_COD_EST]  || '').trim();
    const qtde   = Number(row[RCFG.BL_QTDE_EST] || 0);
    const tipo   = String(row[RCFG.BL_TIPO]     || '').trim();
    if (!sku || tipo !== 'Composto' || !codEst) continue;
    if (!mapaCompostos[sku]) mapaCompostos[sku] = [];
    mapaCompostos[sku].push({ codEst, qtde });
    if (!mapaBL_prod[codEst]) mapaBL_prod[codEst] = { produto: '', saldo: 0 };
  }

  // Custos + Proporções para desmembramento de faturamento
  const custSS  = SpreadsheetApp.openById(RCFG.ID_CUSTOS);
  const custAba = _getAba(custSS, RCFG.ABA_CUSTOS);
  if (!custAba) {
    const disponiveis = custSS.getSheets().map(s => '"' + s.getName() + '"').join(', ');
    throw new Error('Aba Custos "' + RCFG.ABA_CUSTOS + '" nao encontrada.\nAbas disponiveis: ' + disponiveis);
  }
  const custDados  = custAba.getDataRange().getValues();
  const mapaCustos = {};
  for (let i = 1; i < custDados.length; i++) {
    const sku   = String(custDados[i][RCFG.CUST_SKU]   || '').trim();
    const custo = parseFloat(String(custDados[i][RCFG.CUST_VALOR] || '0').replace(',', '.')) || 0;
    if (sku && !mapaCustos[sku]) mapaCustos[sku] = custo;
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

  return { blDados, mapaBL_prod, mapaCompostos, proporcoes };
}

// ============================================================
// CARREGAR ESTOQUE
// ============================================================
function _carregarEstoque(blDados, apenasSimples) {
  const estoque = {};
  for (let i = 1; i < blDados.length; i++) {
    const row  = blDados[i];
    const sku  = String(row[RCFG.BL_SKU]  || '').trim();
    const tipo = String(row[RCFG.BL_TIPO] || '').trim();
    if (!sku) continue;
    if (tipo === 'Composto') continue;
    if (apenasSimples && tipo !== 'Simples') continue;
    const saldo = Number(row[RCFG.BL_SALDO_PAD] || 0) +
                  Number(row[RCFG.BL_SALDO_ARM] || 0) +
                  Number(row[RCFG.BL_SALDO_CHG] || 0);
    if (estoque[sku] === undefined) estoque[sku] = saldo;
  }
  return estoque;
}

// ============================================================
// CARREGAR VENDAS DO GE FINANCE
// ============================================================
function _carregarVendas(desmembrar, blInfo) {
  const geSS  = SpreadsheetApp.openById(RCFG.ID_GE);
  const geAba = _getAba(geSS, RCFG.ABA_GE);
  if (!geAba) {
    const disponiveis = geSS.getSheets().map(s => '"' + s.getName() + '"').join(', ');
    throw new Error('Aba GE Finance "' + RCFG.ABA_GE + '" nao encontrada.\nAbas disponiveis: ' + disponiveis);
  }
  const geDados = geAba.getDataRange().getValues();

  let dataMax = new Date(0);
  for (let i = 1; i < geDados.length; i++) {
    const d = new Date(geDados[i][RCFG.GE_DATA]);
    if (!isNaN(d) && d > dataMax) dataMax = d;
  }
  dataMax.setHours(23, 59, 59, 999);

  const diasPorLinha = geDados.map((row, i) => {
    if (i === 0) return -1;
    const d = new Date(row[RCFG.GE_DATA]);
    return isNaN(d) ? -1 : Math.floor((dataMax - d) / 86400000);
  });

  const mapaCompostos = (desmembrar && blInfo.mapaCompostos) || {};
  const proporcoes    = (desmembrar && blInfo.proporcoes)    || {};
  const mapaBL_prod   = (desmembrar && blInfo.mapaBL_prod)   || {};

  const vendas   = {};
  const prodInfo = {};

  function acum(sku, nomeProd, dias, qtd, fat, caixa, margem, canal) {
    if (!vendas[sku]) {
      vendas[sku]   = { qtdSem: [0,0,0,0,0,0,0], cons: _novoCons3() };
      prodInfo[sku] = nomeProd;
    }
    PER_SEM.forEach((p, idx) => {
      if (dias >= p.de && dias < p.ate) vendas[sku].qtdSem[idx] += qtd;
    });
    PER_CON.forEach(p => {
      if (dias < p.de || dias >= p.ate) return;
      const c = vendas[sku].cons[p.nome];
      c.qtd    += qtd;
      c.fat    += fat;
      c.caixa  += caixa;
      c.margem += margem;
      if (CANAIS_ML.includes(canal))     c.fatML     += fat;
      if (CANAIS_SHOPEE.includes(canal)) c.fatShopee += fat;
      c.porCanal[canal] = (c.porCanal[canal] || 0) + fat;
    });
  }

  for (let i = 1; i < geDados.length; i++) {
    const dias = diasPorLinha[i];
    if (dias < 0 || dias >= 90) continue;

    const skuGE    = String(geDados[i][RCFG.GE_SKU]     || '').trim();
    const nomeProd = String(geDados[i][RCFG.GE_PRODUTO]  || '').trim();
    const canal    = String(geDados[i][RCFG.GE_CANAL]    || '').trim();
    const qtd      = Number(geDados[i][RCFG.GE_QTD]      || 0);
    const fat      = Number(geDados[i][RCFG.GE_FAT]      || 0);
    const caixa    = Number(geDados[i][RCFG.GE_CAIXA]    || 0);
    const margem   = Number(geDados[i][RCFG.GE_MARGEM]   || 0);

    if (!skuGE) continue;
    const canalNorm = CANAIS_TODOS.find(c => c.toLowerCase() === canal.toLowerCase());
    if (!canalNorm) continue;

    if (desmembrar && mapaCompostos[skuGE]) {
      mapaCompostos[skuGE].forEach(comp => {
        const prop = (proporcoes[skuGE] && proporcoes[skuGE][comp.codEst]) || 0;
        const nome = (mapaBL_prod[comp.codEst] && mapaBL_prod[comp.codEst].produto) || '';
        acum(comp.codEst, nome, dias, qtd * comp.qtde, fat * prop, caixa * prop, margem * prop, canalNorm);
      });
    } else {
      acum(skuGE, nomeProd, dias, qtd, fat, caixa, margem, canalNorm);
    }
  }

  return { vendas, prodInfo };
}

function _novoCons3() {
  const novo = () => ({ qtd:0, fat:0, caixa:0, margem:0, fatML:0, fatShopee:0, porCanal:{},
                         curvaFat:'', curvaQtd:'', curva2:'', pctFat:0 });
  return { '0-30': novo(), '30-60': novo(), '60-90': novo() };
}

// ============================================================
// CONSTRUIR ARRAY DE ITENS
// ============================================================
function _construirItens(vendas, prodInfo, estoque, mapaBL_prod) {
  return Object.keys(vendas).map(sku => ({
    sku,
    produto:   prodInfo[sku] || (mapaBL_prod[sku] && mapaBL_prod[sku].produto) || '',
    estoque:   estoque[sku] !== undefined ? estoque[sku] : 0,
    qtdSem:    vendas[sku].qtdSem,
    cons:      vendas[sku].cons,
    _urgencia: 0,
  }));
}

// ============================================================
// CALCULAR ABC + CURVA 2 + URGÊNCIA + ORDENAR
// ============================================================
function _calcularTudo(itens) {
  if (itens.length === 0) return [];

  PER_CON.forEach(p => {
    const nome = p.nome;

    const byFat  = [...itens].sort((a, b) => b.cons[nome].fat - a.cons[nome].fat);
    const totFat = byFat.reduce((s, x) => s + x.cons[nome].fat, 0);
    let cumFat   = 0;
    byFat.forEach(item => {
      cumFat += item.cons[nome].fat;
      const pct = totFat > 0 ? cumFat / totFat : 1;
      item.cons[nome].curvaFat = pct <= RCFG.ABC_A ? 'A' : pct <= RCFG.ABC_B ? 'B' : 'C';
      item.cons[nome].pctFat   = pct;
    });

    const byQtd  = [...itens].sort((a, b) => b.cons[nome].qtd - a.cons[nome].qtd);
    const totQtd = byQtd.reduce((s, x) => s + x.cons[nome].qtd, 0);
    let cumQtd   = 0;
    byQtd.forEach(item => {
      cumQtd += item.cons[nome].qtd;
      const pct = totQtd > 0 ? cumQtd / totQtd : 1;
      item.cons[nome].curvaQtd = pct <= RCFG.ABC_A ? 'A' : pct <= RCFG.ABC_B ? 'B' : 'C';
    });

    itens.forEach(item => {
      const c = item.cons[nome];
      c.curva2 = (c.curvaFat || 'C') + (c.curvaQtd || 'C');
    });
  });

  itens.forEach(item => {
    const c = item.cons['0-30'];
    if (c.qtd === 0 && c.fat === 0) { item._urgencia = -99999; return; }
    const c2Peso    = C2_PESO[c.curva2] || 1;
    const dailyRate = (c.qtd / 30) || 0.001;
    const diasEst   = item.estoque / dailyRate;
    item._urgencia  = c2Peso * 10000 - diasEst;
  });

  itens.sort((a, b) => b._urgencia - a._urgencia);
  return itens;
}

// ============================================================
// ESCREVER ABA (49 colunas)
// ============================================================
// Layout das colunas:
//  A(1):   SKU
//  B-H(2-8): Qtd semanal 0-7, 7-15, 15-30, 30-45, 45-60, 60-75, 75-90
//  I(9):   Qtd 0-30
//  J(10):  Fat 0-30
//  K(11):  Caixa 0-30
//  L(12):  Margem R$ 0-30
//  M(13):  Margem % 0-30
//  N(14):  Qtd 30-60
//  O(15):  Fat 30-60
//  P(16):  Caixa 30-60
//  Q(17):  Margem R$ 30-60
//  R(18):  Margem % 30-60
//  S(19):  Qtd 60-90
//  T(20):  Fat 60-90
//  U(21):  Caixa 60-90
//  V(22):  Margem R$ 60-90
//  W(23):  Margem % 60-90
//  X(24):  % Fat ML 0-30
//  Y(25):  % Fat Shopee 0-30
//  Z(26):  Curva Fat 0-30
//  AA(27): % Acum Fat 0-30
//  AB(28): Curva Fat 30-60
//  AC(29): % Acum Fat 30-60
//  AD(30): Curva Fat 60-90
//  AE(31): % Acum Fat 60-90
//  AF(32): Curva 2 0-30
//  AG(33): Curva 2 30-60
//  AH(34): Curva 2 60-90
//  ── Estoque ──
//  AI(35): Estoque BL
//  AJ(36): Pendente Mapeamento
//  AK(37): Full
//  AL(38): Resultado ((Estoque + Pendente Mapeamento) - Full)
//  AM(39): Compras Pendentes
//  AN(40): Resultado Final ((Estoque + Pendente + Compras) - Full)
//  ── Analítico ──
//  AO(41): Conta que mais vende (0-30)
//  AP(42): Var% Fat 0-30 vs 30-60
//  ── Diferenças Entre Períodos (base = Resultado Final) ──
//  AQ(43): 0-7  − Resultado Final
//  AR(44): 0-30 − Resultado Final
//  AS(45): 30-60 − Resultado Final
//  AT(46): 60-90 − Resultado Final
//  ── Diferenças Períodos Somados (base = Resultado Final) ──
//  AU(47): Acum 0-30 − Resultado Final
//  AV(48): Acum 0-60 − Resultado Final
//  AW(49): Acum 0-90 − Resultado Final
// ============================================================
function _escreverAba(aba, itens, pivotPend, pivotFull, pivotComp) {
  const HEADER = RCFG.HEADER_ROWS;
  const PRIMA  = HEADER + 1;  // linha 6
  const NCOLS  = 49;

  // Limpar dados e formatação anteriores
  const lastRow = aba.getLastRow();
  if (lastRow >= PRIMA) {
    aba.getRange(PRIMA, 1, lastRow - HEADER, NCOLS).clearContent();
    aba.getRange(PRIMA, 1, lastRow - HEADER, NCOLS).setBackground(null);
    aba.getRange(PRIMA, 1, lastRow - HEADER, NCOLS).setFontColor(null);
  }
  if (itens.length === 0) return;

  const rows = itens.map(item => {
    const c030  = item.cons['0-30'];
    const c3060 = item.cons['30-60'];
    const c6090 = item.cons['60-90'];

    const m030  = c030.fat  > 0 ? c030.margem  / c030.fat  : 0;
    const m3060 = c3060.fat > 0 ? c3060.margem / c3060.fat : 0;
    const m6090 = c6090.fat > 0 ? c6090.margem / c6090.fat : 0;

    const pctML     = c030.fat > 0 ? c030.fatML     / c030.fat : 0;
    const pctShopee = c030.fat > 0 ? c030.fatShopee / c030.fat : 0;

    // Conta que mais vende em 0-30
    let maisVende = '';
    let maxFat = 0;
    for (const canal in c030.porCanal) {
      if (c030.porCanal[canal] > maxFat) { maxFat = c030.porCanal[canal]; maisVende = canal; }
    }

    // Var% faturamento 0-30 vs 30-60
    // Positivo = 0-30 maior (produto crescendo), negativo = produto em queda
    const varFat = c3060.fat !== 0
      ? (c030.fat - c3060.fat) / Math.abs(c3060.fat)
      : (c030.fat > 0 ? 1 : 0);

    // Colunas de estoque composto
    const pendMapeamento = pivotPend[item.sku] || 0;
    const full           = pivotFull[item.sku] || 0;
    const comprasPend    = pivotComp[item.sku] || 0;
    const resultado      = (item.estoque + pendMapeamento) - full;
    const resultadoFinal = (item.estoque + pendMapeamento + comprasPend) - full;

    // Diferenças usando Resultado Final como base de estoque disponível
    const dif07   = item.qtdSem[0]                                - resultadoFinal;  // AQ
    const dif030  = c030.qtd                                      - resultadoFinal;  // AR / AU
    const dif3060 = c3060.qtd                                     - resultadoFinal;  // AS
    const dif6090 = c6090.qtd                                     - resultadoFinal;  // AT
    const dif060  = c030.qtd + c3060.qtd                          - resultadoFinal;  // AV
    const dif090  = c030.qtd + c3060.qtd + c6090.qtd              - resultadoFinal;  // AW

    return [
      item.sku,              // 1:  A
      item.qtdSem[0],        // 2:  B
      item.qtdSem[1],        // 3:  C
      item.qtdSem[2],        // 4:  D
      item.qtdSem[3],        // 5:  E
      item.qtdSem[4],        // 6:  F
      item.qtdSem[5],        // 7:  G
      item.qtdSem[6],        // 8:  H
      c030.qtd,              // 9:  I
      c030.fat,              // 10: J
      c030.caixa,            // 11: K
      c030.margem,           // 12: L
      m030,                  // 13: M  (%)
      c3060.qtd,             // 14: N
      c3060.fat,             // 15: O
      c3060.caixa,           // 16: P
      c3060.margem,          // 17: Q
      m3060,                 // 18: R  (%)
      c6090.qtd,             // 19: S
      c6090.fat,             // 20: T
      c6090.caixa,           // 21: U
      c6090.margem,          // 22: V
      m6090,                 // 23: W  (%)
      pctML,                 // 24: X  (%)
      pctShopee,             // 25: Y  (%)
      c030.curvaFat  || '',  // 26: Z
      c030.pctFat    || 0,   // 27: AA (%)
      c3060.curvaFat || '',  // 28: AB
      c3060.pctFat   || 0,   // 29: AC (%)
      c6090.curvaFat || '',  // 30: AD
      c6090.pctFat   || 0,   // 31: AE (%)
      c030.curva2    || '',  // 32: AF
      c3060.curva2   || '',  // 33: AG
      c6090.curva2   || '',  // 34: AH
      item.estoque,          // 35: AI
      pendMapeamento,        // 36: AJ
      full,                  // 37: AK
      resultado,             // 38: AL
      comprasPend,           // 39: AM
      resultadoFinal,        // 40: AN
      maisVende,             // 41: AO
      varFat,                // 42: AP (%)
      dif07,                 // 43: AQ
      dif030,                // 44: AR
      dif3060,               // 45: AS
      dif6090,               // 46: AT
      dif030,                // 47: AU  Acum 0-30 = mesmo que AR
      dif060,                // 48: AV
      dif090,                // 49: AW
    ];
  });

  // ── Escrever dados ──
  aba.getRange(PRIMA, 1, rows.length, NCOLS).setValues(rows);

  // SKU como texto
  aba.getRange(PRIMA, 1, rows.length, 1).setNumberFormat('@');

  // Colunas percentuais: M(13), R(18), W(23), X(24), Y(25), AA(27), AC(29), AE(31), AP(42)
  [13, 18, 23, 24, 25, 27, 29, 31, 42].forEach(col => {
    aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('0.0%');
  });

  // ── Colorir Curva 2: AF(32), AG(33), AH(34) ──
  _colorirCurva2Coluna(aba, itens, PRIMA, 32, '0-30');
  _colorirCurva2Coluna(aba, itens, PRIMA, 33, '30-60');
  _colorirCurva2Coluna(aba, itens, PRIMA, 34, '60-90');

  // ── Colorir diferenças: AQ(43)..AW(49) ──
  _colorirDiferencas(aba, itens, PRIMA, [43, 44, 45, 46, 47, 48, 49], rows);

  // ── Linha vermelha claro: Margem % 0-30 < 11,5% ──
  const RED_ROW = '#ffcdd2';
  const redRanges = [];
  itens.forEach((item, i) => {
    const c = item.cons['0-30'];
    if (c.fat > 0 && c.margem / c.fat < RCFG.MARGEM_MIN) {
      redRanges.push(aba.getRange(PRIMA + i, 1, 1, NCOLS).getA1Notation());
    }
  });
  if (redRanges.length > 0) aba.getRangeList(redRanges).setBackground(RED_ROW);

  SpreadsheetApp.flush();
}

// ============================================================
// COLORIR COLUNAS DE DIFERENÇA
// diff = qtd_vendida − Resultado Final
//   diff > 0 → vendas superam o disponível → CRÍTICO (vermelho)
//   diff < 0 → disponível cobre as vendas  → OK (verde)
//   diff = 0 → exatamente igual            → amarelo
// ============================================================
function _colorirDiferencas(aba, itens, primaLinha, cols, rows) {
  const COR_CRITICO = '#ef9a9a';
  const COR_OK      = '#c8e6c9';
  const COR_ZERO    = '#fff9c4';

  cols.forEach(col => {
    const colIdx = col - 1;
    const critico = [], ok = [], zero = [];

    itens.forEach((item, i) => {
      const val = rows[i][colIdx];
      if (val === 0) {
        zero.push(aba.getRange(primaLinha + i, col).getA1Notation());
      } else if (val > 0) {
        critico.push(aba.getRange(primaLinha + i, col).getA1Notation());
      } else {
        ok.push(aba.getRange(primaLinha + i, col).getA1Notation());
      }
    });

    if (critico.length) aba.getRangeList(critico).setBackground(COR_CRITICO);
    if (ok.length)      aba.getRangeList(ok).setBackground(COR_OK);
    if (zero.length)    aba.getRangeList(zero).setBackground(COR_ZERO);
  });
}

// ============================================================
// COLORIR CURVA 2
// ============================================================
function _colorirCurva2Coluna(aba, itens, primaLinha, col, nomePer) {
  const BG = {
    'AA':'#1b5e20','AB':'#388e3c','AC':'#a5d6a7',
    'BA':'#f57f17','BB':'#fbc02d','BC':'#fff176',
    'CA':'#b71c1c','CB':'#e57373','CC':'#ffcdd2',
  };
  const FG = {
    'AA':'#ffffff','AB':'#ffffff','AC':'#000000',
    'BA':'#ffffff','BB':'#000000','BC':'#000000',
    'CA':'#ffffff','CB':'#000000','CC':'#000000',
  };
  const grupos = {};

  itens.forEach((item, i) => {
    const c2 = item.cons[nomePer] && item.cons[nomePer].curva2;
    if (!c2 || !BG[c2]) return;
    if (!grupos[c2]) grupos[c2] = [];
    grupos[c2].push(aba.getRange(primaLinha + i, col).getA1Notation());
  });

  Object.keys(grupos).forEach(c2 => {
    if (!grupos[c2].length) return;
    aba.getRangeList(grupos[c2]).setBackground(BG[c2]).setFontColor(FG[c2]);
  });
}
