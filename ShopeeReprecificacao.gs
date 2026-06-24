// ============================================================
// SHOPEE REPRECIFICAÇÃO — v1
// Consolida produtos, estoque BL, custos e vendas da Shopee.
//
// Abas locais esperadas (criadas por criarAbas()):
//   "Produtos Shopee"   — exportação "Editar em Massa"
//   "Promoções Shopee"  — exportação de descontos/promoções
//   "Vendas Shopee"     — exportação "Meus Pedidos" (acumulado 90 d)
//   "Novo Custo"        — SKU (A) | Novo Custo (B)
//   "Configurações"     — chave (A) | valor (B)
//
// Aba de saída: "Reprecificação"
// Layout: A–AN (40 colunas)
//   A-E   Dados do produto (nome, ID, variação, skus)
//   F     SKU Final
//   G     Preço Cheio   H Preço Promo   I Preço Final   J Est. Shopee
//   K     Est. BaseLinker   L Tipo
//   M     Comissão   N Repasse   O Taxa %   P Custo   Q Margem R$   R Margem %
//   S     Preço Reverso   T Margem R$ (reverso)   U Margem % (reverso)
//   V     Preço Novo [MANUAL]
//   W     Novo Custo
//   X     Comissão/V   Y Repasse/V   Z Margem R$/V   AA Repasse %/V   AB Margem %/V
//   AC-AH  Vendas: Receita 0-30 | Qtd 0-30 | Receita 30-60 | Qtd 30-60 | Receita 60-90 | Qtd 60-90
//   AI-AK  ABC 0-30 | 30-60 | 60-90
//   AL-AN  AABBCC 0-30 | 30-60 | 60-90
// ============================================================

const SRCFG = {
  ID_BL:      '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL:     'estoque',
  ID_CUSTOS:  '1C9Z4vT9SaamEGJ_hCWXdOm8vWOMhxw_vAb049mvtrP0',
  ABA_CUSTOS: 'Soma Composições',

  ABA_PRODUTOS:   'Produtos Shopee',
  ABA_PROMOCOES:  'Promoções Shopee',
  ABA_VENDAS:     'Vendas Shopee',
  ABA_NOVO_CUSTO: 'Novo Custo',
  ABA_CONFIG:     'Configurações',
  ABA_SAIDA:      'Reprecificação',

  BL_SKU: 1, BL_COD_EST: 2, BL_QTDE_EST: 3,
  BL_SALDO_PAD: 4, BL_SALDO_ARM: 5, BL_SALDO_CHG: 6, BL_TIPO: 7,

  CUST_SKU: 0, CUST_VALOR: 1,

  ABC_A: 0.70,
  ABC_B: 0.90,
  HEADER_ROWS: 2,
  NCOLS: 40,
};

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛍️ Shopee Reprecificação')
    .addItem('🔄 Atualizar Reprecificação', 'gerarReprecificacao')
    .addSeparator()
    .addItem('⚙️ Criar abas de entrada',    'criarAbas')
    .addToUi();
}

// ============================================================
// CRIAR ABAS DE ENTRADA — setup inicial
// ============================================================
function criarAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = [
    {
      nome: SRCFG.ABA_PRODUTOS,
      header: ['Nome do Produto','ID do Item','Variação','SKU da Variação','SKU do Usuário','Preço Normal','Estoque'],
      nota: 'Cole aqui a exportação "Editar em Massa" da Shopee. Cabeçalho na linha 1.',
    },
    {
      nome: SRCFG.ABA_PROMOCOES,
      header: ['SKU do Usuário','SKU da Variação','ID do Item','Preço Promocional'],
      nota: 'Cole aqui a exportação de promoções/descontos. Cabeçalho na linha 1.',
    },
    {
      nome: SRCFG.ABA_VENDAS,
      header: ['Status do pedido','Data de criação do pedido','Número de referência SKU','Quantidade','Subtotal do produto'],
      nota: 'Cole o export original de Meus Pedidos da Shopee direto nesta aba (pode acumular até 90 dias).',
    },
    {
      nome: SRCFG.ABA_NOVO_CUSTO,
      header: ['SKU','Novo Custo'],
      nota: 'Digite o SKU e o novo custo unitário para simulação.',
    },
    {
      nome: SRCFG.ABA_CONFIG,
      header: ['Configuração','Valor'],
      nota: 'Parâmetros gerais.',
    },
  ];

  abas.forEach(({ nome, header, nota }) => {
    let aba = ss.getSheetByName(nome);
    if (!aba) aba = ss.insertSheet(nome);
    aba.getRange(1, 1, 1, header.length)
       .setValues([header])
       .setFontWeight('bold')
       .setBackground('#263238')
       .setFontColor('#ffffff');
  });

  const cfg = ss.getSheetByName(SRCFG.ABA_CONFIG);
  if (cfg && cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 2, 2).setValues([
      ['Taxa (%)',       0.06],
      ['Descrição taxa', 'Simples Nacional — ajuste conforme o regime'],
    ]);
    cfg.getRange(2, 2, 1, 1).setNumberFormat('0.00%');
  }

  SpreadsheetApp.getUi().alert(
    '✅ Abas criadas!\n\n' +
    '• "Produtos Shopee"  — cole a exportação "Editar em Massa"\n' +
    '• "Promoções Shopee" — cole a exportação de descontos\n' +
    '• "Vendas Shopee"    — cole Meus Pedidos (até 90 dias acumulado)\n' +
    '• "Novo Custo"       — SKU + novo custo para simulação\n' +
    '• "Configurações"    — taxa % já pré-preenchida com 6%\n\n' +
    'Depois clique em 🔄 Atualizar Reprecificação.'
  );
}

// ============================================================
// PRINCIPAL
// ============================================================
function gerarReprecificacao() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Carregando dados... Aguarde.');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const taxa          = _sr_carregarTaxa(ss);
    const mapaNovoCusto = _sr_carregarNovoCusto(ss);
    const mapaCustos    = _sr_carregarCustos();
    const { mapaCompostos, proporcoes, estBL } = _sr_carregarBL(mapaCustos);
    const mapaPromo     = _sr_carregarPromocoes(ss);
    const mapaVendas    = _sr_carregarVendas(ss);
    const produtos      = _sr_carregarProdutos(ss);

    if (!produtos.length) {
      ui.alert('❌ Nenhum produto encontrado em "' + SRCFG.ABA_PRODUTOS + '".\nCole a exportação e tente novamente.');
      return;
    }

    const itens = _sr_montarItens(
      produtos, estBL, mapaCustos, mapaCompostos, proporcoes,
      mapaPromo, mapaVendas, mapaNovoCusto, taxa
    );

    _sr_calcularCurvas(itens);
    _sr_escreverAba(ss, itens, taxa);

    ui.alert('✅ Reprecificação atualizada! ' + itens.length + ' variações processadas.');
  } catch (e) {
    ui.alert('❌ Erro: ' + e.message + '\n\n' + e.stack);
  }
}

// ============================================================
// CARREGAR TAXA (Configurações!B2 que contenha "taxa")
// ============================================================
function _sr_carregarTaxa(ss) {
  const aba = ss.getSheetByName(SRCFG.ABA_CONFIG);
  if (!aba) return 0.06;
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const chave = String(dados[i][0] || '').toLowerCase();
    if (chave.includes('taxa') && !chave.includes('descri')) {
      const v = dados[i][1];
      if (typeof v === 'number') return v < 1 ? v : v / 100;
      const n = parseFloat(String(v).replace(',', '.'));
      return isNaN(n) ? 0.06 : (n < 1 ? n : n / 100);
    }
  }
  return 0.06;
}

// ============================================================
// CARREGAR NOVO CUSTO
// ============================================================
function _sr_carregarNovoCusto(ss) {
  const aba  = ss.getSheetByName(SRCFG.ABA_NOVO_CUSTO);
  const mapa = {};
  if (!aba || aba.getLastRow() < 2) return mapa;
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const sku   = String(dados[i][0] || '').trim();
    const custo = parseFloat(String(dados[i][1] || '0').replace(',', '.')) || 0;
    if (sku) mapa[sku] = custo;
  }
  return mapa;
}

// ============================================================
// CARREGAR CUSTOS (Soma Composições)
// ============================================================
function _sr_carregarCustos() {
  const ss  = SpreadsheetApp.openById(SRCFG.ID_CUSTOS);
  const aba = _sr_getAba(ss, SRCFG.ABA_CUSTOS);
  if (!aba) throw new Error('Aba Custos "' + SRCFG.ABA_CUSTOS + '" não encontrada.');
  const dados = aba.getDataRange().getValues();
  const mapa  = {};
  for (let i = 1; i < dados.length; i++) {
    const sku   = String(dados[i][SRCFG.CUST_SKU]   || '').trim();
    const custo = parseFloat(String(dados[i][SRCFG.CUST_VALOR] || '0').replace(',', '.')) || 0;
    if (sku && !(sku in mapa)) mapa[sku] = custo;
  }
  return mapa;
}

// ============================================================
// CARREGAR BASELINKER — estoque + estrutura de kits
// ============================================================
function _sr_carregarBL(mapaCustos) {
  const ss  = SpreadsheetApp.openById(SRCFG.ID_BL);
  const aba = _sr_getAba(ss, SRCFG.ABA_BL);
  if (!aba) throw new Error('Aba BL "' + SRCFG.ABA_BL + '" não encontrada.');

  const dados         = aba.getDataRange().getValues();
  const estBL         = {};
  const mapaCompostos = {};

  for (let i = 1; i < dados.length; i++) {
    const row    = dados[i];
    const sku    = String(row[SRCFG.BL_SKU]      || '').trim();
    const tipo   = String(row[SRCFG.BL_TIPO]     || '').trim();
    const codEst = String(row[SRCFG.BL_COD_EST]  || '').trim();
    const qtde   = Number(row[SRCFG.BL_QTDE_EST] || 0);
    const saldo  = Number(row[SRCFG.BL_SALDO_PAD] || 0)
                 + Number(row[SRCFG.BL_SALDO_ARM] || 0)
                 + Number(row[SRCFG.BL_SALDO_CHG] || 0);
    if (!sku) continue;

    if (tipo === 'Composto') {
      if (!codEst) continue;
      if (!mapaCompostos[sku]) mapaCompostos[sku] = [];
      mapaCompostos[sku].push({ codEst, qtde });
    } else {
      estBL[sku] = (estBL[sku] || 0) + saldo;
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

  return { mapaCompostos, proporcoes, estBL };
}

// ============================================================
// CARREGAR PRODUTOS SHOPEE (detecção de colunas por cabeçalho)
// ============================================================
function _sr_carregarProdutos(ss) {
  const aba = ss.getSheetByName(SRCFG.ABA_PRODUTOS);
  if (!aba || aba.getLastRow() < 2) return [];

  const dados = aba.getDataRange().getValues();
  const cols  = _sr_detectarColunas(dados[0], {
    nome:       ['nome do produto', 'nome do item', 'nome'],
    idItem:     ['id do item', 'item id', 'id'],
    variacao:   ['variacao', 'modelo', 'variation', 'nome do modelo'],
    skuVar:     ['sku da variacao', 'sku variacao', 'codigo da variacao', 'sku de variacao'],
    skuUsr:     ['sku do usuario', 'sku usuario', 'sku'],
    preco:      ['preco normal', 'preco original', 'preco', 'price'],
    estoque:    ['estoque', 'quantidade em estoque', 'quantidade', 'qty', 'stock'],
  });

  const produtos = [];
  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    const get = (k) => cols[k] >= 0 ? String(row[cols[k]] || '').trim() : '';
    const getN = (k) => {
      if (cols[k] < 0) return 0;
      const v = row[cols[k]];
      return typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.')) || 0;
    };

    const nome   = get('nome');
    const skuVar = get('skuVar');
    const skuUsr = get('skuUsr');
    if (!nome && !skuVar && !skuUsr) continue; // linha vazia

    produtos.push({
      nome:      nome,
      idItem:    get('idItem'),
      variacao:  get('variacao'),
      skuVar:    skuVar,
      skuUsr:    skuUsr,
      precoNorm: getN('preco'),
      estShopee: getN('estoque'),
    });
  }
  return produtos;
}

// ============================================================
// CARREGAR PROMOÇÕES SHOPEE
// ============================================================
function _sr_carregarPromocoes(ss) {
  const aba  = ss.getSheetByName(SRCFG.ABA_PROMOCOES);
  const mapa = {};
  if (!aba || aba.getLastRow() < 2) return mapa;

  const dados = aba.getDataRange().getValues();
  const cols  = _sr_detectarColunas(dados[0], {
    skuUsr:    ['sku do usuario', 'sku usuario', 'sku'],
    skuVar:    ['sku da variacao', 'sku variacao'],
    precoPromo:['preco promocional', 'preco especial', 'preco de desconto', 'preco promo'],
  });

  for (let i = 1; i < dados.length; i++) {
    const row      = dados[i];
    const getN = (k) => {
      if (cols[k] < 0) return 0;
      const v = row[cols[k]];
      return typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.')) || 0;
    };
    const skuUsr = cols.skuUsr >= 0 ? String(row[cols.skuUsr] || '').trim() : '';
    const skuVar = cols.skuVar >= 0 ? String(row[cols.skuVar] || '').trim() : '';
    const preco  = getN('precoPromo');
    if (!preco) continue;
    if (skuUsr) mapa[skuUsr] = preco;
    if (skuVar) mapa[skuVar] = preco;
  }
  return mapa;
}

// ============================================================
// CARREGAR VENDAS SHOPEE — acumulado, bin 0-30/30-60/60-90 dias
// Referência: dataMax = data mais recente nos dados (não "hoje")
// Status aceitos: Concluído, Entregue, Enviado, Order Received, A Enviar
// ============================================================
const SR_STATUS_OK = ['concluido', 'entregue', 'enviado', 'order received', 'a enviar'];

function _sr_carregarVendas(ss) {
  const aba  = ss.getSheetByName(SRCFG.ABA_VENDAS);
  const mapa = {}; // sku → { fat030, qtd030, fat3060, qtd3060, fat6090, qtd6090 }
  if (!aba || aba.getLastRow() < 2) return mapa;

  const dados = aba.getDataRange().getValues();
  const cols  = _sr_detectarColunas(dados[0], {
    status:  ['status do pedido', 'status'],
    data:    ['data de criacao do pedido', 'data de criacao', 'data do pedido', 'data criacao', 'data'],
    sku:     ['numero de referencia sku', 'referencia sku', 'sku'],
    qtd:     ['quantidade', 'qtd', 'qty'],
    receita: ['subtotal do produto', 'subtotal', 'receita', 'valor'],
  });

  // Primeiro passo: encontrar dataMax em todos os dados válidos
  let dataMax = new Date(0);
  for (let i = 1; i < dados.length; i++) {
    const d = cols.data >= 0 ? _sr_parseData(dados[i][cols.data]) : null;
    if (d && d > dataMax) dataMax = d;
  }
  dataMax = new Date(dataMax.getFullYear(), dataMax.getMonth(), dataMax.getDate());

  // Segundo passo: acumular vendas relativas a dataMax
  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];

    // Filtro de status
    if (cols.status >= 0) {
      const st = String(row[cols.status] || '').normalize('NFD')
        .replace(/[̀-ͯ]/g, '').toLowerCase().trim();
      if (!SR_STATUS_OK.some(s => st.includes(s))) continue;
    }

    const d = cols.data >= 0 ? _sr_parseData(row[cols.data]) : null;
    if (!d) continue;
    const dNorm = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const dias  = Math.floor((dataMax - dNorm) / 86400000);
    if (dias < 0 || dias >= 90) continue;

    const sku = cols.sku >= 0 ? String(row[cols.sku] || '').trim() : '';
    if (!sku) continue;

    const getN = (k) => {
      if (cols[k] < 0) return 0;
      const v = row[cols[k]];
      return typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.')) || 0;
    };
    const qtd = getN('qtd');
    const rec = getN('receita');

    if (!mapa[sku]) mapa[sku] = { fat030: 0, qtd030: 0, fat3060: 0, qtd3060: 0, fat6090: 0, qtd6090: 0 };
    const v = mapa[sku];
    if      (dias < 30) { v.fat030  += rec; v.qtd030  += qtd; }
    else if (dias < 60) { v.fat3060 += rec; v.qtd3060 += qtd; }
    else                { v.fat6090 += rec; v.qtd6090 += qtd; }
  }
  return mapa;
}

// ============================================================
// MONTAR ITENS — enriquecer produtos com todas as métricas
// ============================================================
function _sr_montarItens(produtos, estBL, mapaCustos, mapaCompostos, proporcoes,
                          mapaPromo, mapaVendas, mapaNovoCusto, taxa) {
  return produtos.map(p => {
    const skuFinal   = p.skuUsr || p.skuVar;
    const precoPromo = mapaPromo[skuFinal] || mapaPromo[p.skuVar] || 0;
    const precoFinal = precoPromo > 0 ? precoPromo : p.precoNorm;

    // Tipo e estoque BL
    // Custo: mapaNovoCusto é fonte primária; mapaCustos (Soma Composições) é fallback
    let tipo, estoqueBL, custo;
    if (mapaCompostos[skuFinal]) {
      tipo = 'Composto';
      const comps = mapaCompostos[skuFinal];
      estoqueBL = comps.reduce((min, c) => {
        const disp = c.qtde > 0 ? Math.floor((estBL[c.codEst] || 0) / c.qtde) : 0;
        return Math.min(min, disp);
      }, Infinity);
      if (!isFinite(estoqueBL)) estoqueBL = 0;
      custo = comps.reduce((s, c) => {
        const custoComp = mapaNovoCusto[c.codEst] || mapaCustos[c.codEst] || 0;
        return s + custoComp * c.qtde;
      }, 0);
    } else {
      tipo      = 'Simples';
      estoqueBL = estBL[skuFinal] || 0;
      custo     = mapaNovoCusto[skuFinal] || mapaCustos[skuFinal] || 0;
    }

    // Comissão e margens com preço atual
    const comissao  = _sr_calcComissao(precoFinal);
    const repasse   = precoFinal - comissao;
    const margemR   = precoFinal - comissao - precoFinal * taxa - custo;
    const margemPct = precoFinal > 0 ? margemR / precoFinal : 0;

    // Preço reverso: preço P tal que margem % = margemPct com novo custo
    const novoCusto      = mapaNovoCusto[skuFinal] || 0;
    const precoReverso   = novoCusto > 0
      ? _sr_calcPrecoReverso(margemPct, novoCusto, taxa)
      : 0;
    const comReverso     = precoReverso > 0 ? _sr_calcComissao(precoReverso) : 0;
    const margemRReverso = precoReverso > 0
      ? precoReverso - comReverso - precoReverso * taxa - novoCusto
      : 0;
    const margemPctReverso = precoReverso > 0 && precoReverso > 0
      ? margemRReverso / precoReverso
      : 0;

    // Vendas
    const v = mapaVendas[skuFinal] || mapaVendas[p.skuVar] || {};
    const fat030  = v.fat030  || 0;  const qtd030  = v.qtd030  || 0;
    const fat3060 = v.fat3060 || 0;  const qtd3060 = v.qtd3060 || 0;
    const fat6090 = v.fat6090 || 0;  const qtd6090 = v.qtd6090 || 0;

    return {
      nome: p.nome, idItem: p.idItem, variacao: p.variacao,
      skuVar: p.skuVar, skuUsr: p.skuUsr, skuFinal,
      precoNorm: p.precoNorm, precoPromo, precoFinal,
      estShopee: p.estShopee, estoqueBL, tipo,
      comissao, repasse, taxa, custo, margemR, margemPct,
      precoReverso, margemRReverso, margemPctReverso,
      novoCusto,
      fat030, qtd030, fat3060, qtd3060, fat6090, qtd6090,
    };
  });
}

// ============================================================
// COMISSÃO POR FAIXA DE PREÇO
// ============================================================
function _sr_calcComissao(preco) {
  if (!preco || preco <= 0) return 0;
  if (preco <= 79.99)  return 0.20 * preco + 4;
  if (preco <= 99.99)  return 0.14 * preco + 16;
  if (preco <= 199.99) return 0.14 * preco + 20;
  return 0.14 * preco + 26; // ≥ R$200
}

// ============================================================
// PREÇO REVERSO — encontra P tal que margem% = alvo com novo custo
// Resolve por faixa: P = (novoCusto + b) / (1 - taxa - margemAlvo - a)
// ============================================================
function _sr_calcPrecoReverso(margemAlvo, novoCusto, taxa) {
  if (novoCusto <= 0) return 0;
  const tiers = [
    { min: 0,    max: 79.99,   a: 0.20, b: 4  },
    { min: 80,   max: 99.99,   a: 0.14, b: 16 },
    { min: 100,  max: 199.99,  a: 0.14, b: 20 },
    { min: 200,  max: Infinity, a: 0.14, b: 26 },
  ];
  for (const tier of tiers) {
    const denom = 1 - taxa - margemAlvo - tier.a;
    if (denom <= 0) continue;
    const P = (novoCusto + tier.b) / denom;
    if (P >= tier.min - 0.005 && P <= tier.max + 0.005) return P;
  }
  // Fallback: faixa mais alta
  const t     = tiers[tiers.length - 1];
  const denom = 1 - taxa - margemAlvo - t.a;
  return denom > 0 ? (novoCusto + t.b) / denom : 0;
}

// ============================================================
// CURVAS ABC + AABBCC por período (baseadas em receita)
// ============================================================
function _sr_calcularCurvas(itens) {
  const periodos = [
    { fatKey: 'fat030',  abcKey: 'abc030',  aabbccKey: 'aabbcc030'  },
    { fatKey: 'fat3060', abcKey: 'abc3060', aabbccKey: 'aabbcc3060' },
    { fatKey: 'fat6090', abcKey: 'abc6090', aabbccKey: 'aabbcc6090' },
  ];

  periodos.forEach(({ fatKey, abcKey, aabbccKey }) => {
    const sorted = [...itens].sort((a, b) => b[fatKey] - a[fatKey]);
    const total  = sorted.reduce((s, x) => s + x[fatKey], 0);

    let cum = 0;
    sorted.forEach(item => {
      cum += item[fatKey];
      const pct   = total > 0 ? cum / total : 1;
      item[abcKey] = pct <= SRCFG.ABC_A ? 'A' : pct <= SRCFG.ABC_B ? 'B' : 'C';
    });

    ['A', 'B', 'C'].forEach(grp => {
      const grpItems = sorted.filter(x => x[abcKey] === grp);
      const grpTotal = grpItems.reduce((s, x) => s + x[fatKey], 0);
      let grpCum = 0;
      grpItems.forEach(item => {
        grpCum += item[fatKey];
        const sub = grpTotal > 0 ? grpCum / grpTotal : 1;
        item[aabbccKey] = grp + (sub <= SRCFG.ABC_A ? 'A' : sub <= SRCFG.ABC_B ? 'B' : 'C');
      });
    });
  });
}

// ============================================================
// ESCREVER ABA DE SAÍDA
// ============================================================
function _sr_escreverAba(ss, itens, taxa) {
  let aba = ss.getSheetByName(SRCFG.ABA_SAIDA);
  if (!aba) aba = ss.insertSheet(SRCFG.ABA_SAIDA);

  const PRIMA = SRCFG.HEADER_ROWS + 1;
  const NCOLS = SRCFG.NCOLS;

  // Cabeçalho: limpa e reescreve totalmente
  aba.getRange(1, 1, SRCFG.HEADER_ROWS, NCOLS).clearContent().clearFormat();

  // Dados: apenas conteúdo (preserva formatação manual do usuário)
  const lastRow = aba.getLastRow();
  if (lastRow >= PRIMA) {
    aba.getRange(PRIMA, 1, lastRow - PRIMA + 1, NCOLS).clearContent();
    // Resetar cores nas colunas de ABC/AABBCC (AI=35..AN=40)
    for (let c = 35; c <= 40; c++) {
      aba.getRange(PRIMA, c, lastRow - PRIMA + 1, 1)
         .setBackground(null).setFontColor(null);
    }
    // Resetar margem % colunas (R=18, AB=28) para detectar verde/amarelo/vermelho
    [18, 28].forEach(c => {
      aba.getRange(PRIMA, c, lastRow - PRIMA + 1, 1).setBackground(null).setFontColor(null);
    });
  }

  _sr_escreverCabecalho(aba);

  if (itens.length === 0) { SpreadsheetApp.flush(); return; }

  // ── Formata SKU como texto antes de escrever ──
  aba.getRange(PRIMA, 6, itens.length, 1).setNumberFormat('@');

  // ── Colunas A-U (1-21): valores calculados ──
  const rowsAU = itens.map(it => [
    it.nome,                                    // A 1
    it.idItem,                                  // B 2
    it.variacao,                                // C 3
    String(it.skuVar),                          // D 4
    String(it.skuUsr),                          // E 5
    String(it.skuFinal),                        // F 6
    it.precoNorm,                               // G 7
    it.precoPromo > 0 ? it.precoPromo : '',     // H 8
    it.precoFinal,                              // I 9
    it.estShopee,                               // J 10
    it.estoqueBL,                               // K 11
    it.tipo,                                    // L 12
    it.comissao,                                // M 13
    it.repasse,                                 // N 14
    it.taxa,                                    // O 15
    it.custo,                                   // P 16
    it.margemR,                                 // Q 17
    it.margemPct,                               // R 18
    it.precoReverso > 0 ? it.precoReverso : '', // S 19
    it.precoReverso > 0 ? it.margemRReverso:'', // T 20
    it.precoReverso > 0 ? it.margemPctReverso:'',// U 21
  ]);
  aba.getRange(PRIMA, 1, itens.length, 21).setValues(rowsAU);

  // ── Coluna W (23): novo custo ──
  const rowsW = itens.map(it => [it.novoCusto > 0 ? it.novoCusto : '']);
  aba.getRange(PRIMA, 23, itens.length, 1).setValues(rowsW);

  // ── Colunas X-AB (24-28): fórmulas que dependem de V (22) e W (23) ──
  const fmls = itens.map((_, idx) => {
    const r  = PRIMA + idx;
    const cv = `V${r}`;
    const cw = `W${r}`;
    const co = `O${r}`;
    const com = `IF(${cv}<=79.99,0.2*${cv}+4,IF(${cv}<=99.99,0.14*${cv}+16,IF(${cv}<=199.99,0.14*${cv}+20,0.14*${cv}+26)))`;
    return [
      `=IF(OR(${cv}="",${cv}=0),"",${com})`,                                         // X 24 comissão
      `=IF(OR(${cv}="",${cv}=0),"",${cv}-X${r})`,                                    // Y 25 repasse
      `=IF(OR(${cv}="",${cv}=0),"",${cv}-X${r}-${cv}*${co}-IF(${cw}="",P${r},${cw}))`, // Z 26 margem R$
      `=IF(OR(${cv}="",${cv}=0),"",Y${r}/${cv})`,                                    // AA 27 repasse %
      `=IF(OR(${cv}="",${cv}=0),"",Z${r}/${cv})`,                                    // AB 28 margem %
    ];
  });
  aba.getRange(PRIMA, 24, itens.length, 5).setFormulas(fmls);

  // ── Colunas AC-AH (29-34): vendas ──
  const rowsVendas = itens.map(it => [
    it.fat030,  it.qtd030,    // AC AD
    it.fat3060, it.qtd3060,   // AE AF
    it.fat6090, it.qtd6090,   // AG AH
  ]);
  aba.getRange(PRIMA, 29, itens.length, 6).setValues(rowsVendas);

  // ── Colunas AI-AN (35-40): ABC e AABBCC ──
  const rowsCurvas = itens.map(it => [
    it.abc030   || '', it.abc3060   || '', it.abc6090   || '',  // AI AJ AK
    it.aabbcc030 || '', it.aabbcc3060 || '', it.aabbcc6090 || '', // AL AM AN
  ]);
  aba.getRange(PRIMA, 35, itens.length, 6).setValues(rowsCurvas);

  // ── Formatação numérica ──
  // Moeda: G I M N P Q S T (e as de preço)
  [7, 9, 13, 14, 16, 17, 19, 20].forEach(c => {
    aba.getRange(PRIMA, c, itens.length, 1).setNumberFormat('R$ #,##0.00');
  });
  // Percentual: O R U (taxa, margem, margem reverso)
  [15, 18, 21].forEach(c => {
    aba.getRange(PRIMA, c, itens.length, 1).setNumberFormat('0.00%');
  });
  // Receita vendas: AC AE AG
  [29, 31, 33].forEach(c => {
    aba.getRange(PRIMA, c, itens.length, 1).setNumberFormat('R$ #,##0.00');
  });

  // ── Colorir margem % atual (R=18) ──
  _sr_colorirMargem(aba, itens, PRIMA, 18, it => it.margemPct);

  // ── Colorir ABC (AI-AK = cols 35-37) ──
  _sr_colorirABC(aba, itens, PRIMA,
    [35, 36, 37], ['abc030', 'abc3060', 'abc6090']);

  // ── Colorir AABBCC (AL-AN = cols 38-40) ──
  _sr_colorirAABBCC(aba, itens, PRIMA,
    [38, 39, 40], ['aabbcc030', 'aabbcc3060', 'aabbcc6090']);

  // Data atualização (coluna AP = 42)
  aba.getRange(1, 42).setValue('Atualizado em');
  aba.getRange(2, 42).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  );

  SpreadsheetApp.flush();
}

// ============================================================
// CABEÇALHO DA ABA DE SAÍDA
// ============================================================
function _sr_escreverCabecalho(aba) {
  const NCOLS = SRCFG.NCOLS;
  const h1    = Array(NCOLS).fill('');
  const h2    = Array(NCOLS).fill('');

  // Linha 1 — grupos
  h1[0]  = 'Dados do Produto';
  h1[5]  = 'Preços';
  h1[9]  = 'Estoque';
  h1[11] = 'Análise de Margem (Preço Atual)';
  h1[18] = 'Preço Reverso';
  h1[21] = 'Preço Novo';
  h1[28] = 'Vendas Shopee';
  h1[34] = 'Curva ABC';
  h1[37] = 'Curva AABBCC';

  // Linha 2 — nomes de colunas (base 0)
  const names = [
    'Nome do Produto',    // A 0
    'ID Shopee',          // B 1
    'Variação',           // C 2
    'SKU Variação',       // D 3
    'SKU Usuário',        // E 4
    'SKU Final',          // F 5
    'Preço Cheio',        // G 6
    'Preço Promo',        // H 7
    'Preço Final',        // I 8
    'Est. Shopee',        // J 9
    'Est. BaseLinker',    // K 10
    'Tipo',               // L 11
    'Comissão R$',        // M 12
    'Repasse R$',         // N 13
    'Taxa %',             // O 14
    'Custo R$',           // P 15
    'Margem R$',          // Q 16
    'Margem %',           // R 17
    'Preço Reverso',      // S 18
    'Margem R$ (rev.)',   // T 19
    'Margem % (rev.)',    // U 20
    'Preço Novo ✏️',      // V 21
    'Novo Custo',         // W 22
    'Comissão/V',         // X 23
    'Repasse/V',          // Y 24
    'Margem R$/V',        // Z 25
    'Repasse %/V',        // AA 26
    'Margem %/V',         // AB 27
    'Receita 0-30d',      // AC 28
    'Qtd 0-30d',          // AD 29
    'Receita 30-60d',     // AE 30
    'Qtd 30-60d',         // AF 31
    'Receita 60-90d',     // AG 32
    'Qtd 60-90d',         // AH 33
    'ABC 0-30d',          // AI 34
    'ABC 30-60d',         // AJ 35
    'ABC 60-90d',         // AK 36
    'AABBCC 0-30d',       // AL 37
    'AABBCC 30-60d',      // AM 38
    'AABBCC 60-90d',      // AN 39
  ];
  names.forEach((n, i) => { h2[i] = n; });

  aba.getRange(1, 1, 2, NCOLS).setValues([h1, h2]);

  // Cores por grupo (col de 1-indexed)
  const grupos = [
    { de:1,  ate:5,  bg:'#455a64' }, // Dados do produto
    { de:6,  ate:10, bg:'#1565c0' }, // Preços + estoque
    { de:11, ate:11, bg:'#546e7a' }, // Tipo
    { de:12, ate:18, bg:'#2e7d32' }, // Análise margem atual
    { de:19, ate:21, bg:'#6a1b9a' }, // Preço reverso
    { de:22, ate:22, bg:'#e65100' }, // Preço Novo (manual)
    { de:23, ate:28, bg:'#ad1457' }, // Colunas V-dependentes
    { de:29, ate:34, bg:'#00695c' }, // Vendas
    { de:35, ate:37, bg:'#1976d2' }, // ABC
    { de:38, ate:40, bg:'#00838f' }, // AABBCC
  ];

  grupos.forEach(g => {
    const w = g.ate - g.de + 1;
    aba.getRange(1, g.de, 2, w)
       .setBackground(g.bg)
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  });

  // Merges linha 1
  aba.getRange(1, 1,  1, 5).merge();   // Dados do Produto
  aba.getRange(1, 6,  1, 5).merge();   // Preços + Estoque
  aba.getRange(1, 12, 1, 7).merge();   // Análise de Margem
  aba.getRange(1, 19, 1, 3).merge();   // Preço Reverso
  aba.getRange(1, 23, 1, 6).merge();   // Preço Novo bloco
  aba.getRange(1, 29, 1, 6).merge();   // Vendas
  aba.getRange(1, 35, 1, 3).merge();   // ABC
  aba.getRange(1, 38, 1, 3).merge();   // AABBCC

  // Linha Preço Novo (V col 22) — destaque especial
  aba.getRange(2, 22, 1, 1).setBackground('#ff6f00').setFontColor('#ffffff').setFontWeight('bold');

  // Dimensões
  aba.setRowHeight(1, 28);
  aba.setRowHeight(2, 40);
  aba.setColumnWidth(1, 200);  // Nome
  aba.setColumnWidth(6, 110);  // SKU Final
  aba.setColumnWidth(22, 100); // Preço Novo
  aba.setFrozenRows(2);
  aba.setFrozenColumns(6); // Freezar até SKU Final
}

// ============================================================
// COLORIR MARGEM % (verde ≥15%, amarelo 10-15%, vermelho <10%)
// ============================================================
function _sr_colorirMargem(aba, itens, prima, col, getFn) {
  const letter  = _sr_colLetter(col);
  const grupos  = { verde: [], amarelo: [], vermelho: [] };
  itens.forEach((it, i) => {
    const val = getFn(it);
    if (val === '' || val === null) return;
    const a1 = letter + (prima + i);
    if      (val >= 0.15) grupos.verde.push(a1);
    else if (val >= 0.10) grupos.amarelo.push(a1);
    else                  grupos.vermelho.push(a1);
  });
  if (grupos.verde.length)    aba.getRangeList(grupos.verde).setBackground('#c8e6c9').setFontColor('#1b5e20');
  if (grupos.amarelo.length)  aba.getRangeList(grupos.amarelo).setBackground('#fff9c4').setFontColor('#f57f17');
  if (grupos.vermelho.length) aba.getRangeList(grupos.vermelho).setBackground('#ffcdd2').setFontColor('#b71c1c');
}

// ============================================================
// COLORIR ABC
// ============================================================
function _sr_colorirABC(aba, itens, prima, cols, keys) {
  const BG = { 'A': '#c8e6c9', 'B': '#fff9c4', 'C': '#ffcdd2' };
  const FG = { 'A': '#1b5e20', 'B': '#f57f17', 'C': '#b71c1c' };
  cols.forEach((col, idx) => {
    const key    = keys[idx];
    const letter = _sr_colLetter(col);
    const grupos = { 'A': [], 'B': [], 'C': [] };
    itens.forEach((it, i) => {
      const l = it[key];
      if (l && grupos[l]) grupos[l].push(letter + (prima + i));
    });
    Object.keys(grupos).forEach(l => {
      if (grupos[l].length) aba.getRangeList(grupos[l]).setBackground(BG[l]).setFontColor(FG[l]);
    });
  });
}

// ============================================================
// COLORIR AABBCC
// ============================================================
function _sr_colorirAABBCC(aba, itens, prima, cols, keys) {
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
    const letter = _sr_colLetter(col);
    const grupos = {};
    itens.forEach((it, i) => {
      const l = it[key];
      if (!l || !BG[l]) return;
      if (!grupos[l]) grupos[l] = [];
      grupos[l].push(letter + (prima + i));
    });
    Object.keys(grupos).forEach(l => {
      aba.getRangeList(grupos[l]).setBackground(BG[l]).setFontColor(FG[l]);
    });
  });
}

// ============================================================
// HELPERS
// ============================================================
function _sr_parseData(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  // ISO: 2025-06-14 ou 2025-06-14 10:30
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(parseInt(iso[1]), parseInt(iso[2]) - 1, parseInt(iso[3]));
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function _sr_colLetter(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function _sr_getAba(ss, nome) {
  let aba = ss.getSheetByName(nome);
  if (aba) return aba;
  const norm = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const t = norm(nome);
  for (const a of ss.getSheets()) {
    if (norm(a.getName()) === t) return a;
  }
  return null;
}

function _sr_detectarColunas(headerRow, buscas) {
  const norm = s => String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const headers = headerRow.map(norm);
  const result  = {};
  for (const campo in buscas) {
    const aliases = buscas[campo].map(norm);
    let idx = -1;
    for (const alias of aliases) {
      idx = headers.findIndex(h => h.includes(alias));
      if (idx >= 0) break;
    }
    result[campo] = idx;
  }
  return result;
}
