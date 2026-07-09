// ============================================================
// MOVIMENTAÇÃO DE ESTOQUE — Corretiva / Planejada / Excesso
// Fontes: GE Finance (vendas 0-5/7/10/15d) + BaseLinker API
//
// Requer:
//   BaseLinkerAPI.gs → _bl_call()
//   BASELINKER_API_KEY em Script Properties
// ============================================================

const ME = {
  // Planilhas externas
  ID_GE:  '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE: 'Dados Completos',
  ID_BL:  '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL: 'estoque',

  // Colunas GE (0-indexed)
  GE_SKU:   1,   // B
  GE_QTD:   2,   // C
  GE_DATA:  3,   // D
  GE_CANAL: 7,   // H

  // Colunas BL estoque (0-indexed)
  BL_SKU:      1,
  BL_COD_EST:  2,
  BL_QTDE_EST: 3,
  BL_TIPO:     7,

  // BaseLinker
  INVENTORY_ID:    39947,
  WH_PADRAO:       'bl_44285',
  WH_ARMAZENAMENTO:'bl_50394',
  WH_CHEGOU:       'bl_51442',

  // Períodos de vendas (dias)
  PERIODOS: [5, 7, 10, 15],

  // Status de pedidos em andamento com pagamento confirmado
  STATUS_IDS: [
    252551,  // Novos pedidos
    268797,  // NF Emitida
    268798,  // Erro NF
    268799,  // Erro Etiqueta
    268813,  // [SEP] ML Flex
    268814,  // [SEP] ML Agência
    268806,  // [SEP] Shopee Direta
    268805,  // [SEP] Shopee Xpress
    268800,  // [EMB] ML Flex
    268801,  // [EMB] ML Agência
    268803,  // [EMB] Shopee Direta
    268804,  // [EMB] Shopee Xpress
    252552,  // [EXP] ML Flex
    268792,  // [EXP] ML Agência
    268794,  // [EXP] Shopee Direta
    268795,  // [EXP] Shopee Xpress
  ],

  // Métodos de envio agrupados (correspondência em delivery_method)
  METODOS: ['ML Flex', 'ML Agência', 'Shopee Direta', 'Shopee Xpress'],

  // Curvas ABC
  ABC_A: 0.70,
  ABC_B: 0.90,
};

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Movimentação Estoque')
    .addItem('🔴 Corretiva',         'meGerarCorretiva')
    .addItem('🟡 Planejada',         'meGerarPlanejada')
    .addItem('🟢 Excesso',           'meGerarExcesso')
    .addSeparator()
    .addItem('🔄 Atualizar Todas',   'meGerarTodas')
    .addToUi();
}

function meGerarCorretiva() { _me_gerarAba('Corretiva'); }
function meGerarPlanejada()  { _me_gerarAba('Planejada'); }
function meGerarExcesso()    { _me_gerarAba('Excesso'); }

function meGerarTodas() {
  _me_gerarAba('Corretiva');
  _me_gerarAba('Planejada');
  _me_gerarAba('Excesso');
}

// ============================================================
// ENTRADA PRINCIPAL
// ============================================================
function _me_gerarAba(tipoAba) {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando "' + tipoAba + '"... Aguarde (pode levar ~1 min).');

  try {
    const { mapaCompostos } = _me_carregarEstruturaBL();
    const vendas            = _me_carregarVendas(mapaCompostos);
    const { stockMap, locMap, dimMap, pidParaSku } = _me_carregarDadosAPI(vendas, mapaCompostos);
    const pedidos           = _me_carregarPedidos();
    const itens             = _me_montarItens(vendas, stockMap, locMap, dimMap, pedidos, mapaCompostos);
    _me_calcularCurvas(itens);
    _me_escreverAba(tipoAba, itens);
    ui.alert('✅ "' + tipoAba + '" atualizada! ' + itens.length + ' SKUs processados.');
  } catch (e) {
    ui.alert('❌ Erro: ' + e.message + '\n\n' + e.stack);
  }
}

// ============================================================
// 1. CARREGAR ESTRUTURA BL (Simples / PAI / Composto)
// ============================================================
function _me_carregarEstruturaBL() {
  const ss   = SpreadsheetApp.openById(ME.ID_BL);
  const aba  = ss.getSheetByName(ME.ABA_BL);
  if (!aba) throw new Error('Aba BL "' + ME.ABA_BL + '" não encontrada em ' + ME.ID_BL);
  const rows = aba.getDataRange().getValues();

  const mapaCompostos = {};  // skuPAI → [{ codEst, qtde }]

  for (let i = 1; i < rows.length; i++) {
    const r    = rows[i];
    const sku  = String(r[ME.BL_SKU]     || '').trim();
    const tipo = String(r[ME.BL_TIPO]    || '').trim();
    if (!sku || tipo !== 'Composto') continue;
    const cod  = String(r[ME.BL_COD_EST]  || '').trim();
    const qtde = Number(r[ME.BL_QTDE_EST] || 0);
    if (!cod) continue;
    if (!mapaCompostos[sku]) mapaCompostos[sku] = [];
    mapaCompostos[sku].push({ codEst: cod, qtde });
  }

  return { mapaCompostos };
}

// ============================================================
// 2. CARREGAR VENDAS GE (0-5, 0-7, 0-10, 0-15 dias)
//    Exclui linhas onde Canal contém "FULL"
//    Desmembra Compostos em seus componentes Simples
// ============================================================
function _me_carregarVendas(mapaCompostos) {
  const ss   = SpreadsheetApp.openById(ME.ID_GE);
  const aba  = ss.getSheetByName(ME.ABA_GE);
  if (!aba) throw new Error('Aba GE "' + ME.ABA_GE + '" não encontrada em ' + ME.ID_GE);
  const rows = aba.getDataRange().getValues();

  // Descobre a data mais recente na planilha
  let dataMax = new Date(0);
  for (let i = 1; i < rows.length; i++) {
    const d = _me_parseData(rows[i][ME.GE_DATA]);
    if (d && d > dataMax) dataMax = d;
  }
  dataMax.setHours(0, 0, 0, 0);

  // vendas: sku → { qtd5, qtd7, qtd10, qtd15 }
  const vendas = {};
  const maxDias = ME.PERIODOS[ME.PERIODOS.length - 1]; // 15

  function acum(sku, dias, qty) {
    if (!vendas[sku]) vendas[sku] = { qtd5: 0, qtd7: 0, qtd10: 0, qtd15: 0 };
    if (dias < 5)  vendas[sku].qtd5  += qty;
    if (dias < 7)  vendas[sku].qtd7  += qty;
    if (dias < 10) vendas[sku].qtd10 += qty;
    if (dias < 15) vendas[sku].qtd15 += qty;
  }

  for (let i = 1; i < rows.length; i++) {
    const r     = rows[i];
    const canal = String(r[ME.GE_CANAL] || '').toUpperCase();
    if (canal.includes('FULL')) continue;

    const d = _me_parseData(r[ME.GE_DATA]);
    if (!d) continue;
    const dNorm = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const dias  = Math.floor((dataMax - dNorm) / 86400000);
    if (dias < 0 || dias >= maxDias) continue;

    const skuGE = String(r[ME.GE_SKU] || '').trim();
    const qty   = Number(r[ME.GE_QTD] || 0);
    if (!skuGE || !qty) continue;

    if (mapaCompostos[skuGE]) {
      // PAI/Composto → distribui pelas peças
      mapaCompostos[skuGE].forEach(comp => {
        acum(comp.codEst, dias, qty * comp.qtde);
      });
    } else {
      acum(skuGE, dias, qty);
    }
  }

  return vendas;
}

// ============================================================
// 3. CARREGAR ESTOQUE / LOCALIZAÇÕES / DIMENSÕES VIA API
//    getInventoryProductsList → SKU→pid
//    getInventoryProductsData (lotes 1000) → stock, locations, dimensions
// ============================================================
function _me_carregarDadosAPI(vendas, mapaCompostos) {
  // 3a. Mapeia todos os product IDs do inventário
  const pidParaSku = {};  // pid → sku
  const skuParaPid = {};  // sku → pid
  let page = 1;
  while (true) {
    const r = _bl_call('getInventoryProductsList', {
      inventory_id: ME.INVENTORY_ID,
      page: page,
    });
    const prods = r.products || {};
    const entries = Object.entries(prods);
    if (entries.length === 0) break;
    entries.forEach(([pid, info]) => {
      const sku = String(info.sku || '').trim();
      pidParaSku[pid] = sku;
      if (sku && !skuParaPid[sku]) skuParaPid[sku] = pid;
    });
    if (entries.length < 1000) break;
    page++;
    Utilities.sleep(200);
  }

  // 3b. Coleta todos os PIDs relevantes (apenas SKUs Simples que nos interessam)
  const todosSkus = new Set([
    ...Object.keys(vendas),
    // Filhos dos compostos
    ...Object.values(mapaCompostos).flat().map(c => c.codEst),
  ]);
  // Exclui SKUs PAI (Composto) — só precisa dos Simples
  Object.keys(mapaCompostos).forEach(pai => todosSkus.delete(pai));

  const pidsList = [...todosSkus]
    .map(sku => skuParaPid[sku])
    .filter(Boolean)
    .map(Number);

  // Também inclui todos os produtos do inventário para ter localizações completas
  // (produtos sem venda também podem aparecer no excesso)
  const todosPids = Object.keys(pidParaSku).map(Number);

  // 3c. Busca dados em lotes de 1000
  const LOTE = 1000;
  const stockMap = {};  // sku → { pad, arm, chg }
  const locMap   = {};  // sku → string da melhor localização
  const dimMap   = {};  // sku → { peso, vol }

  const whs = [ME.WH_PADRAO, ME.WH_ARMAZENAMENTO, ME.WH_CHEGOU];

  for (let i = 0; i < todosPids.length; i += LOTE) {
    const lote = todosPids.slice(i, i + LOTE);
    const r    = _bl_call('getInventoryProductsData', {
      inventory_id: ME.INVENTORY_ID,
      products: lote,
    });
    Object.entries(r.products || {}).forEach(([pid, p]) => {
      const sku = pidParaSku[pid] || String(pid);
      if (!sku) return;

      const s = p.stock || {};
      stockMap[sku] = {
        pad: Number(s[ME.WH_PADRAO]        || 0),
        arm: Number(s[ME.WH_ARMAZENAMENTO] || 0),
        chg: Number(s[ME.WH_CHEGOU]        || 0),
      };

      locMap[sku] = _me_melhorLocalizacao(p.locations || {}, whs);

      const h = Number(p.height || 0);
      const w = Number(p.width  || 0);
      const l = Number(p.length || 0);
      dimMap[sku] = {
        peso: Number(p.weight || 0),
        vol:  h && w && l ? h * w * l : 0,
      };
    });

    if (i + LOTE < todosPids.length) Utilities.sleep(300);
  }

  return { stockMap, locMap, dimMap, pidParaSku };
}

// ============================================================
// 4. CARREGAR PEDIDOS ABERTOS (agrupados por método de envio)
//    Filtra payment_done > 0 (pagamento confirmado)
// ============================================================
function _me_carregarPedidos() {
  // pedidoMap: sku → { mlFlex, mlAgencia, shopDireta, shopXpress, total }
  const pedidoMap = {};

  function _addPedido(sku, metodo, qty) {
    if (!pedidoMap[sku]) {
      pedidoMap[sku] = { mlFlex: 0, mlAgencia: 0, shopDireta: 0, shopXpress: 0, total: 0 };
    }
    const m = pedidoMap[sku];
    const norm = String(metodo || '').toLowerCase();
    if      (norm.includes('flex'))   m.mlFlex    += qty;
    else if (norm.includes('agên') || norm.includes('agen')) m.mlAgencia += qty;
    else if (norm.includes('direta')) m.shopDireta += qty;
    else if (norm.includes('xpress')) m.shopXpress += qty;
    m.total += qty;
  }

  for (const statusId of ME.STATUS_IDS) {
    let idFrom = 0;
    let tentativas = 0;
    while (tentativas < 50) {
      tentativas++;
      const r     = _bl_call('getOrders', { status_id: statusId, id_from: idFrom });
      const batch = r.orders || [];
      if (batch.length === 0) break;

      batch.forEach(pedido => {
        if (!(Number(pedido.payment_done) > 0)) return; // só pagamento confirmado
        const metodo = pedido.delivery_method || '';
        (pedido.products || []).forEach(prod => {
          const sku = String(prod.sku || '').trim();
          const qty = Number(prod.quantity) || 0;
          if (sku && qty) _addPedido(sku, metodo, qty);
        });
      });

      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }
    Utilities.sleep(100);
  }

  return pedidoMap;
}

// ============================================================
// 5. MONTAR ITENS — combina todas as fontes
// ============================================================
function _me_montarItens(vendas, stockMap, locMap, dimMap, pedidoMap, mapaCompostos) {
  // Universo de SKUs: todos que temos estoque OU vendas (excluindo PAI/Composto)
  const allSkus = new Set([
    ...Object.keys(vendas),
    ...Object.keys(stockMap),
  ]);
  Object.keys(mapaCompostos).forEach(pai => allSkus.delete(pai)); // remove PAIs

  const itens = [];

  allSkus.forEach(sku => {
    const v = vendas[sku]    || { qtd5: 0, qtd7: 0, qtd10: 0, qtd15: 0 };
    const s = stockMap[sku]  || { pad: 0, arm: 0, chg: 0 };
    const d = dimMap[sku]    || { peso: 0, vol: 0 };
    const p = pedidoMap[sku] || { mlFlex: 0, mlAgencia: 0, shopDireta: 0, shopXpress: 0, total: 0 };
    const loc = locMap[sku]  || '';

    const total     = s.pad + s.arm + s.chg;
    const diferenca = s.pad - p.total;              // Padrão − pedidos abertos
    const movimentar = diferenca < 0 ? -diferenca : 0; // qty a mover de Arm/Chg → Padrão
    const excesso    = s.pad - v.qtd15;             // Padrão − vendas 15d (excesso no picking)

    itens.push({
      sku, loc,
      pad: s.pad, arm: s.arm, chg: s.chg, total,
      qtd5: v.qtd5, qtd7: v.qtd7, qtd10: v.qtd10, qtd15: v.qtd15,
      mlFlex: p.mlFlex, mlAgencia: p.mlAgencia,
      shopDireta: p.shopDireta, shopXpress: p.shopXpress,
      pedTotal: p.total,
      diferenca, movimentar, excesso,
      peso: d.peso, vol: d.vol,
      // curvas preenchidas depois
      abc: '', aabbcc: '',
    });
  });

  return itens;
}

// ============================================================
// 6. CURVAS ABC + AABBCC (baseado em vendas 0-15)
// ============================================================
function _me_calcularCurvas(itens) {
  const sorted = [...itens].sort((a, b) => b.qtd15 - a.qtd15);
  const total  = sorted.reduce((s, x) => s + x.qtd15, 0);

  let cum = 0;
  sorted.forEach(item => {
    cum += item.qtd15;
    const pct = total > 0 ? cum / total : 1;
    item.abc = pct <= ME.ABC_A ? 'A' : pct <= ME.ABC_B ? 'B' : 'C';
  });

  ['A', 'B', 'C'].forEach(grp => {
    const gi    = sorted.filter(x => x.abc === grp);
    const gTot  = gi.reduce((s, x) => s + x.qtd15, 0);
    let gCum = 0;
    gi.forEach(item => {
      gCum += item.qtd15;
      const sp = gTot > 0 ? gCum / gTot : 1;
      const sl = sp <= ME.ABC_A ? 'A' : sp <= ME.ABC_B ? 'B' : 'C';
      item.aabbcc = grp + sl;
    });
  });
}

// ============================================================
// 7. ESCREVER ABA
// ============================================================
function _me_escreverAba(tipoAba, todosItens) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   aba = ss.getSheetByName(tipoAba);
  if (!aba) aba = ss.insertSheet(tipoAba);

  // ── Filtro e ordenação por tipo de aba ──
  let itens;
  if (tipoAba === 'Corretiva') {
    // Padrão < pedidos abertos → mover URGENTE
    itens = todosItens
      .filter(x => x.diferenca < 0)
      .sort((a, b) => (a.loc || 'ZZZZ').localeCompare(b.loc || 'ZZZZ') || a.sku.localeCompare(b.sku));

  } else if (tipoAba === 'Planejada') {
    // Padrão < vendas 0-7 (vai faltar nos próximos 7 dias)
    // OU diferença < 0 mas com total suficiente no armazém
    itens = todosItens
      .filter(x => x.pad < x.qtd7 || x.diferenca < 0)
      .sort((a, b) => {
        // ABC primeiro (A > B > C), depois maior urgência
        const abcOrd = a.abc.localeCompare(b.abc);
        if (abcOrd !== 0) return abcOrd;
        return a.diferenca - b.diferenca; // mais negativo = mais urgente
      });

  } else {
    // Excesso: Padrão > vendas 0-15 (mais de 15 dias de venda no picking)
    itens = todosItens
      .filter(x => x.excesso > 0 && x.pad > 0)
      .sort((a, b) => b.excesso - a.excesso);
  }

  // ── Cabeçalho ──
  const NCOLS = 21;
  aba.clearContents();
  aba.clearFormats();

  const h1 = ['', '', 'Estoque', '', '', '', 'Vendas (vendidos)', '', '', '',
               'Pedidos Abertos', '', '', '', '', '', '', 'Dimensões', '', 'Curva', ''];
  const h2 = ['', '',
               'Padrão','Arm.','Chegou','Total',
               '0-5','0-7','0-10','0-15',
               'ML Flex','ML Agência','Shopee Direta','Shopee Xpress','Total',
               tipoAba === 'Excesso' ? 'Excesso Pick' : 'Diferença',
               tipoAba === 'Excesso' ? 'Retornar'     : 'Movimentar',
               'Peso (kg)','Vol (cm³)','ABC','AABBCC'];
  const h3 = ['SKU','Localização',
               'Padrão','Arm.','Chegou','Total',
               '0-5d','0-7d','0-10d','0-15d',
               'ML Flex','ML Agência','Shp.Direta','Shp.Xpress','Total',
               tipoAba === 'Excesso' ? 'Excesso' : 'Diferença',
               tipoAba === 'Excesso' ? 'Retornar' : 'Movimentar',
               'kg','cm³','ABC','AABBCC'];

  aba.getRange(1, 1, 1, NCOLS).setValues([h1]);
  aba.getRange(2, 1, 1, NCOLS).setValues([h2]);
  aba.getRange(3, 1, 1, NCOLS).setValues([h3]);

  _me_estilizarCabecalho(aba, tipoAba, NCOLS);

  if (itens.length === 0) { SpreadsheetApp.flush(); return; }

  // Formato texto no SKU antes de escrever
  aba.getRange(4, 1, itens.length, 1).setNumberFormat('@');

  const col16 = tipoAba === 'Excesso' ? 'excesso' : 'diferenca';
  const col17 = tipoAba === 'Excesso' ? 'excesso'  : 'movimentar';

  const linhas = itens.map(x => [
    String(x.sku),                                     // A  SKU
    x.loc,                                             // B  Localização
    x.pad, x.arm, x.chg, x.total,                     // C-F Estoque
    x.qtd5, x.qtd7, x.qtd10, x.qtd15,                 // G-J Vendas
    x.mlFlex, x.mlAgencia, x.shopDireta, x.shopXpress, x.pedTotal, // K-O Pedidos
    tipoAba === 'Excesso' ? x.excesso : x.diferenca,   // P
    tipoAba === 'Excesso' ? x.excesso : x.movimentar,  // Q
    x.peso, x.vol,                                     // R-S Dimensões
    x.abc, x.aabbcc,                                   // T-U Curvas
  ]);

  aba.getRange(4, 1, linhas.length, NCOLS).setValues(linhas);

  // Colorir curvas ABC
  const BG_ABC = { A: '#c8e6c9', B: '#fff9c4', C: '#ffcdd2' };
  const FG_ABC = { A: '#1b5e20', B: '#f57f17', C: '#b71c1c' };
  const grupos = { A: [], B: [], C: [] };
  itens.forEach((x, i) => {
    if (x.abc && BG_ABC[x.abc]) grupos[x.abc].push('T' + (4 + i));
  });
  ['A', 'B', 'C'].forEach(g => {
    if (grupos[g].length) {
      aba.getRangeList(grupos[g]).setBackground(BG_ABC[g]).setFontColor(FG_ABC[g]);
    }
  });

  // Colorir Diferença/Excesso: vermelho se negativo (Corretiva/Planejada)
  if (tipoAba !== 'Excesso') {
    const vermelhos = [];
    itens.forEach((x, i) => { if (x.diferenca < 0) vermelhos.push('P' + (4 + i)); });
    if (vermelhos.length) aba.getRangeList(vermelhos).setBackground('#ffcdd2');
  }

  // Data/hora de atualização
  aba.getRange(1, NCOLS + 1).setValue('Atualização');
  aba.getRange(2, NCOLS + 1).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  );

  SpreadsheetApp.flush();
}

// ============================================================
// 8. ESTILIZAR CABEÇALHO
// ============================================================
function _me_estilizarCabecalho(aba, tipoAba, NCOLS) {
  const CORES = {
    Corretiva: '#d32f2f',
    Planejada: '#f57f17',
    Excesso:   '#388e3c',
  };
  const cor = CORES[tipoAba] || '#455a64';

  // Linha 1: grupos
  const grupos = [
    [1,  2,  '#455a64'],  // SKU + Localização
    [3,  6,  '#1565c0'],  // Estoque
    [7,  10, '#00838f'],  // Vendas
    [11, 15, '#4527a0'],  // Pedidos
    [16, 17, cor],        // Diferença/Movimentar ou Excesso/Retornar
    [18, 19, '#4e342e'],  // Dimensões
    [20, 21, '#558b2f'],  // Curvas
  ];
  grupos.forEach(([de, ate, bg]) => {
    aba.getRange(1, de, 1, ate - de + 1)
       .setBackground(bg).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
    if (ate > de) aba.getRange(1, de, 1, ate - de + 1).merge();
  });

  // Linha 2 e 3: sub-cabeçalhos
  aba.getRange(2, 1, 2, NCOLS)
     .setBackground('#263238').setFontColor('#ffffff')
     .setFontWeight('bold').setHorizontalAlignment('center')
     .setVerticalAlignment('middle').setWrap(true);

  // Alturas
  aba.setRowHeight(1, 25);
  aba.setRowHeight(2, 25);
  aba.setRowHeight(3, 40);

  // Larguras
  aba.setColumnWidth(1, 110);  // SKU
  aba.setColumnWidth(2, 140);  // Localização

  // Congelar
  aba.setFrozenRows(3);
  aba.setFrozenColumns(2);
}

// ============================================================
// HELPERS
// ============================================================

// Seleciona a melhor localização entre os armazéns ativos
// Prioridade: menor Block (B1 < B2) → menor Floor (A1 < A2)
function _me_melhorLocalizacao(locsObj, warehouseIds) {
  let bestRaw     = '';
  let bestBlock   = 9999;
  let bestFloor   = 9999;

  warehouseIds.forEach(wid => {
    const loc = String(locsObj[wid] || '').trim();
    if (!loc) return;
    const parts = loc.split('-');
    // Formato: A9-B2-RA-C3-A2
    if (parts.length < 5) return;
    const blockNum = parseInt(parts[1].substring(1)) || 9999;
    const floorNum = parseInt(parts[4].substring(1)) || 9999;
    if (blockNum < bestBlock || (blockNum === bestBlock && floorNum < bestFloor)) {
      bestBlock = blockNum;
      bestFloor = floorNum;
      bestRaw   = loc;
    }
  });

  return bestRaw;
}

// Parse de data em DD/MM/YYYY ou Date
function _me_parseData(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
