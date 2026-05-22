// ============================================================
// PLANEJAMENTO PICKING — BaseLinker → Google Sheets
// Pedidos dos últimos 7 dias (D-1 até D-7)
// ============================================================

const PK_BL_TOKEN   = '8004176-8026704-5DUYJBOPVCCE3W6VATUJEEAJY8P7Z4YS2IHQCWEU8YAM2RR74VA1N2RE95PVYWGZ';
const PK_BL_URL     = 'https://api.baselinker.com/connector.php';
// Slack webhook URL — store via: File > Project properties > Script properties
// Key: SLACK_HOOK  Value: https://hooks.slack.com/services/...
const PK_SLACK_HOOK = PropertiesService.getScriptProperties().getProperty('SLACK_HOOK') || '';

// Métodos de envio a EXCLUIR (Full)
const PK_EXCLUIR_METODOS = ['ME2 - Mercado Envios Full', 'Fulfilled by Shopee'];

// Status a EXCLUIR — IDs exatos do BaseLinker
const PK_STATUS_EXCLUIR = new Set([
  '252554',  // Cancelado
  '268816',  // Pedidos Fulfillment
  '268817',  // [Full] Separação
  '268818',  // [Full] Montagem
  '268820',  // [Full] Aguardando coleta
  '268821',  // [Full] Coletado
  '269423',  // [Full] Separando
]);

const PK_INVENTORY_ID      = '39947';
const PK_WH_PADRAO         = 'bl_44285';
const PK_WH_ARMAZENAMENTO  = 'bl_50394';
const PK_WH_CHEGOU         = 'bl_51442';

const PK_ABA_ANALISE  = 'Análise Picking';
const PK_ABA_EXCESSO  = 'Picking em Excesso';
const PK_ABC = [0.50, 0.80];

const PK_CURVE_WEIGHTS = {
  'AA': 1.0, 'AB': 1.4, 'AC': 1.8,
  'BA': 2.5, 'BB': 3.2, 'BC': 4.0,
  'CA': 5.0, 'CB': 6.5, 'CC': 8.0,
};

const PK_CABECALHO = ['SKU','Quantidade','Status Base','ID Pedido base','ID Marketplace','Data da Venda'];

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Picking')
    .addItem('🔄 Atualizar base de vendas (Página1)', 'atualizarPicking')
    .addItem('📊 Atualizar análise de movimentação', 'atualizarAnalise')
    .addItem('📦 Atualizar picking em excesso', 'atualizarExcesso')
    .addSeparator()
    .addItem('▶ Ativar atualização automática (3x/dia)', 'ativarTriggerPicking')
    .addItem('⏹ Desativar atualização automática', 'desativarTriggerPicking')
    .addToUi();
}

// ============================================================
// FUNÇÃO PRINCIPAL — base de vendas (Página1)
// ============================================================
function atualizarPicking() {
  const inicio = new Date();
  let erroMsg  = null;
  let totalLinhas = 0;

  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getActiveSheet();
    const statusMap = _pkMapaStatus();

    const agora     = new Date();
    const inicioDia = new Date(agora);
    inicioDia.setHours(0, 0, 0, 0);
    const tsTo   = Math.floor(inicioDia.getTime() / 1000) - 1;
    const tsFrom = Math.floor(inicioDia.getTime() / 1000) - (7 * 86400);

    const rows    = [];
    let dataAtual = tsFrom;
    let continuar = true;

    while (continuar) {
      const resp    = _pkAPI('getOrders', { date_confirmed_from: dataAtual, get_unconfirmed_orders: false });
      const pedidos = resp.orders || [];
      if (pedidos.length === 0) break;

      for (const pedido of pedidos) {
        if (pedido.date_confirmed > tsTo) { continuar = false; break; }

        // Filtro de status por ID exato
        if (PK_STATUS_EXCLUIR.has(String(pedido.order_status_id))) continue;

        // Filtro de método de envio: pula Full
        const metodo = String(pedido.delivery_method || '');
        if (PK_EXCLUIR_METODOS.includes(metodo)) continue;

        const statusNome = statusMap[String(pedido.order_status_id)] || '';
        const idExterno  = String(pedido.external_order_id || pedido.shop_order_id || '');
        const dataVenda  = pedido.date_confirmed
          ? new Date(pedido.date_confirmed * 1000).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }) : '';

        for (const prod of (pedido.products || [])) {
          const sku = String(prod.sku || prod.product_id || '').trim();
          if (!sku) continue;
          rows.push([sku, Number(prod.quantity) || 0, statusNome, String(pedido.order_id), idExterno, dataVenda]);
        }
      }

      if (continuar && pedidos.length >= 100) {
        dataAtual = pedidos[pedidos.length - 1].date_confirmed + 1;
        Utilities.sleep(300);
      } else { continuar = false; }
    }

    aba.getRange(1, 1, aba.getMaxRows(), 1).setNumberFormat('@');
    SpreadsheetApp.flush();
    aba.clearContents();
    aba.getRange(1, 1, 1, PK_CABECALHO.length).setValues([PK_CABECALHO]);
    if (rows.length > 0) aba.getRange(2, 1, rows.length, PK_CABECALHO.length).setValues(rows);
    aba.getRange(1, 1, 1, PK_CABECALHO.length)
       .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, PK_CABECALHO.length);
    SpreadsheetApp.flush();
    totalLinhas = rows.length;

  } catch (e) {
    erroMsg = e.message;
    Logger.log('Erro atualizarPicking: ' + e.stack);
  }

  const duracao = ((new Date() - inicio) / 1000).toFixed(1);
  if (erroMsg) {
    _pkSlack('❌ *Picking — Erro na atualização*\nErro: `' + erroMsg + '`\nDuração: ' + duracao + 's');
  } else {
    _pkSlack('✅ *Picking — Atualizado com sucesso*\n📦 ' + totalLinhas + ' linhas importadas\n⏱ Duração: ' + duracao + 's');
  }
}

// ============================================================
// TRIGGERS — 3x por dia (06h, 11h, 17h BRT)
// ============================================================
function ativarTriggerPicking() {
  desativarTriggerPicking(false);
  [9, 14, 20].forEach(hora => ScriptApp.newTrigger('atualizarPicking').timeBased().atHour(hora).everyDays(1).create());
  SpreadsheetApp.getUi().alert('✅ Atualização automática ativada!\nHorários BRT: 06:00 · 11:00 · 17:00');
}

function desativarTriggerPicking(mostrarAlerta) {
  if (mostrarAlerta === undefined) mostrarAlerta = true;
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'atualizarPicking').forEach(t => ScriptApp.deleteTrigger(t));
  if (mostrarAlerta) SpreadsheetApp.getUi().alert('⏹ Atualização automática desativada.');
}

// ============================================================
// HELPERS
// ============================================================
function _pkAPI(metodo, params) {
  const resp = UrlFetchApp.fetch(PK_BL_URL, {
    method: 'post',
    headers: { 'X-BLToken': PK_BL_TOKEN },
    payload: { method: metodo, parameters: JSON.stringify(params || {}) },
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (!data || data.status !== 'SUCCESS')
    throw new Error('[' + metodo + '] ' + (data && data.error_message ? data.error_message : JSON.stringify(data)));
  return data;
}

function _pkMapaStatus() {
  const resp = _pkAPI('getOrderStatusList', {});
  const mapa = {};
  for (const s of (resp.statuses || [])) mapa[String(s.id)] = s.name;
  return mapa;
}

function _pkSlack(texto) {
  if (!PK_SLACK_HOOK || !PK_SLACK_HOOK.startsWith('https://')) { Logger.log('Slack: ' + texto); return; }
  try {
    UrlFetchApp.fetch(PK_SLACK_HOOK, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: texto }), muteHttpExceptions: true });
  } catch (e) { Logger.log('Slack erro: ' + e.message); }
}

// ============================================================
// ANÁLISE DE MOVIMENTAÇÃO — gera/atualiza aba "Análise Picking"
// ============================================================
function atualizarAnalise() {
  const inicio = new Date();
  let erroMsg  = null;
  let totalLinhas = 0;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Ler Página1 e agregar por SKU
    const abaBase = ss.getSheets()[0];
    const dados   = abaBase.getDataRange().getValues();
    const skuAgg  = {};
    for (let i = 1; i < dados.length; i++) {
      const sku = String(dados[i][0]).replace(/^'/, '').trim();
      if (!sku) continue;
      const qty    = Number(dados[i][1]) || 0;
      const status = String(dados[i][2]).trim();
      if (!skuAgg[sku]) skuAgg[sku] = { totalQty: 0, movQty: 0 };
      skuAgg[sku].totalQty += qty;
      if (status === 'Movimentação') skuAgg[sku].movQty += qty;
    }

    const skusVendidos = Object.keys(skuAgg);
    if (skusVendidos.length === 0) throw new Error('Página1 está vazia. Atualize a base de vendas primeiro.');

    // 2. Carregar mapa SKU → productId do BaseLinker
    const skuToId = {};
    let page = 1;
    while (true) {
      const resp    = _pkAPI('getInventoryProductsList', { inventory_id: PK_INVENTORY_ID, page });
      const entries = Object.entries(resp.products || {});
      if (entries.length === 0) break;
      entries.forEach(([pid, p]) => { const s = String(p.sku || '').trim(); if (s) skuToId[s] = pid; });
      if (entries.length < 1000) break;
      page++;
    }

    // 3. Buscar estoque por productId
    const productIds = skusVendidos.map(s => skuToId[s]).filter(Boolean);
    const stockMap   = {};
    const LOTE = 1000;
    for (let i = 0; i < productIds.length; i += LOTE) {
      const resp = _pkAPI('getInventoryProductsData', { inventory_id: PK_INVENTORY_ID, products: productIds.slice(i, i + LOTE) });
      Object.entries(resp.products || {}).forEach(([pid, d]) => {
        const s = d.stock || {};
        stockMap[pid] = { pad: Number(s[PK_WH_PADRAO] || 0), arm: Number(s[PK_WH_ARMAZENAMENTO] || 0), chg: Number(s[PK_WH_CHEGOU] || 0) };
      });
    }

    // 4. Calcular curva AA sobre todos os SKUs vendidos
    const todosItens = skusVendidos.map(sku => ({ sku, qty: skuAgg[sku].totalQty }));
    _atribuirCurvaAA(todosItens);
    const curvaMap = {};
    todosItens.forEach(item => { curvaMap[item.sku] = item.curva; });

    // 5. Montar linhas filtradas
    const rows = [];
    skusVendidos.forEach(sku => {
      const agg = skuAgg[sku];
      const pid = skuToId[sku];
      const est = pid ? (stockMap[pid] || { pad:0, arm:0, chg:0 }) : { pad:0, arm:0, chg:0 };
      const B = agg.totalQty, C = agg.movQty;
      const D = est.pad, E = est.arm, F = est.chg;
      const G = D + E + F;
      if (G <= 0 || B - D <= 0) return;
      const H     = Math.min(G, B - D);
      const dias  = B > 0 ? Math.max(0, D) * 7 / B : 0;
      const curva = curvaMap[sku] || 'CC';
      const peso  = PK_CURVE_WEIGHTS[curva] || 8.0;
      const score = (dias + 1) * peso / B;
      rows.push({ sku, B, C, D, E, F, G, H, dias, curva, score });
    });

    // 6. Ordenar por Score de Urgência (menor = mais urgente)
    rows.sort((a, b) => a.score - b.score);

    // 7. Escrever na aba
    let aba = ss.getSheetByName(PK_ABA_ANALISE);
    if (!aba) aba = ss.insertSheet(PK_ABA_ANALISE);
    aba.clearContents();
    aba.clearFormats();

    const h1 = ['SKU','Qntd. Vend.','Qntd. Status\nMovimentação','Estoque','','','','Qntd.\nMovimentar','Dias de\nCobertura','Curva','Score','Última Atualização'];
    aba.getRange(1, 1, 1, 12).setValues([h1]);
    aba.getRange(1, 4, 1, 4).merge();
    const h2 = ['','','','Padrão','Armazenamento','Chegou','Saldo Total','','','','',''];
    aba.getRange(2, 1, 1, 12).setValues([h2]);
    aba.getRange(1, 1, 2, 12)
       .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold')
       .setHorizontalAlignment('center').setVerticalAlignment('middle');
    aba.setFrozenRows(2);

    const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    if (rows.length > 0) {
      aba.getRange(3, 1, rows.length, 1).setNumberFormat('@');
      const dataRows = rows.map(r => [r.sku, r.B, r.C, r.D, r.E, r.F, r.G, r.H, r.dias, r.curva, r.score, '']);
      aba.getRange(3, 1, rows.length, 12).setValues(dataRows);
      aba.getRange(3, 9,  rows.length, 1).setNumberFormat('0.0');
      aba.getRange(3, 11, rows.length, 1).setNumberFormat('0.000');
    }
    aba.getRange(3, 12).setValue(agora);
    aba.autoResizeColumns(1, 12);
    SpreadsheetApp.flush();
    totalLinhas = rows.length;

  } catch (e) {
    erroMsg = e.message;
    Logger.log('Erro atualizarAnalise: ' + e.stack);
  }

  const duracao = ((new Date() - inicio) / 1000).toFixed(1);
  if (erroMsg) {
    _pkSlack('❌ *Análise Picking — Erro*\nErro: `' + erroMsg + '`\nDuração: ' + duracao + 's');
  } else {
    _pkSlack('✅ *Análise Picking — Atualizada*\n📊 ' + totalLinhas + ' SKUs | ⏱ ' + duracao + 's');
  }
}

// ============================================================
// CURVA AA — classifica itens em 2 níveis ABC por volume
// ============================================================
function _atribuirCurvaAA(items) {
  if (!items || items.length === 0) return;

  function classificar(grupo, nivel) {
    if (grupo.length === 0) return;
    grupo.sort((a, b) => b.qty - a.qty);
    const total = grupo.reduce((s, x) => s + x.qty, 0);
    let cum = 0;
    grupo.forEach(item => {
      cum += item.qty;
      const pct = total > 0 ? cum / total : 1;
      item['_l' + nivel] = pct <= PK_ABC[0] ? 'A' : pct <= PK_ABC[1] ? 'B' : 'C';
    });
  }

  classificar(items, 1);
  for (const l1 of ['A','B','C']) classificar(items.filter(x => x._l1 === l1), 2);
  items.forEach(item => { item.curva = (item._l1 || 'C') + (item._l2 || 'C'); });
}

// ============================================================
// PICKING EM EXCESSO — SKUs com estoque padrão > demanda 7 dias
// ============================================================
function atualizarExcesso() {
  const inicio = new Date();
  let erroMsg  = null;
  let totalLinhas = 0;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Ler Página1 e agregar por SKU
    const abaBase = ss.getSheets()[0];
    const dados   = abaBase.getDataRange().getValues();
    const skuAgg  = {};
    for (let i = 1; i < dados.length; i++) {
      const sku = String(dados[i][0]).replace(/^'/, '').trim();
      if (!sku) continue;
      const qty = Number(dados[i][1]) || 0;
      if (!skuAgg[sku]) skuAgg[sku] = { totalQty: 0 };
      skuAgg[sku].totalQty += qty;
    }

    const skusVendidos = Object.keys(skuAgg);
    if (skusVendidos.length === 0) throw new Error('Página1 está vazia. Atualize a base de vendas primeiro.');

    // 2. Carregar mapa SKU → productId do BaseLinker
    const skuToId = {};
    let page = 1;
    while (true) {
      const resp    = _pkAPI('getInventoryProductsList', { inventory_id: PK_INVENTORY_ID, page });
      const entries = Object.entries(resp.products || {});
      if (entries.length === 0) break;
      entries.forEach(([pid, p]) => { const s = String(p.sku || '').trim(); if (s) skuToId[s] = pid; });
      if (entries.length < 1000) break;
      page++;
    }

    // 3. Buscar estoque por productId
    const productIds = skusVendidos.map(s => skuToId[s]).filter(Boolean);
    const stockMap   = {};
    const LOTE = 1000;
    for (let i = 0; i < productIds.length; i += LOTE) {
      const resp = _pkAPI('getInventoryProductsData', { inventory_id: PK_INVENTORY_ID, products: productIds.slice(i, i + LOTE) });
      Object.entries(resp.products || {}).forEach(([pid, d]) => {
        const s = d.stock || {};
        stockMap[pid] = Number(s[PK_WH_PADRAO] || 0);
      });
    }

    // 4. Calcular curva AA
    const todosItens = skusVendidos.map(sku => ({ sku, qty: skuAgg[sku].totalQty }));
    _atribuirCurvaAA(todosItens);
    const curvaMap = {};
    todosItens.forEach(item => { curvaMap[item.sku] = item.curva; });

    // 5. Filtrar apenas SKUs com excesso (estoque padrão > vendas 7 dias)
    const rows = [];
    skusVendidos.forEach(sku => {
      const B   = skuAgg[sku].totalQty;
      const pid = skuToId[sku];
      const D   = pid ? (stockMap[pid] || 0) : 0;
      const excesso = D - B;
      if (excesso <= 0) return;
      const curva = curvaMap[sku] || 'CC';
      rows.push({ sku, B, D, excesso, curva });
    });

    // 6. Ordenar por excesso decrescente (mais excesso primeiro)
    rows.sort((a, b) => b.excesso - a.excesso);

    // 7. Escrever na aba
    let aba = ss.getSheetByName(PK_ABA_EXCESSO);
    if (!aba) aba = ss.insertSheet(PK_ABA_EXCESSO);
    aba.clearContents();
    aba.clearFormats();

    const cabecalho = ['SKU','Qntd. Vend. 7d','Estoque Padrão','Excesso','Curva','Última Atualização'];
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    aba.getRange(1, 1, 1, cabecalho.length)
       .setBackground('#e8430a').setFontColor('#ffffff').setFontWeight('bold')
       .setHorizontalAlignment('center');
    aba.setFrozenRows(1);

    const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    if (rows.length > 0) {
      aba.getRange(2, 1, rows.length, 1).setNumberFormat('@');
      const dataRows = rows.map((r, i) => [r.sku, r.B, r.D, r.excesso, r.curva, i === 0 ? agora : '']);
      aba.getRange(2, 1, rows.length, cabecalho.length).setValues(dataRows);
    } else {
      aba.getRange(2, 6).setValue(agora);
    }
    aba.autoResizeColumns(1, cabecalho.length);
    SpreadsheetApp.flush();
    totalLinhas = rows.length;

  } catch (e) {
    erroMsg = e.message;
    Logger.log('Erro atualizarExcesso: ' + e.stack);
  }

  const duracao = ((new Date() - inicio) / 1000).toFixed(1);
  if (erroMsg) {
    _pkSlack('❌ *Picking em Excesso — Erro*\nErro: `' + erroMsg + '`\nDuração: ' + duracao + 's');
  } else {
    _pkSlack('✅ *Picking em Excesso — Atualizado*\n📦 ' + totalLinhas + ' SKUs em excesso | ⏱ ' + duracao + 's');
  }
}

// ============================================================
// UTILITÁRIO — lista todos os status (use para diagnóstico)
// ============================================================
function listarTodosStatus() {
  const resp = _pkAPI('getOrderStatusList', {});
  const nomes = (resp.statuses || []).map(s => s.id + ' → ' + s.name).join('\n');
  SpreadsheetApp.getUi().alert(nomes);
}

// ============================================================
// UTILITÁRIO — inspeciona dados completos de um produto por SKU
// ============================================================
function inspecionarProduto() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('Inspetor de Produto', 'Digite o SKU do produto:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const skuBusca = res.getResponseText().trim();
  if (!skuBusca) return;

  // Encontrar o productId pelo SKU
  let productId = null;
  let page = 1;
  while (!productId) {
    const resp    = _pkAPI('getInventoryProductsList', { inventory_id: PK_INVENTORY_ID, page });
    const entries = Object.entries(resp.products || {});
    if (entries.length === 0) break;
    for (const [pid, p] of entries) {
      if (String(p.sku || '').trim() === skuBusca) { productId = pid; break; }
    }
    if (entries.length < 1000) break;
    page++;
  }

  if (!productId) {
    ui.alert('SKU "' + skuBusca + '" não encontrado no inventário.');
    return;
  }

  const resp = _pkAPI('getInventoryProductsData', { inventory_id: PK_INVENTORY_ID, products: [productId] });
  const produto = (resp.products || {})[productId];
  if (!produto) {
    ui.alert('Produto encontrado (ID ' + productId + ') mas sem dados retornados.');
    return;
  }

  // Logar JSON completo no console e exibir resumo no alerta
  Logger.log('=== PRODUTO: ' + skuBusca + ' (ID: ' + productId + ') ===\n' + JSON.stringify(produto, null, 2));

  const location = produto.location || produto.locations || produto.storage_location || '(campo não encontrado)';
  const estoque  = JSON.stringify(produto.stock || {});
  ui.alert(
    'SKU: ' + skuBusca + '\nID BaseLinker: ' + productId +
    '\n\nlocation: ' + JSON.stringify(location) +
    '\nstock: ' + estoque +
    '\n\n⚠ JSON completo disponível em: Visualizar > Logs'
  );
}
