// ============================================================
// PLANEJAMENTO PICKING — BaseLinker → Google Sheets
// Pedidos dos últimos 7 dias (D-1 até D-7)
// Excluindo: Cancelados, ME2 - Mercado Envios Full, Fulfilled by Shopee
// ============================================================

const PK_BL_TOKEN   = '8004176-8026704-5DUYJBOPVCCE3W6VATUJEEAJY8P7Z4YS2IHQCWEU8YAM2RR74VA1N2RE95PVYWGZ';
const PK_BL_URL     = 'https://api.baselinker.com/connector.php';
const PK_SLACK_HOOK = 'COLE_AQUI_O_WEBHOOK_DO_SLACK'; // substitua pela URL do seu Incoming Webhook

// Métodos de envio a EXCLUIR (Full)
const PK_EXCLUIR_METODOS = ['ME2 - Mercado Envios Full', 'Fulfilled by Shopee'];

// Inventário e armazéns BaseLinker
const PK_INVENTORY_ID      = '39947';
const PK_WH_PADRAO         = 'bl_44285';
const PK_WH_ARMAZENAMENTO  = 'bl_50394';
const PK_WH_CHEGOU         = 'bl_51442';

// Aba de análise
const PK_ABA_ANALISE = 'Análise Picking';

// Thresholds curva ABC por nível (% acumulado do volume)
// A: 0–50% | B: 50–80% | C: 80–100%
const PK_ABC = [0.50, 0.80];

const PK_CABECALHO = [
  'SKU',            // A
  'Quantidade',     // B
  'Status Base',    // C
  'ID Pedido base', // D
  'ID Marketplace', // E
  'Data da Venda',  // F
];

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Picking')
    .addItem('🔄 Atualizar base de vendas (Página1)', 'atualizarPicking')
    .addItem('📊 Atualizar análise de movimentação', 'atualizarAnalise')
    .addSeparator()
    .addItem('▶ Ativar atualização automática (3x/dia)', 'ativarTriggerPicking')
    .addItem('⏹ Desativar atualização automática', 'desativarTriggerPicking')
    .addToUi();
}

// ============================================================
// FUNÇÃO PRINCIPAL
// ============================================================
function atualizarPicking() {
  const inicio = new Date();
  let erroMsg  = null;
  let totalLinhas = 0;

  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getActiveSheet();

    // Mapa de status: id → nome
    const statusMap = _pkMapaStatus();

    // Período: D-7 00:00 BRT até D-1 23:59:59 BRT
    // Apps Script roda em UTC; BRT = UTC-3 → deslocar 3h
    const agora     = new Date();
    const inicioDia = new Date(agora);
    inicioDia.setHours(0, 0, 0, 0);

    // D-1 23:59:59 = início de hoje - 1s
    const tsTo   = Math.floor(inicioDia.getTime() / 1000) - 1;
    // D-7 00:00   = início de hoje - 7 dias
    const tsFrom = Math.floor(inicioDia.getTime() / 1000) - (7 * 86400);

    // ---- Buscar pedidos ----
    const rows     = [];
    let dataAtual  = tsFrom;
    let continuar  = true;

    while (continuar) {
      const resp    = _pkAPI('getOrders', {
        date_confirmed_from:    dataAtual,
        get_unconfirmed_orders: false,
      });
      const pedidos = resp.orders || [];
      if (pedidos.length === 0) break;

      for (const pedido of pedidos) {
        // Parar quando ultrapassar o fim do período
        if (pedido.date_confirmed > tsTo) {
          continuar = false;
          break;
        }

        // Filtro de status: pula Cancelado
        const statusNome = statusMap[String(pedido.order_status_id)] || '';
        if (statusNome.toLowerCase().includes('cancelad')) continue;

        // Filtro de método de envio: pula Full
        const metodo = String(pedido.delivery_method || '');
        if (PK_EXCLUIR_METODOS.includes(metodo)) continue;

        // ID Marketplace = external_order_id ou shop_order_id
        const idExterno = String(
          pedido.external_order_id || pedido.shop_order_id || ''
        );

        // Data formatada em BRT
        const dataVenda = pedido.date_confirmed
          ? new Date(pedido.date_confirmed * 1000)
              .toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' })
          : '';

        for (const prod of (pedido.products || [])) {
          const sku = String(prod.sku || prod.product_id || '').trim();
          if (!sku) continue;
          rows.push([
            sku,
            Number(prod.quantity) || 0,
            statusNome,
            String(pedido.order_id),
            idExterno,
            dataVenda,
          ]);
        }
      }

      // Próxima página: avança pelo timestamp do último pedido
      if (continuar && pedidos.length >= 100) {
        dataAtual = pedidos[pedidos.length - 1].date_confirmed + 1;
        Utilities.sleep(300);
      } else {
        continuar = false;
      }
    }

    // ---- Escrever na planilha (batch) ----
    // Formatar coluna A inteira como texto ANTES de qualquer escrita
    aba.getRange(1, 1, aba.getMaxRows(), 1).setNumberFormat('@');
    SpreadsheetApp.flush();

    aba.clearContents();
    aba.getRange(1, 1, 1, PK_CABECALHO.length).setValues([PK_CABECALHO]);

    if (rows.length > 0) {
      aba.getRange(2, 1, rows.length, PK_CABECALHO.length).setValues(rows);
    }

    // Formatar cabeçalho
    aba.getRange(1, 1, 1, PK_CABECALHO.length)
       .setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');
    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, PK_CABECALHO.length);

    SpreadsheetApp.flush();
    totalLinhas = rows.length;

  } catch (e) {
    erroMsg = e.message;
    Logger.log('Erro atualizarPicking: ' + e.stack);
  }

  // ---- Notificação Slack ----
  const duracao = ((new Date() - inicio) / 1000).toFixed(1);
  if (erroMsg) {
    _pkSlack(
      '❌ *Picking — Erro na atualização*\n' +
      'Erro: `' + erroMsg + '`\n' +
      'Duração: ' + duracao + 's'
    );
  } else {
    _pkSlack(
      '✅ *Picking — Atualizado com sucesso*\n' +
      '📦 ' + totalLinhas + ' linhas importadas\n' +
      '⏱ Duração: ' + duracao + 's'
    );
  }
}

// ============================================================
// TRIGGERS — 3x por dia
// Horários UTC: 09h, 14h, 20h = 06h, 11h, 17h BRT
// ============================================================
function ativarTriggerPicking() {
  desativarTriggerPicking(false);
  [9, 14, 20].forEach(hora => {
    ScriptApp.newTrigger('atualizarPicking')
      .timeBased()
      .atHour(hora)
      .everyDays(1)
      .create();
  });
  SpreadsheetApp.getUi().alert(
    '✅ Atualização automática ativada!\n' +
    'Horários BRT: 06:00 · 11:00 · 17:00'
  );
}

function desativarTriggerPicking(mostrarAlerta) {
  if (mostrarAlerta === undefined) mostrarAlerta = true;
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'atualizarPicking')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (mostrarAlerta)
    SpreadsheetApp.getUi().alert('⏹ Atualização automática desativada.');
}

// ============================================================
// HELPERS
// ============================================================
function _pkAPI(metodo, params) {
  const resp = UrlFetchApp.fetch(PK_BL_URL, {
    method: 'post',
    headers: { 'X-BLToken': PK_BL_TOKEN },
    payload: {
      method:     metodo,
      parameters: JSON.stringify(params || {}),
    },
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (!data || data.status !== 'SUCCESS') {
    throw new Error(
      '[' + metodo + '] ' +
      (data && data.error_message ? data.error_message : JSON.stringify(data))
    );
  }
  return data;
}

function _pkMapaStatus() {
  const resp = _pkAPI('getOrderStatusList', {});
  const mapa = {};
  for (const s of (resp.statuses || [])) {
    mapa[String(s.id)] = s.name;
  }
  return mapa;
}

function _pkSlack(texto) {
  if (!PK_SLACK_HOOK || !PK_SLACK_HOOK.startsWith('https://')) {
    Logger.log('Slack (sem webhook): ' + texto);
    return;
  }
  try {
    UrlFetchApp.fetch(PK_SLACK_HOOK, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify({ text: texto }),
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('Slack erro: ' + e.message);
  }
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

    // ── 1. Ler Página1 e agregar por SKU ──────────────────────
    const abaBase  = ss.getSheets()[0];
    const dados    = abaBase.getDataRange().getValues();
    const skuAgg   = {}; // sku → { totalQty, movQty }

    for (let i = 1; i < dados.length; i++) {
      const sku    = String(dados[i][0]).replace(/^'/, '').trim();
      if (!sku) continue;
      const qty    = Number(dados[i][1]) || 0;
      const status = String(dados[i][2]).trim();

      if (!skuAgg[sku]) skuAgg[sku] = { totalQty: 0, movQty: 0 };
      skuAgg[sku].totalQty += qty;
      if (status === 'Movimentação') skuAgg[sku].movQty += qty;
    }

    const skusVendidos = Object.keys(skuAgg);
    if (skusVendidos.length === 0) throw new Error('Página1 está vazia. Atualize a base de vendas primeiro.');

    // ── 2. Carregar mapa SKU → productId do BaseLinker ────────
    const skuToId = {};
    let page = 1;
    while (true) {
      const resp     = _pkAPI('getInventoryProductsList', { inventory_id: PK_INVENTORY_ID, page });
      const products = resp.products || {};
      const entries  = Object.entries(products);
      if (entries.length === 0) break;
      entries.forEach(([pid, p]) => {
        const s = String(p.sku || '').trim();
        if (s) skuToId[s] = pid;
      });
      if (entries.length < 1000) break;
      page++;
    }

    // ── 3. Buscar estoque por productId ───────────────────────
    const productIds = skusVendidos.map(s => skuToId[s]).filter(Boolean);
    const stockMap   = {}; // productId → { pad, arm, chg }
    const LOTE = 1000;
    for (let i = 0; i < productIds.length; i += LOTE) {
      const resp = _pkAPI('getInventoryProductsData', {
        inventory_id: PK_INVENTORY_ID,
        products: productIds.slice(i, i + LOTE),
      });
      Object.entries(resp.products || {}).forEach(([pid, d]) => {
        const s = d.stock || {};
        stockMap[pid] = {
          pad: Number(s[PK_WH_PADRAO]        || 0),
          arm: Number(s[PK_WH_ARMAZENAMENTO] || 0),
          chg: Number(s[PK_WH_CHEGOU]        || 0),
        };
      });
    }

    // ── 4. Calcular curva AAA sobre TODOS os SKUs vendidos ────
    const todosItens = skusVendidos.map(sku => ({ sku, qty: skuAgg[sku].totalQty }));
    _atribuirCurvaAAA(todosItens);
    const curvaMap = {};
    todosItens.forEach(item => { curvaMap[item.sku] = item.curva; });

    // ── 5. Montar linhas filtradas ────────────────────────────
    const rows = [];
    skusVendidos.forEach(sku => {
      const agg = skuAgg[sku];
      const pid = skuToId[sku];
      const est = pid ? (stockMap[pid] || { pad:0, arm:0, chg:0 }) : { pad:0, arm:0, chg:0 };

      const B = agg.totalQty;
      const C = agg.movQty;
      const D = est.pad;
      const E = est.arm;
      const F = est.chg;
      const G = D + E + F;

      // Filtro: só aparece se G > 0 e D < B (estoque Padrão não cobre demanda)
      if (G <= 0 || B - D <= 0) return;

      const H = Math.min(G, B - D);        // Qntd. Movimentar
      const I = B > 0 ? D / B : 0;         // % Cobertura = Padrão / Vendido
      const curva = curvaMap[sku] || 'CCC';

      rows.push({ sku, B, C, D, E, F, G, H, I, curva });
    });

    // ── 6. Ordenar: curva AAA→CCC, depois menor cobertura ─────
    const CURVE_ORDER = {};
    let idx = 0;
    for (const l1 of ['A','B','C'])
      for (const l2 of ['A','B','C'])
        for (const l3 of ['A','B','C'])
          CURVE_ORDER[l1+l2+l3] = idx++;

    rows.sort((a, b) => {
      const dc = (CURVE_ORDER[a.curva] || 0) - (CURVE_ORDER[b.curva] || 0);
      return dc !== 0 ? dc : a.I - b.I;
    });

    // ── 7. Escrever na aba ────────────────────────────────────
    let aba = ss.getSheetByName(PK_ABA_ANALISE);
    if (!aba) aba = ss.insertSheet(PK_ABA_ANALISE);

    aba.clearContents();
    aba.clearFormats();

    // Cabeçalho linha 1 (com merge de Estoque D:G)
    const h1 = ['SKU','Qntd. Vend.','Qntd. Status\nMovimentação',
                 'Estoque','','','',
                 'Qntd.\nMovimentar','% de\nCobertura','Curva','','Última Atualização'];
    aba.getRange(1, 1, 1, 12).setValues([h1]);
    aba.getRange(1, 4, 1, 4).merge(); // mescla D1:G1 → "Estoque"

    // Cabeçalho linha 2 (sub-títulos dos armazéns)
    const h2 = ['','','','Padrão','Armazenamento','Chegou','Saldo Total','','','','',''];
    aba.getRange(2, 1, 1, 12).setValues([h2]);

    // Formatar os dois cabeçalhos
    aba.getRange(1, 1, 2, 12)
       .setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
    aba.setFrozenRows(2);

    // Dados (a partir da linha 3)
    const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    if (rows.length > 0) {
      // Formatar coluna A como texto puro
      aba.getRange(3, 1, rows.length, 1).setNumberFormat('@');

      const dataRows = rows.map(r => [
        r.sku, r.B, r.C, r.D, r.E, r.F, r.G, r.H, r.I, r.curva, '', agora,
      ]);
      aba.getRange(3, 1, rows.length, 12).setValues(dataRows);

      // Formatar coluna I (% Cobertura) como percentual
      aba.getRange(3, 9, rows.length, 1).setNumberFormat('0.0%');
    }

    aba.autoResizeColumns(1, 12);
    SpreadsheetApp.flush();
    totalLinhas = rows.length;

  } catch (e) {
    erroMsg = e.message;
    Logger.log('Erro atualizarAnalise: ' + e.stack);
  }

  // ── Slack ─────────────────────────────────────────────────
  const duracao = ((new Date() - inicio) / 1000).toFixed(1);
  if (erroMsg) {
    _pkSlack('❌ *Análise Picking — Erro*\nErro: `' + erroMsg + '`\nDuração: ' + duracao + 's');
  } else {
    _pkSlack('✅ *Análise Picking — Atualizada*\n📊 ' + totalLinhas + ' SKUs | ⏱ ' + duracao + 's');
  }
}

// ============================================================
// CURVA AAA — classifica itens em 3 níveis ABC por volume
// Modifica cada item adicionando a propriedade .curva (ex: "AAB")
// ============================================================
function _atribuirCurvaAAA(items) {
  if (!items || items.length === 0) return;

  function classificar(grupo, nivel) {
    if (grupo.length === 0) return;
    grupo.sort((a, b) => b.qty - a.qty);
    const total = grupo.reduce((s, x) => s + x.qty, 0);
    let cum = 0;
    grupo.forEach(item => {
      cum += item.qty;
      const pct = total > 0 ? cum / total : 1;
      const letra = pct <= PK_ABC[0] ? 'A' : pct <= PK_ABC[1] ? 'B' : 'C';
      item['_l' + nivel] = letra;
    });
  }

  // Nível 1
  classificar(items, 1);

  // Nível 2 — dentro de cada grupo do nível 1
  for (const l1 of ['A', 'B', 'C']) {
    classificar(items.filter(x => x._l1 === l1), 2);
  }

  // Nível 3 — dentro de cada grupo do nível 2
  for (const l1 of ['A', 'B', 'C']) {
    for (const l2 of ['A', 'B', 'C']) {
      classificar(items.filter(x => x._l1 === l1 && x._l2 === l2), 3);
    }
  }

  // Montar string final
  items.forEach(item => {
    item.curva = (item._l1 || 'C') + (item._l2 || 'C') + (item._l3 || 'C');
  });
}
