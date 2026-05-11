// ============================================================
// MOVIMENTAÇÃO 2026 — BaseLinker → Google Sheets
// Puxa pedidos com status "Movimentação" e estoques por armazém
// ============================================================

const BL_TOKEN        = '8004176-8026704-5DUYJBOPVCCE3W6VATUJEEAJY8P7Z4YS2IHQCWEU8YAM2RR74VA1N2RE95PVYWGZ';
const BL_URL          = 'https://api.baselinker.com/connector.php';
const INVENTORY_ID    = '39947';
const WH_PADRAO       = 'bl_44285';
const WH_ARMAZENAMENTO = 'bl_50394';
const WH_CHEGOU       = 'bl_51442';

// ============================================================
// MAPEAMENTO order_source_id → nome do canal
// Preencha após rodar "Descobrir IDs dos Canais"
// ============================================================
const CANAL_MAP = {
  '24365': 'ML Humble',
  '24367': 'ML Najumi',
  '24370': 'ML Edmotos',
  '24401': 'ML Moto Vibe',
  '24411': 'ML Sky',
  '24366': 'Shopee Humble',
  '24368': 'Shopee Najumi',
  '24369': 'Shopee Sky',
};

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Movimentação')
    .addItem('🔄 Atualizar dados',          'sincronizarMovimentacao')
    .addSeparator()
    .addItem('🔍 Descobrir IDs dos Canais', 'descobrirCanais')
    .addItem('🔍 Diagnóstico BaseLinker',   'diagnosticarBL')
    .addToUi();
}

// ============================================================
// CHAMADA À API DO BASELINKER
// ============================================================
function blPost(method, params) {
  const resp = UrlFetchApp.fetch(BL_URL, {
    method: 'post',
    payload: {
      token: BL_TOKEN,
      method: method,
      parameters: JSON.stringify(params || {}),
    },
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (data.status !== 'SUCCESS') {
    throw new Error('[' + method + '] ' + (data.error_message || JSON.stringify(data)));
  }
  return data;
}

// ============================================================
// DESCOBRIR IDs DOS CANAIS
// Varre os últimos pedidos e lista todos os order_source_id
// únicos encontrados para você preencher o CANAL_MAP acima
// ============================================================
function descobrirCanais() {
  const ui = SpreadsheetApp.getUi();
  try {
    const fontes = {}; // source_id → { source, exemplo_pedido }
    let idFrom = 0;
    let totalPedidos = 0;

    // Varre até 500 pedidos recentes para coletar todos os source_ids
    while (totalPedidos < 500) {
      const r     = blPost('getOrders', { id_from: idFrom });
      const batch = r.orders || [];
      if (batch.length === 0) break;

      batch.forEach(p => {
        const key = String(p.order_source_id || 'sem_id');
        if (!fontes[key]) {
          fontes[key] = {
            source:   p.order_source || '',
            exemplo:  p.order_id,
          };
        }
      });

      totalPedidos += batch.length;
      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }

    let msg = '📋 IDs de canais encontrados em ' + totalPedidos + ' pedidos:\n\n';
    msg += 'order_source_id | order_source | ex. pedido\n';
    msg += '─────────────────────────────────────────\n';
    Object.keys(fontes).sort().forEach(id => {
      const f = fontes[id];
      msg += id + ' | ' + f.source + ' | #' + f.exemplo + '\n';
    });
    msg += '\nPreencha o CANAL_MAP no código com esses IDs.';

    ui.alert('🔍 Canais encontrados', msg, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ============================================================
// DIAGNÓSTICO — mostra campos do 1º pedido encontrado
// Útil para descobrir qual campo tem o nome do canal
// ============================================================
function diagnosticarBL() {
  const ui = SpreadsheetApp.getUi();
  try {
    // Pega qualquer pedido recente para inspecionar campos
    const r = blPost('getOrders', { id_from: 0 });
    const pedido = (r.orders || [])[0];
    if (!pedido) { ui.alert('Nenhum pedido encontrado.'); return; }

    const campos = [
      'order_id', 'order_source', 'order_source_id', 'order_source_user',
      'order_source_user_id', 'delivery_method', 'date_confirmed',
      'user_login', 'extra_field_1', 'extra_field_2', 'order_status_id',
    ];
    let msg = '📋 Campos do 1º pedido encontrado:\n\n';
    campos.forEach(c => { msg += c + ': ' + JSON.stringify(pedido[c]) + '\n'; });

    const prod = (pedido.products || [])[0];
    if (prod) {
      msg += '\n📦 1º produto do pedido:\n';
      ['product_id', 'sku', 'name', 'quantity'].forEach(c => {
        msg += c + ': ' + JSON.stringify(prod[c]) + '\n';
      });
    }
    ui.alert('🔍 Diagnóstico', msg, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ============================================================
// SINCRONIZAR MOVIMENTAÇÃO
// ============================================================
function sincronizarMovimentacao() {
  const ui  = SpreadsheetApp.getUi();
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const normStr = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();

  try {
    // ── 1. Buscar ID do status "Movimentação" ─────────────────
    const statusResp = blPost('getOrderStatusList', {});
    const statusObj  = statusResp.statuses.find(s => normStr(s.name).includes('movimenta'));
    if (!statusObj) {
      throw new Error('Status "Movimentação" não encontrado.\nDisponíveis: ' +
        statusResp.statuses.map(s => '"' + s.name + '"').join(', '));
    }

    // ── 2. Buscar pedidos (com paginação) ─────────────────────
    const pedidos = [];
    let idFrom = 0;
    while (true) {
      const r     = blPost('getOrders', { status_id: statusObj.id, id_from: idFrom });
      const batch = r.orders || [];
      pedidos.push(...batch);
      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }
    if (pedidos.length === 0) {
      ui.alert('Nenhum pedido com status "' + statusObj.name + '" encontrado.');
      return;
    }

    // ── 3. Coletar product_ids únicos dos pedidos ─────────────
    const prodIds = [...new Set(
      pedidos.flatMap(p => (p.products || [])
        .map(pr => String(pr.product_id || ''))
        .filter(Boolean))
    )];

    // ── 4. Buscar estoque via getInventoryProductsData ────────
    // Usa os IDs de armazém e inventário fixos (descobertos no código existente)
    // stockMap: product_id → { pad, arm, chg }
    const stockMap = {};
    const LOTE = 1000;
    for (let i = 0; i < prodIds.length; i += LOTE) {
      const lote = prodIds.slice(i, i + LOTE);
      const res  = blPost('getInventoryProductsData', {
        inventory_id: INVENTORY_ID,
        products: lote,
      });
      Object.entries(res.products || {}).forEach(([pid, d]) => {
        const stock = d.stock || {};
        stockMap[pid] = {
          pad: Number(stock[WH_PADRAO]        || 0),
          arm: Number(stock[WH_ARMAZENAMENTO] || 0),
          chg: Number(stock[WH_CHEGOU]        || 0),
        };
      });
    }

    // ── 5. Montar linhas ──────────────────────────────────────
    const linhas = [];
    pedidos.forEach(pedido => {
      // Canal de venda: usa CANAL_MAP se preenchido, senão mostra source + id bruto
      const sourceId = String(pedido.order_source_id || '');
      const canal    = CANAL_MAP[sourceId]
                    || (pedido.order_source + (sourceId ? ' / ' + sourceId : ''));
      const metodo = pedido.delivery_method   || '';
      const data   = pedido.date_confirmed ? new Date(pedido.date_confirmed * 1000) : '';

      (pedido.products || []).forEach(prod => {
        const sku = prod.sku      || '';
        const qty = Number(prod.quantity) || 0;
        const pid = String(prod.product_id || '');
        const est = stockMap[pid] || { pad: 0, arm: 0, chg: 0 };

        linhas.push([
          statusObj.name,              // A: Status BaseLinker
          canal,                       // B: Canal de venda
          metodo,                      // C: Método de envio
          sku,                         // D: SKU
          qty,                         // E: Quantidade
          data,                        // F: Data do pedido
          est.pad,                     // G: Estoque Padrão
          est.arm,                     // H: Estoque Armazenamento
          est.chg,                     // I: Estoque Chegou
          est.pad + est.arm + est.chg, // J: Total
        ]);
      });
    });

    // ── 6. Escrever na planilha ───────────────────────────────
    const PRIMA   = 3;
    const lastRow = aba.getLastRow();
    if (lastRow >= PRIMA) {
      aba.getRange(PRIMA, 1, lastRow - PRIMA + 1, 10).clearContent();
    }
    if (linhas.length > 0) {
      // SKU (col D = 4) como texto antes de escrever os dados
      aba.getRange(PRIMA, 4, linhas.length, 1).setNumberFormat('@');
      aba.getRange(PRIMA, 1, linhas.length, 10).setValues(linhas);
      aba.getRange(PRIMA, 6, linhas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    }

    SpreadsheetApp.flush();
    ui.alert('✅ ' + linhas.length + ' linhas importadas de ' + pedidos.length + ' pedidos.');

  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}
