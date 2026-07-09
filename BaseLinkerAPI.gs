// ============================================================
// BASELINKER API — helper + diagnóstico
//
// Adicione em Extensões → Apps Script → ⚙️ Propriedades do script:
//   BASELINKER_API_KEY  →  {userId}-{accountId}-{token}
// ============================================================

const BL_API_URL = 'https://api.baselinker.com/connector.php';

// ── Chamada genérica à API BaseLinker ───────────────────────
function _bl_call(method, params) {
  const key = PropertiesService.getScriptProperties().getProperty('BASELINKER_API_KEY');
  if (!key) throw new Error('BASELINKER_API_KEY não configurada nas propriedades do script.');

  const resp = UrlFetchApp.fetch(BL_API_URL, {
    method:      'post',
    contentType: 'application/x-www-form-urlencoded',
    payload:     'token=' + key +
                 '&method=' + method +
                 '&parameters=' + JSON.stringify(params || {}),
    muteHttpExceptions: true,
  });

  const data = JSON.parse(resp.getContentText());
  if (data.status !== 'SUCCESS') {
    throw new Error('[BL:' + method + '] ' + (data.error_message || JSON.stringify(data)));
  }
  return data;
}

// ============================================================
// DIAGNÓSTICO — rode isso primeiro para mapear IDs reais
// Execute: _bl_diagnosticarAPI()
// ============================================================
function _bl_diagnosticarAPI() {
  const ui  = SpreadsheetApp.getUi();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   out = '';

  // 1. Inventários (catálogo de produtos)
  try {
    const r = _bl_call('getInventories', {});
    const invs = r.inventories || [];
    out += '══ INVENTÁRIOS (' + invs.length + ') ══\n';
    invs.forEach(i => {
      out += '  id=' + i.inventory_id + ' nome="' + i.name + '"\n';
    });
  } catch(e) { out += '❌ getInventories: ' + e.message + '\n'; }

  // 2. Armazéns/Estoques (storages)
  try {
    const r = _bl_call('getStoragesList', {});
    const storages = r.storages || [];
    out += '\n══ ARMAZÉNS / STORAGES (' + storages.length + ') ══\n';
    storages.forEach(s => {
      out += '  storage_id="' + s.storage_id + '" nome="' + s.name + '"\n';
    });
  } catch(e) { out += '❌ getStoragesList: ' + e.message + '\n'; }

  // 3. Status de pedidos
  try {
    const r = _bl_call('getOrderStatusList', {});
    const statuses = r.statuses || [];
    out += '\n══ STATUS DE PEDIDOS (' + statuses.length + ') ══\n';
    statuses.forEach(s => {
      out += '  id=' + s.id + ' nome="' + s.name + '"\n';
    });
  } catch(e) { out += '❌ getOrderStatusList: ' + e.message + '\n'; }

  // 4. Localizações — estrutura (primeira página, 1 inventário)
  try {
    // Primeiro tenta sem inventory_id (usa o padrão)
    const r = _bl_call('getInventoryLocations', {});
    const locs = r.locations || [];
    out += '\n══ LOCALIZAÇÕES — amostra (' + locs.length + ' retornadas) ══\n';
    locs.slice(0, 10).forEach(l => {
      out += '  loc_id=' + l.location_id +
             ' nome="' + l.name + '"' +
             ' warehouse="' + (l.warehouse_id || l.storage_id || '-') + '"' +
             ' desc="' + (l.description || '') + '"\n';
    });
    if (locs.length > 10) out += '  ... e mais ' + (locs.length - 10) + '\n';

    // Mostra keys do primeiro objeto para entender todos os campos
    if (locs.length > 0) {
      out += '  Campos disponíveis: ' + Object.keys(locs[0]).join(', ') + '\n';
    }
  } catch(e) { out += '❌ getInventoryLocations: ' + e.message + '\n'; }

  // 5. Amostra de 1 pedido para ver estrutura de campos
  try {
    const r = _bl_call('getOrders', {
      status_id:       0,  // 0 = todos
      date_from:       Math.floor(Date.now()/1000) - 2*86400, // últimas 48h
      get_unconfirmed: false,
      page:            1,
    });
    const orders = r.orders || [];
    out += '\n══ PEDIDO — amostra (1 de ' + orders.length + ' nas últimas 48h) ══\n';
    if (orders.length > 0) {
      const o = orders[0];
      out += '  order_id=' + o.order_id + '\n';
      out += '  status_id=' + o.order_status_id + '\n';
      out += '  payment_done=' + o.payment_done + '\n';
      out += '  delivery_method="' + o.delivery_method + '"\n';
      out += '  Todos os campos: ' + Object.keys(o).join(', ') + '\n';
      // Produtos do pedido
      const prods = o.products || [];
      out += '  Produtos (' + prods.length + '):\n';
      prods.slice(0, 3).forEach(p => {
        out += '    sku="' + p.sku + '" qty=' + p.quantity +
               ' storage="' + (p.storage || '-') + '"\n';
      });
      if (prods.length > 0) {
        out += '  Campos do produto: ' + Object.keys(prods[0]).join(', ') + '\n';
      }
    }
  } catch(e) { out += '❌ getOrders: ' + e.message + '\n'; }

  // 6. Amostra de produto para ver dimensões
  try {
    // Busca 1 produto para ver campos de dimensão
    const rList = _bl_call('getInventoryProductsList', {
      inventory_id: '',  // deixa vazio para padrão
      page:         1,
    });
    const prods = rList.products ? Object.entries(rList.products) : [];
    out += '\n══ PRODUTO — amostra de dimensões ══\n';
    if (prods.length > 0) {
      const [pid] = prods[0];
      const rData = _bl_call('getInventoryProductsData', {
        inventory_id: '',
        products:     [Number(pid)],
      });
      const pData = rData.products ? rData.products[pid] : null;
      if (pData) {
        out += '  product_id=' + pid + ' sku="' + pData.sku + '"\n';
        out += '  weight=' + pData.weight +
               ' height=' + pData.height +
               ' width='  + pData.width  +
               ' length=' + pData.length + '\n';
        out += '  Campos: ' + Object.keys(pData).join(', ') + '\n';
      }
    }
  } catch(e) { out += '❌ getInventoryProductsData: ' + e.message + '\n'; }

  // 7. Amostra de estoque por armazém
  try {
    const rList = _bl_call('getInventoryProductsList', {
      inventory_id: '',
      page:         1,
    });
    const prods = rList.products ? Object.entries(rList.products) : [];
    if (prods.length > 0) {
      const [pid] = prods[0];
      const rStock = _bl_call('getInventoryProductsStock', {
        inventory_id: '',
        products:     [Number(pid)],
      });
      const pStock = rStock.products ? rStock.products[pid] : null;
      out += '\n══ ESTOQUE — estrutura de armazéns (1 produto) ══\n';
      out += '  product_id=' + pid + '\n';
      if (pStock) {
        out += '  Armazéns e estoques:\n';
        Object.entries(pStock).forEach(([key, val]) => {
          out += '    "' + key + '" = ' + JSON.stringify(val) + '\n';
        });
      }
    }
  } catch(e) { out += '❌ getInventoryProductsStock: ' + e.message + '\n'; }

  // Grava resultado em aba oculta
  let aba = ss.getSheetByName('_bl_diagnostico');
  if (!aba) { aba = ss.insertSheet('_bl_diagnostico'); aba.hideSheet(); }
  aba.clearContents();
  const chunks = [];
  for (let i = 0; i < out.length; i += 50000) chunks.push([out.substring(i, i+50000)]);
  if (chunks.length) aba.getRange(1,1,chunks.length,1).setValues(chunks);

  ui.alert(
    '✅ Diagnóstico BaseLinker\n\nResultado gravado em "_bl_diagnostico".\n\n' +
    out.substring(0, 1200) + (out.length > 1200 ? '\n...(ver aba para completo)' : '')
  );
}
