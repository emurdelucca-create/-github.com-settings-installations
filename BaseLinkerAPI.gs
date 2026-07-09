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

  // 4. Warehouses do inventário (Padrão / Armazenamento / Chegou)
  try {
    const r = _bl_call('getInventoryWarehousesList', { inventory_id: 39947 });
    const whs = r.warehouses || [];
    out += '\n══ WAREHOUSES DO INVENTÁRIO (' + whs.length + ') ══\n';
    whs.forEach(w => {
      out += '  id=' + (w.warehouse_id || w.id) +
             ' nome="' + w.name + '"' +
             ' desc="' + (w.description || '') + '"\n';
    });
    if (whs.length > 0) out += '  Campos: ' + Object.keys(whs[0]).join(', ') + '\n';
  } catch(e) { out += '❌ getInventoryWarehousesList: ' + e.message + '\n'; }

  // 4b. Localizações — tenta métodos alternativos
  const locMethods = [
    ['getInventoryLocationsList',    { inventory_id: 39947 }],
    ['getInventoryLocationsGroups',  { inventory_id: 39947 }],
    ['getInventoryProductLocations', { inventory_id: 39947 }],
  ];
  for (const [m, p] of locMethods) {
    try {
      const r = _bl_call(m, p);
      out += '\n✅ ' + m + ' FUNCIONOU:\n' + JSON.stringify(r).substring(0, 500) + '\n';
      break;
    } catch(e) {
      out += '❌ ' + m + ': ' + e.message + '\n';
    }
  }

  // 5. Pedidos — testa sem data, depois com status específico
  const orderTries = [
    { status_id: 268797, get_unconfirmed: false },         // NF Emitida
    { status_id: 268806, get_unconfirmed: false },         // [SEP] Shopee Direta
    { status_id: 0,      get_unconfirmed: false, page: 1 },// todos confirmados
    { status_id: 0,      get_unconfirmed: true,  page: 1 },// todos incluindo não confirmados
  ];
  for (const params of orderTries) {
    try {
      const r = _bl_call('getOrders', params);
      const orders = r.orders || [];
      out += '\n══ PEDIDO params=' + JSON.stringify(params) + ' → ' + orders.length + ' ══\n';
      if (orders.length > 0) {
        const o = orders[0];
        out += '  order_id='       + o.order_id          + '\n';
        out += '  status_id='      + o.order_status_id   + '\n';
        out += '  payment_done='   + o.payment_done      + '\n';
        out += '  delivery_method="' + o.delivery_method + '"\n';
        out += '  date_add='       + o.date_add          + '\n';
        out += '  date_confirmed=' + (o.date_confirmed || '-') + '\n';
        out += '  Campos pedido: ' + Object.keys(o).join(', ') + '\n';
        const prods = o.products || [];
        if (prods.length > 0) {
          out += '  Produto[0]: sku="' + prods[0].sku + '" qty=' + prods[0].quantity + '\n';
          out += '  Campos produto: ' + Object.keys(prods[0]).join(', ') + '\n';
        }
        break;
      }
    } catch(e) { out += '❌ getOrders ' + JSON.stringify(params) + ': ' + e.message + '\n'; }
  }

  // 6. Produto — dimensões + STOCK + LOCATIONS (tudo dentro de getInventoryProductsData)
  try {
    const rList = _bl_call('getInventoryProductsList', { inventory_id: 39947, page: 1 });
    const allProds = rList.products ? Object.entries(rList.products) : [];
    const pids = allProds.slice(0, 3).map(([id]) => Number(id));
    const rData = _bl_call('getInventoryProductsData', { inventory_id: 39947, products: pids });
    out += '\n══ PRODUTO — stock + locations ══\n';
    pids.forEach(pid => {
      const p = rData.products ? rData.products[String(pid)] : null;
      if (!p) return;
      out += '\n  product_id=' + pid + ' sku="' + p.sku + '"\n';
      out += '  Medidas: weight=' + p.weight + ' h=' + p.height + ' w=' + p.width + ' l=' + p.length + '\n';
      out += '  STOCK: '     + JSON.stringify(p.stock     || {}) + '\n';
      out += '  LOCATIONS: ' + JSON.stringify(p.locations || {}) + '\n';
    });
  } catch(e) { out += '❌ getInventoryProductsData (stock+loc): ' + e.message + '\n'; }

  // 7. getInventoryProductsStock — parâmetro correto
  try {
    const rStock = _bl_call('getInventoryProductsStock', {
      inventory_id: 39947,
      products:     [83984413],
    });
    out += '\n══ getInventoryProductsStock (produto 83984413) ══\n';
    const pStock = rStock.products ? rStock.products['83984413'] : rStock;
    out += JSON.stringify(pStock).substring(0, 800) + '\n';
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
