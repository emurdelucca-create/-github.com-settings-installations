// ============================================================
// INSPEÇÃO DA API DE PEDIDOS SHOPEE
// Mostra o JSON bruto de get_order_list + get_order_detail
// para você decidir quais campos guardar na base de dados.
//
// Como usar:
//   Apps Script → executar _sr_inspecionarPedidosAPI()
//   OU adicionar ao menu e chamar pela planilha.
// ============================================================

function _sr_inspecionarPedidosAPI() {
  const ui = SpreadsheetApp.getUi();

  const temToken = !!PropertiesService.getScriptProperties().getProperty('SHOPEE_ACCESS_TOKEN');
  if (!temToken) { ui.alert('❌ Token Shopee não configurado.'); return; }

  const agora   = Math.floor(Date.now() / 1000);
  const de      = agora - 3 * 86400; // últimos 3 dias
  const ate     = agora;

  // ── 1. get_order_list ──────────────────────────────────────
  // Retorna a lista de pedidos no período.
  // Campos retornados por padrão (sem response_optional_fields):
  //   order_sn, region, currency, cod, total_amount, order_status,
  //   shipping_carrier, payment_method, estimated_shipping_fee,
  //   message_to_seller, create_time, update_time, days_to_ship,
  //   ship_by_date, buyer_user_id, buyer_username, invoice_data
  let listResp;
  try {
    listResp = _shopeeGet('/api/v2/order/get_order_list', {
      time_range_field: 'create_time',
      time_from:        de,
      time_to:          ate,
      page_size:        3,          // só 3 para não poluir o diagnóstico
      response_fields:  'order_sn,order_status,create_time,update_time,total_amount',
    });
  } catch (e) {
    ui.alert('❌ Erro get_order_list: ' + e.message);
    return;
  }

  const pedidos = listResp.order_list || [];
  const sns     = pedidos.map(p => p.order_sn).filter(Boolean);

  let saida = '══ get_order_list (3 pedidos) ══\n';
  saida += JSON.stringify(listResp, null, 2).substring(0, 2000);

  if (!sns.length) {
    ui.alert(saida + '\n\n⚠️ Nenhum pedido nos últimos 3 dias.');
    _sr_gravarInspecao(saida);
    return;
  }

  // ── 2. get_order_detail ────────────────────────────────────
  // response_optional_fields disponíveis (adicione/remova conforme precisar):
  //   buyer_user_id      → ID do comprador
  //   buyer_username     → username do comprador
  //   pay_time           → timestamp do pagamento
  //   total_amount       → valor total do pedido
  //   actual_shipping_fee → frete real cobrado
  //   payment_method     → forma de pagamento
  //   shipping_carrier   → transportadora
  //   note               → observação do comprador
  //   recipient_address  → endereço de entrega (objeto)
  //   package_list       → lista de pacotes com tracking
  //   item_list          → ← PRINCIPAL: itens do pedido (veja abaixo)
  //
  // Dentro de item_list cada item tem:
  //   item_id            → ID do produto pai
  //   item_name          → nome do produto
  //   item_sku           → SKU de referência (pai)
  //   model_id           → ID da variação (Variante ID)
  //   model_name         → nome da variação
  //   model_sku          → SKU da variação (texto definido pelo vendedor)
  //   model_quantity_purchased → quantidade vendida
  //   model_original_price     → preço original
  //   model_discounted_price   → preço pago (com desconto/promo)
  //   wholesale           → se foi venda atacado (bool)
  //   is_main_item        → se é o item principal do bundle
  //   add_on_deal         → se é brinde/add-on
  let detResp;
  try {
    detResp = _shopeeGet('/api/v2/order/get_order_detail', {
      order_sn_list:           sns.slice(0, 2).join(','),
      response_optional_fields: [
        'buyer_user_id',
        'buyer_username',
        'pay_time',
        'total_amount',
        'actual_shipping_fee',
        'payment_method',
        'shipping_carrier',
        'note',
        'package_list',
        'item_list',
      ].join(','),
    });
  } catch (e) {
    saida += '\n\n❌ Erro get_order_detail: ' + e.message;
    _sr_gravarInspecao(saida);
    ui.alert(saida.substring(0, 1500));
    return;
  }

  saida += '\n\n══ get_order_detail (até 2 pedidos) ══\n';
  saida += JSON.stringify(detResp, null, 2);

  _sr_gravarInspecao(saida);

  // Mostra resumo no alert (JSON completo está na aba _sr_inspecao_pedidos)
  const resumo = _sr_resumirPedidos(detResp.order_list || []);
  ui.alert(
    '✅ JSON gravado na aba "_sr_inspecao_pedidos"\n\n' +
    'Campos encontrados por pedido:\n' + resumo +
    '\n\nAbra a aba para ver o JSON completo.'
  );
}

// ── Grava o JSON bruto numa aba oculta para leitura fácil ──
function _sr_gravarInspecao(texto) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   aba = ss.getSheetByName('_sr_inspecao_pedidos');
  if (!aba) {
    aba = ss.insertSheet('_sr_inspecao_pedidos');
    aba.hideSheet();
  }
  aba.clearContents();
  // Quebra em linhas de 50 000 chars (limite de célula do Sheets)
  const chunks = [];
  for (let i = 0; i < texto.length; i += 50000) {
    chunks.push([texto.substring(i, i + 50000)]);
  }
  aba.getRange(1, 1, chunks.length, 1).setValues(chunks);
}

// ── Resumo dos campos/valores de cada pedido para o alert ──
function _sr_resumirPedidos(orders) {
  if (!orders.length) return '(nenhum pedido)';
  return orders.map(o => {
    const itens = (o.item_list || []).map(it =>
      '    • ' + it.item_name + '\n' +
      '      model_sku="' + it.model_sku + '" model_id=' + it.model_id + '\n' +
      '      qty=' + it.model_quantity_purchased +
      ' preço=' + it.model_discounted_price +
      ' orig=' + it.model_original_price
    ).join('\n');
    return (
      '📦 ' + o.order_sn + ' [' + o.order_status + ']\n' +
      '  create=' + new Date(o.create_time * 1000).toLocaleString('pt-BR') + '\n' +
      '  pay_time=' + (o.pay_time ? new Date(o.pay_time * 1000).toLocaleString('pt-BR') : '-') + '\n' +
      '  total=R$' + (o.total_amount || '?') + '\n' +
      '  pagamento=' + (o.payment_method || '?') + '\n' +
      '  frete_real=R$' + (o.actual_shipping_fee || '?') + '\n' +
      '  itens (' + (o.item_list || []).length + '):\n' + itens
    );
  }).join('\n\n');
}
