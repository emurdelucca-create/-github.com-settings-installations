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

// Status considerados no cálculo de estoque insuficiente
const STATUS_PICKING = [
  'Novos pedidos', 'Agendados Meli',
  'NF Emitida', 'Erro NF', 'Erro Etiqueta', 'Movimentação',
];

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Movimentação')
    .addItem('🔄 Atualizar dados',                    'sincronizarMovimentacao')
    .addItem('🔄 Atualizar Estoque Insuficiente',     'sincronizarEstoqueInsuficiente')
    .addSeparator()
    .addItem('🔍 Descobrir IDs dos Canais',           'descobrirCanais')
    .addItem('🔍 Diagnóstico BaseLinker',             'diagnosticarBL')
    .addItem('🔍 Investigar Estoque Insuficiente',    'investigarEstoqueInsuficiente')
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
// INVESTIGAR ESTOQUE INSUFICIENTE
// Busca pedidos em Separação com SKU alvo e exibe todos os campos
// de produto para identificar o sinal de "Estoque insuficiente"
// ============================================================
function investigarEstoqueInsuficiente() {
  const ui     = SpreadsheetApp.getUi();
  const SKU_ALVO = '03470-S'; // SKU a investigar

  try {
    // ── 1. Descobrir todos os status "[SEP]" ──────────────────
    const statusResp = blPost('getOrderStatusList', {});
    const sepStatuses = statusResp.statuses.filter(s =>
      String(s.name).toUpperCase().includes('[SEP]')
    );
    if (sepStatuses.length === 0) {
      const nomes = statusResp.statuses.map(s => '"' + s.name + '"').join(', ');
      throw new Error('Nenhum status [SEP] encontrado. Disponíveis: ' + nomes);
    }

    // ── 2. Varrer pedidos em todos os status [SEP] ────────────
    let pedidosViu = 0;
    const achados  = []; // { pedido, prods_alvo, statusNome }

    for (const sepStatus of sepStatuses) {
      let idFrom = 0;
      while (true) {
        const r     = blPost('getOrders', { status_id: sepStatus.id, id_from: idFrom });
        const batch = r.orders || [];
        if (batch.length === 0) break;
        pedidosViu += batch.length;

        batch.forEach(pedido => {
          const prods_alvo = (pedido.products || []).filter(pr =>
            String(pr.sku || '').trim().toUpperCase() === SKU_ALVO.toUpperCase()
          );
          if (prods_alvo.length > 0) {
            achados.push({ pedido, prods_alvo, statusNome: sepStatus.name });
          }
        });

        if (achados.length >= 5 || batch.length < 100) break;
        idFrom = batch[batch.length - 1].order_id;
      }
      if (achados.length >= 5) break;
    }

    if (achados.length === 0) {
      const nomesStatus = sepStatuses.map(s => '"' + s.name + '"').join(', ');
      ui.alert('ℹ️ Nenhum pedido nos status ' + nomesStatus + ' com SKU ' + SKU_ALVO +
               ' encontrado nos ' + pedidosViu + ' pedidos analisados.');
      return;
    }

    // ── 3. Montar relatório com todos os campos relevantes ────
    const nomesEncontrados = [...new Set(achados.map(a => a.statusNome))].join(', ');
    let msg = '📋 SKU: ' + SKU_ALVO + ' | Status: ' + nomesEncontrados +
              ' | ' + achados.length + ' pedido(s) encontrado(s) em ' + pedidosViu + ' analisados\n';
    msg += '═'.repeat(60) + '\n\n';

    // Campos do produto que podem sinalizar "Estoque Insuficiente"
    const CAMPOS_PROD = [
      'product_id', 'sku', 'name', 'quantity',
      'location',        // hipótese principal: vazio = sem estoque
      'location_id',
      'bundle_id',       // se é componente de kit
      'bundle_name',
      'pick_state',      // estado de separação/picking
      'weight',
      'stock_location',
      'warehouse_id',
    ];

    achados.forEach((item, idx) => {
      const p = item.pedido;
      msg += '─── Pedido #' + p.order_id + ' [' + item.statusNome + '] ───────────────────────\n';
      msg += 'Data: ' + (p.date_confirmed ? new Date(p.date_confirmed * 1000).toLocaleString('pt-BR') : '—') + '\n';
      msg += 'Canal: ' + (CANAL_MAP[String(p.order_source_id)] || p.order_source + '/' + p.order_source_id) + '\n\n';

      item.prods_alvo.forEach((prod, pi) => {
        msg += '  Produto ' + (pi + 1) + ' (SKU: ' + prod.sku + ' | qty: ' + prod.quantity + '):\n';
        CAMPOS_PROD.forEach(campo => {
          const val = prod[campo];
          if (val !== undefined) {
            msg += '    ' + campo + ': ' + JSON.stringify(val) + '\n';
          }
        });

        // Campos extras desconhecidos — mostrar tudo que não seja padrão esperado
        const CAMPOS_COMUNS = new Set([
          'product_id','sku','name','price_brutto','price_netto','tax_rate',
          'quantity','quantity_returned','product_description','product_description_long',
          'weight','ean','warehouse_id','location','location_id',
          'bundle_id','bundle_name','pick_state','storage','storage_id',
          'auction_id','order_product_id','attributes','components',
        ]);
        const extras = Object.keys(prod).filter(k => !CAMPOS_COMUNS.has(k));
        if (extras.length > 0) {
          msg += '    [extras]: ';
          extras.forEach(k => { msg += k + '=' + JSON.stringify(prod[k]) + '; '; });
          msg += '\n';
        }
        msg += '\n';
      });

      // Mostrar TODOS os produtos do pedido para comparar location
      if ((p.products || []).length > 1) {
        msg += '  Outros produtos do pedido (location rápido):\n';
        (p.products || []).filter(pr => pr.sku !== SKU_ALVO).forEach(pr => {
          msg += '    SKU ' + pr.sku + ' → location: ' + JSON.stringify(pr.location) +
                 ' | pick_state: ' + JSON.stringify(pr.pick_state) + '\n';
        });
      }
      msg += '\n';
    });

    // Exibe em alerta (pode ser longo — GAS suporta até ~4000 chars)
    if (msg.length > 3800) {
      msg = msg.substring(0, 3800) + '\n...[truncado — veja Logs]';
      Logger.log(msg);
    }
    ui.alert('🔍 Investigação Estoque Insuficiente', msg, ui.ButtonSet.OK);

  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ============================================================
// SINCRONIZAR ESTOQUE INSUFICIENTE
// Agrega qtd de TODOS os status ativos por SKU,
// mostra SKUs onde Estoque Padrão < Qtd. Necessária,
// pinta vermelho se Saldo Total também for insuficiente
// ============================================================
function sincronizarEstoqueInsuficiente() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let aba = ss.getSheetByName('Estoque Insuficiente');
  if (!aba) aba = ss.insertSheet('Estoque Insuficiente');

  try {
    // ── 1. Todos os status do print ───────────────────────────
    const statusResp = blPost('getOrderStatusList', {});
    const statusIds  = statusResp.statuses
      .filter(s =>
        STATUS_PICKING.includes(s.name) ||
        String(s.name).toUpperCase().includes('[SEP]')
      )
      .map(s => s.id);

    if (statusIds.length === 0) {
      throw new Error('Nenhum status encontrado. Verifique STATUS_PICKING no código.');
    }

    // ── 2a. Coletar SKUs presentes em [SEP] ──────────────────
    const sepIds = statusResp.statuses
      .filter(s => String(s.name).toUpperCase().includes('[SEP]'))
      .map(s => s.id);

    const skusEmSeparacao = new Set();
    for (const sid of sepIds) {
      let idFrom = 0;
      while (true) {
        const r     = blPost('getOrders', { status_id: sid, id_from: idFrom });
        const batch = r.orders || [];
        if (batch.length === 0) break;
        batch.forEach(pedido => {
          (pedido.products || []).forEach(prod => {
            const sku = String(prod.sku || '').trim();
            if (sku) skusEmSeparacao.add(sku);
          });
        });
        if (batch.length < 100) break;
        idFrom = batch[batch.length - 1].order_id;
      }
    }

    // ── 2b. Agregar Qtd. Necessária de TODOS os status ────────
    // skuMap: sku → { totalQty, productId, oldestTs }
    const skuMap = {};

    for (const sid of statusIds) {
      let idFrom = 0;
      while (true) {
        const r     = blPost('getOrders', { status_id: sid, id_from: idFrom });
        const batch = r.orders || [];
        if (batch.length === 0) break;

        batch.forEach(pedido => {
          const ts = pedido.date_confirmed || pedido.date_add || 0;

          (pedido.products || []).forEach(prod => {
            const sku = String(prod.sku || '').trim();
            if (!sku) return;
            const qty = Number(prod.quantity) || 0;
            const pid = String(prod.product_id || '');

            if (!skuMap[sku]) {
              skuMap[sku] = { totalQty: 0, productId: pid, oldestTs: ts };
            }
            skuMap[sku].totalQty += qty;
            if (ts > 0 && (skuMap[sku].oldestTs === 0 || ts < skuMap[sku].oldestTs)) {
              skuMap[sku].oldestTs = ts;
            }
          });
        });

        if (batch.length < 100) break;
        idFrom = batch[batch.length - 1].order_id;
      }
    }

    // ── 3. Buscar estoque ─────────────────────────────────────
    const allPids  = [...new Set(Object.values(skuMap).map(d => d.productId).filter(Boolean))];
    const stockMap = {};
    const LOTE     = 1000;
    for (let i = 0; i < allPids.length; i += LOTE) {
      const res = blPost('getInventoryProductsData', {
        inventory_id: INVENTORY_ID,
        products: allPids.slice(i, i + LOTE),
      });
      Object.entries(res.products || {}).forEach(([pid, d]) => {
        const s = d.stock || {};
        stockMap[pid] = {
          pad: Number(s[WH_PADRAO]        || 0),
          arm: Number(s[WH_ARMAZENAMENTO] || 0),
          chg: Number(s[WH_CHEGOU]        || 0),
        };
      });
    }

    // ── 4. Filtrar (Padrão < Necessário) e montar linhas ──────
    const linhas    = [];
    const negativos = []; // índices (base 0) com saldo total negativo

    Object.entries(skuMap).forEach(([sku, data]) => {
      if (!skusEmSeparacao.has(sku)) return;  // só aparece se estiver em [SEP]
      const est = stockMap[data.productId] || { pad: 0, arm: 0, chg: 0 };
      if (est.pad >= data.totalQty) return; // picking suficiente, ignora

      const saldoTotal = est.pad + est.arm + est.chg;
      const saldoGeral = saldoTotal - data.totalQty;
      const dataAntiga = data.oldestTs > 0 ? new Date(data.oldestTs * 1000) : '';

      if (saldoGeral < 0) negativos.push(linhas.length);

      linhas.push([
        sku,           // A: SKU
        data.totalQty, // B: Qtd. Necessária
        est.pad,       // C: Estoque Padrão
        est.arm,       // D: Estoque Armazenamento
        est.chg,       // E: Estoque Chegou
        saldoTotal,    // F: Saldo Total
        saldoGeral,    // G: Saldo Total − Qtd. Necessária
        dataAntiga,    // H: Data da venda mais antiga
      ]);
    });

    // ── 5. Escrever na aba ────────────────────────────────────
    const CABECALHO = [[
      'SKU', 'Qtd. Necessária', 'Estoque Padrão', 'Estoque Armazenamento',
      'Estoque Chegou', 'Saldo Total', 'Saldo Total - Qtd. Necessária', 'Data da venda mais antiga',
    ]];

    aba.clearContents();
    aba.getRange(1, 1, 1, 8).setValues(CABECALHO);

    if (linhas.length > 0) {
      const dataRange = aba.getRange(2, 1, linhas.length, 8);
      dataRange.setValues(linhas);
      dataRange.setBackground(null);
      aba.getRange(2, 1, linhas.length, 1).setNumberFormat('@');
      aba.getRange(2, 8, linhas.length, 1).setNumberFormat('dd/mm/yyyy');

      negativos.forEach(idx => {
        aba.getRange(idx + 2, 1, 1, 8).setBackground('#FF9999');
      });
    }

    SpreadsheetApp.flush();
    ui.alert('✅ ' + linhas.length + ' SKU(s) com falta no picking. ' +
             negativos.length + ' com falta no estoque geral (vermelho).');

  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ============================================================
// SINCRONIZAR MOVIMENTAÇÃO
// Cria abas "Movimentação C/ Estoque" e "Movimentação S/ Estoque"
// Uma linha por SKU | colunas agrupadas por método de envio
// Ordenado pela data do pedido mais antigo (crescente)
// ============================================================
function sincronizarMovimentacao() {
  const ui      = SpreadsheetApp.getUi();
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const normStr = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();

  try {
    // ── 1. Status "Movimentação" ──────────────────────────────
    const statusResp = blPost('getOrderStatusList', {});
    const statusObj  = statusResp.statuses.find(s => normStr(s.name).includes('movimenta'));
    if (!statusObj) throw new Error('Status "Movimentação" não encontrado.');

    // ── 2. Buscar todos os pedidos ────────────────────────────
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

    // ── 3. Agregar por SKU ────────────────────────────────────
    // skuMap: sku → { productId, global:{qty,oldestTs,canals}, methods:{metodo:{qty,oldestTs,canals}} }
    const skuMap = {};

    pedidos.forEach(pedido => {
      const sourceId = String(pedido.order_source_id || '');
      const canal    = CANAL_MAP[sourceId] || (pedido.order_source + (sourceId ? '/' + sourceId : ''));
      const metodo   = pedido.delivery_method || 'Sem método';
      const ts       = pedido.date_confirmed || pedido.date_add || 0;

      (pedido.products || []).forEach(prod => {
        const sku = String(prod.sku || '').trim();
        if (!sku) return;
        const qty = Number(prod.quantity) || 0;
        const pid = String(prod.product_id || '');

        if (!skuMap[sku]) {
          skuMap[sku] = { productId: pid, global: { qty: 0, oldestTs: 0, canals: {} }, methods: {} };
        }
        const d = skuMap[sku];

        d.global.qty += qty;
        if (ts > 0 && (d.global.oldestTs === 0 || ts < d.global.oldestTs)) d.global.oldestTs = ts;
        d.global.canals[canal] = (d.global.canals[canal] || 0) + qty;

        if (!d.methods[metodo]) d.methods[metodo] = { qty: 0, oldestTs: 0, canals: {} };
        d.methods[metodo].qty += qty;
        if (ts > 0 && (d.methods[metodo].oldestTs === 0 || ts < d.methods[metodo].oldestTs)) {
          d.methods[metodo].oldestTs = ts;
        }
        d.methods[metodo].canals[canal] = (d.methods[metodo].canals[canal] || 0) + qty;
      });
    });

    // ── 4. Buscar estoque ─────────────────────────────────────
    const allPids  = [...new Set(Object.values(skuMap).map(d => d.productId).filter(Boolean))];
    const stockMap = {};
    const LOTE     = 1000;
    for (let i = 0; i < allPids.length; i += LOTE) {
      const res = blPost('getInventoryProductsData', {
        inventory_id: INVENTORY_ID,
        products: allPids.slice(i, i + LOTE),
      });
      Object.entries(res.products || {}).forEach(([pid, d]) => {
        const s = d.stock || {};
        stockMap[pid] = {
          pad: Number(s[WH_PADRAO]        || 0),
          arm: Number(s[WH_ARMAZENAMENTO] || 0),
          chg: Number(s[WH_CHEGOU]        || 0),
        };
      });
    }

    // ── 5. Filtrar (Padrão < 0) e separar ────────────────────
    const allMethods = [...new Set(
      Object.values(skuMap).flatMap(d => Object.keys(d.methods))
    )].sort();

    const comEstoque = []; // Padrão < 0 e Saldo Total >= 0
    const semEstoque = []; // Padrão < 0 e Saldo Total <  0

    Object.entries(skuMap).forEach(([sku, data]) => {
      const est = stockMap[data.productId] || { pad: 0, arm: 0, chg: 0 };
      if (est.pad >= 0) return; // picking OK, ignora
      const saldoTotal = est.pad + est.arm + est.chg;
      const row = { sku, est, saldoTotal, data };
      if (saldoTotal >= 0) comEstoque.push(row);
      else                 semEstoque.push(row);
    });

    // Ordenar pelo pedido mais antigo (crescente)
    const sortFn = (a, b) => (a.data.global.oldestTs || Infinity) - (b.data.global.oldestTs || Infinity);
    comEstoque.sort(sortFn);
    semEstoque.sort(sortFn);

    // ── 6. Escrever nas abas ──────────────────────────────────
    _escreverAbaMovimentacao(ss, 'Movimentação C/ Estoque', comEstoque, allMethods);
    _escreverAbaMovimentacao(ss, 'Movimentação S/ Estoque', semEstoque, allMethods);

    // Remover aba original "Movimentação" (foi renomeada pelo usuário)
    ['Movimentação', 'Movimentacao'].forEach(nome => {
      const abaOld = ss.getSheetByName(nome);
      if (abaOld) ss.deleteSheet(abaOld);
    });

    SpreadsheetApp.flush();
    ui.alert('✅ C/ Estoque: ' + comEstoque.length + ' SKUs | S/ Estoque: ' + semEstoque.length + ' SKUs');

  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ============================================================
// HELPER — escreve uma aba de movimentação com blocos por método
// ============================================================
function _escreverAbaMovimentacao(ss, nomeAba, rows, allMethods) {
  let aba = ss.getSheetByName(nomeAba);
  if (aba) {
    aba.getRange(1, 1, aba.getMaxRows(), aba.getMaxColumns()).breakApart();
    aba.clearContents();
    aba.clearFormats();
  } else {
    aba = ss.insertSheet(nomeAba);
  }

  if (rows.length === 0) return;

  const COLS     = 7;
  const SEP      = 1;
  const STEP     = COLS + SEP; // 8
  const DATA_ROW = 3;

  const SUB = ['SKU', 'Qtd.', 'Data mais antiga', 'Est. Padrão', 'Est. Armazenamento', 'Est. Chegou', 'Saldo Total'];

  const totalCols = STEP * (1 + allMethods.length);

  // Linha 1: nomes dos grupos
  const row1 = new Array(totalCols).fill('');
  row1[0] = 'Geral (Todos os Métodos)';
  allMethods.forEach((m, i) => { row1[STEP * (i + 1)] = m; });
  aba.getRange(1, 1, 1, totalCols).setValues([row1]);

  // Mesclar cada grupo na linha 1 (exatamente 7 colunas)
  aba.getRange(1, 1, 1, COLS).merge();
  allMethods.forEach((_, i) => {
    aba.getRange(1, STEP * (i + 1) + 1, 1, COLS).merge();
  });

  // Linha 2: sub-cabeçalhos
  const row2 = new Array(totalCols).fill('');
  for (let b = 0; b <= allMethods.length; b++) {
    SUB.forEach((h, j) => { row2[b * STEP + j] = h; });
  }
  aba.getRange(2, 1, 1, totalCols).setValues([row2]);

  // Bloco Geral — escrito independentemente
  const gData = rows.map(({ sku, est, saldoTotal, data }) => {
    const d = data.global.oldestTs > 0 ? new Date(data.global.oldestTs * 1000) : '';
    return [sku, data.global.qty, d, est.pad, est.arm, est.chg, saldoTotal];
  });
  aba.getRange(DATA_ROW, 1, gData.length, COLS).setValues(gData);
  aba.getRange(DATA_ROW, 1, gData.length, 1).setNumberFormat('@');
  aba.getRange(DATA_ROW, 3, gData.length, 1).setNumberFormat('dd/mm/yyyy');

  // Cada bloco de método — independente, sem linhas vazias
  allMethods.forEach((m, i) => {
    const startCol = STEP * (i + 1) + 1; // 1-based

    const mRows = rows
      .filter(r => r.data.methods[m])
      .sort((a, b) => (a.data.methods[m].oldestTs || Infinity) - (b.data.methods[m].oldestTs || Infinity));

    if (mRows.length === 0) return;

    const mData = mRows.map(({ sku, est, saldoTotal, data }) => {
      const md = data.methods[m];
      const d  = md.oldestTs > 0 ? new Date(md.oldestTs * 1000) : '';
      return [sku, md.qty, d, est.pad, est.arm, est.chg, saldoTotal];
    });

    aba.getRange(DATA_ROW, startCol, mData.length, COLS).setValues(mData);
    aba.getRange(DATA_ROW, startCol,     mData.length, 1).setNumberFormat('@');
    aba.getRange(DATA_ROW, startCol + 2, mData.length, 1).setNumberFormat('dd/mm/yyyy');
  });
}

