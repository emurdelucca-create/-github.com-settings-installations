// ============================================================
// VISÃO GERAL SEPARAÇÃO
// Preenche abas "Datas" e "SKUs" com pedidos dos 8 status
// configurados em VG_STATUS_ALVO, usando date_confirmed.
//
// Pré-requisito: BASELINKER_API_KEY em
//   Extensões → Apps Script → ⚙ Propriedades do script
// ============================================================

const VG_BL_URL = 'https://api.baselinker.com/connector.php';

const VG_STATUS_ALVO = [
  'NF Emitida',
  'Erro NF',
  'Erro Etiqueta',
  'Movimentação',
  '[SEP] ML Flex',
  '[SEP] ML Agência',
  '[SEP] Shopee Direta',
  '[SEP] Shopee Xpress',
];

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Visão Geral')
    .addItem('🔄 Atualizar Datas e SKUs', 'vg_atualizar')
    .addToUi();
}

// ── Chamada à API BaseLinker ──────────────────────────────────
function vg_bl(method, params) {
  const token = PropertiesService.getScriptProperties().getProperty('BASELINKER_API_KEY');
  if (!token) throw new Error(
    'Configure BASELINKER_API_KEY em Extensões → Apps Script → ⚙ Propriedades do script.'
  );
  const resp = UrlFetchApp.fetch(VG_BL_URL, {
    method: 'post',
    payload: { token, method, parameters: JSON.stringify(params || {}) },
    muteHttpExceptions: true,
  });
  const d = JSON.parse(resp.getContentText());
  if (d.status !== 'SUCCESS') {
    throw new Error('[' + method + '] ' + (d.error_message || JSON.stringify(d)));
  }
  return d;
}

// ── Função principal ──────────────────────────────────────────
function vg_atualizar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  try {
    // 1. Descobrir IDs dos status alvo
    const statusResp = vg_bl('getOrderStatusList', {});
    const nameToId   = {};
    (statusResp.statuses || []).forEach(s => { nameToId[s.name] = s.id; });

    const encontrados    = VG_STATUS_ALVO.filter(n =>  nameToId[n]);
    const naoEncontrados = VG_STATUS_ALVO.filter(n => !nameToId[n]);

    if (!encontrados.length) {
      ui.alert('❌ Nenhum dos status configurados foi encontrado no BaseLinker.');
      return;
    }
    if (naoEncontrados.length) {
      ss.toast('Aviso — status não encontrados: ' + naoEncontrados.join(', '), '⚠', 10);
    }

    ss.toast('Buscando pedidos em ' + encontrados.length + ' status…', '📊 Visão Geral', 600);

    // 2. Buscar todos os pedidos e agregar por data
    //    dateMap: 'YYYY-MM-DD' → { count: N, skus: { sku: qty } }
    //    statusOldest: statusName → menor timestamp encontrado
    const dateMap      = {};
    const statusOldest = {};

    for (const nome of encontrados) {
      const sid   = nameToId[nome];
      let idFrom  = 0;

      while (true) {
        const r     = vg_bl('getOrders', { status_id: sid, id_from: idFrom });
        const batch = r.orders || [];
        if (!batch.length) break;

        for (const pedido of batch) {
          const ts = pedido.date_confirmed;
          if (!ts) continue;

          // Rastrear a data mais antiga por status
          if (!statusOldest[nome] || ts < statusOldest[nome]) {
            statusOldest[nome] = ts;
          }

          // Agrupar por data (fuso horário da planilha)
          const ds = Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd');
          if (!dateMap[ds]) dateMap[ds] = { count: 0, skus: {} };
          dateMap[ds].count++;

          for (const prod of (pedido.products || [])) {
            const sku = String(prod.sku || '').trim();
            if (!sku) continue;
            const qty = Number(prod.quantity) || 0;
            dateMap[ds].skus[sku] = (dateMap[ds].skus[sku] || 0) + qty;
          }
        }

        if (batch.length < 100) break;
        idFrom = batch[batch.length - 1].order_id;
      }
    }

    // 3. Ordenar datas (YYYY-MM-DD ordena lexicograficamente = cronológica)
    const sortedDates = Object.keys(dateMap).sort();
    if (!sortedDates.length) {
      ui.alert('Nenhum pedido com date_confirmed encontrado nos status selecionados.');
      return;
    }

    // 4. Gravar abas
    _vg_escreverDatas(ss, tz, sortedDates, dateMap, statusOldest);
    _vg_escreverSKUs(ss, sortedDates, dateMap);

    SpreadsheetApp.flush();
    ui.alert('✅ Concluído!\n' + sortedDates.length + ' datas encontradas\n' + encontrados.length + ' status analisados');

  } catch (e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ── Escrever aba "Datas" ──────────────────────────────────────
// A2: "Data"          B2: "Qtd. Pedidos"
// A3 em diante: data  B3 em diante: contagem de pedidos
//
// D2: "Status"        E2: "Data mais antiga"
// D3:D10: nomes dos 8 status
// E3:E10: data mais antiga de cada status
function _vg_escreverDatas(ss, tz, sortedDates, dateMap, statusOldest) {
  let aba = ss.getSheetByName('Datas');
  if (!aba) aba = ss.insertSheet('Datas');
  aba.clearContents();
  aba.clearFormats();

  // Cabeçalhos — linha 2
  aba.getRange('A2').setValue('Data');
  aba.getRange('B2').setValue('Qtd. Pedidos');
  aba.getRange('D2').setValue('Status');
  aba.getRange('E2').setValue('Data mais antiga');

  // A3:B? — todas as datas + contagem, da mais antiga para a mais nova
  const linhasDatas = sortedDates.map(ds => {
    const [y, m, d] = ds.split('-').map(Number);
    return [new Date(y, m - 1, d), dateMap[ds].count];
  });
  if (linhasDatas.length) {
    aba.getRange(3, 1, linhasDatas.length, 2).setValues(linhasDatas);
    aba.getRange(3, 1, linhasDatas.length, 1).setNumberFormat('dd/mm/yyyy');
  }

  // D3:E10 — os 8 status + data mais antiga de cada
  const linhasStatus = VG_STATUS_ALVO.map(nome => {
    const ts = statusOldest[nome];
    if (!ts) return [nome, ''];
    const ds = Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd');
    const [y, m, d] = ds.split('-').map(Number);
    return [nome, new Date(y, m - 1, d)];
  });
  aba.getRange(3, 4, linhasStatus.length, 2).setValues(linhasStatus);
  aba.getRange(3, 5, linhasStatus.length, 1).setNumberFormat('dd/mm/yyyy');
}

// ── Escrever aba "SKUs" ───────────────────────────────────────
// Blocos de 2 colunas com 1 coluna de espaço entre cada bloco:
//   Data mais antiga → A/B  (cols 1-2)
//   Segunda data     → D/E  (cols 4-5)
//   Terceira data    → G/H  (cols 7-8)  ... e assim por diante
//
// Linha 1: data (cabeçalho do bloco)
// Linha 2: "SKU" | "Qtd."
// Linha 3+: SKU | quantidade total — ordenado por qty decrescente
//           (somente SKUs com pedido naquela data)
function _vg_escreverSKUs(ss, sortedDates, dateMap) {
  let aba = ss.getSheetByName('SKUs');
  if (!aba) aba = ss.insertSheet('SKUs');
  aba.clearContents();
  aba.clearFormats();

  const STEP = 3; // 2 colunas de dados + 1 coluna de gap

  sortedDates.forEach((ds, idx) => {
    const col = 1 + idx * STEP; // 1-based: 1, 4, 7, 10, …

    // Linha 1: data do bloco
    const [y, m, d] = ds.split('-').map(Number);
    aba.getRange(1, col).setValue(new Date(y, m - 1, d)).setNumberFormat('dd/mm/yyyy');

    // Linha 2: cabeçalhos
    aba.getRange(2, col    ).setValue('SKU');
    aba.getRange(2, col + 1).setValue('Qtd.');

    // Linha 3+: SKUs com pedidos nessa data, ordenados por quantidade desc
    const skuRows = Object.entries(dateMap[ds].skus)
      .sort((a, b) => b[1] - a[1])   // maior quantidade primeiro
      .map(([sku, qty]) => [sku, qty]);

    if (skuRows.length) {
      aba.getRange(3, col, skuRows.length, 2).setValues(skuRows);
      aba.getRange(3, col, skuRows.length, 1).setNumberFormat('@'); // SKU como texto
    }
  });
}
