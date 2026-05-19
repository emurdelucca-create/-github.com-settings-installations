// ============================================================
// PLANEJAMENTO PICKING — BaseLinker → Google Sheets
// Pedidos dos últimos 7 dias (D-1 até D-7)
// Excluindo: Cancelados, ME2 - Mercado Envios Full, Fulfilled by Shopee
// ============================================================

const PK_BL_TOKEN   = '8004176-8026704-5DUYJBOPVCCE3W6VATUJEEAJY8P7Z4YS2IHQCWEU8YAM2RR74VA1N2RE95PVYWGZ';
const PK_BL_URL     = 'https://api.baselinker.com/connector.php';
const PK_SLACK_HOOK = 'COLE_AQUI_O_WEBHOOK_DO_SLACK'; // preencher após criar o Incoming Webhook

// Métodos de envio a EXCLUIR (Full)
const PK_EXCLUIR_METODOS = ['ME2 - Mercado Envios Full', 'Fulfilled by Shopee'];

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
    .addItem('🔄 Atualizar agora', 'atualizarPicking')
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
  if (!PK_SLACK_HOOK || PK_SLACK_HOOK === 'COLE_AQUI_O_WEBHOOK_DO_SLACK') {
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
