// ============================================================
// VISÃO GERAL SEPARAÇÃO
// Preenche abas "Datas" e "SKUs" com pedidos dos 8 status
// configurados em VG_STATUS_ALVO, usando date_confirmed.
//
// Pré-requisito: BASELINKER_API_KEY em
//   Extensões → Apps Script → ⚙ Propriedades do script
// ============================================================

const VG_BL_URL = 'https://api.baselinker.com/connector.php';

// Fonte de estoque (mesma planilha usada no MovEstoqueDashboard)
const VG_BL_SS_ID   = '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik';
const VG_BL_ABA_EST = 'estoque';
const VG_BL_COL_SKU = 1, VG_BL_COL_PAD = 4, VG_BL_COL_ARM = 5, VG_BL_COL_CHG = 6;

const VG_ABA_RAW = '_vg_raw_'; // aba oculta: Status | Data | OrderID | SKU | Qtd

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

// Paleta dark dashboard
const VG_T = {
  bgBase:    '#0d1117',
  bgCard:    '#161b22',
  bgHeader:  '#1c2333',
  bgAlt:     '#131920',
  accent:    '#1f6feb',
  accentSub: '#388bfd',
  green:     '#3fb950',
  amber:     '#d29922',
  red:       '#f85149',
  cyan:      '#39d353',
  textMain:  '#e6edf3',
  textSub:   '#8b949e',
  textDim:   '#484f58',
  border:    '#30363d',
};

// ── Web App ───────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('VisaoGeralDashboard')
    .setTitle('Visão Geral — Separação')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Lê o mapa de SKU -> saldo total (Padrão+Armazenamento+Chegou) da
// planilha BaseLinker "estoque". SKU não encontrado = saldo 0 (sem estoque).
function _vg_getEstoqueMap() {
  const map = new Map();
  const ss  = SpreadsheetApp.openById(VG_BL_SS_ID);
  const aba = ss.getSheetByName(VG_BL_ABA_EST);
  if (!aba) return map;
  const last = aba.getLastRow();
  if (last < 2) return map;
  const ncols = Math.max(VG_BL_COL_SKU, VG_BL_COL_PAD, VG_BL_COL_ARM, VG_BL_COL_CHG) + 1;
  aba.getRange(2, 1, last - 1, ncols).getValues().forEach(r => {
    const sku = String(r[VG_BL_COL_SKU] || '').trim();
    if (!sku) return;
    const saldo = (Number(r[VG_BL_COL_PAD]) || 0) + (Number(r[VG_BL_COL_ARM]) || 0) + (Number(r[VG_BL_COL_CHG]) || 0);
    map.set(sku, saldo);
  });
  return map;
}

function _vg_fmtDataYMD(ds, tz) {
  const [y, m, d] = ds.split('-').map(Number);
  return Utilities.formatDate(new Date(y, m - 1, d), tz, 'dd/MM/yyyy');
}

// Agrupa pedidos em datas/status/SKUs conforme a regra de estoque:
//   Com Estoque:  unidade = pedido inteiro. Só entra se TODOS os itens
//                 tiverem estoque. Se qualificar, todos os itens do
//                 pedido aparecem.
//   Sem Estoque:  unidade = item. Entra qualquer item sem estoque, de
//                 qualquer pedido (mesmo que o pedido seja misto) — só
//                 os itens sem estoque aparecem, os outros itens desse
//                 mesmo pedido (com estoque) não aparecem aqui.
// Em ambos os casos, "Pedidos por Data" conta pedidos únicos que têm
// pelo menos 1 item qualificado naquela data.
function _vg_montarSecao(pedidosMap, temEstoque, tz, isComEstoque) {
  const datasSet         = {}; // data(yyyy-MM-dd) -> Set(orderId)
  const statusOldestDate = {}; // nome -> data(yyyy-MM-dd) mais antiga
  const skusPorStatus    = { Geral: {} }; // chave -> data -> { sku: qty }
  VG_STATUS_ALVO.forEach(n => { skusPorStatus[n] = {}; });

  pedidosMap.forEach((pedido, orderId) => {
    const { status, data, itens } = pedido;
    let itensRelevantes;
    if (isComEstoque) {
      if (!itens.every(it => temEstoque(it.sku))) return;
      itensRelevantes = itens;
    } else {
      itensRelevantes = itens.filter(it => !temEstoque(it.sku));
      if (!itensRelevantes.length) return;
    }

    if (!datasSet[data]) datasSet[data] = new Set();
    datasSet[data].add(orderId);

    if (!statusOldestDate[status] || data < statusOldestDate[status]) statusOldestDate[status] = data;

    ['Geral', status].forEach(chave => {
      if (!skusPorStatus[chave]) skusPorStatus[chave] = {};
      if (!skusPorStatus[chave][data]) skusPorStatus[chave][data] = {};
      itensRelevantes.forEach(it => {
        skusPorStatus[chave][data][it.sku] = (skusPorStatus[chave][data][it.sku] || 0) + it.qty;
      });
    });
  });

  const datas = Object.keys(datasSet).sort().map(ds => ({
    data: _vg_fmtDataYMD(ds, tz),
    pedidos: datasSet[ds].size,
  }));

  const status = VG_STATUS_ALVO.map(nome => ({
    nome,
    dataAntiga: statusOldestDate[nome] ? _vg_fmtDataYMD(statusOldestDate[nome], tz) : '',
  }));

  const skusFormatado = {};
  Object.keys(skusPorStatus).forEach(chave => {
    const porData = skusPorStatus[chave];
    skusFormatado[chave] = Object.keys(porData).sort().map(ds => ({
      data: _vg_fmtDataYMD(ds, tz),
      itens: Object.entries(porData[ds])
        .map(([sku, qty]) => ({ sku, qty }))
        .sort((a, b) => b.qty - a.qty),
    }));
  });

  return { datas, status, skusPorStatus: skusFormatado };
}

// Lê a aba raw já preenchida e retorna JSON para o dashboard HTML
function vg_getDados() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const tz   = ss.getSpreadsheetTimeZone();
  const hoje = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  const dados = { hoje, timestamp: '', comEstoque: null, semEstoque: null, funcionarios: [] };

  const abaRaw = ss.getSheetByName(VG_ABA_RAW);
  if (abaRaw) {
    const ts = String(abaRaw.getRange('A1').getValue());
    dados.timestamp = ts.replace(/.*atualização:\s*/i, '').trim();

    const last = abaRaw.getLastRow();
    if (last >= 3) {
      const raw = abaRaw.getRange(3, 1, last - 2, 5).getValues();
      const pedidosMap = new Map(); // orderId -> { status, data, itens: [{sku,qty}] }
      raw.forEach(([status, data, orderId, sku, qty]) => {
        if (!orderId || !sku) return;
        const oid = String(orderId);
        if (!pedidosMap.has(oid)) {
          pedidosMap.set(oid, { status: String(status), data: String(data), itens: [] });
        }
        pedidosMap.get(oid).itens.push({ sku: String(sku), qty: Number(qty) || 0 });
      });

      const estoqueMap = _vg_getEstoqueMap();
      const temEstoque = sku => (estoqueMap.get(sku) || 0) > 0;

      dados.comEstoque = _vg_montarSecao(pedidosMap, temEstoque, tz, true);
      dados.semEstoque = _vg_montarSecao(pedidosMap, temEstoque, tz, false);
    }
  }

  // Aba Embalagem
  const abaEmb = ss.getSheetByName('Embalagem');
  if (abaEmb) {
    dados.embalagem = {
      hoje:      Number(abaEmb.getRange('A2').getValue()) || 0,
      picoData:  abaEmb.getRange('A3').getDisplayValue() || '',
      picoQtd:   Number(abaEmb.getRange('A4').getValue()) || 0,
      mes:       String(abaEmb.getRange('A5').getValue() || ''),
    };
    const lastRowEmb = abaEmb.getLastRow();
    if (lastRowEmb >= 2) {
      dados.funcionarios = abaEmb.getRange(2, 4, lastRowEmb - 1, 3).getValues()
        .filter(r => r[0])
        .map(r => ({ nome: String(r[0]), hoje: Number(r[1]) || 0, mes: Number(r[2]) || 0 }));
    }
  }

  return dados;
}

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Visão Geral')
    .addItem('🔄 Atualizar Datas e SKUs', 'vg_atualizar')
    .addSeparator()
    .addItem('⏱ Configurar atualização automática (5 min)', 'vg_configurarGatilho')
    .addItem('🗑 Remover atualização automática', 'vg_removerGatilho')
    .addSeparator()
    .addItem('👥 Criar/resetar aba de Funcionários', 'vg_criarAbaFuncionarios')
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

// ── Gatilho automático ─────────────────────────────────────────
function vg_atualizar_automatico() {
  vg_atualizar(true);
}

function vg_atualizarEmbalagem_automatico() {
  vg_atualizarEmbalagem();
}

function vg_configurarGatilho() {
  const fns = ['vg_atualizar_automatico', 'vg_atualizarEmbalagem_automatico'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (fns.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('vg_atualizar_automatico').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('vg_atualizarEmbalagem_automatico').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert('✅ Atualização automática configurada a cada 30 minutos.');
}

function vg_removerGatilho() {
  const fns = ['vg_atualizar_automatico', 'vg_atualizarEmbalagem_automatico'];
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (fns.includes(t.getHandlerFunction())) { ScriptApp.deleteTrigger(t); removidos++; }
  });
  SpreadsheetApp.getUi().alert(
    removidos > 0 ? '✅ Atualização automática removida.' : 'Nenhum gatilho automático encontrado.'
  );
}

// ── Funcionários: whitelist na aba "Funcionários" ────────────
// Retorna Set com os nomes válidos (lowercase) ou null se a aba
// não existir (sem filtro, aceita qualquer valor).
// Retorna Map: admin_comments (lowercase) → nome do dash (col B)
// null se a aba não existir ou não tiver dados a partir da linha 2.
function _vg_getValidEmployees(ss) {
  const aba = ss.getSheetByName('Funcionários');
  if (!aba || aba.getLastRow() < 2) return null;
  const map = new Map();
  aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues()
    .forEach(([a, b]) => {
      const key     = String(a || '').trim();
      const display = String(b || '').trim() || key;
      if (key) map.set(key.toLowerCase(), display);
    });
  return map.size ? map : null;
}

// Cria (ou recria) a aba "Funcionários" com a lista pré-configurada.
// Col A = valor exato que aparece em admin_comments (ação automática BaseLinker)
// Col B = nome que será exibido no dashboard
// Adicione ou remova linhas diretamente na planilha conforme necessário.
function vg_criarAbaFuncionarios() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const VG_T_LOCAL = VG_T;

  // [col A: admin_comments value, col B: display name no dash]
  const funcionarios = [
    ['Anna Paula Alves', 'Anna Paula'],
    ['Camila Rangel',    'Camila'],
    ['Eduardo',          'Eduardo'],
    ['Eduardo Leite',    'Edu Leite'],
    ['Eiky Eduardo',     'Eiky'],
    ['Freelancer',       'Freelancer'],
    ['Kauã Galindo',     'Kauã'],
    ['Lucas Prisco',     'Lucas'],
    ['Marcos Felipe',    'Marcos'],
    ['Mirian',           'Mirian'],
    ['Nalva',            'Nalva'],
    ['Nicolas Ventura',  'Nicolas'],
    ['Roger',            'Roger'],
    ['Sky Moto',         'Junior'],
    ['Yolanda Ferreira', 'Yolanda'],
  ];

  let aba = ss.getSheetByName('Funcionários');
  if (aba) ss.deleteSheet(aba);
  aba = ss.insertSheet('Funcionários');
  aba.setTabColor(VG_T_LOCAL.accent);

  const totalR = funcionarios.length + 3;
  aba.getRange(1, 1, totalR, 2)
    .setBackground(VG_T_LOCAL.bgBase)
    .setFontColor(VG_T_LOCAL.textMain)
    .setFontFamily('Arial')
    .setFontSize(11);

  aba.getRange(1, 1, 1, 2).setBackground(VG_T_LOCAL.bgHeader);
  aba.getRange('A1').setValue('Nome na BaseLinker')
    .setFontWeight('bold').setFontColor(VG_T_LOCAL.accentSub).setFontSize(11);
  aba.getRange('B1').setValue('Nome no Dash')
    .setFontWeight('bold').setFontColor(VG_T_LOCAL.accentSub).setFontSize(11)
    .setHorizontalAlignment('center');
  aba.setRowHeight(1, 32);

  aba.getRange(2, 1, funcionarios.length, 2).setValues(funcionarios);
  aba.getRange(2, 2, funcionarios.length, 1)
    .setFontColor(VG_T_LOCAL.green).setHorizontalAlignment('center');

  for (let r = 2; r <= funcionarios.length + 1; r++) aba.setRowHeight(r, 28);
  aba.setColumnWidth(1, 220);
  aba.setColumnWidth(2, 140);
  aba.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    '✅ Aba "Funcionários" criada com ' + funcionarios.length + ' nomes.\n' +
    'Coluna A = nome configurado na ação automática do BaseLinker\n' +
    'Coluna B = nome exibido no dashboard\n\n' +
    'Edite diretamente na planilha para adicionar ou remover.'
  );
}

// ── Embalagem: conta todos os pedidos que PASSARAM por embalagem ──
// - Status [EXP]: date_in_status é a data de embalagem (prova direta)
// - Outros status pós-embalagem (Enviado, Entregue, etc.):
//   admin_comments com funcionário válido = prova de que foi embalado;
//   date_in_status = proxy da data (coleta Shopee no mesmo dia)
// - Status [SEP] são ignorados (ainda não foram embalados)
function vg_atualizarEmbalagem() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const tz    = ss.getSpreadsheetTimeZone();
  const hoje  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const mes   = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const props = PropertiesService.getScriptProperties();

  // Início do mês como Unix timestamp para filtro date_confirmed_from
  const [anoN, mesN] = mes.split('-').map(Number);
  const inicioMesSec = Math.floor(new Date(anoN, mesN - 1, 1).getTime() / 1000);

  // Classificar todos os status da conta
  const statusResp  = vg_bl('getOrderStatusList', {});
  const todosStatus = statusResp.statuses || [];
  const expStatuses  = todosStatus.filter(s => s.name.startsWith('[EXP]'));
  const postStatuses = todosStatus.filter(s =>
    !s.name.startsWith('[EXP]') && !s.name.startsWith('[SEP]')
  );

  // ── Aba Embalagem ──────────────────────────────────────────────
  let aba = ss.getSheetByName('Embalagem');
  if (!aba) aba = ss.insertSheet('Embalagem');

  // Detectar virada de dia → salvar contagem anterior no histórico
  const h1Val      = aba.getRange('H1').getValue();
  const storedDate = h1Val instanceof Date
    ? Utilities.formatDate(h1Val, tz, 'yyyy-MM-dd')
    : String(h1Val || '');
  const hLastRow = aba.getLastRow();

  if (storedDate && storedDate !== hoje) {
    const prevCount = hLastRow >= 2
      ? aba.getRange(2, 8, hLastRow - 1, 1).getValues().filter(([id]) => id).length
      : 0;
    const prevMes  = storedDate.slice(0, 7);
    const prevData = JSON.parse(props.getProperty('VG_EMB_MONTH_' + prevMes) || '{}');
    prevData[storedDate] = prevCount;
    props.setProperty('VG_EMB_MONTH_' + prevMes, JSON.stringify(prevData));
  }

  // ── Contagem por todos os status relevantes ────────────────────
  const validEmps  = _vg_getValidEmployees(ss);
  const contadosIds = new Set(); // evita dupla contagem entre status
  const hojeIds    = new Set();  // IDs do dia atual
  const funcHoje   = {};
  const funcMes    = {};

  function processarPedido(pedido, exigirFunc) {
    const oid = String(pedido.order_id);
    if (contadosIds.has(oid)) return;
    const ts = pedido.date_in_status;
    if (!ts) return;
    const ds = Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd');
    if (!ds.startsWith(mes)) return;

    const rawFunc = String(pedido.admin_comments || '').trim();
    const func    = validEmps
      ? (validEmps.get(rawFunc.toLowerCase()) || null)
      : (rawFunc || null);

    if (exigirFunc && !func) return;

    contadosIds.add(oid);
    if (ds === hoje) hojeIds.add(oid);

    if (func) {
      funcMes[func]  = (funcMes[func]  || 0) + 1;
      if (ds === hoje) funcHoje[func] = (funcHoje[func] || 0) + 1;
    }
  }

  // [EXP]: date_in_status = data de embalagem; funcionário não obrigatório
  for (const { id } of expStatuses) {
    let idFrom = 0;
    while (true) {
      const r     = vg_bl('getOrders', { status_id: id, id_from: idFrom });
      const batch = r.orders || [];
      if (!batch.length) break;
      batch.forEach(p => processarPedido(p, false));
      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }
  }

  // Pós-embalagem (Enviado, Entregue, etc.): exige admin_comments como prova
  // date_confirmed_from limita aos pedidos confirmados este mês (performance)
  for (const { id } of postStatuses) {
    let idFrom = 0;
    while (true) {
      const r     = vg_bl('getOrders', {
        status_id: id,
        id_from: idFrom,
        date_confirmed_from: inicioMesSec,
      });
      const batch = r.orders || [];
      if (!batch.length) break;
      batch.forEach(p => processarPedido(p, true));
      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }
  }

  // ── Histórico mensal (Script Properties) ─────────────────────
  const monthKey  = 'VG_EMB_MONTH_' + mes;
  const monthData = JSON.parse(props.getProperty(monthKey) || '{}');
  monthData[hoje] = hojeIds.size;
  props.setProperty(monthKey, JSON.stringify(monthData));

  // ── Métricas ─────────────────────────────────────────────────
  let picoDia = '', picoQtd = 0;
  for (const [d, q] of Object.entries(monthData)) {
    if (q > picoQtd) { picoQtd = q; picoDia = d; }
  }
  const picoDisplay = picoDia
    ? Utilities.formatDate(new Date(picoDia + 'T12:00:00'), tz, 'dd/MM/yyyy')
    : '';
  const mesDisplay = Utilities.formatDate(new Date(), tz, 'MMMM');

  // ── Gravar aba ───────────────────────────────────────────────
  aba.clearContents();
  const agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  aba.getRange('A1').setValue('Atualizado em: ' + agora);
  aba.getRange('A2').setValue(hojeIds.size);
  aba.getRange('A3').setNumberFormat('@').setValue(picoDisplay);
  aba.getRange('A4').setValue(picoQtd);
  aba.getRange('A5').setValue(mesDisplay);

  aba.getRange('D1').setValue('FUNCIONÁRIO');
  aba.getRange('E1').setValue('HOJE');
  aba.getRange('F1').setValue('MÊS');
  const todosFunc = [...new Set([...Object.keys(funcHoje), ...Object.keys(funcMes)])];
  const funcRows  = todosFunc
    .map(nome => [nome, funcHoje[nome] || 0, funcMes[nome] || 0])
    .sort((a, b) => b[1] - a[1] || b[2] - a[2]);
  if (funcRows.length) {
    aba.getRange(2, 4, funcRows.length, 3).setValues(funcRows);
  }

  // Histórico A7+
  const sorted = Object.entries(monthData).sort(([a], [b]) => a.localeCompare(b));
  if (sorted.length) {
    const rows = sorted.map(([ds, q]) => {
      const [y, m, d] = ds.split('-').map(Number);
      return [new Date(y, m - 1, d), q];
    });
    aba.getRange(7, 1, rows.length, 2).setValues(rows);
    aba.getRange(7, 1, rows.length, 1).setNumberFormat('dd/mm/yyyy');
  }

  // IDs do dia em coluna H (H1=data, H2+=IDs) — usado para virada de dia
  aba.getRange('H1').setNumberFormat('@').setValue(hoje);
  const idRows = [...hojeIds].map(id => [id]);
  if (idRows.length) aba.getRange(2, 8, idRows.length, 1).setValues(idRows);

  SpreadsheetApp.flush();
}

// ── Função principal ──────────────────────────────────────────
function vg_atualizar(silencioso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const ui = silencioso ? null : SpreadsheetApp.getUi();

  try {
    const statusResp = vg_bl('getOrderStatusList', {});
    const nameToId   = {};
    (statusResp.statuses || []).forEach(s => { nameToId[s.name] = s.id; });

    const encontrados    = VG_STATUS_ALVO.filter(n =>  nameToId[n]);
    const naoEncontrados = VG_STATUS_ALVO.filter(n => !nameToId[n]);

    if (!encontrados.length) {
      if (ui) ui.alert('❌ Nenhum dos status configurados foi encontrado no BaseLinker.');
      return;
    }
    if (naoEncontrados.length) {
      ss.toast('Aviso — status não encontrados: ' + naoEncontrados.join(', '), '⚠', 10);
    }

    ss.toast('Buscando pedidos em ' + encontrados.length + ' status…', '📊 Visão Geral', 600);

    // Coleta granular: uma linha por item de pedido — a classificação
    // por estoque (Com/Sem Estoque) é feita depois, em vg_getDados(),
    // cruzando com a planilha de estoque no momento da visualização.
    const linhasRaw = [];

    for (const nome of encontrados) {
      const sid  = nameToId[nome];
      let idFrom = 0;
      while (true) {
        const r     = vg_bl('getOrders', { status_id: sid, id_from: idFrom });
        const batch = r.orders || [];
        if (!batch.length) break;
        for (const pedido of batch) {
          const ts = pedido.date_confirmed;
          if (!ts) continue;
          const ds = Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd');
          for (const prod of (pedido.products || [])) {
            const sku = String(prod.sku || '').trim();
            if (!sku) continue;
            const qty = Number(prod.quantity) || 0;
            linhasRaw.push([nome, ds, String(pedido.order_id), sku, qty]);
          }
        }
        if (batch.length < 100) break;
        idFrom = batch[batch.length - 1].order_id;
      }
    }

    if (!linhasRaw.length) {
      if (ui) ui.alert('Nenhum pedido com date_confirmed encontrado nos status selecionados.');
      return;
    }

    const agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
    _vg_escreverRaw(ss, linhasRaw, agora);

    SpreadsheetApp.flush();
    const nDatas = new Set(linhasRaw.map(r => r[1])).size;
    if (ui) ui.alert('✅ Concluído!\n' + nDatas + ' datas encontradas\n' + encontrados.length + ' status analisados\n' + linhasRaw.length + ' itens de pedido');

  } catch (e) {
    if (ui) ui.alert('❌ Erro: ' + e.message);
    else ss.toast('❌ Erro na atualização automática: ' + e.message, '📊 Visão Geral', 30);
  }
}

// ── Escrever aba raw (oculta) — Status | Data | OrderID | SKU | Qtd ──
function _vg_escreverRaw(ss, linhasRaw, agora) {
  let aba = ss.getSheetByName(VG_ABA_RAW);
  if (!aba) {
    aba = ss.insertSheet(VG_ABA_RAW);
    aba.hideSheet();
  }
  aba.clearContents();
  aba.clearFormats();
  aba.getRange(1, 1).setValue('Atualizado em: ' + agora);
  aba.getRange(2, 1, 1, 5).setValues([['Status', 'Data', 'OrderID', 'SKU', 'Qtd']]);
  if (linhasRaw.length) {
    aba.getRange(3, 1, linhasRaw.length, 5).setValues(linhasRaw);
  }
}
