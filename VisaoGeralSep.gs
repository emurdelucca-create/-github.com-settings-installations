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
const VG_ABA_ESTOQUE_INSUF = '_vg_estoque_insuf_'; // aba oculta: SKU | QtdInsuficiente | QtdComLocalizacao

// Chave de acesso à página de administração do token da extensão
// (?admin=<chave> na URL do Web App — ver doGet). Só quem tiver essa URL
// completa consegue ver/gerar o token; quem só usa a extensão nunca
// precisa (nem deve) conhecer essa chave.
const VG_ADMIN_KEY = 'XW1M14E39sLuYAJQMiUDPKN5';

const VG_STATUS_ALVO = [
  'NF Emitida',
  'Erro NF',
  'Erro Etiqueta',
  'Movimentação',
  '[SEP] ML Flex',
  '[SEP] ML Agência',
  '[SEP] Shopee Direta',
  '[SEP] Shopee Xpress',
  '[SEP] TikTok',
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
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.admin && params.admin === VG_ADMIN_KEY) {
    return _vg_paginaAdminToken();
  }
  return HtmlService.createHtmlOutputFromFile('VisaoGeralDashboard')
    .setTitle('Visão Geral — Separação')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Página simples (HTML puro, sem depender da planilha nem do editor do
// Apps Script) que mostra o token da extensão, gerando um se ainda não
// existir. Só acessível via ?admin=<VG_ADMIN_KEY> na URL do Web App.
function _vg_paginaAdminToken() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('VG_EXT_TOKEN');
  if (!token) {
    token = Utilities.getUuid();
    props.setProperty('VG_EXT_TOKEN', token);
  }
  const html = `<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * { box-sizing: border-box; }
  body { background:#0d1117; color:#e6edf3; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
         display:flex; align-items:center; justify-content:center; min-height:100vh; margin:0; padding:16px; }
  .card { background:#161b22; border:1px solid #30363d; border-radius:10px; padding:24px; max-width:420px; width:100%; }
  h1 { font-size:15px; color:#388bfd; margin:0 0 6px; }
  p { font-size:12px; color:#8b949e; margin:0 0 16px; line-height:1.5; }
  .token-box { display:flex; gap:8px; }
  input { flex:1; background:#0d1117; border:1px solid #30363d; color:#e6edf3; border-radius:6px; padding:8px 10px; font-size:13px; font-family:'Courier New',monospace; }
  button { background:#388bfd; color:#fff; border:none; border-radius:6px; padding:8px 14px; font-size:12px; font-weight:700; cursor:pointer; }
  button:hover { opacity:0.85; }
  #msg { font-size:11px; color:#3fb950; margin-top:8px; min-height:14px; }
</style></head>
<body>
  <div class="card">
    <h1>🔑 Token da extensão</h1>
    <p>Cole esse valor no popup da extensão "Estoque Insuficiente BaseLinker", no campo Token, em cada dispositivo. Mesmo token para todos.</p>
    <div class="token-box">
      <input id="tk" readonly value="${token}">
      <button onclick="copiar()">Copiar</button>
    </div>
    <div id="msg"></div>
  </div>
  <script>
    function copiar() {
      var el = document.getElementById('tk');
      el.select();
      navigator.clipboard.writeText(el.value).then(function() {
        document.getElementById('msg').textContent = '✅ Copiado!';
      });
    }
  </script>
</body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('Token da extensão')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Recebe capturas da extensão "Estoque Insuficiente" (POST em text/plain
// pra evitar preflight CORS). Corpo esperado:
//   { token: "...", itens: [{ sku, qtdInsuficiente, qtdComLocalizacao }] }
// Cada captura SUBSTITUI os dados anteriores (snapshot do estado atual do
// painel de coleta no momento do clique em "Obter Dados").
function doPost(e) {
  const responder = obj => ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const body = JSON.parse(e.postData.contents);
    const tokenEsperado = PropertiesService.getScriptProperties().getProperty('VG_EXT_TOKEN');
    if (!tokenEsperado || body.token !== tokenEsperado) {
      return responder({ ok: false, error: 'Token inválido.' });
    }

    const itens = Array.isArray(body.itens) ? body.itens : [];
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const tz    = ss.getSpreadsheetTimeZone();
    const agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

    _vg_escreverEstoqueInsuficiente(ss, itens, agora);
    SpreadsheetApp.flush();

    return responder({ ok: true, skus: itens.length });
  } catch (err) {
    return responder({ ok: false, error: err.message });
  }
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
  // chave -> data -> sku -> { qty, orderIds: Set } — orderIds alimenta o
  // popup "ver pedidos" ao passar o mouse no SKU, no HTML.
  const skusPorStatus = { Geral: {} };
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
      const porSku = skusPorStatus[chave][data];
      itensRelevantes.forEach(it => {
        if (!porSku[it.sku]) porSku[it.sku] = { qty: 0, orderIds: new Set() };
        porSku[it.sku].qty += it.qty;
        porSku[it.sku].orderIds.add(orderId);
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
        .map(([sku, info]) => ({ sku, qty: info.qty, orderIds: Array.from(info.orderIds).sort() }))
        .sort((a, b) => b.qty - a.qty),
    }));
  });

  return { datas, status, skusPorStatus: skusFormatado };
}

// Lê a aba raw já preenchida e retorna JSON para o dashboard HTML
// O Google Sheets auto-converte strings "yyyy-MM-dd" gravadas numa
// célula em objetos Date de verdade, mesmo a coluna não tendo sido
// formatada como data explicitamente. String(dataObj) produziria algo
// tipo "Tue Sep 01 2026 00:00:00 GMT-0300 (...)", que o parser de data
// (ds.split('-')) não entende, gerando NaN e caindo no epoch (exibido
// como 31/12/1969 com fuso negativo). Por isso a leitura é defensiva:
// se vier um Date de verdade, reformata para "yyyy-MM-dd" explicitamente.
function _vg_toDateStr(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v || '');
}

function vg_getDados() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const tz   = ss.getSpreadsheetTimeZone();
  const hoje = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  const dados = { hoje, timestamp: '', comEstoque: null, semEstoque: null, funcionarios: [] };

  const abaRaw = ss.getSheetByName(VG_ABA_RAW);
  if (abaRaw) {
    const ts = String(abaRaw.getRange('A1').getValue());
    dados.timestamp = ts.replace(/^Atualizado em:\s*/i, '').trim();

    const last = abaRaw.getLastRow();
    if (last >= 3) {
      const raw = abaRaw.getRange(3, 1, last - 2, 5).getValues();
      const pedidosMap = new Map(); // orderId -> { status, data, itens: [{sku,qty}] }
      raw.forEach(([status, data, orderId, sku, qty]) => {
        if (!orderId || !sku) return;
        const oid = String(orderId);
        if (!pedidosMap.has(oid)) {
          pedidosMap.set(oid, { status: String(status), data: _vg_toDateStr(data, tz), itens: [] });
        }
        pedidosMap.get(oid).itens.push({ sku: String(sku), qty: Number(qty) || 0 });
      });

      const estoqueMap = _vg_getEstoqueMap();
      const temEstoque = sku => (estoqueMap.get(sku) || 0) > 0;

      dados.comEstoque = _vg_montarSecao(pedidosMap, temEstoque, tz, true);
      dados.semEstoque = _vg_montarSecao(pedidosMap, temEstoque, tz, false);
    }
  }

  // Aba Estoque Insuficiente (alimentada pela extensão de captura)
  const abaInsuf = ss.getSheetByName(VG_ABA_ESTOQUE_INSUF);
  dados.estoqueInsuficiente = { timestamp: '', itens: [] };
  if (abaInsuf) {
    const tsInsuf = String(abaInsuf.getRange('A1').getValue());
    dados.estoqueInsuficiente.timestamp = tsInsuf.replace(/^Atualizado em:\s*/i, '').trim();
    const lastInsuf = abaInsuf.getLastRow();
    if (lastInsuf >= 3) {
      dados.estoqueInsuficiente.itens = abaInsuf.getRange(3, 1, lastInsuf - 2, 3).getValues()
        .filter(([sku]) => sku)
        .map(([sku, qtdInsuficiente, qtdComLocalizacao]) => ({
          sku: String(sku),
          qtdInsuficiente: Number(qtdInsuficiente) || 0,
          qtdComLocalizacao: Number(qtdComLocalizacao) || 0,
        }))
        // Só mostra SKUs com estoque insuficiente de fato — um SKU que só
        // apareceu no painel com localização definida não entra aqui.
        .filter(item => item.qtdInsuficiente > 0)
        .sort((a, b) => b.qtdInsuficiente - a.qtdInsuficiente);
    }
  }

  // Aba Embalagem
  const abaEmb = ss.getSheetByName('Embalagem');
  if (abaEmb) {
    const tsEmb = String(abaEmb.getRange('A1').getValue());
    dados.embalagem = {
      timestamp: tsEmb.replace(/^Atualizado em:\s*/i, '').trim(),
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

// Chamada pelo botão "Atualizar Planilha" no HTML — dispara a busca real
// no BaseLinker (a mesma coisa que o gatilho automático de 5 em 5 min
// faz) e já devolve os dados recalculados na mesma chamada, pra não
// precisar de um segundo round-trip.
// silencioso=true é obrigatório aqui: SpreadsheetApp.getUi() não funciona
// quando a função é chamada via google.script.run a partir do Web App
// (não existe diálogo de planilha nesse contexto).
function vg_atualizarEExibir() {
  vg_atualizar(true);
  vg_atualizarEmbalagem();
  return vg_getDados();
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

// Gera (ou reexibe) o token compartilhado que a extensão de captura de
// Estoque Insuficiente usa para autenticar no doPost.
// Propositalmente SEM item de menu na planilha: qualquer pessoa com acesso
// de edição à planilha conseguiria ver o token através de um menu, mas o
// acesso à planilha e o acesso ao token/extensão são coisas independentes
// (quem usa a extensão no depósito não precisa nem deve ter acesso à
// planilha). Para gerar/ver o token, rode esta função direto pelo editor:
// Extensões → Apps Script → selecione "vg_gerarTokenExtensao" → ▶ Executar
// → Ver → Registros de execução (Logger.log mostra o token).
function vg_gerarTokenExtensao() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('VG_EXT_TOKEN');
  if (!token) {
    token = Utilities.getUuid();
    props.setProperty('VG_EXT_TOKEN', token);
  }
  Logger.log('Token da extensão (copie e cole no popup, campo Token): ' + token);
  return token;
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
  ScriptApp.newTrigger('vg_atualizar_automatico').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('vg_atualizarEmbalagem_automatico').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('✅ Atualização automática configurada a cada 5 minutos.');
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

// ── Embalagem: ledger permanente de pedidos embalados ─────────────
// Reescrito porque a versão anterior recontava tudo do zero a cada
// execução, consultando o BaseLinker por status ATUAL — quando um
// pedido "sumia" do BaseLinker (arquivado/fora da janela de retenção),
// ele desaparecia das contagens pra sempre, mesmo tendo sido embalado
// de verdade. Agora: sinal único (admin_comments bate com um
// funcionário da aba "Funcionários", em QUALQUER status), e assim que
// um pedido é visto uma vez, é gravado permanentemente na aba oculta
// _vg_emb_ledger_ com a data de HOJE (a execução roda a cada 5 min, então
// "quando detectamos" e "quando foi embalado" coincidem na prática).
// Um pedido já gravado nunca é reprocessado nem pode sumir das métricas.
const VG_ABA_EMB_LEDGER = '_vg_emb_ledger_'; // OrderID | Funcionario | Data (dd/MM/yyyy)

// Âncora FIXA (não rolante) de onde a varredura de embalagem passa a
// contar — nunca dias antes disso. Guardada em Script Properties, criada
// uma única vez (na primeira execução após esse fix) como "hoje 00:00".
// Sem isso, toda vez que o recurso reiniciasse contaria semanas de
// pedidos antigos de uma vez, carimbando tudo com "hoje" — foi exatamente
// o bug relatado (575 hoje / 2.642 pico / 12.221 no mês, tudo inflado).
// Constrói a meia-noite de HOJE como Date local, com os componentes
// ano/mês/dia extraídos via Utilities.formatDate (respeita o fuso da
// planilha) e montados com "new Date(ano, mes-1, dia)" — o MESMO padrão
// já usado (e comprovadamente correto) em _vg_fmtDataYMD/inicioMesSec.
// Evita de propósito qualquer round-trip por string ISO tipo
// "new Date('2026-09-30T00:00:00')": sem fuso explícito, o V8 pode
// interpretar isso como horário local do AMBIENTE DE EXECUÇÃO (não
// necessariamente o fuso da planilha), o que gerava uma âncora errada e
// deixava o filtro date_confirmed_from praticamente sem efeito — foi
// exatamente o motivo dos números terem piorado (22 mil pedidos "hoje").
function _vg_meiaNoiteHojeSec(tz) {
  const [ano, mes, dia] = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-').map(Number);
  return Math.floor(new Date(ano, mes - 1, dia).getTime() / 1000);
}

function _vg_getEmbInicioSec(props, tz) {
  let v = props.getProperty('VG_EMB_LEDGER_INICIO');
  if (!v) {
    v = String(_vg_meiaNoiteHojeSec(tz));
    props.setProperty('VG_EMB_LEDGER_INICIO', v);
  }
  return Number(v);
}

// Zera o ledger e a aba Embalagem UMA ÚNICA VEZ (controlado por uma flag
// em Script Properties) — limpa os dados poluídos pelas execuções
// anteriores (que varreram o histórico quase inteiro de uma vez, por
// causa do bug acima) e reancora VG_EMB_LEDGER_INICIO em "agora", pra
// passar a contar só daqui pra frente. V3 porque a V2 já rodou com a
// âncora quebrada — precisa resetar de novo com o cálculo corrigido.
function _vg_resetLedgerEmbalagemSeNecessario(ss, props, tz) {
  if (props.getProperty('VG_EMB_LEDGER_RESET_V3')) return;

  const aba = ss.getSheetByName(VG_ABA_EMB_LEDGER);
  if (aba) {
    aba.clearContents();
    aba.getRange(1, 1, 1, 3).setValues([['OrderID', 'Funcionario', 'Data']]);
  }
  const abaEmb = ss.getSheetByName('Embalagem');
  if (abaEmb) abaEmb.clearContents();

  props.setProperty('VG_EMB_LEDGER_INICIO', String(_vg_meiaNoiteHojeSec(tz)));
  props.setProperty('VG_EMB_LEDGER_RESET_V3', '1');
}

function _vg_getLedgerOrderIds(ss) {
  const set = new Set();
  const aba = ss.getSheetByName(VG_ABA_EMB_LEDGER);
  if (!aba) return set;
  const last = aba.getLastRow();
  if (last < 2) return set;
  aba.getRange(2, 1, last - 1, 1).getValues().forEach(([id]) => { if (id) set.add(String(id)); });
  return set;
}

function _vg_appendLedger(ss, novasLinhas) {
  if (!novasLinhas.length) return;
  let aba = ss.getSheetByName(VG_ABA_EMB_LEDGER);
  if (!aba) {
    aba = ss.insertSheet(VG_ABA_EMB_LEDGER);
    aba.hideSheet();
    aba.getRange(1, 1, 1, 3).setValues([['OrderID', 'Funcionario', 'Data']]);
  }
  const startRow = aba.getLastRow() + 1;
  // Texto puro nas duas colunas — evita o Sheets auto-converter e quebrar
  // a leitura depois (mesmo problema já visto em _vg_escreverRaw).
  aba.getRange(startRow, 1, novasLinhas.length, 1).setNumberFormat('@');
  aba.getRange(startRow, 3, novasLinhas.length, 1).setNumberFormat('@');
  aba.getRange(startRow, 1, novasLinhas.length, 3).setValues(novasLinhas);
}

function _vg_toDateStrDMY(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
  return String(v || '').trim();
}

function vg_atualizarEmbalagem() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const tz    = ss.getSpreadsheetTimeZone();
  const props = PropertiesService.getScriptProperties();

  _vg_resetLedgerEmbalagemSeNecessario(ss, props, tz);
  const dateConfirmedFromSec = _vg_getEmbInicioSec(props, tz);

  const validEmps      = _vg_getValidEmployees(ss);
  const idsExistentes  = _vg_getLedgerOrderIds(ss);
  const hojeDisplay     = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  const novasLinhas    = []; // [OrderID, Funcionario, Data]

  const statusResp  = vg_bl('getOrderStatusList', {});
  const todosStatus = statusResp.statuses || [];

  todosStatus.forEach(({ id }) => {
    let idFrom = 0;
    while (true) {
      const r     = vg_bl('getOrders', {
        status_id: id,
        id_from: idFrom,
        date_confirmed_from: dateConfirmedFromSec,
      });
      const batch = r.orders || [];
      if (!batch.length) break;
      batch.forEach(pedido => {
        const oid = String(pedido.order_id);
        if (idsExistentes.has(oid)) return; // já registrado — nunca reprocessa

        const rawFunc = String(pedido.admin_comments || '').trim();
        const func    = validEmps
          ? (validEmps.get(rawFunc.toLowerCase()) || null)
          : (rawFunc || null);
        if (!func) return;

        idsExistentes.add(oid); // evita duplicar se aparecer em 2 status na mesma execução
        novasLinhas.push([oid, func, hojeDisplay]);
      });
      if (batch.length < 100) break;
      idFrom = batch[batch.length - 1].order_id;
    }
  });

  _vg_appendLedger(ss, novasLinhas);
  _vg_recalcularEmbalagemDoLedger(ss, tz);
}

// Recalcula a aba "Embalagem" (a que o dashboard lê) inteiramente a
// partir do ledger — nunca a partir de uma nova consulta ao BaseLinker.
// Isso garante que um pedido que "sumiu" do BaseLinker continua contando
// pra sempre, porque já está salvo aqui.
function _vg_recalcularEmbalagemDoLedger(ss, tz) {
  const hojeDisplay = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  const mesAtual    = Utilities.formatDate(new Date(), tz, 'MM/yyyy');

  const porDia    = {}; // 'dd/MM/yyyy' -> qtd de pedidos
  const funcHoje  = {};
  const funcMes   = {};

  const abaLedger = ss.getSheetByName(VG_ABA_EMB_LEDGER);
  if (abaLedger) {
    const last = abaLedger.getLastRow();
    if (last >= 2) {
      abaLedger.getRange(2, 1, last - 1, 3).getValues().forEach(([orderId, func, dataVal]) => {
        if (!orderId) return;
        const ds = _vg_toDateStrDMY(dataVal, tz);
        if (!ds) return;

        porDia[ds] = (porDia[ds] || 0) + 1;
        if (ds.slice(3) === mesAtual) funcMes[func] = (funcMes[func] || 0) + 1;
        if (ds === hojeDisplay)       funcHoje[func] = (funcHoje[func] || 0) + 1;
      });
    }
  }

  let picoDia = '', picoQtd = 0;
  Object.keys(porDia).forEach(ds => {
    if (ds.slice(3) === mesAtual && porDia[ds] > picoQtd) { picoQtd = porDia[ds]; picoDia = ds; }
  });
  const hojeQtd    = porDia[hojeDisplay] || 0;
  const mesDisplay = Utilities.formatDate(new Date(), tz, 'MMMM');

  let aba = ss.getSheetByName('Embalagem');
  if (!aba) aba = ss.insertSheet('Embalagem');
  aba.clearContents();
  const agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  aba.getRange('A1').setValue('Atualizado em: ' + agora);
  aba.getRange('A2').setValue(hojeQtd);
  aba.getRange('A3').setNumberFormat('@').setValue(picoDia);
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

  // Histórico diário completo (não só o mês corrente) a partir de A7 —
  // ordenado pela data real, não pela string "dd/MM/yyyy".
  const historico = Object.keys(porDia).sort((a, b) => {
    const [da, ma, ya] = a.split('/').map(Number);
    const [db, mb, yb] = b.split('/').map(Number);
    return new Date(ya, ma - 1, da).getTime() - new Date(yb, mb - 1, db).getTime();
  });
  if (historico.length) {
    const rows = historico.map(ds => {
      const [d, m, y] = ds.split('/').map(Number);
      return [new Date(y, m - 1, d), porDia[ds]];
    });
    aba.getRange(7, 1, rows.length, 2).setValues(rows);
    aba.getRange(7, 1, rows.length, 1).setNumberFormat('dd/mm/yyyy');
  }

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

// ── Escrever aba oculta de Estoque Insuficiente (captura via extensão) ──
// Cada chamada SUBSTITUI o conteúdo anterior (snapshot do momento da captura).
function _vg_escreverEstoqueInsuficiente(ss, itens, agora) {
  let aba = ss.getSheetByName(VG_ABA_ESTOQUE_INSUF);
  if (!aba) {
    aba = ss.insertSheet(VG_ABA_ESTOQUE_INSUF);
    aba.hideSheet();
  }
  aba.clearContents();
  aba.clearFormats();
  aba.getRange(1, 1).setValue('Atualizado em: ' + agora);
  aba.getRange(2, 1, 1, 3).setValues([['SKU', 'QtdInsuficiente', 'QtdComLocalizacao']]);
  if (itens.length) {
    const linhas = itens.map(it => [
      String(it.sku || '').trim(),
      Number(it.qtdInsuficiente) || 0,
      Number(it.qtdComLocalizacao) || 0,
    ]).filter(([sku]) => sku);
    if (linhas.length) aba.getRange(3, 1, linhas.length, 3).setValues(linhas);
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
    // Formata a coluna B (Data) como texto puro ANTES de escrever, para
    // que o Sheets não auto-converta a string "yyyy-MM-dd" num objeto
    // Date (ver _vg_toDateStr para o efeito colateral disso na leitura).
    aba.getRange(3, 2, linhasRaw.length, 1).setNumberFormat('@');
    aba.getRange(3, 1, linhasRaw.length, 5).setValues(linhasRaw);
  }
}
