// ============================================================
// PEDIDOS DE COMPRA — Web App autônomo
//
// Deploy: Extensões → Apps Script → Implantar → Nova implantação
//   Tipo: App da Web | Executar como: Eu | Acesso: Organização
//
// Propriedades do Script necessárias (Configurações ⚙ → Propriedades):
//   BLING_CLIENT_ID     — Client ID do app em developer.bling.com.br
//   BLING_CLIENT_SECRET — Client Secret do app
//
// Os tokens abaixo são gerenciados automaticamente pelo código:
//   BLING_ACCESS_TOKEN  — renovado automaticamente via refresh
//   BLING_REFRESH_TOKEN — salvo após primeira autorização
//   BLING_TOKEN_EXPIRES — timestamp de expiração
// ============================================================

const PC_SS_COMPRAS_ID  = '1GG6EenKiOO1K8XN0JrnD8W_PicQxzxkhsyDrU4yJa4w';
const PC_SS_CONTROLE_ID = '1eNMjC-iGBCAdVbJYSnP_Y4p7hJAzF6Gx';
const PC_ABA_COMPRAS    = 'Com Desmembramento';
const PC_GID_CONTROLE   = 1005476335;

const PC_IDX_SKU    = 0;   // col A
const PC_IDX_FORNEC = 54;  // col BC
const PC_IDX_QTD    = 56;  // col BE

// Colunas de quantidade por empresa (0-indexed)
const PC_EMPRESAS = [
  { nome: 'HUMBLE',    col: 57 },  // col BF
  { nome: 'NAJUMI',    col: 59 },  // col BH
  { nome: 'SKY',       col: 61 },  // col BJ
  { nome: 'EDMOTOS',   col: 63 },  // col BL
  { nome: 'MOTO VIBE', col: 65 },  // col BN
];
const PC_NCOLS_COMP = 66;  // lê até col BN (1-indexed)

const PC_CTRL_COD_FORN = 10; // col K
const PC_CTRL_SKU      = 11; // col L
const PC_CTRL_PREC     = 16; // col Q

const BLING_REDIRECT = 'https://www.google.com';
const BLING_TOKEN_URL = 'https://www.bling.com.br/Api/v3/oauth/token';
const BLING_API_BASE  = 'https://www.bling.com.br/Api/v3';

// ── Web App ──────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('PedidosCompra.html')
    .setTitle('Pedidos de Compra')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── OAuth 2.0 — Bling v3 ─────────────────────────────────────

function _pc_props() {
  return PropertiesService.getScriptProperties();
}

// Retorna access_token válido; renova automaticamente se expirado
function _pc_getToken() {
  const p       = _pc_props();
  const access  = p.getProperty('BLING_ACCESS_TOKEN');
  const refresh = p.getProperty('BLING_REFRESH_TOKEN');
  const expires = parseInt(p.getProperty('BLING_TOKEN_EXPIRES') || '0');

  if (!refresh) throw new Error('Bling não autorizado. Use o botão "Conectar ao Bling" no app.');

  // Renova se faltam menos de 5 minutos para expirar
  if (!access || Date.now() >= expires - 300000) {
    return _pc_refreshToken(refresh);
  }
  return access;
}

function _pc_refreshToken(refreshToken) {
  const p           = _pc_props();
  const clientId    = p.getProperty('BLING_CLIENT_ID');
  const clientSecret = p.getProperty('BLING_CLIENT_SECRET');
  if (!clientId || !clientSecret) throw new Error('BLING_CLIENT_ID / BLING_CLIENT_SECRET não configurados');

  const cred = Utilities.base64Encode(clientId + ':' + clientSecret);
  const res  = UrlFetchApp.fetch(BLING_TOKEN_URL, {
    method: 'POST',
    headers: { 'Authorization': 'Basic ' + cred, 'Content-Type': 'application/x-www-form-urlencoded' },
    payload: 'grant_type=refresh_token&refresh_token=' + encodeURIComponent(refreshToken),
    muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('Refresh falhou: ' + (body.error_description || JSON.stringify(body)));

  const newAccess   = body.access_token;
  const newRefresh  = body.refresh_token || refreshToken;
  const newExpires  = Date.now() + (body.expires_in || 3600) * 1000;

  p.setProperty('BLING_ACCESS_TOKEN',  newAccess);
  p.setProperty('BLING_REFRESH_TOKEN', newRefresh);
  p.setProperty('BLING_TOKEN_EXPIRES', String(newExpires));
  return newAccess;
}

// Retorna a URL de autorização para o usuário abrir no navegador
function pc_getAuthUrl() {
  const clientId = _pc_props().getProperty('BLING_CLIENT_ID');
  if (!clientId) return { ok: false, error: 'BLING_CLIENT_ID não configurado nas Propriedades do Script' };

  const url = 'https://www.bling.com.br/Api/v3/oauth/authorize' +
    '?response_type=code' +
    '&client_id=' + encodeURIComponent(clientId) +
    '&redirect_uri=' + encodeURIComponent(BLING_REDIRECT) +
    '&state=' + Date.now();
  return { ok: true, url };
}

// Troca o código de autorização por access_token + refresh_token
function pc_trocarCodigo(rawInput) {
  try {
    // Aceita tanto o código puro quanto a URL completa
    let code = String(rawInput || '').trim();
    const match = code.match(/[?&]code=([^&]+)/);
    if (match) code = decodeURIComponent(match[1]);
    if (!code) throw new Error('Código inválido');

    const p           = _pc_props();
    const clientId    = p.getProperty('BLING_CLIENT_ID');
    const clientSecret = p.getProperty('BLING_CLIENT_SECRET');
    if (!clientId || !clientSecret) throw new Error('BLING_CLIENT_ID / BLING_CLIENT_SECRET não configurados');

    const cred = Utilities.base64Encode(clientId + ':' + clientSecret);
    const res  = UrlFetchApp.fetch(BLING_TOKEN_URL, {
      method: 'POST',
      headers: { 'Authorization': 'Basic ' + cred, 'Content-Type': 'application/x-www-form-urlencoded' },
      payload: 'grant_type=authorization_code' +
               '&code='         + encodeURIComponent(code) +
               '&redirect_uri=' + encodeURIComponent(BLING_REDIRECT),
      muteHttpExceptions: true,
    });
    const body = JSON.parse(res.getContentText() || '{}');
    if (res.getResponseCode() !== 200) throw new Error(body.error_description || JSON.stringify(body));

    p.setProperty('BLING_ACCESS_TOKEN',  body.access_token);
    p.setProperty('BLING_REFRESH_TOKEN', body.refresh_token);
    p.setProperty('BLING_TOKEN_EXPIRES', String(Date.now() + (body.expires_in || 3600) * 1000));
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Verifica se já está autorizado
function pc_checkAuth() {
  const refresh = _pc_props().getProperty('BLING_REFRESH_TOKEN');
  return { autorizado: !!refresh };
}

// ── Helpers de API ────────────────────────────────────────────
function _pc_blingGet(path, params) {
  const token = _pc_getToken();
  let url = BLING_API_BASE + path;
  if (params && Object.keys(params).length) {
    url += '?' + Object.entries(params).map(([k,v]) => encodeURIComponent(k) + '=' + encodeURIComponent(v)).join('&');
  }
  const res  = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');
  if (code < 200 || code >= 300) {
    throw new Error('Bling ' + code + ': ' + (body?.error?.message || JSON.stringify(body).slice(0, 250)));
  }
  return body;
}

function _pc_blingPost(path, payload) {
  const token = _pc_getToken();
  const res   = UrlFetchApp.fetch(BLING_API_BASE + path, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json', 'Accept': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');
  if (code < 200 || code >= 300) {
    // Monta mensagem de erro com o máximo de detalhe possível
    const err   = body?.error || body;
    const msg   = err?.message || '';
    const fields = (err?.fields || []).map(f => f.msg || f.message || JSON.stringify(f)).join('; ');
    const detail = [msg, fields].filter(Boolean).join(' | ') || JSON.stringify(body).slice(0, 400);
    throw new Error('Bling ' + code + ': ' + detail);
  }
  return body;
}

// ── Planilhas ────────────────────────────────────────────────
function _pc_abaByGid(ss, gid) {
  return ss.getSheets().find(s => s.getSheetId() === gid) || ss.getSheets()[0];
}

function pc_carregarDados() {
  try {
    const ssCompras  = SpreadsheetApp.openById(PC_SS_COMPRAS_ID);
    const abaCompras = ssCompras.getSheetByName(PC_ABA_COMPRAS);
    if (!abaCompras) throw new Error('Aba "' + PC_ABA_COMPRAS + '" não encontrada');

    const lastRow = abaCompras.getLastRow();
    if (lastRow < 2) return { ok: true, fornecedores: {} };

    const raw = abaCompras.getRange(2, 1, lastRow - 1, PC_NCOLS_COMP).getValues();

    const ssCtrl   = SpreadsheetApp.openById(PC_SS_CONTROLE_ID);
    const abaCtrl  = _pc_abaByGid(ssCtrl, PC_GID_CONTROLE);
    const ctrlLast = abaCtrl.getLastRow();
    const precoMap = {};
    if (ctrlLast > 1) {
      abaCtrl.getRange(2, 1, ctrlLast - 1, 17).getValues().forEach(r => {
        const sku   = String(r[PC_CTRL_SKU]  || '').trim();
        const preco = Number(r[PC_CTRL_PREC] || 0);
        if (sku && preco > 0) precoMap[sku] = preco;
      });
    }

    const fornecedores = {};
    raw.forEach(r => {
      const sku  = String(r[PC_IDX_SKU]    || '').trim();
      const forn = String(r[PC_IDX_FORNEC] || '').trim();
      const qtd  = Number(r[PC_IDX_QTD]   || 0);
      if (!sku || !forn || !(qtd > 0)) return;
      // Quantidade por empresa
      const qtdEmpresas = {};
      PC_EMPRESAS.forEach(e => { qtdEmpresas[e.nome] = Number(r[e.col] || 0); });
      if (!fornecedores[forn]) fornecedores[forn] = [];
      fornecedores[forn].push({ sku, qtd, preco: precoMap[sku] || 0, qtdEmpresas });
    });

    return { ok: true, fornecedores };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Produtos — nomes via planilha BaseLinker ─────────────────
const PC_SS_BASELINKER_ID = '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik';
const PC_ABA_ESTOQUE      = 'estoque';
const PC_BL_COL_SKU       = 1;  // col B
const PC_BL_COL_NOME      = 8;  // col I

function pc_buscarProdutosBling(skus) {
  try {
    const ss  = SpreadsheetApp.openById(PC_SS_BASELINKER_ID);
    const aba = ss.getSheetByName(PC_ABA_ESTOQUE);
    if (!aba) throw new Error('Aba "' + PC_ABA_ESTOQUE + '" não encontrada na planilha BaseLinker');

    const last   = aba.getLastRow();
    const nomeMap = {};
    if (last > 1) {
      aba.getRange(2, 1, last - 1, PC_BL_COL_NOME + 1).getValues().forEach(r => {
        const sku  = String(r[PC_BL_COL_SKU]  || '').trim();
        const nome = String(r[PC_BL_COL_NOME] || '').trim();
        if (sku) nomeMap[sku] = nome || sku;
      });
    }

    const produtos = {};
    for (const sku of skus) {
      // id = null aqui; será resolvido no Bling só ao criar o pedido
      produtos[sku] = { id: null, nome: nomeMap[sku] || sku };
    }
    return { ok: true, produtos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Importar último cód. fornecedor por SKU (Controle de Compras) ─
// Lê a planilha Controle de Compras e retorna o ÚLTIMO código de fornecedor
// encontrado para cada SKU solicitado (a linha mais recente vence).
function pc_importarCodsFornecedor(skus) {
  try {
    const ssCtrl  = SpreadsheetApp.openById(PC_SS_CONTROLE_ID);
    const aba     = _pc_abaByGid(ssCtrl, PC_GID_CONTROLE);
    const lastRow = aba.getLastRow();
    if (lastRow < 2) return { ok: true, codigos: {} };

    const skuSet = new Set(skus.map(s => String(s).trim()));
    const ncols  = Math.max(PC_CTRL_COD_FORN, PC_CTRL_SKU) + 1;
    const raw    = aba.getRange(2, 1, lastRow - 1, ncols).getValues();

    // Percorre de cima para baixo — a última linha de cada SKU substitui as anteriores
    const codigos = {};
    raw.forEach(r => {
      const sku = String(r[PC_CTRL_SKU]      || '').trim();
      const cod = String(r[PC_CTRL_COD_FORN] || '').trim();
      if (sku && cod && skuSet.has(sku)) codigos[sku] = cod;
    });

    return { ok: true, codigos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Carregar qtd por empresa (para atualizar abas existentes) ─
function pc_carregarQtdEmpresas(skus) {
  try {
    const ssCompras  = SpreadsheetApp.openById(PC_SS_COMPRAS_ID);
    const abaCompras = ssCompras.getSheetByName(PC_ABA_COMPRAS);
    if (!abaCompras) throw new Error('Aba "' + PC_ABA_COMPRAS + '" não encontrada');

    const lastRow = abaCompras.getLastRow();
    if (lastRow < 2) return { ok: true, qtds: {} };

    const raw     = abaCompras.getRange(2, 1, lastRow - 1, PC_NCOLS_COMP).getValues();
    const skuSet  = new Set(skus.map(s => String(s).trim()));
    const qtds    = {};

    raw.forEach(r => {
      const sku = String(r[PC_IDX_SKU] || '').trim();
      if (!sku || !skuSet.has(sku)) return;
      const qtdEmpresas = {};
      PC_EMPRESAS.forEach(e => { qtdEmpresas[e.nome] = Number(r[e.col] || 0); });
      qtds[sku] = qtdEmpresas;
    });

    return { ok: true, qtds };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── CSV Fornecedores — persistência no servidor ───────────────
// Salva a lista de fornecedores (importada via CSV) nas Script Properties.
// Divide em chunks de 8 KB para respeitar o limite de 9 KB por propriedade.
function pc_salvarFornecedoresCSV(lista) {
  try {
    const p    = _pc_props();
    const json = JSON.stringify(lista);
    const CHUNK = 8000;
    const total = Math.ceil(json.length / CHUNK);
    for (let i = 0; i < total; i++) {
      p.setProperty('CSV_FORN_' + i, json.slice(i * CHUNK, (i + 1) * CHUNK));
    }
    // Apaga chunks antigos que possam ter sobrado
    for (let i = total; i < total + 20; i++) {
      if (p.getProperty('CSV_FORN_' + i) !== null) p.deleteProperty('CSV_FORN_' + i);
      else break;
    }
    p.setProperty('CSV_FORN_COUNT', String(total));
    p.setProperty('CSV_FORN_DATA',  Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Carrega a lista de fornecedores salva pelo CSV import
function pc_carregarFornecedoresCSV() {
  try {
    const p     = _pc_props();
    const total = parseInt(p.getProperty('CSV_FORN_COUNT') || '0');
    if (!total) return { ok: true, fornecedores: [], dataAtualizacao: null };
    let json = '';
    for (let i = 0; i < total; i++) {
      json += (p.getProperty('CSV_FORN_' + i) || '');
    }
    const lista = JSON.parse(json);
    const data  = p.getProperty('CSV_FORN_DATA') || '';
    return { ok: true, fornecedores: lista, dataAtualizacao: data };
  } catch(e) {
    return { ok: true, fornecedores: [], dataAtualizacao: null };
  }
}

// ── Bling — Fornecedores ─────────────────────────────────────
const PC_PROP_FORN_CACHE = 'BLING_FORNECEDORES_CACHE';
const PC_PROP_FORN_DATA  = 'BLING_FORNECEDORES_DATA';

// Carrega fornecedores do cache (Script Properties) — não bate na API
function pc_carregarFornecedoresCache() {
  const p     = _pc_props();
  const raw   = p.getProperty(PC_PROP_FORN_CACHE);
  const data  = p.getProperty(PC_PROP_FORN_DATA) || '';
  if (!raw) return { ok: true, fornecedores: [], dataAtualizacao: null };
  try {
    return { ok: true, fornecedores: JSON.parse(raw), dataAtualizacao: data };
  } catch(e) {
    return { ok: true, fornecedores: [], dataAtualizacao: null };
  }
}

// Busca todos os contatos no Bling, salva no cache e retorna a lista
function pc_atualizarFornecedores() {
  try {
    const todos  = [];
    let pagina   = 1;
    let continua = true;
    while (continua) {
      // Sem filtro de criterio: retorna todos os contatos cadastrados
      const res   = _pc_blingGet('/contatos', { pagina, limite: 100 });
      const items = res.data || [];
      items.forEach(c => {
        const nome = (c.nome || c.fantasia || '').trim();
        if (c.id && nome) todos.push({ id: String(c.id), nome });
      });
      continua = items.length === 100;
      pagina++;
      if (continua) Utilities.sleep(250);
    }
    todos.sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' }));

    const dataStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const p = _pc_props();
    p.setProperty(PC_PROP_FORN_CACHE, JSON.stringify(todos));
    p.setProperty(PC_PROP_FORN_DATA,  dataStr);
    return { ok: true, fornecedores: todos, dataAtualizacao: dataStr };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Mantido para compatibilidade — agora chama pc_atualizarFornecedores
function pc_buscarFornecedoresBling() {
  return pc_atualizarFornecedores();
}

// ── Cache de IDs de produto ───────────────────────────────────
const PC_PROP_PROD_CACHE = 'BLING_PRODUTO_IDS_CACHE';

function _pc_lerCacheProdutos() {
  const raw = _pc_props().getProperty(PC_PROP_PROD_CACHE);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch(e) { return {}; }
}

function _pc_salvarCacheProdutos(mapa) {
  _pc_props().setProperty(PC_PROP_PROD_CACHE, JSON.stringify(mapa));
}

// ── Bling — Criar pedidos ────────────────────────────────────
function pc_criarPedidos(pedidos) {
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const resultados = [];

  // Carrega cache de IDs de produto uma vez só
  const idCache = _pc_lerCacheProdutos();
  let cacheModificado = false;

  // Coleta todos os SKUs sem ID entre todos os pedidos
  const skusSemId = [];
  pedidos.forEach(ped => ped.itens.forEach(it => {
    if (!it.produtoId && idCache[it.sku]) {
      it.produtoId = idCache[it.sku];  // resolve pelo cache
    } else if (!it.produtoId) {
      skusSemId.push(it.sku);
    }
  }));

  // Busca no Bling apenas os SKUs ainda sem ID (deduplicados)
  const skusUnicos = [...new Set(skusSemId)];
  for (const sku of skusUnicos) {
    try {
      const r = _pc_blingGet('/produtos', { codigo: sku, pagina: 1, limite: 1 });
      const p = (r.data || [])[0];
      if (p) {
        idCache[sku]     = String(p.id);
        cacheModificado  = true;
        // Aplica de volta nos itens
        pedidos.forEach(ped => ped.itens.forEach(it => {
          if (it.sku === sku && !it.produtoId) it.produtoId = idCache[sku];
        }));
      }
    } catch(e) { /* ignora; vai falhar abaixo no item sem ID */ }
    Utilities.sleep(180);
  }

  if (cacheModificado) _pc_salvarCacheProdutos(idCache);

  for (const ped of pedidos) {
    try {
      if (!ped.fornecedorId) throw new Error('Fornecedor não selecionado');

      // Verifica se todos os itens têm ID
      for (const it of ped.itens) {
        if (!it.produtoId) throw new Error('Produto não encontrado no Bling para SKU: ' + it.sku);
      }

      const payload = {
        data:       hoje,
        fornecedor: { id: Number(ped.fornecedorId) },
        itens: ped.itens.map(it => {
          const item = {
            produto:    { id: Number(it.produtoId) },
            descricao:  String(it.nome  || it.sku),
            quantidade: Number(it.qtd)   || 0,
            valor:      Number(it.preco) || 0,
          };
          if (it.codFornecedor && String(it.codFornecedor).trim()) {
            item.codigo = String(it.codFornecedor).trim();
          }
          return item;
        }),
      };
      // Quando separado por empresa, grava o nome nos campos de observação
      if (ped.empresa) {
        payload.observacoes         = ped.empresa;
        payload.observacoesInternas = ped.empresa;
      }

      Logger.log('Payload pedido (%s): %s', ped.fornecedorNome, JSON.stringify(payload).slice(0, 800));
      const res      = _pc_blingPost('/pedidos/compras', payload);
      const idPedido = res.data?.id || res.id || '?';
      const numero   = res.data?.numero || res.data?.numeroOrdem || '';
      resultados.push({ ok: true, fornecedor: ped.fornecedorNome, idPedido, numero, numItens: ped.itens.length });
    } catch(e) {
      resultados.push({ ok: false, fornecedor: ped.fornecedorNome, error: e.message });
    }
    Utilities.sleep(300);
  }
  return { resultados };
}
