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

const PC_CTRL_SKU   = 11;  // col L
const PC_CTRL_PREC  = 16;  // col Q

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
    throw new Error('Bling ' + code + ': ' + (body?.error?.message || JSON.stringify(body).slice(0, 250)));
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

    const raw = abaCompras.getRange(2, 1, lastRow - 1, 57).getValues();

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
      if (!fornecedores[forn]) fornecedores[forn] = [];
      fornecedores[forn].push({ sku, qtd, preco: precoMap[sku] || 0 });
    });

    return { ok: true, fornecedores };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Bling — Produtos ─────────────────────────────────────────
function pc_buscarProdutosBling(skus) {
  const produtos = {};
  for (const sku of skus) {
    try {
      const res  = _pc_blingGet('/produtos', { codigo: sku, pagina: 1, limite: 1 });
      const prod = (res.data || [])[0];
      produtos[sku] = prod ? { id: String(prod.id), nome: prod.nome || prod.descricao || sku } : null;
    } catch(e) {
      produtos[sku] = null;
    }
    Utilities.sleep(200);
  }
  return { ok: true, produtos };
}

// ── Bling — Fornecedores ─────────────────────────────────────
function pc_buscarFornecedoresBling() {
  try {
    const todos  = [];
    let pagina   = 1;
    let continua = true;
    while (continua) {
      const res   = _pc_blingGet('/contatos', { criterio: 4, pagina, limite: 100 });
      const items = res.data || [];
      items.forEach(c => {
        const nome = (c.nome || c.fantasia || '').trim();
        if (c.id && nome) todos.push({ id: String(c.id), nome });
      });
      continua = items.length === 100;
      pagina++;
      if (continua) Utilities.sleep(220);
    }
    todos.sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' }));
    return { ok: true, fornecedores: todos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Bling — Criar pedidos ────────────────────────────────────
function pc_criarPedidos(pedidos) {
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const resultados = [];

  for (const ped of pedidos) {
    try {
      if (!ped.fornecedorId) throw new Error('Fornecedor não selecionado');

      const payload = {
        data:       hoje,
        fornecedor: { id: Number(ped.fornecedorId) },
        situacao:   { id: 1 },
        itens: ped.itens.map(it => ({
          produto:    { id: Number(it.produtoId) },
          quantidade: Number(it.qtd)   || 0,
          valor:      Number(it.preco) || 0,
          ipi:        Number(it.ipi)   || 0,
          ...(it.codFornecedor ? { codigo: String(it.codFornecedor) } : {}),
        })),
      };

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
