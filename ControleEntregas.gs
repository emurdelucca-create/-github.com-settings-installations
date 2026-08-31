// ============================================================
// CONTROLE DE ENTREGAS — Web App
// Planilha: 1RrDp-TBmNPVAw9ffixgxR0GQAtb5cih0FEeASnBY6GU
// ============================================================

const CE_SS_ID    = '1RrDp-TBmNPVAw9ffixgxR0GQAtb5cih0FEeASnBY6GU';
const CE_ABA_NFS  = 'NFs';

// CNPJ dos destinatários (sem pontuação)
const CE_EMPRESAS_CNPJ = {
  '43339573000157': 'HUMBLE',
  '44373737000125': 'NAJUMI',
  '49640848000174': 'SKY',
  '54518222000120': 'EDMOTOS',
  '60743800000124': 'MOTO VIBE',
};

// BaseLinker — SKUs
const CE_SS_BL_ID   = '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik';
const CE_BL_SKU_COL = 1;  // col B (0-indexed)
const CE_BL_NOM_COL = 7;  // col H (0-indexed)

// Controle de Compras — mapa cód.fornecedor → SKU
const CE_CTRL_SS_ID     = '1eNMjC-iGBCAdVbJYSnP_Y4p7hJAzF6Gx';
const CE_CTRL_GID       = 1005476335;
const CE_CTRL_COD_FORN  = 10;  // col K (0-indexed)
const CE_CTRL_SKU       = 11;  // col L (0-indexed)

// Bling
const CE_BLING_TOKEN_URL = 'https://www.bling.com.br/Api/v3/oauth/token';
const CE_BLING_API       = 'https://www.bling.com.br/Api/v3';
const CE_BLING_REDIRECT  = 'https://www.google.com';

// Colunas da sheet NFs (1-indexed para GAS getRange)
// A=ID  B=PedidoBling  C=Fornecedor  D=NNF  E=Empresa  F=QtdVolume
// G=DataEntrega  H=LancouContas  I=ImportouXML  J=Responsavel
// K=Status  L=DataCriacao  M=ItensJSON
const CE_NCOLS = 13;

// ── Web App ──────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('ControleEntregas.html')
    .setTitle('Controle de Entregas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Script Properties ────────────────────────────────────────
function _ce_props() { return PropertiesService.getScriptProperties(); }

// ── Planilha ─────────────────────────────────────────────────
function _ce_ss() { return SpreadsheetApp.openById(CE_SS_ID); }

function _ce_aba() {
  const ss  = _ce_ss();
  let   aba = ss.getSheetByName(CE_ABA_NFS);
  if (!aba) {
    aba = ss.insertSheet(CE_ABA_NFS);
    aba.getRange(1, 1, 1, CE_NCOLS).setValues([[
      'ID','Pedido Bling','Fornecedor','Nº NF','Empresa Nossa',
      'Qtd. Volume NF','Data Entrega Agendada','Lançou Contas?',
      'Importou XML Bling?','Responsável','Status','Data Criação','Itens JSON',
    ]]);
    aba.setFrozenRows(1);
  }
  return aba;
}

function _ce_json(s) {
  try { return JSON.parse(s || '[]'); } catch(e) { return []; }
}

function _ce_rowToObj(r) {
  return {
    id:          String(r[0]  || ''),
    pedidoBling: String(r[1]  || ''),
    fornecedor:  String(r[2]  || ''),
    nNF:         String(r[3]  || ''),
    empresa:     String(r[4]  || ''),
    qtdVolume:   String(r[5]  || ''),
    dataEntrega: r[6] ? (r[6] instanceof Date
                  ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'dd/MM/yyyy')
                  : String(r[6])) : '',
    lancouContas:String(r[7]  || 'Não'),
    importouXML: String(r[8]  || 'Não'),
    responsavel: String(r[9]  || ''),
    status:      String(r[10] || 'Aguardando Chegar'),
    dataCriacao: r[11] ? (r[11] instanceof Date
                  ? Utilities.formatDate(r[11], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
                  : String(r[11])) : '',
    itens:       _ce_json(r[12]),
  };
}

// Salva nova NF
function ce_salvarNF(dados) {
  try {
    const aba   = _ce_aba();
    const id    = 'NF_' + Date.now();
    const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    aba.appendRow([
      id,
      String(dados.pedidoBling  || ''),
      String(dados.fornecedor   || ''),
      String(dados.nNF          || ''),
      String(dados.empresa      || ''),
      String(dados.qtdVolume    || ''),
      String(dados.dataEntrega  || ''),
      String(dados.lancouContas || 'Não'),
      String(dados.importouXML  || 'Não'),
      String(dados.responsavel  || ''),
      String(dados.status       || 'Aguardando Chegar'),
      agora,
      JSON.stringify(dados.itens || []),
    ]);
    return { ok: true, id };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Carrega todas as NFs
function ce_carregarNFs() {
  try {
    const aba  = _ce_aba();
    const last = aba.getLastRow();
    if (last < 2) return { ok: true, nfs: [] };
    const raw = aba.getRange(2, 1, last - 1, CE_NCOLS).getValues();
    const nfs = raw.filter(r => r[0]).map(_ce_rowToObj);
    return { ok: true, nfs };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Atualiza um campo de uma NF existente
function ce_atualizarCampo(id, campo, valor) {
  try {
    const MAPA = {
      pedidoBling:  2, fornecedor: 3, nNF: 4, empresa: 5,
      qtdVolume:    6, dataEntrega: 7, lancouContas: 8,
      importouXML:  9, responsavel: 10, status: 11, itens: 13,
    };
    const col = MAPA[campo];
    if (!col) return { ok: false, error: 'Campo desconhecido: ' + campo };

    const aba  = _ce_aba();
    const last = aba.getLastRow();
    if (last < 2) return { ok: false, error: 'Sem dados' };

    const ids = aba.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
    const idx = ids.indexOf(String(id));
    if (idx === -1) return { ok: false, error: 'NF não encontrada' };

    const val = (campo === 'itens') ? JSON.stringify(valor) : String(valor);
    aba.getRange(idx + 2, col).setValue(val);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Bling OAuth ───────────────────────────────────────────────
function _ce_getToken() {
  const p       = _ce_props();
  const access  = p.getProperty('BLING_ACCESS_TOKEN');
  const refresh = p.getProperty('BLING_REFRESH_TOKEN');
  const expires = parseInt(p.getProperty('BLING_TOKEN_EXPIRES') || '0');
  if (!refresh) throw new Error('Bling não autorizado');
  if (!access || Date.now() >= expires - 300000) return _ce_refreshToken(refresh);
  return access;
}

function _ce_refreshToken(rt) {
  const p  = _ce_props();
  const ci = p.getProperty('BLING_CLIENT_ID');
  const cs = p.getProperty('BLING_CLIENT_SECRET');
  if (!ci || !cs) throw new Error('BLING_CLIENT_ID/BLING_CLIENT_SECRET não configurados');
  const cred = Utilities.base64Encode(ci + ':' + cs);
  const res  = UrlFetchApp.fetch(CE_BLING_TOKEN_URL, {
    method: 'POST',
    headers: { Authorization: 'Basic ' + cred, 'Content-Type': 'application/x-www-form-urlencoded' },
    payload: 'grant_type=refresh_token&refresh_token=' + encodeURIComponent(rt),
    muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('Refresh falhou: ' + (body.error_description || JSON.stringify(body)));
  p.setProperties({
    BLING_ACCESS_TOKEN:  body.access_token,
    BLING_REFRESH_TOKEN: body.refresh_token || rt,
    BLING_TOKEN_EXPIRES: String(Date.now() + (body.expires_in || 3600) * 1000),
  });
  return body.access_token;
}

function ce_getAuthUrl() {
  const ci = _ce_props().getProperty('BLING_CLIENT_ID');
  if (!ci) return { ok: false, error: 'BLING_CLIENT_ID não configurado nas Propriedades do Script' };
  const url = 'https://www.bling.com.br/Api/v3/oauth/authorize?response_type=code' +
    '&client_id=' + encodeURIComponent(ci) +
    '&redirect_uri=' + encodeURIComponent(CE_BLING_REDIRECT) +
    '&state=' + Date.now();
  return { ok: true, url };
}

function ce_trocarCodigo(rawInput) {
  try {
    let code = String(rawInput || '').trim();
    const m  = code.match(/[?&]code=([^&]+)/);
    if (m) code = decodeURIComponent(m[1]);
    const p  = _ce_props();
    const ci = p.getProperty('BLING_CLIENT_ID');
    const cs = p.getProperty('BLING_CLIENT_SECRET');
    if (!ci || !cs) throw new Error('BLING_CLIENT_ID/BLING_CLIENT_SECRET não configurados');
    const cred = Utilities.base64Encode(ci + ':' + cs);
    const res  = UrlFetchApp.fetch(CE_BLING_TOKEN_URL, {
      method: 'POST',
      headers: { Authorization: 'Basic ' + cred, 'Content-Type': 'application/x-www-form-urlencoded' },
      payload: 'grant_type=authorization_code&code=' + encodeURIComponent(code) + '&redirect_uri=' + encodeURIComponent(CE_BLING_REDIRECT),
      muteHttpExceptions: true,
    });
    const body = JSON.parse(res.getContentText() || '{}');
    if (res.getResponseCode() !== 200) throw new Error(body.error_description || JSON.stringify(body));
    p.setProperties({
      BLING_ACCESS_TOKEN:  body.access_token,
      BLING_REFRESH_TOKEN: body.refresh_token,
      BLING_TOKEN_EXPIRES: String(Date.now() + (body.expires_in || 3600) * 1000),
    });
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function ce_checkAuth() {
  return { autorizado: !!_ce_props().getProperty('BLING_REFRESH_TOKEN') };
}

// ── Bling — Buscar pedido de compra por número ────────────────
function ce_buscarPedidoBling(numeroPedido) {
  try {
    const token = _ce_getToken();
    const numero = String(numeroPedido).trim();

    // Tenta busca direta por número
    const r1 = UrlFetchApp.fetch(CE_BLING_API + '/pedidos/compras?numero=' + encodeURIComponent(numero) + '&pagina=1&limite=50', {
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
      muteHttpExceptions: true,
    });
    let pedidos = (JSON.parse(r1.getContentText() || '{}').data || []);

    // Se não encontrou direto, percorre páginas (até 10) buscando pelo número
    if (!pedidos.length) {
      for (let pg = 1; pg <= 10 && !pedidos.length; pg++) {
        Utilities.sleep(200);
        const rp = UrlFetchApp.fetch(CE_BLING_API + '/pedidos/compras?pagina=' + pg + '&limite=100', {
          headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
          muteHttpExceptions: true,
        });
        const items = JSON.parse(rp.getContentText() || '{}').data || [];
        if (!items.length) break;
        pedidos = items.filter(p => String(p.numero || p.numeroOrdem || '') === numero);
      }
    }

    if (!pedidos.length) return { ok: false, error: 'Pedido nº ' + numero + ' não encontrado no Bling' };

    // Busca detalhes do primeiro resultado
    const ped = pedidos[0];
    Utilities.sleep(200);
    const r2 = UrlFetchApp.fetch(CE_BLING_API + '/pedidos/compras/' + ped.id, {
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
      muteHttpExceptions: true,
    });
    const det = (JSON.parse(r2.getContentText() || '{}').data) || {};

    const itens = (det.itens || []).map(it => ({
      produtoId: String(it.produto?.id    || ''),
      sku:       String(it.produto?.codigo || ''),
      descricao: String(it.descricao       || ''),
      codForn:   String(it.codigo          || ''),
      qtd:       Number(it.quantidade      || 0),
      preco:     Number(it.valor           || 0),
    }));

    return {
      ok: true,
      pedido: { id: String(ped.id), numero, fornecedor: det.fornecedor?.nome || '' },
      itens,
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── BaseLinker — SKUs ─────────────────────────────────────────
function ce_buscarSKUsBL() {
  try {
    const ss   = SpreadsheetApp.openById(CE_SS_BL_ID);
    const aba  = ss.getSheets()[0];
    const last = aba.getLastRow();
    if (last < 2) return { ok: true, skus: [] };
    const ncols = Math.max(CE_BL_SKU_COL, CE_BL_NOM_COL) + 1;
    const raw   = aba.getRange(2, 1, last - 1, ncols).getValues();
    const skus  = [];
    raw.forEach(r => {
      const sku  = String(r[CE_BL_SKU_COL] || '').trim();
      const nome = String(r[CE_BL_NOM_COL]  || '').trim();
      if (sku && nome) skus.push({ sku, nome });
    });
    return { ok: true, skus };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Re-sincroniza qtdPedido de todas as NFs que ainda têm itens com qtdPedido=0
function ce_ressincronizarQtds() {
  try {
    const aba  = _ce_aba();
    const last = aba.getLastRow();
    if (last < 2) return { ok: true, atualizadas: 0 };

    const raw = aba.getRange(2, 1, last - 1, CE_NCOLS).getValues();

    // Mapa codForn → SKU do Controle de Compras
    let codFornMap = {};
    try { const r = ce_buscarMapCodFornSKU(); if (r.ok) codFornMap = r.map; } catch(e) {}

    let atualizadas = 0;

    for (let i = 0; i < raw.length; i++) {
      const r = raw[i];
      if (!r[0]) continue;
      const pedidoBling = String(r[1] || '').trim();
      if (!pedidoBling) continue;

      const itens = _ce_json(r[12]);
      if (!itens.length) continue;
      if (!itens.some(it => !it.qtdPedido)) continue; // todos já preenchidos

      try {
        Utilities.sleep(300);
        const res = ce_buscarPedidoBling(pedidoBling);
        if (!res.ok || !res.itens.length) continue;

        const bmap = {};
        res.itens.forEach(bi => {
          if (bi.codForn) bmap[bi.codForn.trim().toLowerCase()] = bi;
          if (bi.sku)     bmap[bi.sku.trim().toLowerCase()]     = bi;
        });

        let mudou = false;
        const novos = itens.map((it, idx) => {
          if (it.qtdPedido) return it;
          const key = (it.cProd || '').trim().toLowerCase();
          let bi = bmap[key] || null;
          if (!bi) {
            const sk = codFornMap[(it.cProd||'').trim()] ||
                       codFornMap[(it.cProd||'').trim().toUpperCase()];
            if (sk) bi = bmap[sk.trim().toLowerCase()] || null;
          }
          if (!bi && res.itens[idx]) bi = res.itens[idx];
          if (bi && bi.qtd) { mudou = true; return { ...it, qtdPedido: bi.qtd }; }
          return it;
        });

        if (mudou) {
          aba.getRange(i + 2, 13).setValue(JSON.stringify(novos));
          atualizadas++;
        }
      } catch(e) { continue; }
    }

    return { ok: true, atualizadas };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Mapa CNPJ → empresa (para uso no frontend via JSON)
function ce_getEmpresas() {
  return { ok: true, empresas: CE_EMPRESAS_CNPJ };
}

// Mapa cód.fornecedor → SKU (lê planilha Controle de Compras, cols K e L)
function ce_buscarMapCodFornSKU() {
  try {
    const ss  = SpreadsheetApp.openById(CE_CTRL_SS_ID);
    const aba = ss.getSheets().find(s => s.getSheetId() === CE_CTRL_GID);
    if (!aba) return { ok: true, map: {} };
    const last = aba.getLastRow();
    if (last < 2) return { ok: true, map: {} };
    const ncols = CE_CTRL_SKU + 1;
    const raw   = aba.getRange(2, 1, last - 1, ncols).getValues();
    const map   = {};
    raw.forEach(r => {
      const cod = String(r[CE_CTRL_COD_FORN] || '').trim();
      const sku = String(r[CE_CTRL_SKU]       || '').trim();
      if (cod && sku) map[cod] = sku;
    });
    return { ok: true, map };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
