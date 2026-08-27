// ============================================================
// PEDIDOS DE COMPRA — Web App autônomo
//
// Deploy: Extensões → Apps Script → Implantar → Nova implantação
//   Tipo: App da Web | Executar como: Eu | Acesso: Organização
//
// Propriedades do Script necessárias:
//   BLING_API_TOKEN — Bearer token OAuth 2 da API Bling v3
// ============================================================

const PC_SS_COMPRAS_ID  = '1GG6EenKiOO1K8XN0JrnD8W_PicQxzxkhsyDrU4yJa4w';
const PC_SS_CONTROLE_ID = '1eNMjC-iGBCAdVbJYSnP_Y4p7hJAzF6Gx';
const PC_ABA_COMPRAS    = 'Com Desmembramento';
const PC_GID_CONTROLE   = 1005476335;

// Índices 0-based em getValues()
const PC_IDX_SKU    = 0;   // col A — Base Compras
const PC_IDX_FORNEC = 54;  // col BC
const PC_IDX_QTD    = 56;  // col BE

const PC_CTRL_SKU   = 11;  // col L — Controle de Compras
const PC_CTRL_PREC  = 16;  // col Q

// ── Web App ──────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('PedidosCompra')
    .setTitle('Pedidos de Compra')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Helpers internos ─────────────────────────────────────────
function _pc_abaByGid(ss, gid) {
  return ss.getSheets().find(s => s.getSheetId() === gid) || ss.getSheets()[0];
}

function _pc_blingReq(method, path, params, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('BLING_API_TOKEN');
  if (!token) throw new Error('BLING_API_TOKEN não configurado nas Propriedades do Script');

  let url = 'https://www.bling.com.br/Api/v3' + path;
  if (params && Object.keys(params).length) {
    url += '?' + Object.entries(params)
      .map(([k, v]) => encodeURIComponent(k) + '=' + encodeURIComponent(v))
      .join('&');
  }

  const opts = {
    method: method.toLowerCase(),
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
    muteHttpExceptions: true,
  };
  if (payload) {
    opts.headers['Content-Type'] = 'application/json';
    opts.payload = JSON.stringify(payload);
  }

  const res  = UrlFetchApp.fetch(url, opts);
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');
  if (code < 200 || code >= 300) {
    const msg = body?.error?.message || body?.errors?.[0]?.msg || JSON.stringify(body).slice(0, 300);
    throw new Error('Bling HTTP ' + code + ': ' + msg);
  }
  return body;
}

// ── Carregar dados das planilhas ─────────────────────────────
function pc_carregarDados() {
  try {
    const ssCompras  = SpreadsheetApp.openById(PC_SS_COMPRAS_ID);
    const abaCompras = ssCompras.getSheetByName(PC_ABA_COMPRAS);
    if (!abaCompras) throw new Error('Aba "' + PC_ABA_COMPRAS + '" não encontrada');

    const lastRow = abaCompras.getLastRow();
    if (lastRow < 2) return { ok: true, fornecedores: {}, precoMap: {} };

    const raw = abaCompras.getRange(2, 1, lastRow - 1, 57).getValues();

    // Último preço por SKU no Controle de Compras
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

    // Agrupar por fornecedor — apenas linhas com BE > 0
    const fornecedores = {};
    raw.forEach(r => {
      const sku  = String(r[PC_IDX_SKU]    || '').trim();
      const forn = String(r[PC_IDX_FORNEC] || '').trim();
      const qtd  = Number(r[PC_IDX_QTD]   || 0);
      if (!sku || !forn || !(qtd > 0)) return;
      if (!fornecedores[forn]) fornecedores[forn] = [];
      fornecedores[forn].push({ sku, qtd, preco: precoMap[sku] || 0 });
    });

    return { ok: true, fornecedores, precoMap };
  } catch(e) {
    return { ok: false, error: e.message + '\n' + (e.stack || '') };
  }
}

// ── Buscar nomes e IDs de produtos no Bling ──────────────────
function pc_buscarProdutosBling(skus) {
  const produtos = {};
  for (const sku of skus) {
    try {
      const res  = _pc_blingReq('GET', '/produtos', { codigo: sku, pagina: 1, limite: 1 });
      const prod = (res.data || [])[0];
      produtos[sku] = prod ? { id: prod.id, nome: prod.nome || prod.descricao || sku } : null;
    } catch(e) {
      produtos[sku] = null;
    }
    Utilities.sleep(180);
  }
  return { ok: true, produtos };
}

// ── Buscar fornecedores cadastrados no Bling ─────────────────
function pc_buscarFornecedoresBling() {
  try {
    const todos  = [];
    let pagina   = 1;
    let continua = true;
    while (continua) {
      // criterio=4 → fornecedor no Bling API v3
      const res   = _pc_blingReq('GET', '/contatos', { pagina, limite: 100, criterio: 4 });
      const items = res.data || [];
      items.forEach(c => {
        const nome = (c.nome || c.fantasia || '').trim();
        if (c.id && nome) todos.push({ id: c.id, nome });
      });
      continua = items.length === 100;
      pagina++;
      if (continua) Utilities.sleep(200);
    }
    todos.sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' }));
    return { ok: true, fornecedores: todos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Criar pedidos no Bling ───────────────────────────────────
// pedidos = [{
//   fornecedorId: number, fornecedorNome: string,
//   itens: [{produtoId, sku, codFornecedor, qtd, preco, ipi}]
// }]
function pc_criarPedidos(pedidos) {
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const resultados = [];

  for (const ped of pedidos) {
    try {
      if (!ped.fornecedorId) throw new Error('Fornecedor não selecionado no Bling');

      const payload = {
        data: hoje,
        fornecedor: { id: Number(ped.fornecedorId) },
        itens: ped.itens.map(it => {
          const item = {
            produto:    { id: Number(it.produtoId) },
            quantidade: Number(it.qtd)   || 0,
            valor:      Number(it.preco) || 0,
            ipi:        Number(it.ipi)   || 0,
          };
          if (it.codFornecedor) item.codigo = String(it.codFornecedor);
          return item;
        }),
        situacao: { id: 1 }, // Em aberto / Orçamento
      };

      const res = _pc_blingReq('POST', '/pedidos/compras', null, payload);
      const idPedido = res.data?.id || res.id || '?';
      resultados.push({
        ok: true,
        fornecedor: ped.fornecedorNome,
        idPedido: idPedido,
        numItens: ped.itens.length,
      });
    } catch(e) {
      resultados.push({ ok: false, fornecedor: ped.fornecedorNome, error: e.message });
    }
    Utilities.sleep(300);
  }

  return { resultados };
}
