// ============================================================
// PEDIDOS DE COMPRA — Web App autônomo
//
// Deploy: Extensões → Apps Script → Implantar → Nova implantação
//   Tipo: App da Web | Executar como: Eu | Acesso: Organização
//
// Propriedades do Script necessárias:
//   BLING_API_KEY — API key do usuário API (Bling → Usuários → API key)
//                   Essa chave não expira, diferente do token OAuth v3.
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
  return HtmlService.createHtmlOutputFromFile('PedidosCompra.html')
    .setTitle('Pedidos de Compra')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Helpers internos ─────────────────────────────────────────
function _pc_abaByGid(ss, gid) {
  return ss.getSheets().find(s => s.getSheetId() === gid) || ss.getSheets()[0];
}

function _pc_apiKey() {
  const k = PropertiesService.getScriptProperties().getProperty('BLING_API_KEY');
  if (!k) throw new Error('BLING_API_KEY não configurado nas Propriedades do Script');
  return k;
}

// GET para a API v2 do Bling — retorna body.retorno
function _pc_blingGet(path, params) {
  const qs = Object.entries(Object.assign({ apikey: _pc_apiKey() }, params || {}))
    .map(([k, v]) => encodeURIComponent(k) + '=' + encodeURIComponent(v))
    .join('&');

  const res  = UrlFetchApp.fetch('https://bling.com.br/Api/v2' + path + '?' + qs,
    { muteHttpExceptions: true });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');

  if (body.retorno?.erros) {
    const e = body.retorno.erros;
    throw new Error('Bling: ' + (Array.isArray(e) ? e.map(x => x.erro || JSON.stringify(x)).join('; ') : JSON.stringify(e)));
  }
  if (code < 200 || code >= 300) throw new Error('Bling HTTP ' + code);
  return body.retorno;
}

// POST para a API v2 do Bling — payload é uma string XML
function _pc_blingPost(path, xml) {
  const url = 'https://bling.com.br/Api/v2' + path + '?apikey=' + encodeURIComponent(_pc_apiKey());
  const res  = UrlFetchApp.fetch(url, {
    method:      'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload:     'xml=' + encodeURIComponent(xml),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');

  if (body.retorno?.erros) {
    const e = body.retorno.erros;
    throw new Error('Bling: ' + (Array.isArray(e) ? e.map(x => x.erro || JSON.stringify(x)).join('; ') : JSON.stringify(e)));
  }
  if (code < 200 || code >= 300) throw new Error('Bling HTTP ' + code);
  return body.retorno;
}

function _pc_escXml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── Carregar dados das planilhas ─────────────────────────────
function pc_carregarDados() {
  try {
    const ssCompras  = SpreadsheetApp.openById(PC_SS_COMPRAS_ID);
    const abaCompras = ssCompras.getSheetByName(PC_ABA_COMPRAS);
    if (!abaCompras) throw new Error('Aba "' + PC_ABA_COMPRAS + '" não encontrada');

    const lastRow = abaCompras.getLastRow();
    if (lastRow < 2) return { ok: true, fornecedores: {} };

    const raw = abaCompras.getRange(2, 1, lastRow - 1, 57).getValues();

    // Último preço por SKU — Controle de Compras col Q
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

    return { ok: true, fornecedores };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Buscar nomes e IDs de produtos no Bling (API v2) ─────────
function pc_buscarProdutosBling(skus) {
  const produtos = {};
  for (const sku of skus) {
    try {
      const ret  = _pc_blingGet('/produtos/json/', { codigo: sku });
      const list = ret?.produtos || [];
      // v2 pode retornar produto com match parcial — filtra pelo código exato
      const match = list.find(p => String(p.produto?.codigo || '').trim().toUpperCase() === sku.toUpperCase());
      const prod  = match ? match.produto : (list[0]?.produto || null);
      produtos[sku] = prod ? { id: String(prod.id), nome: prod.descricao || sku } : null;
    } catch(e) {
      produtos[sku] = null;
    }
    Utilities.sleep(220);
  }
  return { ok: true, produtos };
}

// ── Buscar fornecedores cadastrados no Bling (API v2) ─────────
function pc_buscarFornecedoresBling() {
  try {
    const todos  = [];
    let pagina   = 1;
    let continua = true;
    while (continua) {
      // tipo[]=F → somente fornecedores
      const ret  = _pc_blingGet('/contatos/json/', { 'tipo[]': 'F', pagina });
      const list = ret?.contatos || [];
      list.forEach(item => {
        const c = item.contato;
        if (!c) return;
        const nome = (c.nome || c.fantasia || '').trim();
        if (c.id && nome) todos.push({ id: String(c.id), nome });
      });
      // v2 retorna até 25 por página por padrão
      continua = list.length >= 25;
      pagina++;
      if (continua) Utilities.sleep(220);
    }
    todos.sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' }));
    return { ok: true, fornecedores: todos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Criar pedidos no Bling (API v2 — XML) ────────────────────
function pc_criarPedidos(pedidos) {
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const resultados = [];

  for (const ped of pedidos) {
    try {
      if (!ped.fornecedorId) throw new Error('Fornecedor não selecionado no Bling');

      const itensXml = ped.itens.map(it =>
        `<item>` +
        `<idproduto>${_pc_escXml(it.produtoId)}</idproduto>` +
        `<quantidade>${Number(it.qtd) || 0}</quantidade>` +
        `<valor>${Number(it.preco || 0).toFixed(2)}</valor>` +
        `<ipi>${Number(it.ipi || 0).toFixed(2)}</ipi>` +
        (it.codFornecedor ? `<codigofornecedor>${_pc_escXml(it.codFornecedor)}</codigofornecedor>` : '') +
        `</item>`
      ).join('');

      const xml =
        `<?xml version="1.0" encoding="UTF-8"?>` +
        `<pedidocompra>` +
        `<data>${hoje}</data>` +
        `<fornecedor><id>${_pc_escXml(ped.fornecedorId)}</id></fornecedor>` +
        `<itens>${itensXml}</itens>` +
        `</pedidocompra>`;

      const ret     = _pc_blingPost('/pedidocompra/json/', xml);
      const pedList = ret?.pedidocompra || [];
      const pedData = pedList[0]?.pedidocompra || pedList[0] || {};
      const idPed   = pedData.id     || '?';
      const numPed  = pedData.numero || pedData.numeroOrdem || '';

      resultados.push({
        ok:        true,
        fornecedor: ped.fornecedorNome,
        idPedido:  idPed,
        numero:    numPed,
        numItens:  ped.itens.length,
      });
    } catch(e) {
      resultados.push({ ok: false, fornecedor: ped.fornecedorNome, error: e.message });
    }
    Utilities.sleep(300);
  }

  return { resultados };
}
