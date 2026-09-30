const $ = id => document.getElementById(id);

async function carregarConfig() {
  const { webAppUrl, token } = await chrome.storage.local.get(['webAppUrl', 'token']);
  if (webAppUrl) $('webAppUrl').value = webAppUrl;
  if (token) $('token').value = token;
}

async function salvarConfig() {
  await chrome.storage.local.set({
    webAppUrl: $('webAppUrl').value.trim(),
    token: $('token').value.trim(),
  });
}

function log(msg, cls) {
  const el = $('status');
  el.textContent = msg;
  el.className = cls || '';
}

// Injetada dentro da aba de produto da Shopee (contexto de página de
// verdade — é isso que evita o bloqueio anti-bot que a chamada direta do
// servidor do Apps Script sofre). Precisa ser autossuficiente: o Chrome
// serializa essa função e reexecuta dentro da página, sem acesso a nada
// do escopo do popup.js.
async function capturarProdutoShopee(itemId, shopId) {
  try {
    const resp = await fetch(
      'https://shopee.com.br/api/v4/item/get?itemid=' + itemId + '&shopid=' + shopId,
      { headers: { 'Accept': 'application/json' }, credentials: 'include' }
    );
    const data = await resp.json();
    const item = data.item || (data.data && data.data.item) || null;
    if (!item) return { ok: false, error: 'Resposta sem campo "item".', raw: data };

    // Preço: tenta o formato com variações (models[]) primeiro; cai pro
    // formato "produto único" se não tiver variação. Campos de preço da
    // Shopee vêm multiplicados por 100000 (ex: 15390000 = R$153,90).
    let bruto = null, final = null;
    const modelos = Array.isArray(item.models) && item.models.length ? item.models : null;
    if (modelos) {
      modelos.forEach(m => {
        const atual  = (m.price != null ? m.price : m.current_price) / 100000;
        const origem = (m.price_before_discount != null ? m.price_before_discount : (m.original_price != null ? m.original_price : m.price)) / 100000;
        if (bruto === null || origem > bruto) bruto = origem;
        if (final === null || atual  < final) final = atual;
      });
    } else {
      final = (item.price_min != null ? item.price_min : item.price) / 100000;
      const origBase = item.price_before_discount_max != null ? item.price_before_discount_max
        : (item.price_before_discount != null ? item.price_before_discount : item.price);
      bruto = origBase / 100000;
    }

    const vendidoAcumulado = item.historical_sold != null ? item.historical_sold : (item.sold || 0);

    return { ok: true, precoBruto: bruto, precoFinal: final, vendidoAcumulado, rawDebug: JSON.stringify(item).substring(0, 500) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Abre o link em segundo plano, espera carregar, injeta a captura, e
// fecha a aba — tudo sem tirar o foco do que o usuário está fazendo.
function esperarAbaCarregar(tabId) {
  return new Promise(resolve => {
    const timeout = setTimeout(() => { chrome.tabs.onUpdated.removeListener(listener); resolve(); }, 15000);
    function listener(id, info) {
      if (id === tabId && info.status === 'complete') {
        clearTimeout(timeout);
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    }
    chrome.tabs.onUpdated.addListener(listener);
  });
}

async function capturarItem(item) {
  // Aba em PRIMEIRO plano (active:true) — a Shopee parece tratar aba em
  // segundo plano como suspeita (mesmo erro de bloqueio anti-bot do
  // servidor apareceu numa aba invisível). Rouba o foco por um instante,
  // mas evita o bloqueio.
  const tab = await chrome.tabs.create({ url: item.link, active: true });
  try {
    await esperarAbaCarregar(tab.id);
    await new Promise(res => setTimeout(res, 1500)); // dá tempo do JS da página rodar de verdade
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: capturarProdutoShopee,
      args: [item.itemId, item.shopId],
    });
    return result;
  } finally {
    chrome.tabs.remove(tab.id).catch(() => {});
  }
}

$('salvar').addEventListener('click', async () => {
  await salvarConfig();
  log('✅ Configuração salva.', 'ok');
});

$('capturar').addEventListener('click', async () => {
  const webAppUrl = $('webAppUrl').value.trim();
  const token = $('token').value.trim();
  if (!webAppUrl || !token) {
    log('⚠ Preencha a URL do Web App e o Token antes de continuar.', 'err');
    return;
  }
  await salvarConfig();

  const btn = $('capturar');
  btn.disabled = true;
  $('debug').style.display = 'none';

  try {
    log('📋 Buscando lista de concorrentes na planilha...', '');
    const respLista = await fetch(webAppUrl + '?token=' + encodeURIComponent(token));
    const lista = await respLista.json();
    if (!lista.ok) throw new Error(lista.error || 'Falha ao buscar a lista.');

    const itens = lista.itens || [];
    if (!itens.length) throw new Error('Nenhum concorrente cadastrado (preencha a coluna "Link Anúncio" na planilha).');

    const capturas = [];
    let primeiroDebug = '';
    for (let i = 0; i < itens.length; i++) {
      log('🔍 Capturando ' + (i + 1) + '/' + itens.length + '... (' + itens[i].itemId + ')', '');
      const r = await capturarItem(itens[i]);
      if (r && r.ok) {
        capturas.push({ itemId: itens[i].itemId, shopId: itens[i].shopId, precoBruto: r.precoBruto, precoFinal: r.precoFinal, vendidoAcumulado: r.vendidoAcumulado });
        if (!primeiroDebug) primeiroDebug = r.rawDebug || '';
      } else {
        console.warn('Falha ao capturar', itens[i], r);
      }
      await new Promise(res => setTimeout(res, 400)); // folga entre aberturas de aba
    }

    if (!capturas.length) throw new Error('Nenhuma captura teve sucesso — veja o console da extensão (botão direito no popup → Inspecionar) para detalhes.');

    log('📤 Enviando ' + capturas.length + ' captura(s) pra planilha...', '');
    const respEnvio = await fetch(webAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ token, capturas }),
    });
    const dataEnvio = await respEnvio.json();
    if (!dataEnvio.ok) throw new Error(dataEnvio.error || 'Erro ao gravar na planilha.');

    log('✅ ' + capturas.length + '/' + itens.length + ' concorrente(s) capturado(s) e gravado(s) na planilha!', 'ok');
    if (primeiroDebug) {
      $('debug').textContent = 'Debug (1º item, JSON bruto cortado):\n' + primeiroDebug;
      $('debug').style.display = 'block';
    }
  } catch (e) {
    log('❌ ' + e.message, 'err');
  } finally {
    btn.disabled = false;
  }
});

carregarConfig();
