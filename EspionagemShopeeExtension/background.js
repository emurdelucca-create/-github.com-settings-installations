// Service worker — roda a captura de verdade, sobrevivendo ao popup
// fechar (o Chrome fecha o popup sozinho sempre que ele perde o foco,
// o que acontece justamente quando abrimos a aba do concorrente em
// primeiro plano). O progresso vai sendo salvo em chrome.storage.local,
// e o popup só lê esse status — nunca faz o trabalho ele mesmo.

async function salvarStatus(status) {
  await chrome.storage.local.set({ espStatus: status });
}

// Injetada dentro da aba de produto da Shopee (contexto de página de
// verdade — é isso que evita o bloqueio anti-bot que a chamada direta do
// servidor do Apps Script sofre). Precisa ser autossuficiente: o Chrome
// serializa essa função e reexecuta dentro da página.
async function capturarProdutoShopee(itemId, shopId) {
  try {
    const resp = await fetch(
      'https://shopee.com.br/api/v4/item/get?itemid=' + itemId + '&shopid=' + shopId,
      { headers: { 'Accept': 'application/json' }, credentials: 'include' }
    );
    const data = await resp.json();
    const item = data.item || (data.data && data.data.item) || null;
    if (!item) return { ok: false, error: 'Resposta sem campo "item".', raw: data };

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
  const tab = await chrome.tabs.create({ url: item.link, active: true });
  try {
    await esperarAbaCarregar(tab.id);
    await new Promise(res => setTimeout(res, 1500));
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

async function rodarCaptura(webAppUrl, token) {
  await salvarStatus({ fase: 'rodando', mensagem: '📋 Buscando lista de concorrentes na planilha...' });

  try {
    const respLista = await fetch(webAppUrl + '?token=' + encodeURIComponent(token));
    const lista = await respLista.json();
    if (!lista.ok) throw new Error(lista.error || 'Falha ao buscar a lista.');

    const itens = lista.itens || [];
    if (!itens.length) throw new Error('Nenhum concorrente cadastrado (preencha a coluna "Link Anúncio" na planilha).');

    const capturas = [];
    let primeiroDebug = '';
    for (let i = 0; i < itens.length; i++) {
      await salvarStatus({ fase: 'rodando', mensagem: '🔍 Capturando ' + (i + 1) + '/' + itens.length + '... (' + itens[i].itemId + ')' });
      const r = await capturarItem(itens[i]);
      if (r && r.ok) {
        capturas.push({ itemId: itens[i].itemId, shopId: itens[i].shopId, precoBruto: r.precoBruto, precoFinal: r.precoFinal, vendidoAcumulado: r.vendidoAcumulado });
        if (!primeiroDebug) primeiroDebug = r.rawDebug || '';
      } else {
        console.warn('Falha ao capturar', itens[i], r);
        if (!primeiroDebug && r) primeiroDebug = 'ERRO: ' + (r.error || '') + '\n' + JSON.stringify(r.raw || {}).substring(0, 500);
      }
      await new Promise(res => setTimeout(res, 400));
    }

    if (!capturas.length) {
      await salvarStatus({ fase: 'erro', mensagem: '❌ Nenhuma captura teve sucesso.', debug: primeiroDebug });
      return;
    }

    await salvarStatus({ fase: 'rodando', mensagem: '📤 Enviando ' + capturas.length + ' captura(s) pra planilha...' });
    const respEnvio = await fetch(webAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ token, capturas }),
    });
    const dataEnvio = await respEnvio.json();
    if (!dataEnvio.ok) throw new Error(dataEnvio.error || 'Erro ao gravar na planilha.');

    await salvarStatus({
      fase: 'ok',
      mensagem: '✅ ' + capturas.length + '/' + itens.length + ' concorrente(s) capturado(s) e gravado(s) na planilha!',
      debug: primeiroDebug,
    });
  } catch (e) {
    await salvarStatus({ fase: 'erro', mensagem: '❌ ' + e.message });
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'iniciarCaptura') {
    rodarCaptura(msg.webAppUrl, msg.token);
    sendResponse({ ok: true }); // só confirma que começou — o resultado vem via storage
  }
  return true;
});
