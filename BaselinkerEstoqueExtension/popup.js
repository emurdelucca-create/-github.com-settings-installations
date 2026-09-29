// URL padrão do Web App já publicado (Visão Geral Separação). Pode ser
// sobrescrita no campo, caso o deployment mude no futuro.
const URL_PADRAO = 'https://script.google.com/macros/s/AKfycbzmYYhnemfE7-IAFQW_cvbFA4BDWbeXqmuJ-TNx4vPxc-KGl0JOdzITbF-1H7a_da4T/exec';

const $ = id => document.getElementById(id);

async function carregarConfig() {
  const { webAppUrl, token } = await chrome.storage.local.get(['webAppUrl', 'token']);
  $('webAppUrl').value = webAppUrl || URL_PADRAO;
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

// Injetada na aba do BaseLinker via chrome.scripting.executeScript — precisa
// ser autossuficiente (sem referenciar nada do escopo do popup.js), pois o
// Chrome serializa e reexecuta essa função dentro da página alvo.
async function scrapeInsuficientes() {
  function encontrarContainerRolavel(el) {
    let node = el;
    while (node && node !== document.body) {
      const style = getComputedStyle(node);
      if (node.scrollHeight > node.clientHeight + 10 && /(auto|scroll)/.test(style.overflowY)) {
        return node;
      }
      node = node.parentElement;
    }
    return document.scrollingElement || document.documentElement;
  }

  function coletarLinhas() {
    const linhas = Array.from(document.querySelectorAll('tr.pick_pack_product_row'));
    return linhas.map(row => {
      const skuEl = row.querySelector('.product-info-text');
      const skuMatch = skuEl && skuEl.textContent.match(/SKU:\s*(.+)/);
      // Produtos com variação vêm como "SKU: 25018-S | Atributos: Cor: Preto"
      // no mesmo texto — como nenhum SKU tem espaço, cortar no primeiro
      // espaço já isola o código real, descartando os atributos.
      const sku = skuMatch ? skuMatch[1].trim().split(' ')[0] : '';
      const qtyEl = row.querySelector('.pick_pack_product_total_quantity');
      const qty = qtyEl ? (parseInt(qtyEl.textContent.trim(), 10) || 0) : 0;
      const locEl = row.querySelector('.locations_container .lbl_location');
      const locValue = locEl ? (locEl.getAttribute('location_value') || '') : '';
      const insuficiente = locValue.trim() === 'Estoque insuficiente';
      return { sku, qty, insuficiente };
    }).filter(r => r.sku);
  }

  const primeiraLinha = document.querySelector('tr.pick_pack_product_row');
  if (!primeiraLinha) {
    return { ok: false, error: 'Nenhum item encontrado. Abra o painel "Coleta de pedidos" antes de clicar em Obter Dados.', itens: [] };
  }

  const container = encontrarContainerRolavel(primeiraLinha);

  // Rola até o fim repetidamente até a altura do container parar de
  // crescer por algumas checagens seguidas (cobre listas com carregamento
  // incremental ao rolar, sem depender de um número fixo de pedidos).
  let ultimaAltura = -1;
  let estavel = 0;
  const maxIteracoes = 400;
  for (let i = 0; i < maxIteracoes; i++) {
    container.scrollTop = container.scrollHeight;
    await new Promise(r => setTimeout(r, 220));
    const alturaAtual = container.scrollHeight;
    if (alturaAtual === ultimaAltura) {
      estavel++;
      if (estavel >= 3) break;
    } else {
      estavel = 0;
    }
    ultimaAltura = alturaAtual;
  }
  container.scrollTop = 0; // cortesia: devolve o painel pro topo

  const linhas = coletarLinhas();
  const agregado = {}; // sku -> { qtdInsuficiente, qtdComLocalizacao }
  linhas.forEach(r => {
    if (!agregado[r.sku]) agregado[r.sku] = { qtdInsuficiente: 0, qtdComLocalizacao: 0 };
    if (r.insuficiente) agregado[r.sku].qtdInsuficiente += r.qty;
    else agregado[r.sku].qtdComLocalizacao += r.qty;
  });
  const itens = Object.entries(agregado).map(([sku, v]) => ({ sku, qtdInsuficiente: v.qtdInsuficiente, qtdComLocalizacao: v.qtdComLocalizacao }));

  return { ok: true, itens, totalLinhas: linhas.length };
}

$('salvar').addEventListener('click', async () => {
  await salvarConfig();
  log('✅ Configuração salva.', 'ok');
});

$('obter').addEventListener('click', async () => {
  const webAppUrl = $('webAppUrl').value.trim();
  const token = $('token').value.trim();
  if (!webAppUrl || !token) {
    log('⚠ Preencha a URL do Web App e o Token antes de continuar.', 'err');
    return;
  }
  await salvarConfig();

  $('obter').disabled = true;
  log('🔍 Lendo painel e rolando a lista até o fim...', '');

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab || !/panel-u\.baselinker\.com/.test(tab.url || '')) {
      throw new Error('Abra o painel de Coleta de Pedidos do BaseLinker na aba ativa antes de clicar.');
    }

    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: scrapeInsuficientes,
    });

    if (!result || !result.ok) {
      throw new Error((result && result.error) || 'Falha ao ler o painel.');
    }

    log('📤 Enviando ' + result.itens.length + ' SKUs (' + result.totalLinhas + ' itens lidos)...', '');

    const resp = await fetch(webAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ token, itens: result.itens }),
    });
    const data = await resp.json();
    if (!data.ok) throw new Error(data.error || 'Erro desconhecido no servidor.');

    log('✅ Enviado! ' + data.skus + ' SKUs gravados na planilha.', 'ok');
  } catch (e) {
    log('❌ ' + e.message, 'err');
  } finally {
    $('obter').disabled = false;
  }
});

carregarConfig();
