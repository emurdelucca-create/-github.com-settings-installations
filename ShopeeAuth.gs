// ============================================================
// SHOPEE AUTH + API — OAuth 2.0 + Produtos + Promoções
//
// Credenciais — adicionar em Configurações do projeto → Propriedades do script:
//   SHOPEE_PARTNER_KEY  → Live API Partner Key (shpk6c57...)
//
// Tokens gerados automaticamente e armazenados nas mesmas propriedades:
//   SHOPEE_ACCESS_TOKEN   SHOPEE_REFRESH_TOKEN
//   SHOPEE_SHOP_ID        SHOPEE_TOKEN_EXPIRES
// ============================================================

const SHOPEE_BASE       = 'https://partner.shopeemobile.com';
const SHOPEE_PARTNER_ID = 2037491;

// ============================================================
// GERAR LINK DE AUTORIZAÇÃO
// ============================================================
function gerarLinkAutorizacaoShopee() {
  const ui   = SpreadsheetApp.getUi();
  const pkey = PropertiesService.getScriptProperties().getProperty('SHOPEE_PARTNER_KEY');
  if (!pkey) {
    ui.alert(
      '❌ SHOPEE_PARTNER_KEY não configurada.\n\n' +
      'Adicione em:\nExtensões → Apps Script → ⚙️ Configurações → Propriedades do script\n\n' +
      'Chave: SHOPEE_PARTNER_KEY\nValor: shpk6c57...'
    );
    return;
  }

  const ts   = Math.floor(Date.now() / 1000);
  const path = '/api/v2/shop/auth_partner';
  const sign = _shopeeSign(SHOPEE_PARTNER_ID + path + ts, pkey);
  const url  = SHOPEE_BASE + path
    + '?partner_id=' + SHOPEE_PARTNER_ID
    + '&timestamp='  + ts
    + '&sign='        + sign
    + '&redirect='    + encodeURIComponent('https://localhost');

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:13px;padding:12px;line-height:1.7">' +
    '<p><b>Passo 1.</b> Clique no link abaixo e faça login com a conta <b>Shopee Humble</b>:</p>' +
    '<p><a href="' + url + '" target="_blank" ' +
    'style="background:#e65100;color:#fff;padding:8px 16px;border-radius:4px;' +
    'text-decoration:none;font-size:14px">👉 Autorizar app na Shopee</a></p>' +
    '<p><b>Passo 2.</b> Após autorizar, o browser vai abrir uma página de erro ' +
    '(localhost não existe — isso é normal).</p>' +
    '<p><b>Passo 3.</b> Copie a URL completa da barra de endereço do browser.<br>' +
    'Ela terá um aspecto como:<br>' +
    '<code style="font-size:11px">https://localhost/?code=XXXX&shop_id=YYYYY</code></p>' +
    '<p><b>Passo 4.</b> Feche esta janela e clique em <b>💾 Salvar Token</b> no menu.</p>' +
    '</div>'
  ).setWidth(500).setHeight(280);

  ui.showModalDialog(html, '🔑 Passo 1 — Autorizar Shopee Humble');
}

// ============================================================
// DIALOG PARA COLAR A URL E SALVAR TOKEN
// ============================================================
function mostrarDialogSalvarToken() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; font-size: 13px; padding: 14px; }
      label { font-weight: bold; display: block; margin-top: 10px; }
      input  { width: 100%; padding: 7px; box-sizing: border-box;
               font-size: 13px; border: 1px solid #ccc; border-radius: 3px; }
      button { margin-top: 14px; background: #e65100; color: #fff; border: none;
               padding: 10px 24px; font-size: 14px; border-radius: 4px; cursor: pointer; }
      button:hover { background: #bf360c; }
      .hint { color: #777; font-size: 11px; margin-top: 2px; }
    </style>
    <p>Cole a URL que apareceu após autorizar (ou só o code + shop_id):</p>
    <label>URL completa (opcional)</label>
    <input id="url" placeholder="https://localhost/?code=XXXX&shop_id=YYYYY" />
    <label>code</label>
    <input id="code" placeholder="ex: 6f4a1b2c3d..." />
    <label>shop_id</label>
    <input id="shopid" placeholder="ex: 123456789" />
    <p class="hint">Se colar a URL completa acima, os campos são preenchidos automaticamente.</p>
    <button onclick="enviar()">💾 Salvar Token</button>
    <script>
      document.getElementById('url').addEventListener('input', function() {
        try {
          const p = new URL(this.value.trim()).searchParams;
          const c = p.get('code'), s = p.get('shop_id');
          if (c) document.getElementById('code').value = c;
          if (s) document.getElementById('shopid').value = s;
        } catch(e) {}
      });
      function enviar() {
        const code   = document.getElementById('code').value.trim();
        const shopId = document.getElementById('shopid').value.trim();
        if (!code || !shopId) { alert('Preencha o code e o shop_id.'); return; }
        document.querySelector('button').textContent = '⏳ Salvando...';
        google.script.run
          .withSuccessHandler(msg => { alert(msg); google.script.host.close(); })
          .withFailureHandler(err => {
            alert('Erro: ' + err.message);
            document.querySelector('button').textContent = '💾 Salvar Token';
          })
          .trocarCodigoPorToken(code, parseInt(shopId));
      }
    </script>
  `).setWidth(480).setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, '💾 Passo 2 — Salvar Token Shopee Humble');
}

// ============================================================
// TROCAR CÓDIGO POR ACCESS_TOKEN + REFRESH_TOKEN
// ============================================================
function trocarCodigoPorToken(code, shopId) {
  const pkey = _shopeePartnerKey();
  const path = '/api/v2/auth/token/get';
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _shopeeSign(SHOPEE_PARTNER_ID + path + ts, pkey);

  const resp = UrlFetchApp.fetch(
    SHOPEE_BASE + path
    + '?partner_id=' + SHOPEE_PARTNER_ID
    + '&timestamp='  + ts
    + '&sign='        + sign,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ code, shop_id: shopId, partner_id: SHOPEE_PARTNER_ID }),
      muteHttpExceptions: true,
    }
  );

  const data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error(data.message || JSON.stringify(data));

  const props     = PropertiesService.getScriptProperties();
  const expiresAt = (Math.floor(Date.now() / 1000) + (data.expire_in || 14400)).toString();
  props.setProperty('SHOPEE_ACCESS_TOKEN',  data.access_token);
  props.setProperty('SHOPEE_REFRESH_TOKEN', data.refresh_token);
  props.setProperty('SHOPEE_SHOP_ID',       String(shopId));
  props.setProperty('SHOPEE_TOKEN_EXPIRES', expiresAt);

  return '✅ Token salvo com sucesso! Conta Shopee Humble autorizada.\nPode fechar esta janela e usar o menu 🔄 Atualizar Reprecificação.';
}

// ============================================================
// STATUS DO TOKEN (para diagnóstico)
// ============================================================
function verificarStatusToken() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('SHOPEE_ACCESS_TOKEN');
  const expires = parseInt(props.getProperty('SHOPEE_TOKEN_EXPIRES') || '0');
  const shopId  = props.getProperty('SHOPEE_SHOP_ID');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) {
    SpreadsheetApp.getUi().alert('❌ Nenhum token salvo. Execute "Autorizar Shopee Humble" no menu.');
    return;
  }

  const minutos = Math.round((expires - now) / 60);
  const msg = shopId
    ? '✅ Token ativo\n\nShop ID: ' + shopId + '\nExpira em: ' + (minutos > 0 ? minutos + ' minutos' : 'EXPIRADO — será renovado automaticamente')
    : '⚠️ Token salvo mas shop_id não encontrado.';

  SpreadsheetApp.getUi().alert('🔑 Status do Token', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// OBTER TOKEN VÁLIDO (renova se necessário)
// ============================================================
function _shopeeGetToken() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('SHOPEE_ACCESS_TOKEN');
  const refresh = props.getProperty('SHOPEE_REFRESH_TOKEN');
  const expires = parseInt(props.getProperty('SHOPEE_TOKEN_EXPIRES') || '0');
  const shopId  = props.getProperty('SHOPEE_SHOP_ID');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) throw new Error('Token Shopee não encontrado. Use o menu 🔑 Autorizar Shopee Humble.');

  // Renova se faltam menos de 5 minutos
  if (now >= expires - 300) {
    return _shopeeRenovarToken(refresh, parseInt(shopId));
  }
  return { token, shopId };
}

// ============================================================
// RENOVAR ACCESS_TOKEN VIA REFRESH_TOKEN
// ============================================================
function _shopeeRenovarToken(refreshToken, shopId) {
  const pkey = _shopeePartnerKey();
  const path = '/api/v2/auth/access_token/get';
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _shopeeSign(SHOPEE_PARTNER_ID + path + ts, pkey);

  const resp = UrlFetchApp.fetch(
    SHOPEE_BASE + path
    + '?partner_id=' + SHOPEE_PARTNER_ID
    + '&timestamp='  + ts
    + '&sign='        + sign,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        refresh_token: refreshToken,
        shop_id:       shopId,
        partner_id:    SHOPEE_PARTNER_ID,
      }),
      muteHttpExceptions: true,
    }
  );

  const data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error('Falha ao renovar token: ' + (data.message || JSON.stringify(data)));

  const props     = PropertiesService.getScriptProperties();
  const expiresAt = (Math.floor(Date.now() / 1000) + (data.expire_in || 14400)).toString();
  props.setProperty('SHOPEE_ACCESS_TOKEN',  data.access_token);
  props.setProperty('SHOPEE_REFRESH_TOKEN', data.refresh_token || refreshToken);
  props.setProperty('SHOPEE_TOKEN_EXPIRES', expiresAt);

  return { token: data.access_token, shopId: String(shopId) };
}

// ============================================================
// GET AUTENTICADO
// ============================================================
function _shopeeGet(path, params) {
  const { token, shopId } = _shopeeGetToken();
  const pkey = _shopeePartnerKey();
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _shopeeSign(
    SHOPEE_PARTNER_ID + path + ts + token + parseInt(shopId),
    pkey
  );

  const base = SHOPEE_BASE + path
    + '?partner_id='   + SHOPEE_PARTNER_ID
    + '&timestamp='    + ts
    + '&access_token=' + token
    + '&shop_id='      + shopId
    + '&sign='          + sign;

  const qs = params
    ? '&' + Object.entries(params)
        .filter(([, v]) => v !== undefined && v !== null && v !== '')
        .map(([k, v]) => k + '=' + encodeURIComponent(v))
        .join('&')
    : '';

  const resp = UrlFetchApp.fetch(base + qs, { muteHttpExceptions: true });
  const data = JSON.parse(resp.getContentText());
  if (data.error && data.error !== '') {
    throw new Error('[' + path + '] ' + (data.message || data.error));
  }
  return data.response || data;
}

// ============================================================
// BUSCAR PRODUTOS VIA API
// Retorna array idêntico ao de _sr_carregarProdutos()
// ============================================================
function _sr_buscarProdutosAPI() {
  const produtos = [];
  const allIds   = [];
  let offset     = 0;

  // Passo 1: listar todos os item_id ativos
  while (true) {
    const r = _shopeeGet('/api/v2/product/get_item_list', {
      offset:      offset,
      page_size:   100,
      item_status: 'NORMAL',
    });
    (r.item || []).forEach(i => allIds.push(i.item_id));
    if (!r.has_next_page) break;
    offset += 100;
    Utilities.sleep(300);
  }

  if (!allIds.length) return produtos;

  // Passo 2: detalhes em lotes de 50
  for (let i = 0; i < allIds.length; i += 50) {
    const batch = allIds.slice(i, i + 50);
    const r = _shopeeGet('/api/v2/product/get_item_base_info', {
      item_id_list:    batch.join(','),
      need_tax_info:   false,
      need_complaint_policy: false,
    });

    (r.item_list || []).forEach(item => {
      const nome   = item.item_name   || '';
      const idItem = String(item.item_id || '');

      if (item.has_model) {
        // Com variações
        (item.model_list || []).forEach(model => {
          const preco = _shopeePreco(model.price_info);
          const est   = _shopeeEstoque(model.stock_info_v2);
          produtos.push({
            nome,
            idItem,
            variacao:  model.model_name || '',
            skuVar:    String(model.model_sku || '').trim(),
            skuUsr:    String(model.model_sku || '').trim(),
            precoNorm: preco,
            estShopee: est,
          });
        });
      } else {
        // Sem variação
        const preco = _shopeePreco(item.price_info);
        const est   = _shopeeEstoque(item.stock_info_v2);
        produtos.push({
          nome,
          idItem,
          variacao:  '',
          skuVar:    String(item.item_sku || '').trim(),
          skuUsr:    String(item.item_sku || '').trim(),
          precoNorm: preco,
          estShopee: est,
        });
      }
    });

    if (i + 50 < allIds.length) Utilities.sleep(300);
  }

  return produtos;
}

// ============================================================
// BUSCAR PREÇOS, PROMOÇÕES E ESTOQUE VIA API
// Retorna mapa com entradas por múltiplas chaves para garantir matching:
//   "idItem|modelId"  — variação com ID numérico (principal)
//   "idItem|modelSku" — variação com SKU texto (fallback)
//   "idItem|"         — produto sem variação
// Cada entrada: { precoNorm (original_price), precoPromo (current_price
// se diferente), estShopee }
// ============================================================
function _sr_buscarPrecosAPI() {
  const mapa   = {};
  const allIds = [];
  let offset   = 0;

  while (true) {
    const r = _shopeeGet('/api/v2/product/get_item_list', {
      offset:      offset,
      page_size:   100,
      item_status: 'NORMAL',
    });
    (r.item || []).forEach(i => allIds.push(i.item_id));
    if (!r.has_next_page) break;
    offset += 100;
    Utilities.sleep(300);
  }

  if (!allIds.length) return mapa;

  const _registrar = (key, priceInfo, stockInfo) => {
    if (!key) return;
    const pi = priceInfo && priceInfo[0];
    const precoOrig  = pi ? (pi.original_price || 0) : 0;
    const precoAtual = pi ? (pi.current_price  || 0) : 0;
    mapa[key] = {
      precoNorm:  precoOrig  || precoAtual,
      precoPromo: (precoAtual > 0 && precoAtual < precoOrig) ? precoAtual : 0,
      estShopee:  _shopeeEstoque(stockInfo),
    };
  };

  for (let i = 0; i < allIds.length; i += 50) {
    const batch = allIds.slice(i, i + 50);
    const r = _shopeeGet('/api/v2/product/get_item_base_info', {
      item_id_list:          batch.join(','),
      need_tax_info:         false,
      need_complaint_policy: false,
    });

    (r.item_list || []).forEach(item => {
      const idItem = String(item.item_id || '');
      if (item.has_model) {
        // Chave = idItem + '|' + modelId  (quando model_id disponível)
        //       = idItem + '|' + modelSku (sempre — garante match mesmo sem model_id
        //         e evita colisão de SKU duplicado em anúncios diferentes)
        (item.model_list || []).forEach(model => {
          const modelId  = model.model_id && model.model_id !== 0
                           ? String(model.model_id) : '';
          const modelSku = String(model.model_sku || '').trim();
          if (modelId)  _registrar(idItem + '|' + modelId,  model.price_info, model.stock_info_v2);
          if (modelSku) _registrar(idItem + '|' + modelSku, model.price_info, model.stock_info_v2);
        });
      } else {
        // Produto sem variação: chave = idItem + '|'  e  idItem + '|' + itemSku
        _registrar(idItem + '|', item.price_info, item.stock_info_v2);
        const itemSku = String(item.item_sku || '').trim();
        if (itemSku) _registrar(idItem + '|' + itemSku, item.price_info, item.stock_info_v2);
      }
    });

    if (i + 50 < allIds.length) Utilities.sleep(300);
  }

  return mapa;
}

// ============================================================
// BUSCAR PROMOÇÕES ATIVAS VIA API
// Retorna mapa: sku → precoPromo
// ============================================================
function _sr_buscarPromocoesAPI() {
  const mapa = {};
  let pageNo = 1;

  while (true) {
    let r;
    try {
      r = _shopeeGet('/api/v2/discount/get_discount_list', {
        discount_status: 'ongoing',
        page_size:       100,
        page_no:         pageNo,
      });
    } catch (e) {
      break; // sem promoções ativas ou sem permissão
    }

    const lista = r.discount_list || [];
    if (!lista.length) break;

    lista.forEach(d => {
      try {
        // Detalhes de cada desconto — paginado por item
        let pn = 1;
        while (true) {
          const det = _shopeeGet('/api/v2/discount/get_discount_detail', {
            discount_id: d.discount_id,
            page_size:   100,
            page_no:     pn,
          });
          (det.item_list || []).forEach(item => {
            (item.model_list && item.model_list.length
              ? item.model_list
              : [{ model_sku: item.item_sku, model_promotion_price: item.item_promotion_price }]
            ).forEach(model => {
              const sku   = String(model.model_sku || '').trim();
              const preco = model.model_promotion_price || 0;
              if (sku && preco > 0) mapa[sku] = preco;
            });
          });
          if (!det.more) break;
          pn++;
          Utilities.sleep(200);
        }
      } catch (e) { /* desconto sem detalhes, ignorar */ }
    });

    if (!r.more) break;
    pageNo++;
    Utilities.sleep(400);
  }

  return mapa;
}

// ============================================================
// HELPERS INTERNOS
// ============================================================
function _shopeeSign(message, partnerKey) {
  return Utilities.computeHmacSha256Signature(message, partnerKey)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
    .join('');
}

function _shopeePartnerKey() {
  const k = PropertiesService.getScriptProperties().getProperty('SHOPEE_PARTNER_KEY');
  if (!k) throw new Error('SHOPEE_PARTNER_KEY não configurada nas propriedades do script.');
  return k;
}

function _shopeePreco(priceInfo) {
  if (!priceInfo || !priceInfo.length) return 0;
  return priceInfo[0].current_price || priceInfo[0].original_price || 0;
}

function _shopeeEstoque(stockInfo) {
  if (!stockInfo || !stockInfo.summary_info) return 0;
  return stockInfo.summary_info.total_available_stock || 0;
}
