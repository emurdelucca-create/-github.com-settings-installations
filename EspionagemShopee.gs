// ============================================================
// ESPIONAGEM SHOPEE
// - "Meu anúncio" (loja Najumi): API autenticada (Open Platform)
// - "Anúncio Concorrente": API pública, via extensão de navegador
//   (a API pública é bloqueada por IP quando chamada do servidor do
//   Apps Script — testado e confirmado; só funciona de dentro de um
//   navegador de verdade)
//
// Propriedades do Script necessárias (Configurações ⚙ → Propriedades):
//   NAJUMI_PARTNER_KEY — mesma Live API Partner Key já usada no projeto
//                        "Carteira Shopee Najumi" (projetos diferentes
//                        não compartilham propriedades, por isso precisa
//                        configurar de novo aqui)
//
// Tokens desta loja, gerenciados automaticamente:
//   ESP_NAJUMI_ACCESS_TOKEN   ESP_NAJUMI_REFRESH_TOKEN
//   ESP_NAJUMI_SHOP_ID        ESP_NAJUMI_TOKEN_EXPIRES
// ============================================================

const ESP_NAJUMI_BASE    = 'https://partner.shopeemobile.com';
const ESP_PARTNER_ID     = 2038327; // mesmo app da loja Najumi

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🕵️ Espionagem Shopee')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔑 Autorizar Shopee Najumi')
      .addItem('1️⃣  Gerar link de autorização', 'esp_gerarLinkAutorizacao')
      .addItem('2️⃣  Salvar token (colar URL)',   'esp_mostrarDialogSalvarToken')
      .addItem('🔍 Verificar status do token',   'esp_verificarStatusToken'))
    .addSeparator()
    .addItem('🧪 Testar API — meu anúncio (log)', 'esp_testarMeuAnuncio')
    .addToUi();
}

// ── Helpers de assinatura HMAC ──────────────────────────────
function _espShopeeSign(message, partnerKey) {
  return Utilities.computeHmacSha256Signature(message, partnerKey)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
    .join('');
}

function _espPartnerKey() {
  const k = PropertiesService.getScriptProperties().getProperty('NAJUMI_PARTNER_KEY');
  if (!k) throw new Error('NAJUMI_PARTNER_KEY não configurada nas propriedades do script.');
  return k;
}

// ── PASSO 1: gerar link de autorização ──────────────────────
function esp_gerarLinkAutorizacao() {
  const ui   = SpreadsheetApp.getUi();
  const pkey = _espPartnerKey();

  const ts   = Math.floor(Date.now() / 1000);
  const path = '/api/v2/shop/auth_partner';
  const sign = _espShopeeSign(ESP_PARTNER_ID + path + ts, pkey);
  const url  = ESP_NAJUMI_BASE + path
    + '?partner_id=' + ESP_PARTNER_ID
    + '&timestamp='  + ts
    + '&sign='        + sign
    + '&redirect='    + encodeURIComponent('https://localhost');

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:13px;padding:12px;line-height:1.7">' +
    '<p><b>Passo 1.</b> Clique no link abaixo e faça login com a conta <b>najumimotopecas</b>:</p>' +
    '<p><a href="' + url + '" target="_blank" ' +
    'style="background:#e65100;color:#fff;padding:8px 16px;border-radius:4px;' +
    'text-decoration:none;font-size:14px">👉 Autorizar app na Shopee (Najumi)</a></p>' +
    '<p><b>Passo 2.</b> Após autorizar, o browser vai abrir uma página de erro ' +
    '(localhost não existe — isso é normal).</p>' +
    '<p><b>Passo 3.</b> Copie a URL completa da barra de endereço do browser.<br>' +
    'Ela terá um aspecto como:<br>' +
    '<code style="font-size:11px">https://localhost/?code=XXXX&shop_id=YYYYY</code></p>' +
    '<p><b>Passo 4.</b> Feche esta janela e clique em <b>💾 Salvar token</b> no menu.</p>' +
    '</div>'
  ).setWidth(500).setHeight(280);

  ui.showModalDialog(html, '🔑 Passo 1 — Autorizar Shopee Najumi');
}

// ── PASSO 2: dialog para colar a URL e salvar token ─────────
function esp_mostrarDialogSalvarToken() {
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
    <button onclick="enviar()">💾 Salvar token</button>
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
            document.querySelector('button').textContent = '💾 Salvar token';
          })
          .esp_trocarCodigoPorToken(code, parseInt(shopId));
      }
    </script>
  `).setWidth(480).setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, '💾 Passo 2 — Salvar Token Shopee Najumi');
}

// ── Trocar código por access_token + refresh_token ──────────
function esp_trocarCodigoPorToken(code, shopId) {
  const pkey = _espPartnerKey();
  const path = '/api/v2/auth/token/get';
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _espShopeeSign(ESP_PARTNER_ID + path + ts, pkey);

  const resp = UrlFetchApp.fetch(
    ESP_NAJUMI_BASE + path
    + '?partner_id=' + ESP_PARTNER_ID
    + '&timestamp='  + ts
    + '&sign='        + sign,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ code, shop_id: shopId, partner_id: ESP_PARTNER_ID }),
      muteHttpExceptions: true,
    }
  );

  const data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error(data.message || JSON.stringify(data));

  const props     = PropertiesService.getScriptProperties();
  const expiresAt = (Math.floor(Date.now() / 1000) + (data.expire_in || 14400)).toString();
  props.setProperty('ESP_NAJUMI_ACCESS_TOKEN',  data.access_token);
  props.setProperty('ESP_NAJUMI_REFRESH_TOKEN', data.refresh_token);
  props.setProperty('ESP_NAJUMI_SHOP_ID',       String(shopId));
  props.setProperty('ESP_NAJUMI_TOKEN_EXPIRES', expiresAt);

  return '✅ Token salvo com sucesso! Conta Shopee Najumi autorizada.';
}

// ── Status do token (diagnóstico) ───────────────────────────
function esp_verificarStatusToken() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('ESP_NAJUMI_ACCESS_TOKEN');
  const expires = parseInt(props.getProperty('ESP_NAJUMI_TOKEN_EXPIRES') || '0');
  const shopId  = props.getProperty('ESP_NAJUMI_SHOP_ID');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) {
    SpreadsheetApp.getUi().alert('❌ Nenhum token salvo. Execute "Autorizar Shopee Najumi" no menu.');
    return;
  }

  const minutos = Math.round((expires - now) / 60);
  const msg = shopId
    ? '✅ Token ativo\n\nShop ID: ' + shopId + '\nExpira em: ' + (minutos > 0 ? minutos + ' minutos' : 'EXPIRADO — será renovado automaticamente')
    : '⚠️ Token salvo mas shop_id não encontrado.';

  SpreadsheetApp.getUi().alert('🔑 Status do Token', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── Obter token válido (renova se necessário) ───────────────
function _espGetToken() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('ESP_NAJUMI_ACCESS_TOKEN');
  const refresh = props.getProperty('ESP_NAJUMI_REFRESH_TOKEN');
  const expires = parseInt(props.getProperty('ESP_NAJUMI_TOKEN_EXPIRES') || '0');
  const shopId  = props.getProperty('ESP_NAJUMI_SHOP_ID');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) throw new Error('Token Shopee Najumi não encontrado. Use o menu 🔑 Autorizar Shopee Najumi.');

  if (now >= expires - 300) {
    return _espRenovarToken(refresh, parseInt(shopId));
  }
  return { token, shopId };
}

function _espRenovarToken(refreshToken, shopId) {
  const pkey = _espPartnerKey();
  const path = '/api/v2/auth/access_token/get';
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _espShopeeSign(ESP_PARTNER_ID + path + ts, pkey);

  const resp = UrlFetchApp.fetch(
    ESP_NAJUMI_BASE + path
    + '?partner_id=' + ESP_PARTNER_ID
    + '&timestamp='  + ts
    + '&sign='        + sign,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        refresh_token: refreshToken,
        shop_id:       shopId,
        partner_id:    ESP_PARTNER_ID,
      }),
      muteHttpExceptions: true,
    }
  );

  const data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error('Falha ao renovar token: ' + (data.message || JSON.stringify(data)));

  const props     = PropertiesService.getScriptProperties();
  const expiresAt = (Math.floor(Date.now() / 1000) + (data.expire_in || 14400)).toString();
  props.setProperty('ESP_NAJUMI_ACCESS_TOKEN',  data.access_token);
  props.setProperty('ESP_NAJUMI_REFRESH_TOKEN', data.refresh_token || refreshToken);
  props.setProperty('ESP_NAJUMI_TOKEN_EXPIRES', expiresAt);

  return { token: data.access_token, shopId: String(shopId) };
}

// ── GET autenticado (loja Najumi) ───────────────────────────
function _espShopeeGet(path, params) {
  const { token, shopId } = _espGetToken();
  const pkey = _espPartnerKey();
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _espShopeeSign(
    ESP_PARTNER_ID + path + ts + token + parseInt(shopId),
    pkey
  );

  const base = ESP_NAJUMI_BASE + path
    + '?partner_id='   + ESP_PARTNER_ID
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
// DIAGNÓSTICO — testa o "meu anúncio" (item_id 41655343417) pra ver
// quais campos a API autenticada devolve: preços por variação e
// qualquer contador de vendas disponível. Rode e cole o resultado aqui.
// ============================================================
function esp_testarMeuAnuncio() {
  const ITEM_ID_TESTE = 41655343417;

  Logger.log('=== get_item_base_info ===');
  const base = _espShopeeGet('/api/v2/product/get_item_base_info', { item_id_list: ITEM_ID_TESTE });
  Logger.log(JSON.stringify(base, null, 2).substring(0, 3000));

  Logger.log('\n=== get_model_list (variações/preços) ===');
  const modelos = _espShopeeGet('/api/v2/product/get_model_list', { item_id: ITEM_ID_TESTE });
  Logger.log(JSON.stringify(modelos, null, 2).substring(0, 3000));
}
