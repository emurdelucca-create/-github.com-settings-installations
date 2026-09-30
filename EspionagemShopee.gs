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
    .addItem('📐 Criar/resetar layout da planilha', 'esp_criarLayout')
    .addItem('🔄 Atualizar "Meu anúncio" (todas as linhas)', 'esp_atualizarMeusAnuncios')
    .addSeparator()
    .addItem('🧪 Testar API — meu anúncio (log)', 'esp_testarMeuAnuncio')
    .addToUi();
}

// ── Layout da planilha ───────────────────────────────────────
// A:G = Meu anúncio | H:N = Anúncio Concorrente | O = Link Anúncio
const ESP_COL = {
  MEU_ID: 1, MEU_BRUTO: 2, MEU_FINAL: 3, MEU_7: 4, MEU_15: 5, MEU_15_30: 6, MEU_30: 7,
  CONC_ID: 8, CONC_BRUTO: 9, CONC_FINAL: 10, CONC_7: 11, CONC_15: 12, CONC_15_30: 13, CONC_30: 14,
  LINK: 15,
};
const ESP_PRIMEIRA_LINHA_DADOS = 4;

function esp_criarLayout() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName('Espionagem');
  if (!aba) aba = ss.insertSheet('Espionagem');

  aba.getRange(1, 1, 3, 15).clearContent();
  aba.getRange('A1:G1').merge().setValue('Meu anúncio').setFontWeight('bold').setHorizontalAlignment('center');
  aba.getRange('H1:N1').merge().setValue('Anúncio Concorrente').setFontWeight('bold').setHorizontalAlignment('center');

  const headers2 = ['ID produto', 'Preço Bruto', 'Preço Final', 'Qntd Vend.', '', '', '',
                     'ID produto', 'Preço Bruto', 'Preço Final', 'Qntd Vend.', '', '', '', 'Link Anúncio'];
  aba.getRange(2, 1, 1, 15).setValues([headers2]).setFontWeight('bold');
  aba.getRange('D2:G2').merge().setValue('Qntd Vend.').setHorizontalAlignment('center');
  aba.getRange('K2:N2').merge().setValue('Qntd Vend.').setHorizontalAlignment('center');

  const headers3 = ['', '', '', '0-7', '0-15', '15-30', '0-30',
                     '', '', '', '0-7', '0-15', '15-30', '0-30', ''];
  aba.getRange(3, 1, 1, 15).setValues([headers3]).setFontWeight('bold').setHorizontalAlignment('center');

  aba.setFrozenRows(3);
  for (let c = 1; c <= 15; c++) aba.setColumnWidth(c, c === 15 ? 320 : 100);

  SpreadsheetApp.getUi().alert(
    '✅ Layout criado na aba "Espionagem".\n\n' +
    'Preencha a partir da linha 4:\n' +
    '• Coluna A: ID do produto (item_id) do SEU anúncio\n' +
    '• Coluna O: link completo do anúncio do concorrente\n' +
    '(o shop_id e item_id do concorrente são extraídos automaticamente do link)'
  );
}

// ── Extrai shop_id e item_id de um link da Shopee ───────────
// Formato padrão: .../produto-i.{shopid}.{itemid}?...
function _esp_parseLink(link) {
  const m = String(link || '').match(/-i\.(\d+)\.(\d+)/);
  if (!m) return null;
  return { shopId: Number(m[1]), itemId: Number(m[2]) };
}

// ── Janela de datas: sempre até o FIM DE ONTEM (hoje fica de fora por
// estar incompleto). 0-7/0-15/0-30 = acumulado; 15-30 = janela própria
// (dias 15 a 30 atrás, não soma com os outros).
function _esp_janelasData(tz) {
  const [ano, mes, dia] = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-').map(Number);
  const fimOntem = new Date(ano, mes - 1, dia); // meia-noite de HOJE = fim de ontem (exclusivo)
  const fimOntemSec = Math.floor(fimOntem.getTime() / 1000) - 1;
  const diasSec = d => d * 24 * 60 * 60;
  return {
    fimOntemSec,
    ini7:  fimOntemSec - diasSec(7)  + 1,
    ini15: fimOntemSec - diasSec(15) + 1,
    ini30: fimOntemSec - diasSec(30) + 1,
    ini15_30_inicio: fimOntemSec - diasSec(30) + 1,
    ini15_30_fim:    fimOntemSec - diasSec(15),
  };
}

// ── Varre TODOS os pedidos dos últimos 30 dias uma única vez e monta
// um mapa item_id -> [{ qty, createTime }] — usado pra calcular as 4
// janelas de qualquer item sem refazer a varredura por linha.
function _esp_mapaVendasPorItem(tz) {
  const jan = _esp_janelasData(tz);
  const mapa = {}; // item_id -> [{qty, createTime}]

  // get_order_list só aceita até 15 dias de intervalo por chamada —
  // quebra a janela de 30 dias em pedaços de no máximo 14 dias (folga
  // de 1 dia pra não esbarrar em arredondamento de segundo).
  const MAX_DIAS_POR_CHAMADA = 14;
  const SEG_POR_DIA = 24 * 60 * 60;
  const orderSns = [];

  for (let inicioChunk = jan.ini30; inicioChunk <= jan.fimOntemSec; inicioChunk += MAX_DIAS_POR_CHAMADA * SEG_POR_DIA) {
    const fimChunk = Math.min(inicioChunk + MAX_DIAS_POR_CHAMADA * SEG_POR_DIA - 1, jan.fimOntemSec);

    let cursor = '';
    let mais = true;
    while (mais) {
      const r = _espShopeeGet('/api/v2/order/get_order_list', {
        time_range_field: 'create_time',
        time_from: inicioChunk,
        time_to: fimChunk,
        page_size: 100,
        cursor: cursor,
      });
      (r.order_list || []).forEach(o => orderSns.push(o.order_sn));
      mais = !!r.more;
      cursor = r.next_cursor || '';
      if (!cursor) break;
    }
  }

  // get_order_detail aceita até 50 order_sn por chamada.
  for (let i = 0; i < orderSns.length; i += 50) {
    const lote = orderSns.slice(i, i + 50);
    const r = _espShopeeGet('/api/v2/order/get_order_detail', {
      order_sn_list: lote.join(','),
      response_optional_fields: 'item_list,order_status',
    });
    (r.order_list || []).forEach(pedido => {
      if (pedido.order_status === 'CANCELLED') return;
      (pedido.item_list || []).forEach(it => {
        if (!mapa[it.item_id]) mapa[it.item_id] = [];
        mapa[it.item_id].push({ qty: it.model_quantity_purchased || 0, createTime: pedido.create_time });
      });
    });
  }

  return { mapa, jan };
}

function _esp_somarJanela(vendas, inicioSec, fimSec) {
  return vendas
    .filter(v => v.createTime >= inicioSec && v.createTime <= fimSec)
    .reduce((acc, v) => acc + v.qty, 0);
}

// ── Preço bruto (maior original_price entre variações) e final
// (menor current_price entre variações), via get_model_list ──
function _esp_precosItem(itemId) {
  const r = _espShopeeGet('/api/v2/product/get_model_list', { item_id: itemId });
  const modelos = r.model || [];
  if (!modelos.length) return { bruto: null, final: null };

  let bruto = null, final = null;
  modelos.forEach(m => {
    (m.price_info || []).forEach(p => {
      if (bruto === null || p.original_price > bruto) bruto = p.original_price;
      if (final === null || p.current_price < final)  final = p.current_price;
    });
  });
  return { bruto, final };
}

// ── Atualiza a coluna "Meu anúncio" (A:G) de todas as linhas ────
function esp_atualizarMeusAnuncios() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tz  = ss.getSpreadsheetTimeZone();
  const aba = ss.getSheetByName('Espionagem');
  if (!aba) { SpreadsheetApp.getUi().alert('Crie o layout primeiro (menu → 📐 Criar/resetar layout).'); return; }

  const lastRow = aba.getLastRow();
  if (lastRow < ESP_PRIMEIRA_LINHA_DADOS) return;

  const linhas = aba.getRange(ESP_PRIMEIRA_LINHA_DADOS, ESP_COL.MEU_ID, lastRow - ESP_PRIMEIRA_LINHA_DADOS + 1, 1).getValues();
  const { mapa, jan } = _esp_mapaVendasPorItem(tz);

  linhas.forEach((row, i) => {
    const itemId = row[0];
    if (!itemId) return;
    const linhaPlanilha = ESP_PRIMEIRA_LINHA_DADOS + i;

    const precos = _esp_precosItem(itemId);
    const vendas = mapa[itemId] || [];
    const v7     = _esp_somarJanela(vendas, jan.ini7,  jan.fimOntemSec);
    const v15    = _esp_somarJanela(vendas, jan.ini15, jan.fimOntemSec);
    const v15_30 = _esp_somarJanela(vendas, jan.ini15_30_inicio, jan.ini15_30_fim);
    const v30    = _esp_somarJanela(vendas, jan.ini30, jan.fimOntemSec);

    aba.getRange(linhaPlanilha, ESP_COL.MEU_BRUTO, 1, 6).setValues([[precos.bruto, precos.final, v7, v15, v15_30, v30]]);
  });

  SpreadsheetApp.flush();
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

// Busca especificamente por qualquer campo relacionado a "venda/sold" na
// resposta do get_item_base_info (o log anterior cortou antes de chegar
// lá). Também lista TODAS as chaves de primeiro nível de cada item, pra
// não depender de eu adivinhar o nome exato do campo.
function esp_testarCamposVenda() {
  const ITEM_ID_TESTE = 41655343417;
  const base = _espShopeeGet('/api/v2/product/get_item_base_info', { item_id_list: ITEM_ID_TESTE });
  const item = (base.item_list || [])[0];
  if (!item) { Logger.log('Nenhum item retornado.'); return; }

  Logger.log('=== Todas as chaves de primeiro nível do item ===');
  Logger.log(Object.keys(item).join(', '));

  Logger.log('\n=== Campos com "sold" ou "sale" no nome (qualquer nível) ===');
  const encontrados = [];
  (function buscar(obj, caminho) {
    if (obj === null || typeof obj !== 'object') return;
    Object.keys(obj).forEach(k => {
      const novoCaminho = caminho + '.' + k;
      if (/sold|sale/i.test(k)) encontrados.push(novoCaminho + ' = ' + JSON.stringify(obj[k]));
      buscar(obj[k], novoCaminho);
    });
  })(item, 'item');
  Logger.log(encontrados.length ? encontrados.join('\n') : '(nenhum campo com "sold"/"sale" encontrado)');
}

// Testa a API de pedidos: get_order_list (últimos 7 dias) + get_order_detail
// de um pedido, pra confirmar se dá pra somar quantidade vendida por
// item_id a partir do histórico de pedidos (sem precisar de um campo de
// "sold" pronto).
function esp_testarPedidos() {
  const agora = Math.floor(Date.now() / 1000);
  const seteDiasAtras = agora - 7 * 24 * 60 * 60;

  Logger.log('=== get_order_list (últimos 7 dias) ===');
  const lista = _espShopeeGet('/api/v2/order/get_order_list', {
    time_range_field: 'create_time',
    time_from: seteDiasAtras,
    time_to: agora,
    page_size: 20,
  });
  Logger.log(JSON.stringify(lista, null, 2).substring(0, 2000));

  const primeiro = (lista.order_list || [])[0];
  if (!primeiro) { Logger.log('\nNenhum pedido encontrado nos últimos 7 dias.'); return; }

  Logger.log('\n=== get_order_detail (do primeiro pedido: ' + primeiro.order_sn + ') ===');
  const detalhe = _espShopeeGet('/api/v2/order/get_order_detail', {
    order_sn_list: primeiro.order_sn,
    response_optional_fields: 'item_list,total_amount,order_status',
  });
  Logger.log(JSON.stringify(detalhe, null, 2).substring(0, 3000));
}
