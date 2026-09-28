// ============================================================
// SHOPEE WALLET — Loja Najumi (najumimotopecas)
// Mapeia transações da carteira Shopee (descontos, receitas) por pedido
//
// Projeto Apps Script standalone, separado do projeto Shopee Humble
// (ShopeeAuth.gs/ShopeeReprecificacao.gs) — projetos diferentes não
// compartilham código nem Propriedades do Script, então este arquivo
// é autossuficiente.
//
// Propriedades do Script necessárias (Configurações ⚙ → Propriedades):
//   NAJUMI_PARTNER_KEY  — Live API Partner Key do app Open Platform
//                         (a mesma chave usada na loja Humble)
//
// Tokens desta loja são gerenciados automaticamente pelo código:
//   SW_NAJUMI_ACCESS_TOKEN   SW_NAJUMI_REFRESH_TOKEN
//   SW_NAJUMI_SHOP_ID        SW_NAJUMI_TOKEN_EXPIRES
// ============================================================

const SW_NAJUMI_BASE       = 'https://partner.shopeemobile.com';
const SHOPEE_PARTNER_ID    = 2038327;

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Carteira Shopee Najumi')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔑 Autorizar Shopee Najumi')
      .addItem('1️⃣  Gerar link de autorização', 'gerarLinkAutorizacaoShopeeNajumi')
      .addItem('2️⃣  Salvar token (colar URL)',   'mostrarDialogSalvarTokenNajumi')
      .addItem('🔍 Verificar status do token',   'verificarStatusTokenNajumi'))
    .addSeparator()
    .addItem('🧪 Testar API da carteira (log)', 'sw_testarWalletAPI')
    .addItem('🔄 Sincronizar carteira agora',   'sw_sincronizarCarteira')
    .addItem('⏱️ Ativar sincronização automática', 'sw_ativarTriggerAutomatico')
    .addItem('⏹️ Desativar sincronização automática', 'sw_desativarTriggerAutomatico')
    .addToUi();
}

// ── Helpers de assinatura HMAC (mesmos usados no projeto Humble) ──
function _shopeeSign(message, partnerKey) {
  return Utilities.computeHmacSha256Signature(message, partnerKey)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
    .join('');
}

function _shopeePartnerKey() {
  const k = PropertiesService.getScriptProperties().getProperty('NAJUMI_PARTNER_KEY');
  if (!k) throw new Error('NAJUMI_PARTNER_KEY não configurada nas propriedades do script.');
  return k;
}

// ── PASSO 1: gerar link de autorização ──────────────────────
function gerarLinkAutorizacaoShopeeNajumi() {
  const ui   = SpreadsheetApp.getUi();
  const pkey = _shopeePartnerKey();

  const ts   = Math.floor(Date.now() / 1000);
  const path = '/api/v2/shop/auth_partner';
  const sign = _shopeeSign(SHOPEE_PARTNER_ID + path + ts, pkey);
  const url  = SW_NAJUMI_BASE + path
    + '?partner_id=' + SHOPEE_PARTNER_ID
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
    '<p><b>Passo 4.</b> Feche esta janela e clique em <b>💾 Salvar Token Najumi</b> no menu.</p>' +
    '</div>'
  ).setWidth(500).setHeight(280);

  ui.showModalDialog(html, '🔑 Passo 1 — Autorizar Shopee Najumi');
}

// ── PASSO 2: dialog para colar a URL e salvar token ─────────
function mostrarDialogSalvarTokenNajumi() {
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
    <button onclick="enviar()">💾 Salvar Token Najumi</button>
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
            document.querySelector('button').textContent = '💾 Salvar Token Najumi';
          })
          .trocarCodigoPorTokenNajumi(code, parseInt(shopId));
      }
    </script>
  `).setWidth(480).setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, '💾 Passo 2 — Salvar Token Shopee Najumi');
}

// ── Trocar código por access_token + refresh_token ──────────
function trocarCodigoPorTokenNajumi(code, shopId) {
  const pkey = _shopeePartnerKey();
  const path = '/api/v2/auth/token/get';
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _shopeeSign(SHOPEE_PARTNER_ID + path + ts, pkey);

  const resp = UrlFetchApp.fetch(
    SW_NAJUMI_BASE + path
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
  props.setProperty('SW_NAJUMI_ACCESS_TOKEN',  data.access_token);
  props.setProperty('SW_NAJUMI_REFRESH_TOKEN', data.refresh_token);
  props.setProperty('SW_NAJUMI_SHOP_ID',       String(shopId));
  props.setProperty('SW_NAJUMI_TOKEN_EXPIRES', expiresAt);

  return '✅ Token salvo com sucesso! Conta Shopee Najumi autorizada.';
}

// ── Status do token (diagnóstico) ───────────────────────────
function verificarStatusTokenNajumi() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('SW_NAJUMI_ACCESS_TOKEN');
  const expires = parseInt(props.getProperty('SW_NAJUMI_TOKEN_EXPIRES') || '0');
  const shopId  = props.getProperty('SW_NAJUMI_SHOP_ID');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) {
    SpreadsheetApp.getUi().alert('❌ Nenhum token Najumi salvo. Execute "Autorizar Shopee Najumi" no menu.');
    return;
  }

  const minutos = Math.round((expires - now) / 60);
  const msg = shopId
    ? '✅ Token ativo\n\nShop ID: ' + shopId + '\nExpira em: ' + (minutos > 0 ? minutos + ' minutos' : 'EXPIRADO — será renovado automaticamente')
    : '⚠️ Token salvo mas shop_id não encontrado.';

  SpreadsheetApp.getUi().alert('🔑 Status do Token Najumi', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── Obter token válido (renova se necessário) ───────────────
function _shopeeNajumiGetToken() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('SW_NAJUMI_ACCESS_TOKEN');
  const refresh = props.getProperty('SW_NAJUMI_REFRESH_TOKEN');
  const expires = parseInt(props.getProperty('SW_NAJUMI_TOKEN_EXPIRES') || '0');
  const shopId  = props.getProperty('SW_NAJUMI_SHOP_ID');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) throw new Error('Token Shopee Najumi não encontrado. Use o menu 🔑 Autorizar Shopee Najumi.');

  if (now >= expires - 300) {
    return _shopeeNajumiRenovarToken(refresh, parseInt(shopId));
  }
  return { token, shopId };
}

function _shopeeNajumiRenovarToken(refreshToken, shopId) {
  const pkey = _shopeePartnerKey();
  const path = '/api/v2/auth/access_token/get';
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _shopeeSign(SHOPEE_PARTNER_ID + path + ts, pkey);

  const resp = UrlFetchApp.fetch(
    SW_NAJUMI_BASE + path
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
  if (data.error) throw new Error('Falha ao renovar token Najumi: ' + (data.message || JSON.stringify(data)));

  const props     = PropertiesService.getScriptProperties();
  const expiresAt = (Math.floor(Date.now() / 1000) + (data.expire_in || 14400)).toString();
  props.setProperty('SW_NAJUMI_ACCESS_TOKEN',  data.access_token);
  props.setProperty('SW_NAJUMI_REFRESH_TOKEN', data.refresh_token || refreshToken);
  props.setProperty('SW_NAJUMI_TOKEN_EXPIRES', expiresAt);

  return { token: data.access_token, shopId: String(shopId) };
}

// ── GET autenticado (loja Najumi) ───────────────────────────
function _shopeeNajumiGet(path, params) {
  const { token, shopId } = _shopeeNajumiGetToken();
  const pkey = _shopeePartnerKey();
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _shopeeSign(
    SHOPEE_PARTNER_ID + path + ts + token + parseInt(shopId),
    pkey
  );

  const base = SW_NAJUMI_BASE + path
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
// DIAGNÓSTICO — rode direto no editor GAS, veja Registro de execução.
// Mostra a resposta CRUA da API de transações da carteira, para
// confirmarmos os nomes reais dos campos antes de gravar na planilha.
// ============================================================
function sw_testarWalletAPI() {
  const seteDiasAtras = Math.floor(Date.now() / 1000) - 7 * 24 * 3600;
  const agora          = Math.floor(Date.now() / 1000);

  try {
    const r = _shopeeNajumiGet('/api/v2/payment/get_wallet_transaction_list', {
      create_time_from: seteDiasAtras,
      create_time_to:   agora,
      page_size:        10,
    });
    Logger.log('=== RESPOSTA CRUA get_wallet_transaction_list ===');
    Logger.log(JSON.stringify(r, null, 2));
  } catch (e) {
    Logger.log('ERRO: ' + e.message);
  }
}

// ============================================================
// AINDA NÃO IMPLEMENTADO
// A sincronização real (paginação, campos, gravação na planilha) só
// pode ser escrita com segurança depois de ver a resposta REAL da API
// nesta conta/versão — os nomes de campo mudam entre versões da Shopee
// Open Platform e não vou arriscar código baseado em suposição.
//
// Rode "🧪 Testar API da carteira (log)" no menu, copie o Registro de
// execução e mande — a partir disso essas 3 funções são implementadas
// de uma vez: sincronização, criação da planilha e trigger automático.
// ============================================================
// ============================================================
// SINCRONIZAÇÃO — grava transações novas da carteira na aba
// "Transações", sem duplicar (chave: transaction_id) e sem
// reprocessar o histórico inteiro a cada vez (retoma de onde parou).
// ============================================================
const SW_ABA_TRANSACOES      = 'Transações';
const SW_PROP_ULTIMO_SYNC    = 'SW_NAJUMI_ULTIMO_SYNC_TS';
const SW_JANELA_SEGUNDOS     = 15 * 24 * 3600; // limite seguro por chamada
const SW_PERIODO_INICIAL_DIAS = 90;            // 1ª sincronização (sem histórico salvo)
const SW_MARGEM_SEGURANCA_S  = 300;            // reprocessa 5 min pra trás, dedup cobre o resto

function _sw_abaTransacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(SW_ABA_TRANSACOES);
  if (!aba) {
    aba = ss.insertSheet(SW_ABA_TRANSACOES);
    aba.appendRow(['Data', 'Pedido', 'Tipo', 'Descrição', 'Fluxo', 'Valor', 'Status', 'Saldo Após', 'Comprador', 'Transaction ID']);
    aba.setFrozenRows(1);
  }
  return aba;
}

function _sw_idsExistentes(aba) {
  const last = aba.getLastRow();
  if (last < 2) return new Set();
  const vals = aba.getRange(2, 10, last - 1, 1).getValues(); // coluna J = Transaction ID
  return new Set(vals.map(r => String(r[0])));
}

function sw_sincronizarCarteira() {
  const props = PropertiesService.getScriptProperties();
  const agora = Math.floor(Date.now() / 1000);
  let desde = parseInt(props.getProperty(SW_PROP_ULTIMO_SYNC) || '0');
  desde = desde
    ? desde - SW_MARGEM_SEGURANCA_S
    : agora - SW_PERIODO_INICIAL_DIAS * 24 * 3600;

  const aba        = _sw_abaTransacoes();
  const existentes = _sw_idsExistentes(aba);
  const linhasNovas = [];
  let maxCreateTime = desde;

  let janelaInicio = desde;
  while (janelaInicio < agora) {
    const janelaFim = Math.min(janelaInicio + SW_JANELA_SEGUNDOS, agora);
    let pageNo = 1;

    while (true) {
      let r;
      try {
        r = _shopeeNajumiGet('/api/v2/payment/get_wallet_transaction_list', {
          create_time_from: janelaInicio,
          create_time_to:   janelaFim,
          page_size:        100,
          page_no:          pageNo,
        });
      } catch (e) {
        Logger.log('Erro na janela ' + janelaInicio + '-' + janelaFim + ' página ' + pageNo + ': ' + e.message);
        break;
      }

      const lista = r.transaction_list || [];
      lista.forEach(t => {
        const id = String(t.transaction_id);
        if (t.create_time > maxCreateTime) maxCreateTime = t.create_time;
        if (existentes.has(id)) return;
        existentes.add(id);
        linhasNovas.push([
          new Date(t.create_time * 1000),
          t.order_sn || '',
          t.transaction_type || '',
          (t.description || '').trim(),
          t.money_flow === 'MONEY_IN' ? 'Entrada' : 'Saída',
          t.amount || 0,
          t.status || '',
          t.current_balance || '',
          t.buyer_name || '',
          id,
        ]);
      });

      if (!r.more) break;
      pageNo++;
      Utilities.sleep(250);
    }

    janelaInicio = janelaFim;
    Utilities.sleep(250);
  }

  if (linhasNovas.length) {
    linhasNovas.sort((a, b) => a[0] - b[0]);
    aba.getRange(aba.getLastRow() + 1, 1, linhasNovas.length, linhasNovas[0].length).setValues(linhasNovas);
  }
  props.setProperty(SW_PROP_ULTIMO_SYNC, String(maxCreateTime));

  const msg = linhasNovas.length + ' transação(ões) nova(s) lançada(s) na aba "' + SW_ABA_TRANSACOES + '".';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* rodando via trigger, sem UI */ }
  return msg;
}

// ── Trigger automático (roda a sincronização sozinha) ───────
function sw_ativarTriggerAutomatico() {
  _sw_removerTriggers();
  ScriptApp.newTrigger('sw_sincronizarCarteira')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert('✅ Sincronização automática ativada — roda a cada 1 hora.\nVocê também pode continuar usando "🔄 Sincronizar carteira agora" a qualquer momento.');
}

function sw_desativarTriggerAutomatico() {
  const removidos = _sw_removerTriggers();
  SpreadsheetApp.getUi().alert(removidos > 0 ? '⏹️ Sincronização automática desativada.' : 'Nenhuma sincronização automática estava ativa.');
}

function _sw_removerTriggers() {
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sw_sincronizarCarteira');
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  return triggers.length;
}
