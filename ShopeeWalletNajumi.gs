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
    .addItem('🔢 Conferir total de transações (90d)', 'sw_contarTransacoesAPI')
    .addItem('📊 Analisar tipos de transação (log)', 'sw_analisarTipos')
    .addItem('🔁 Gerar resumo por pedido (múltiplas transações)', 'sw_resumoPorPedido')
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
// CONSTANTES E HELPERS COMPARTILHADOS (contagem + sincronização)
// ============================================================
const SW_ABA_TRANSACOES       = 'Transações';
const SW_PROP_ULTIMO_SYNC     = 'SW_NAJUMI_ULTIMO_SYNC_TS';
const SW_JANELA_SEGUNDOS      = 7 * 24 * 3600; // ponto de partida; se a API recusar por
                                                 // período grande, a largura é reduzida
const SW_PERIODO_INICIAL_DIAS = 90;             // 1ª sincronização (sem histórico salvo)
const SW_MARGEM_SEGURANCA_S   = 300;            // reprocessa 5 min pra trás, dedup cobre o resto
const SW_LIMITE_MS            = 5 * 60 * 1000;  // para antes do limite de 6 min do Apps Script

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

// Pagina uma janela [inicio, fim) por completo, aplicando onTransacao a
// cada item. Lança o erro adiante se a API recusar o período (deixa quem
// chama decidir se reduz a janela e tenta de novo).
function _sw_buscarJanelaPaginada(inicio, fim, onTransacao) {
  let pageNo = 1;
  while (true) {
    const r = _shopeeNajumiGet('/api/v2/payment/get_wallet_transaction_list', {
      create_time_from: inicio,
      create_time_to:   fim,
      page_size:        100,
      page_no:          pageNo,
    });
    (r.transaction_list || []).forEach(onTransacao);
    if (!r.more) return;
    pageNo++;
    Utilities.sleep(250);
  }
}

// ============================================================
// CONTAGEM — pagina a API só somando (sem gravar linhas, sem dedup) e
// compara com o total já lançado na planilha, para conferência.
// A API não expõe um total_count resumido, então isso ainda precisa
// paginar tudo — mas é mais leve que a sincronização real (sem escrita
// na planilha a cada item), e usa o mesmo esquema de lote +
// auto-continuação para não estourar os 6 min do Apps Script.
// ============================================================
const SW_PROP_CONT_CURSOR  = 'SW_NAJUMI_CONT_CURSOR';
const SW_PROP_CONT_TOTAL   = 'SW_NAJUMI_CONT_TOTAL';
const SW_PROP_CONT_LARGURA = 'SW_NAJUMI_CONT_LARGURA';

function sw_contarTransacoesAPI() {
  const props    = PropertiesService.getScriptProperties();
  const agora    = Math.floor(Date.now() / 1000);
  const inicioMs = Date.now();

  const cursorSalvo = props.getProperty(SW_PROP_CONT_CURSOR);
  let janelaInicio = cursorSalvo
    ? parseInt(cursorSalvo)
    : agora - SW_PERIODO_INICIAL_DIAS * 24 * 3600;
  let total = parseInt(props.getProperty(SW_PROP_CONT_TOTAL) || '0');
  let largura = parseInt(props.getProperty(SW_PROP_CONT_LARGURA) || '0') || SW_JANELA_SEGUNDOS;

  let pausouPorTempo = false;
  while (janelaInicio < agora) {
    if (Date.now() - inicioMs > SW_LIMITE_MS) { pausouPorTempo = true; break; }

    let janelaFim = Math.min(janelaInicio + largura, agora);
    let concluida = false;
    while (!concluida) {
      try {
        _sw_buscarJanelaPaginada(janelaInicio, janelaFim, () => { total++; });
        concluida = true;
      } catch (e) {
        if (/time period too large/i.test(e.message) && janelaFim - janelaInicio > 3600) {
          largura   = Math.floor(largura / 2);
          janelaFim = Math.min(janelaInicio + largura, agora);
        } else {
          Logger.log('Erro na janela ' + janelaInicio + '-' + janelaFim + ': ' + e.message);
          concluida = true;
        }
      }
      if (Date.now() - inicioMs > SW_LIMITE_MS) { pausouPorTempo = true; concluida = true; }
    }
    janelaInicio = janelaFim;
    Utilities.sleep(200);
  }

  props.setProperty(SW_PROP_CONT_LARGURA, String(largura));
  props.setProperty(SW_PROP_CONT_TOTAL,   String(total));

  if (pausouPorTempo && janelaInicio < agora) {
    props.setProperty(SW_PROP_CONT_CURSOR, String(janelaInicio));
    ScriptApp.newTrigger('sw_contarTransacoesAPI').timeBased().after(10 * 1000).create();
    Logger.log('Contagem parcial: ' + total + ' até agora — continuando sozinha em ~10s.');
    return;
  }

  props.deleteProperty(SW_PROP_CONT_CURSOR);
  props.deleteProperty(SW_PROP_CONT_TOTAL);
  props.deleteProperty(SW_PROP_CONT_LARGURA);

  const naPlanilha = _sw_abaTransacoes().getLastRow() - 1; // -1 = descarta cabeçalho
  const msg = 'Total na API (90 dias): ' + total + '  |  Total na planilha: ' + naPlanilha +
    (total === naPlanilha ? '  ✅ BATE' : '  ⚠ DIFERENTE');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* rodando via trigger, sem UI */ }
}

// ============================================================
// SINCRONIZAÇÃO — grava transações novas da carteira na aba
// "Transações", sem duplicar (chave: transaction_id) e sem
// reprocessar o histórico inteiro a cada vez (retoma de onde parou).
// ============================================================
const SW_PROP_CURSOR  = 'SW_NAJUMI_CURSOR_SYNC';    // ponto onde uma sincro em andamento parou
const SW_PROP_LARGURA = 'SW_NAJUMI_LARGURA_JANELA'; // largura de janela descoberta, reaproveitada

// Sincroniza em lotes: para antes do limite de execução do Apps Script e,
// se ainda não terminou, agenda um trigger de continuação (~10 s) que
// chama esta mesma função de novo, retomando de onde parou — sem
// precisar que o usuário clique de novo.
function sw_sincronizarCarteira() {
  const props    = PropertiesService.getScriptProperties();
  const agora    = Math.floor(Date.now() / 1000);
  const inicioMs = Date.now();

  const cursorSalvo = props.getProperty(SW_PROP_CURSOR);
  let janelaInicio;
  if (cursorSalvo) {
    janelaInicio = parseInt(cursorSalvo);
  } else {
    const ultimoSync = parseInt(props.getProperty(SW_PROP_ULTIMO_SYNC) || '0');
    janelaInicio = ultimoSync
      ? ultimoSync - SW_MARGEM_SEGURANCA_S
      : agora - SW_PERIODO_INICIAL_DIAS * 24 * 3600;
  }
  let maxCreateTime = janelaInicio;

  const aba        = _sw_abaTransacoes();
  const existentes = _sw_idsExistentes(aba);
  const linhasNovas = [];

  const onTransacao = t => {
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
  };

  // Largura da janela persiste nas Propriedades entre execuções (e entre
  // continuações do mesmo lote): só diminui quando a API recusa por
  // período grande demais, nunca redescobre o limite do zero.
  let largura = parseInt(props.getProperty(SW_PROP_LARGURA) || '0') || SW_JANELA_SEGUNDOS;
  let pausouPorTempo = false;

  while (janelaInicio < agora) {
    if (Date.now() - inicioMs > SW_LIMITE_MS) { pausouPorTempo = true; break; }

    let janelaFim = Math.min(janelaInicio + largura, agora);
    let concluida = false;
    while (!concluida) {
      try {
        _sw_buscarJanelaPaginada(janelaInicio, janelaFim, onTransacao);
        concluida = true;
      } catch (e) {
        if (/time period too large/i.test(e.message) && janelaFim - janelaInicio > 3600) {
          largura   = Math.floor(largura / 2);
          janelaFim = Math.min(janelaInicio + largura, agora);
          Logger.log('Período grande demais, reduzindo janela para ' + Math.round(largura / 3600) + 'h.');
        } else {
          Logger.log('Erro na janela ' + janelaInicio + '-' + janelaFim + ': ' + e.message);
          concluida = true; // erro irrecuperável — não trava aqui, segue para a próxima janela
        }
      }
      if (Date.now() - inicioMs > SW_LIMITE_MS) { pausouPorTempo = true; concluida = true; }
    }
    janelaInicio = janelaFim;
    Utilities.sleep(200);
  }

  if (linhasNovas.length) {
    linhasNovas.sort((a, b) => a[0] - b[0]);
    aba.getRange(aba.getLastRow() + 1, 1, linhasNovas.length, linhasNovas[0].length).setValues(linhasNovas);
  }
  props.setProperty(SW_PROP_LARGURA, String(largura));

  if (pausouPorTempo && janelaInicio < agora) {
    props.setProperty(SW_PROP_CURSOR, String(janelaInicio));
    _sw_agendarContinuacao();
    const msg = linhasNovas.length + ' transação(ões) lançada(s) neste lote — sincronização vai continuar sozinha em ~10s (parou em ' + new Date(janelaInicio * 1000).toLocaleString('pt-BR') + ').';
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* trigger, sem UI */ }
    return msg;
  }

  // Terminou de verdade: consolida o ponto de retomada da próxima sincronização.
  props.deleteProperty(SW_PROP_CURSOR);
  props.setProperty(SW_PROP_ULTIMO_SYNC, String(maxCreateTime));

  const msg = linhasNovas.length + ' transação(ões) nova(s) lançada(s) na aba "' + SW_ABA_TRANSACOES + '". Sincronização concluída.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* rodando via trigger, sem UI */ }
  return msg;
}

// Agenda uma execução única de continuação em ~10s (trigger temporário,
// removido pelo Apps Script sozinho depois de disparar).
function _sw_agendarContinuacao() {
  ScriptApp.newTrigger('sw_sincronizarCarteira')
    .timeBased()
    .after(10 * 1000)
    .create();
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

// ============================================================
// ANÁLISE — agrupa a aba "Transações" por Tipo (coluna C), mostrando
// quantidade, soma de valor, entradas vs saídas e um exemplo de
// descrição por tipo. Só leitura, não altera nada.
// ============================================================
function sw_analisarTipos() {
  const aba = _sw_abaTransacoes();
  const last = aba.getLastRow();
  if (last < 2) {
    Logger.log('Aba "Transações" está vazia.');
    return;
  }

  // Colunas: A Data | B Pedido | C Tipo | D Descrição | E Fluxo | F Valor
  const dados = aba.getRange(2, 1, last - 1, 6).getValues();
  const grupos = {}; // tipo -> { qtd, entradas, saidas, somaValor, exemploDescricao }

  dados.forEach(r => {
    const tipo   = String(r[2] || '(vazio)');
    const desc   = String(r[3] || '');
    const fluxo  = String(r[4] || '');
    const valor  = Number(r[5]) || 0;

    if (!grupos[tipo]) {
      grupos[tipo] = { qtd: 0, entradas: 0, saidas: 0, somaValor: 0, exemploDescricao: desc };
    }
    const g = grupos[tipo];
    g.qtd++;
    g.somaValor += valor;
    if (fluxo === 'Entrada') g.entradas++;
    if (fluxo === 'Saída')   g.saidas++;
  });

  const linhas = Object.entries(grupos)
    .sort((a, b) => b[1].qtd - a[1].qtd)
    .map(([tipo, g]) =>
      tipo + '  |  qtd: ' + g.qtd +
      '  |  entradas: ' + g.entradas + ' / saídas: ' + g.saidas +
      '  |  soma valor: ' + g.somaValor.toFixed(2) +
      '  |  ex: "' + g.exemploDescricao.slice(0, 60) + '"'
    );

  Logger.log('=== TIPOS DE TRANSAÇÃO (' + dados.length + ' linhas analisadas) ===');
  Logger.log(linhas.join('\n'));
}

// ============================================================
// RESUMO POR PEDIDO — agrupa a aba "Transações" por Pedido (order_sn)
// e cria/atualiza a aba "Resumo por Pedido" só com os pedidos que
// tiveram MAIS DE UMA transação (ex: recebeu a venda e depois teve um
// desconto/estorno por devolução) — mostra o saldo final de cada um.
// ============================================================
const SW_ABA_RESUMO_PEDIDO = 'Resumo por Pedido';

// Tipos que indicam a VENDA FOI DESFEITA (devolução/reembolso) — a
// receita original não é lucro real nesse caso.
const SW_TIPOS_DEVOLUCAO = new Set([
  'ADJUSTMENT_FOR_RR_AFTER_ESCROW_VERIFIED',
  'ESCROW_VERIFIED_MINUS',
  'RETURN_COMPENSATION_SERVICE_ADD',
]);
// Tipos que são taxa/ajuste sobre uma venda que de fato aconteceu
// (imposto, taxa de programa, correção administrativa) — não são
// devolução, só reduzem/ajustam o valor líquido recebido.
const SW_TIPOS_TAXA_AJUSTE = new Set([
  'ADJUSTMENT_CENTER_DEDUCT',
  'FBS_FEE_CHARGE_MINUS',
  'ADJUSTMENT_CENTER_ADD',
  'SELLER_COMPENSATE_ADD',
]);

function _sw_classificarMotivo(tipos) {
  const outros = Array.from(tipos).filter(t => t !== 'ESCROW_VERIFIED_ADD');
  if (outros.some(t => SW_TIPOS_DEVOLUCAO.has(t)))   return '↩️ DEVOLUÇÃO (venda desfeita)';
  if (outros.some(t => SW_TIPOS_TAXA_AJUSTE.has(t))) return '💸 TAXA/AJUSTE (venda válida)';
  return '❓ OUTRO';
}

function sw_resumoPorPedido() {
  const abaOrigem = _sw_abaTransacoes();
  const last = abaOrigem.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert('Aba "Transações" está vazia. Sincronize primeiro.');
    return;
  }

  // Colunas: A Data | B Pedido | C Tipo | D Descrição | E Fluxo | F Valor
  const dados = abaOrigem.getRange(2, 1, last - 1, 6).getValues();
  const porPedido = {}; // order_sn -> { qtd, saldo, positivo, negativo, tipos:Set, primeiraData, ultimaData }

  dados.forEach(r => {
    const data   = r[0];
    const pedido = String(r[1] || '').trim();
    const tipo   = String(r[2] || '(vazio)');
    const valor  = Number(r[5]) || 0;
    if (!pedido) return; // ignora transações sem pedido (saque, recarga de ads, etc.)

    if (!porPedido[pedido]) {
      porPedido[pedido] = { qtd: 0, saldo: 0, positivo: 0, negativo: 0, tipos: new Set(), primeiraData: data, ultimaData: data };
    }
    const g = porPedido[pedido];
    g.qtd++;
    g.saldo += valor;
    if (valor > 0) g.positivo += valor; else g.negativo += valor;
    g.tipos.add(tipo);
    if (data < g.primeiraData) g.primeiraData = data;
    if (data > g.ultimaData)   g.ultimaData   = data;
  });

  // Só pedidos com mais de uma transação — é o padrão que indica
  // devolução/estorno/ajuste depois da receita original.
  const linhas = Object.entries(porPedido)
    .filter(([, g]) => g.qtd > 1)
    .map(([pedido, g]) => {
      const pctPerdido = g.positivo > 0 ? (-g.negativo / g.positivo * 100) : 0;
      const status = g.saldo <= 0 ? '🔴 PREJUÍZO TOTAL' : (g.negativo < 0 ? '🟡 PARCIAL' : '🟢 OK');
      return [
        pedido,
        g.qtd,
        Math.round(g.saldo * 100) / 100,
        Math.round(g.positivo * 100) / 100,
        Math.round(g.negativo * 100) / 100,
        Math.round(pctPerdido * 100) / 100,
        status,
        _sw_classificarMotivo(g.tipos),
        Array.from(g.tipos).join(', '),
        g.primeiraData,
        g.ultimaData,
      ];
    })
    .sort((a, b) => a[2] - b[2]); // saldo final ascendente — piores casos primeiro

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let abaResumo = ss.getSheetByName(SW_ABA_RESUMO_PEDIDO);
  if (abaResumo) {
    abaResumo.clearContents();
  } else {
    abaResumo = ss.insertSheet(SW_ABA_RESUMO_PEDIDO);
  }
  abaResumo.appendRow(['Pedido', 'Qtd Transações', 'Saldo Final', 'Receita Bruta', 'Total Descontado', '% Perdido', 'Status', 'Motivo', 'Tipos Envolvidos', 'Primeira Data', 'Última Data']);
  abaResumo.setFrozenRows(1);
  if (linhas.length) {
    abaResumo.getRange(2, 1, linhas.length, linhas[0].length).setValues(linhas);
  }

  const totalPedidosMultiplos = linhas.length;
  const prejuizoTotal = linhas.filter(l => l[2] <= 0).length;
  const msg = totalPedidosMultiplos + ' pedido(s) com mais de uma transação encontrados (' +
    prejuizoTotal + ' com saldo final <= 0). Aba "' + SW_ABA_RESUMO_PEDIDO + '" atualizada.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* sem UI */ }
}
