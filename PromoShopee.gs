// ============================================================
// PROMO SHOPEE — Adicionar itens à promoção ativa (multi-conta)
// Contas: Humble | Najumi | Sky
//
// Estrutura da planilha (mesma da "Comparação de Promos Shopee"):
//   Linha 1: título da aba  |  Linha 2: cabeçalho  |  Linha 3+: dados
//   Col A: ID do produto (item_id)
//   Col D: ID de variação (model_id) — vazio para produto simples
//   Col H: Preço de desconto (promo a aplicar)
//   Col I: Limite de compra por pedido (0 = sem limite)
//   Col J: STATUS — filtra "FALTA NA VIGENTE", grava resultado
//
// Propriedades do script (Extensões → Apps Script → ⚙️ Configurações):
//   SHOPEE_PARTNER_KEY          → Partner Key do app (shpk6c57...)
//   HUMBLE_DISCOUNT_ID          → ID da promoção ativa — Humble
//   NAJUMI_DISCOUNT_ID          → ID da promoção ativa — Najumi
//   SKY_DISCOUNT_ID             → ID da promoção ativa — Sky
//   HUMBLE_ACCESS_TOKEN  + _REFRESH_TOKEN + _SHOP_ID + _TOKEN_EXPIRES
//   NAJUMI_ACCESS_TOKEN  + ...
//   SKY_ACCESS_TOKEN     + ...     (todos gerados automaticamente pelo menu)
// ============================================================

const PS_PARTNER_ID = 2037491;
const PS_BASE       = 'https://partner.shopeemobile.com';
const PS_STORES     = ['Humble', 'Najumi', 'Sky'];

// Colunas (0-based) — ajuste se a planilha mudar
const PS_COL = { ITEM: 0, MODEL: 3, PRECO: 7, LIMITE: 8, STATUS: 9 };
const PS_HDR = 2;                      // linhas de cabeçalho antes dos dados
const PS_PENDING = 'FALTA NA VIGENTE'; // valor de STATUS que será processado

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  const ui   = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📤 Promos Shopee');
  PS_STORES.forEach(s =>
    menu.addItem('🔑 Autorizar ' + s + ' (OAuth)', 'ps_auth_' + s.toLowerCase())
  );
  menu.addSeparator()
      .addItem('📋 Importar tokens existentes (todas as lojas)', 'ps_importarTokens')
      .addSeparator()
      .addItem('🎯 Configurar IDs de Promoção', 'ps_configurarPromos')
      .addSeparator()
      .addItem('📤 Enviar itens — aba atual', 'ps_enviarAbaAtual')
      .addItem('📤 Enviar itens — TODAS as lojas', 'ps_enviarTodas')
      .addToUi();
}

// Wrappers necessários pois o menu não aceita parâmetros
function ps_auth_humble() { _ps_iniciarAuth('Humble'); }
function ps_auth_najumi() { _ps_iniciarAuth('Najumi'); }
function ps_auth_sky()    { _ps_iniciarAuth('Sky');    }

// ============================================================
// AUTORIZAÇÃO OAuth por loja
// ============================================================
function _ps_iniciarAuth(loja) {
  const ui   = SpreadsheetApp.getUi();
  const pkey = PropertiesService.getScriptProperties().getProperty('SHOPEE_PARTNER_KEY');
  if (!pkey) {
    ui.alert('❌ SHOPEE_PARTNER_KEY não configurada.\n\n' +
             'Extensões → Apps Script → ⚙️ Configurações → Propriedades do script');
    return;
  }
  const ts   = Math.floor(Date.now() / 1000);
  const path = '/api/v2/shop/auth_partner';
  const sign = _ps_sign(PS_PARTNER_ID + path + ts, pkey);
  const url  = PS_BASE + path
    + '?partner_id=' + PS_PARTNER_ID + '&timestamp=' + ts + '&sign=' + sign
    + '&redirect=' + encodeURIComponent('https://localhost');

  ui.showModalDialog(
    HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;font-size:13px;padding:12px;line-height:1.9">' +
      '<p><b>1.</b> Clique no link e faça login com a conta <b>Shopee ' + loja + '</b>:</p>' +
      '<p><a href="' + url + '" target="_blank" style="background:#e65100;color:#fff;' +
      'padding:8px 18px;border-radius:4px;text-decoration:none;font-size:14px">' +
      '👉 Autorizar Shopee ' + loja + '</a></p>' +
      '<p><b>2.</b> Após autorizar, o browser abrirá página de erro (localhost) — é normal.</p>' +
      '<p><b>3.</b> Copie a URL completa da barra de endereço e feche esta janela.</p>' +
      '<p><b>4.</b> No menu: <b>📤 Promos Shopee → 🔑 Autorizar ' + loja + '</b> novamente ' +
      'e cole a URL na caixa de diálogo seguinte.</p>' +
      '</div>'
    ).setWidth(500).setHeight(270),
    '🔑 Autorizar Shopee ' + loja + ' — Passo 1'
  );
}

// Segundo passo: exibe dialog para colar a URL e salvar o token
function ps_salvarToken() {
  const loja = SpreadsheetApp.getActiveSheet().getName();
  if (!PS_STORES.includes(loja)) {
    SpreadsheetApp.getUi().alert('Abra a aba da loja (Humble, Najumi ou Sky) antes de executar.');
    return;
  }
  _ps_dialogToken(loja);
}

function _ps_dialogToken(loja) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial;font-size:13px;padding:14px;margin:0}
      label{font-weight:bold;display:block;margin-top:12px}
      input{width:100%;padding:7px;box-sizing:border-box;font-size:13px;
            border:1px solid #ccc;border-radius:3px}
      button{margin-top:16px;background:#e65100;color:#fff;border:none;
             padding:10px 26px;font-size:14px;border-radius:4px;cursor:pointer}
      .hint{color:#888;font-size:11px;margin-top:3px}
    </style>
    <p>Cole a URL que apareceu após autorizar a conta <b>${loja}</b>:</p>
    <label>URL completa da barra de endereço</label>
    <input id="url" placeholder="https://localhost/?code=XXXX&shop_id=YYYYY"/>
    <p class="hint">Os campos abaixo são preenchidos automaticamente ao colar a URL.</p>
    <label>code</label>
    <input id="code" placeholder="6f4a1b2c3d..."/>
    <label>shop_id</label>
    <input id="sid" placeholder="123456789"/>
    <button onclick="enviar()">💾 Salvar Token ${loja}</button>
    <script>
      document.getElementById('url').addEventListener('input', function() {
        try {
          const p = new URL(this.value.trim()).searchParams;
          const c = p.get('code'), s = p.get('shop_id');
          if (c) document.getElementById('code').value = c;
          if (s) document.getElementById('sid').value  = s;
        } catch(e) {}
      });
      function enviar() {
        const code = document.getElementById('code').value.trim();
        const sid  = document.getElementById('sid').value.trim();
        if (!code || !sid) { alert('Preencha o code e o shop_id.'); return; }
        document.querySelector('button').textContent = '⏳ Salvando...';
        google.script.run
          .withSuccessHandler(m => { alert(m); google.script.host.close(); })
          .withFailureHandler(e => {
            alert('Erro: ' + e.message);
            document.querySelector('button').textContent = '💾 Salvar Token ${loja}';
          })
          .ps_trocarToken('${loja}', code, parseInt(sid));
      }
    </script>
  `).setWidth(480).setHeight(370);
  SpreadsheetApp.getUi().showModalDialog(html, '💾 Salvar Token — ' + loja);
}

function ps_trocarToken(loja, code, shopId) {
  const pkey   = _ps_partnerKey();
  const path   = '/api/v2/auth/token/get';
  const ts     = Math.floor(Date.now() / 1000);
  const sign   = _ps_sign(PS_PARTNER_ID + path + ts, pkey);
  const resp   = UrlFetchApp.fetch(
    PS_BASE + path + '?partner_id=' + PS_PARTNER_ID + '&timestamp=' + ts + '&sign=' + sign,
    { method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({ code, shop_id: shopId, partner_id: PS_PARTNER_ID }) }
  );
  const data = JSON.parse(resp.getContentText());
  if (data.error && data.error !== '') throw new Error(data.message || JSON.stringify(data));

  const prefix = loja.toUpperCase() + '_';
  const props  = PropertiesService.getScriptProperties();
  props.setProperty(prefix + 'ACCESS_TOKEN',  data.access_token);
  props.setProperty(prefix + 'REFRESH_TOKEN', data.refresh_token);
  props.setProperty(prefix + 'SHOP_ID',       String(shopId));
  props.setProperty(prefix + 'TOKEN_EXPIRES', String(Math.floor(Date.now() / 1000) + (data.expire_in || 14400)));
  return '✅ Token da conta ' + loja + ' salvo! Shop ID: ' + shopId;
}

// ============================================================
// CONFIGURAR IDs DE PROMOÇÃO
// ============================================================
function ps_configurarPromos() {
  const props = PropertiesService.getScriptProperties();
  const vals  = PS_STORES.map(s => ({
    s, id: props.getProperty(s.toUpperCase() + '_DISCOUNT_ID') || ''
  }));

  const rows = vals.map(v =>
    `<label>${v.s} — Discount ID</label>` +
    `<input id="${v.s}" value="${v.id}" placeholder="ex: 123456789"/>`
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial;font-size:13px;padding:14px;margin:0}
      label{font-weight:bold;display:block;margin-top:12px}
      input{width:100%;padding:7px;box-sizing:border-box;font-size:13px;
            border:1px solid #ccc;border-radius:3px}
      button{margin-top:16px;background:#1565c0;color:#fff;border:none;
             padding:10px 26px;font-size:14px;border-radius:4px;cursor:pointer}
      .hint{color:#888;font-size:11px;margin-top:6px}
    </style>
    <p>Informe o <b>Discount ID</b> da promoção <b>ativa</b> de cada loja.</p>
    <p class="hint">Onde encontrar: Central de Marketing → sua promoção → a URL contém o ID.<br>
    Ou use a API: menu → Listar promoções ativas (se disponível).</p>
    ${rows}
    <button onclick="salvar()">💾 Salvar IDs</button>
    <script>
      function salvar() {
        const ids = {};
        ${vals.map(v => `ids['${v.s}'] = document.getElementById('${v.s}').value.trim();`).join('')}
        document.querySelector('button').textContent = '⏳ Salvando...';
        google.script.run
          .withSuccessHandler(m => { alert(m); google.script.host.close(); })
          .withFailureHandler(e => alert('Erro: ' + e.message))
          .ps_salvarIds(ids);
      }
    </script>
  `).setWidth(420).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, '🎯 IDs de Promoção Ativa');
}

function ps_salvarIds(ids) {
  const props = PropertiesService.getScriptProperties();
  PS_STORES.forEach(s => {
    const v = (ids[s] || '').trim();
    if (v) props.setProperty(s.toUpperCase() + '_DISCOUNT_ID', v);
  });
  return '✅ IDs salvos: ' +
    PS_STORES.filter(s => ids[s]).map(s => s + '=' + ids[s]).join(', ');
}

// ============================================================
// IMPORTAR TOKENS EXISTENTES (sem refazer OAuth)
// Use quando os tokens já estão salvos em outra planilha.
// Cole ACCESS_TOKEN, REFRESH_TOKEN e SHOP_ID de cada loja.
// ============================================================
function ps_importarTokens() {
  const props = PropertiesService.getScriptProperties();

  // Monta estado atual para pré-preencher os campos
  const estado = PS_STORES.map(s => {
    const p = s.toUpperCase() + '_';
    const shopId  = props.getProperty(p + 'SHOP_ID')  || '';
    const expires = parseInt(props.getProperty(p + 'TOKEN_EXPIRES') || '0');
    const now     = Math.floor(Date.now() / 1000);
    const ok      = shopId && expires > now;
    return { s, shopId, ok };
  });

  const linhas = PS_STORES.map(s =>
    `<div class="bloco">
      <div class="titulo">${s}
        <span id="st_${s}" class="tag"></span>
      </div>
      <label>ACCESS_TOKEN</label>
      <input id="${s}_at" placeholder="cole aqui o SHOPEE_ACCESS_TOKEN da planilha ${s}"/>
      <label>REFRESH_TOKEN</label>
      <input id="${s}_rt" placeholder="cole aqui o SHOPEE_REFRESH_TOKEN da planilha ${s}"/>
      <label>SHOP_ID</label>
      <input id="${s}_sid" type="number" placeholder="número (ex: 123456789)"/>
    </div>`
  ).join('');

  const initScript = PS_STORES.map(s => {
    const p   = s.toUpperCase() + '_';
    const sid = props.getProperty(p + 'SHOP_ID') || '';
    const exp = parseInt(props.getProperty(p + 'TOKEN_EXPIRES') || '0');
    const now = Math.floor(Date.now() / 1000);
    const ok  = sid && exp > now;
    return `document.getElementById('${s}_sid').value='${sid}';` +
           `document.getElementById('st_${s}').textContent='${ok ? "✅ token ativo" : sid ? "⚠️ expirado" : "❌ não configurado"}';` +
           `document.getElementById('st_${s}').style.color='${ok ? "#0a6e36" : sid ? "#b85c00" : "#c00"}';`;
  }).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial;font-size:12px;padding:14px;margin:0;overflow-y:auto}
      .bloco{border:1px solid #ddd;border-radius:6px;padding:10px 12px;margin-bottom:12px}
      .titulo{font-weight:bold;font-size:13px;margin-bottom:8px;color:#1a1a1a}
      .tag{font-weight:normal;font-size:11px;margin-left:8px}
      label{display:block;font-weight:bold;margin-top:8px;margin-bottom:2px;font-size:11px;color:#444}
      input{width:100%;padding:5px 7px;box-sizing:border-box;font-size:11px;
            border:1px solid #ccc;border-radius:3px;font-family:monospace}
      .hint{color:#888;font-size:11px;margin-bottom:10px;line-height:1.5}
      button{margin-top:12px;background:#1565c0;color:#fff;border:none;
             padding:9px 24px;font-size:13px;border-radius:4px;cursor:pointer;width:100%}
    </style>
    <p class="hint">
      Cole os valores de <b>SHOPEE_ACCESS_TOKEN</b>, <b>SHOPEE_REFRESH_TOKEN</b> e
      <b>SHOPEE_SHOP_ID</b> de cada planilha existente.<br>
      Vá na planilha da loja → Extensões → Apps Script → ⚙️ Configurações → Propriedades do script.
    </p>
    ${linhas}
    <button onclick="salvar()">💾 Salvar tokens das 3 lojas</button>
    <script>
      ${initScript}
      function salvar() {
        const dados = {};
        ${PS_STORES.map(s =>
          `dados['${s}']={at:document.getElementById('${s}_at').value.trim(),` +
          `rt:document.getElementById('${s}_rt').value.trim(),` +
          `sid:document.getElementById('${s}_sid').value.trim()};`
        ).join('')}
        document.querySelector('button').textContent = '⏳ Salvando...';
        google.script.run
          .withSuccessHandler(m => { alert(m); google.script.host.close(); })
          .withFailureHandler(e => {
            alert('Erro: ' + e.message);
            document.querySelector('button').textContent = '💾 Salvar tokens das 3 lojas';
          })
          .ps_salvarTokensImportados(dados);
      }
    </script>
  `).setWidth(520).setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Importar Tokens Existentes');
}

function ps_salvarTokensImportados(dados) {
  const props = PropertiesService.getScriptProperties();
  // TOKEN_EXPIRES: definimos 4h a partir de agora como estimativa conservadora
  // O token será renovado automaticamente via refresh_token quando necessário
  const expires = String(Math.floor(Date.now() / 1000) + 14400);
  const salvos  = [];

  PS_STORES.forEach(s => {
    const d = dados[s] || {};
    const prefix = s.toUpperCase() + '_';
    if (d.at && d.rt && d.sid) {
      props.setProperty(prefix + 'ACCESS_TOKEN',  d.at);
      props.setProperty(prefix + 'REFRESH_TOKEN', d.rt);
      props.setProperty(prefix + 'SHOP_ID',       d.sid);
      props.setProperty(prefix + 'TOKEN_EXPIRES', expires);
      salvos.push(s + ' (shop_id=' + d.sid + ')');
    }
  });

  if (!salvos.length) throw new Error('Nenhum token foi preenchido. Informe ao menos ACCESS_TOKEN, REFRESH_TOKEN e SHOP_ID para uma loja.');
  return '✅ Tokens salvos: ' + salvos.join(', ');
}

// ============================================================
// ENVIAR ITENS À PROMOÇÃO
// ============================================================
function ps_enviarAbaAtual() {
  const loja = SpreadsheetApp.getActiveSheet().getName();
  if (!PS_STORES.includes(loja)) {
    SpreadsheetApp.getUi().alert('Abra a aba de uma loja (Humble, Najumi ou Sky) antes de executar.');
    return;
  }
  _ps_processarLoja(loja);
}

function ps_enviarTodas() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '📤 Enviar para todas as lojas',
    'Vai adicionar itens com status "' + PS_PENDING + '" nas abas Humble, Najumi e Sky.\n\nConfirmar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  const erros = [];
  PS_STORES.forEach(s => {
    try { _ps_processarLoja(s); }
    catch(e) { erros.push(s + ': ' + e.message); }
  });
  ui.alert(erros.length ? '⚠️ Erros:\n' + erros.join('\n') : '✅ Concluído para todas as lojas!');
}

function _ps_processarLoja(loja) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(loja);
  if (!aba) throw new Error('Aba "' + loja + '" não encontrada.');

  const props      = PropertiesService.getScriptProperties();
  const discountId = parseInt(props.getProperty(loja.toUpperCase() + '_DISCOUNT_ID') || '0');
  if (!discountId) throw new Error(
    'Discount ID da loja ' + loja + ' não configurado.\nUse o menu 🎯 Configurar IDs de Promoção.'
  );

  const lastRow = aba.getLastRow();
  if (lastRow <= PS_HDR) { ui.alert(loja + ': sem dados na aba.'); return; }

  const NUM_COLS  = Math.max(...Object.values(PS_COL)) + 1; // colunas suficientes
  const allData   = aba.getRange(PS_HDR + 1, 1, lastRow - PS_HDR, NUM_COLS).getValues();
  const statusCol = PS_COL.STATUS + 1; // 1-based para getRange

  // Filtra pendentes
  const pendentes = allData
    .map((row, i) => ({ i, row }))
    .filter(({ row }) => String(row[PS_COL.STATUS] || '').trim() === PS_PENDING);

  if (!pendentes.length) {
    ui.alert(loja + ': nenhum item com status "' + PS_PENDING + '".');
    return;
  }

  // Agrupa models pelo mesmo item_id
  const itemMap = {}; // itemId(string) → { purchase_limit, preco, models, rowIdxs, temModel }
  pendentes.forEach(({ i, row }) => {
    const itemId  = String(row[PS_COL.ITEM]  || '').trim();
    const modelId = String(row[PS_COL.MODEL] || '').trim();
    const preco   = _ps_parsePreco(row[PS_COL.PRECO]);
    const limite  = Number(row[PS_COL.LIMITE] || 0);
    if (!itemId || preco <= 0) {
      aba.getRange(PS_HDR + 1 + i, statusCol).setValue('⚠️ item_id ou preço inválido');
      return;
    }
    if (!itemMap[itemId]) {
      itemMap[itemId] = { purchase_limit: limite, preco, models: [], rowIdxs: [], temModel: false };
    }
    if (modelId) {
      itemMap[itemId].models.push({ model_id: parseInt(modelId), model_promotion_price: preco });
      itemMap[itemId].temModel = true;
    }
    itemMap[itemId].rowIdxs.push(i);
  });

  // Monta item_list com rastreamento de rowIdxs
  const itemList = Object.entries(itemMap).map(([itemId, info]) => {
    const entry = {
      item_id:        parseInt(itemId),
      purchase_limit: info.purchase_limit,
      _rowIdxs:       info.rowIdxs,
    };
    if (info.temModel) {
      entry.model_list = info.models;
    } else {
      entry.item_promotion_price = info.preco;
    }
    return entry;
  });

  let adicionados = 0, erros = 0;

  // Envia em lotes de até 50 items por chamada
  for (let i = 0; i < itemList.length; i += 50) {
    const lote      = itemList.slice(i, i + 50);
    const rowMap    = {};
    const payload   = lote.map(({ _rowIdxs, ...item }) => {
      rowMap[item.item_id] = _rowIdxs;
      return item;
    });

    let respData;
    try {
      respData = _ps_post(loja, '/api/v2/discount/add_discount_item', {
        discount_id: discountId,
        item_list:   payload,
      });
    } catch (e) {
      // Erro de transporte — marca todos do lote
      lote.forEach(({ item_id, _rowIdxs }) =>
        _rowIdxs.forEach(r =>
          aba.getRange(PS_HDR + 1 + r, statusCol).setValue('❌ Erro API: ' + e.message)
        )
      );
      erros += lote.length;
      continue;
    }

    // Monta mapa de erros por item_id
    const erroMap = {};
    (respData.error_list || []).forEach(e => {
      erroMap[e.item_id] = e.failed_reason || e.msg || String(e.error_code || 'Erro');
    });

    payload.forEach(item => {
      const msg    = erroMap[item.item_id];
      const rowIdxs = rowMap[item.item_id];
      rowIdxs.forEach(r =>
        aba.getRange(PS_HDR + 1 + r, statusCol)
           .setValue(msg ? '❌ ' + msg : '✅ Adicionado')
      );
      msg ? erros++ : adicionados++;
    });

    if (i + 50 < itemList.length) Utilities.sleep(300);
  }

  ui.alert(
    '📤 ' + loja + ' — Concluído\n\n' +
    '✅ Adicionados: ' + adicionados + ' item(ns)\n' +
    (erros > 0 ? '❌ Com erro: ' + erros + ' item(ns)\n(veja coluna STATUS para detalhes)' : '')
  );
}

// ============================================================
// HELPERS DE AUTENTICAÇÃO MULTI-LOJA
// ============================================================
function _ps_getToken(loja) {
  const prefix  = loja.toUpperCase() + '_';
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty(prefix + 'ACCESS_TOKEN');
  const refresh = props.getProperty(prefix + 'REFRESH_TOKEN');
  const shopId  = props.getProperty(prefix + 'SHOP_ID');
  const expires = parseInt(props.getProperty(prefix + 'TOKEN_EXPIRES') || '0');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) throw new Error(
    'Token da loja ' + loja + ' não encontrado.\nUse o menu 🔑 Autorizar ' + loja + '.'
  );
  if (now >= expires - 300) return _ps_renovarToken(loja, refresh, parseInt(shopId));
  return { token, shopId };
}

function _ps_renovarToken(loja, refreshToken, shopId) {
  const pkey   = _ps_partnerKey();
  const prefix = loja.toUpperCase() + '_';
  const path   = '/api/v2/auth/access_token/get';
  const ts     = Math.floor(Date.now() / 1000);
  const sign   = _ps_sign(PS_PARTNER_ID + path + ts, pkey);
  const resp   = UrlFetchApp.fetch(
    PS_BASE + path + '?partner_id=' + PS_PARTNER_ID + '&timestamp=' + ts + '&sign=' + sign,
    { method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({ refresh_token: refreshToken, shop_id: shopId, partner_id: PS_PARTNER_ID }) }
  );
  const data = JSON.parse(resp.getContentText());
  if (data.error && data.error !== '') throw new Error('Falha ao renovar token ' + loja + ': ' + data.message);

  const props = PropertiesService.getScriptProperties();
  props.setProperty(prefix + 'ACCESS_TOKEN',  data.access_token);
  props.setProperty(prefix + 'REFRESH_TOKEN', data.refresh_token || refreshToken);
  props.setProperty(prefix + 'TOKEN_EXPIRES', String(Math.floor(Date.now() / 1000) + (data.expire_in || 14400)));
  return { token: data.access_token, shopId: String(shopId) };
}

function _ps_post(loja, path, body) {
  const { token, shopId } = _ps_getToken(loja);
  const pkey = _ps_partnerKey();
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _ps_sign(PS_PARTNER_ID + path + ts + token + parseInt(shopId), pkey);
  const url  = PS_BASE + path
    + '?partner_id=' + PS_PARTNER_ID + '&timestamp=' + ts
    + '&access_token=' + token + '&shop_id=' + shopId + '&sign=' + sign;

  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (data.error && data.error !== '') throw new Error('[' + path + '] ' + (data.message || data.error));
  return data.response || data;
}

function _ps_partnerKey() {
  const k = PropertiesService.getScriptProperties().getProperty('SHOPEE_PARTNER_KEY');
  if (!k) throw new Error('SHOPEE_PARTNER_KEY não configurada nas propriedades do script.');
  return k;
}

function _ps_sign(msg, key) {
  return Utilities.computeHmacSha256Signature(msg, key)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function _ps_parsePreco(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/R\$\s*/g, '').replace(/\./g, '').replace(',', '.').trim();
  return parseFloat(s) || 0;
}
