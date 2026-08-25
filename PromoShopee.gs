// ============================================================
// PROMO SHOPEE — Adicionar itens à promoção ativa (multi-conta)
// Contas: Humble | Najumi | Sky — cada uma com app próprio
//
// Estrutura da planilha (mesma da "Comparação de Promos Shopee"):
//   Linha 1: título  |  Linha 2: cabeçalho  |  Linha 3+: dados
//   Col A: item_id   |  Col D: model_id (vazio = produto simples)
//   Col H: Preço de desconto (promo)  |  Col I: Limite/pedido
//   Col J: STATUS — processa "FALTA NA VIGENTE", grava resultado
//
// Propriedades do script — uma por loja (Extensões → Apps Script
// → ⚙️ Configurações → Propriedades do script):
//
//   HUMBLE_PARTNER_ID    → Live Partner_id   do app Humble
//   HUMBLE_PARTNER_KEY   → Live API Partner Key do app Humble
//   HUMBLE_ACCESS_TOKEN  → SHOPEE_ACCESS_TOKEN  da planilha Humble
//   HUMBLE_REFRESH_TOKEN → SHOPEE_REFRESH_TOKEN da planilha Humble
//   HUMBLE_SHOP_ID       → SHOPEE_SHOP_ID       da planilha Humble
//   HUMBLE_DISCOUNT_ID   → ID da promoção ativa — Humble
//   (idem para NAJUMI_ e SKY_)
// ============================================================

const PS_BASE    = 'https://partner.shopeemobile.com';
const PS_STORES  = ['Humble', 'Najumi', 'Sky'];

const PS_COL     = { ITEM: 0, MODEL: 3, PRECO: 7, LIMITE: 8, STATUS: 9 };
const PS_HDR     = 2;
const PS_PENDING = 'FALTA NA VIGENTE';

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  const ui   = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📤 Promos Shopee');
  PS_STORES.forEach(s =>
    menu.addItem('🔑 Autorizar ' + s + ' (OAuth)', 'ps_auth_' + s.toLowerCase())
  );
  menu.addSeparator()
      .addItem('📋 Importar credenciais das planilhas existentes', 'ps_importarCredenciais')
      .addSeparator()
      .addItem('🎯 Configurar IDs de Promoção', 'ps_configurarPromos')
      .addSeparator()
      .addItem('📤 Enviar itens — aba atual',      'ps_enviarAbaAtual')
      .addItem('📤 Enviar itens — TODAS as lojas', 'ps_enviarTodas')
      .addToUi();
}

function ps_auth_humble() { _ps_iniciarAuth('Humble'); }
function ps_auth_najumi() { _ps_iniciarAuth('Najumi'); }
function ps_auth_sky()    { _ps_iniciarAuth('Sky');    }

// ============================================================
// CREDENCIAIS POR LOJA
// ============================================================
function _ps_partnerId(loja) {
  const v = PropertiesService.getScriptProperties().getProperty(loja.toUpperCase() + '_PARTNER_ID');
  if (!v) throw new Error('PARTNER_ID da loja ' + loja + ' não configurado.\nUse o menu 📋 Importar credenciais.');
  return parseInt(v);
}

function _ps_partnerKey(loja) {
  const v = PropertiesService.getScriptProperties().getProperty(loja.toUpperCase() + '_PARTNER_KEY');
  if (!v) throw new Error('PARTNER_KEY da loja ' + loja + ' não configurada.\nUse o menu 📋 Importar credenciais.');
  return v;
}

// ============================================================
// AUTORIZAÇÃO OAuth por loja (abre link + salva token na mesma tela)
// ============================================================
function _ps_iniciarAuth(loja) {
  const ui = SpreadsheetApp.getUi();
  let pid, pkey;
  try {
    pid  = _ps_partnerId(loja);
    pkey = _ps_partnerKey(loja);
  } catch(e) {
    ui.alert('❌ ' + e.message +
      '\n\nPrimeiro configure o Partner_id e Partner Key da loja ' + loja +
      '\nusando o menu: 📋 Importar credenciais das planilhas existentes.');
    return;
  }
  const ts   = Math.floor(Date.now() / 1000);
  const path = '/api/v2/shop/auth_partner';
  const sign = _ps_sign(pid + path + ts, pkey);
  const url  = PS_BASE + path
    + '?partner_id=' + pid + '&timestamp=' + ts + '&sign=' + sign
    + '&redirect=' + encodeURIComponent('https://localhost');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial;font-size:13px;padding:14px;margin:0}
      .passo{margin-bottom:10px}
      b{color:#222}
      label{font-weight:bold;display:block;margin-top:12px;margin-bottom:3px}
      input{width:100%;padding:7px;box-sizing:border-box;font-size:12px;
            border:1px solid #ccc;border-radius:3px;font-family:monospace}
      button{margin-top:14px;background:#e65100;color:#fff;border:none;
             padding:10px;font-size:14px;border-radius:4px;cursor:pointer;width:100%}
      .hint{color:#888;font-size:11px;margin-top:3px}
    </style>
    <div class="passo"><b>Passo 1.</b> Clique no botão e faça login com a conta <b>Shopee ${loja}</b>:</div>
    <a href="${url}" target="_blank" style="display:block;background:#e65100;color:#fff;
       padding:10px;border-radius:4px;text-decoration:none;text-align:center;font-size:14px;font-weight:bold">
       👉 Autorizar Shopee ${loja}
    </a>
    <div class="passo" style="margin-top:10px"><b>Passo 2.</b> Após autorizar, o browser abrirá uma página de erro
    (localhost) — <b>é normal</b>. Copie a URL completa da barra de endereço.</div>
    <div class="passo"><b>Passo 3.</b> Cole a URL aqui embaixo e clique em Salvar:</div>
    <label>URL completa do redirect (localhost/?code=...&shop_id=...)</label>
    <input id="url" placeholder="https://localhost/?code=XXXX&shop_id=YYYYY"/>
    <p class="hint">Os campos abaixo são preenchidos automaticamente ao colar a URL.</p>
    <label>code</label>
    <input id="code" placeholder="código gerado pela Shopee"/>
    <label>shop_id</label>
    <input id="sid" type="number" placeholder="número"/>
    <button onclick="salvar()">💾 Salvar token ${loja}</button>
    <script>
      document.getElementById('url').addEventListener('input', function() {
        try {
          const p = new URL(this.value.trim()).searchParams;
          const c = p.get('code'), s = p.get('shop_id');
          if (c) document.getElementById('code').value = c;
          if (s) document.getElementById('sid').value  = s;
        } catch(e) {}
      });
      function salvar() {
        const code = document.getElementById('code').value.trim();
        const sid  = document.getElementById('sid').value.trim();
        if (!code || !sid) { alert('Cole a URL acima ou preencha o code e shop_id.'); return; }
        document.querySelector('button').textContent = '⏳ Salvando...';
        google.script.run
          .withSuccessHandler(m => { alert(m); google.script.host.close(); })
          .withFailureHandler(e => {
            alert('Erro: ' + e.message);
            document.querySelector('button').textContent = '💾 Salvar token ${loja}';
          })
          .ps_trocarToken('${loja}', code, parseInt(sid));
      }
    </script>
  `).setWidth(500).setHeight(520);
  ui.showModalDialog(html, '🔑 Autorizar Shopee ' + loja);
}

function ps_trocarToken(loja, code, shopId) {
  const pid    = _ps_partnerId(loja);
  const pkey   = _ps_partnerKey(loja);
  const path   = '/api/v2/auth/token/get';
  const ts     = Math.floor(Date.now() / 1000);
  const sign   = _ps_sign(pid + path + ts, pkey);
  const resp   = UrlFetchApp.fetch(
    PS_BASE + path + '?partner_id=' + pid + '&timestamp=' + ts + '&sign=' + sign,
    { method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({ code, shop_id: shopId, partner_id: pid }) }
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
// IMPORTAR CREDENCIAIS (principal fluxo de configuração)
// ============================================================
function ps_importarCredenciais() {
  const props = PropertiesService.getScriptProperties();

  const estado = PS_STORES.map(s => {
    const p      = s.toUpperCase() + '_';
    const shopId = props.getProperty(p + 'SHOP_ID') || '';
    const exp    = parseInt(props.getProperty(p + 'TOKEN_EXPIRES') || '0');
    const pid    = props.getProperty(p + 'PARTNER_ID') || '';
    const now    = Math.floor(Date.now() / 1000);
    const ok     = shopId && exp > now && pid;
    return { s, shopId, pid, ok, expired: shopId && exp <= now };
  });

  const blocos = PS_STORES.map(s => {
    const st = estado.find(e => e.s === s);
    const badge = st.ok ? '✅ configurado' : st.expired ? '⚠️ token expirado' : '❌ não configurado';
    const color = st.ok ? '#0a6e36' : st.expired ? '#b85c00' : '#c00';
    return `
    <div class="bloco">
      <div class="titulo">${s} <span style="font-weight:normal;font-size:11px;color:${color}">${badge}</span></div>
      <label>Live Partner_id <span class="hint">(número — ex: 2037491)</span></label>
      <input id="${s}_pid" type="number" value="${st.pid}" placeholder="Live Partner_id do app ${s}"/>
      <label>Live API Partner Key <span class="hint">(começa com shpk...)</span></label>
      <input id="${s}_pkey" placeholder="Live API Partner Key do app ${s}"/>
      <label>ACCESS_TOKEN <span class="hint">(SHOPEE_ACCESS_TOKEN da planilha ${s})</span></label>
      <input id="${s}_at" placeholder="cole o SHOPEE_ACCESS_TOKEN"/>
      <label>REFRESH_TOKEN <span class="hint">(SHOPEE_REFRESH_TOKEN da planilha ${s})</span></label>
      <input id="${s}_rt" placeholder="cole o SHOPEE_REFRESH_TOKEN"/>
      <label>SHOP_ID <span class="hint">(SHOPEE_SHOP_ID da planilha ${s})</span></label>
      <input id="${s}_sid" type="number" value="${st.shopId}" placeholder="número"/>
    </div>`;
  }).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial;font-size:12px;padding:12px;margin:0;overflow-y:auto}
      .bloco{border:1px solid #ddd;border-radius:6px;padding:10px 12px;margin-bottom:10px;background:#fafafa}
      .titulo{font-weight:bold;font-size:13px;margin-bottom:6px}
      label{display:block;font-weight:bold;margin-top:8px;margin-bottom:2px;color:#333}
      .hint{font-weight:normal;color:#999}
      input{width:100%;padding:5px 7px;box-sizing:border-box;font-size:11px;
            border:1px solid #ccc;border-radius:3px;font-family:monospace}
      button{margin-top:10px;background:#1565c0;color:#fff;border:none;
             padding:9px;font-size:13px;border-radius:4px;cursor:pointer;width:100%}
      .aviso{background:#fff8e1;border:1px solid #ffe082;border-radius:4px;
             padding:8px 10px;font-size:11px;margin-bottom:10px;color:#5d4037}
    </style>
    <p class="aviso">
      📌 Preencha os dados de cada loja.<br>
      <b>Partner_id e Partner Key</b>: página do app Shopee → Live Partner_id / Live API Partner Key.<br>
      <b>ACCESS_TOKEN, REFRESH_TOKEN, SHOP_ID</b>: planilha existente da loja → Extensões → Apps Script → ⚙️ → Propriedades do script.
    </p>
    ${blocos}
    <button onclick="salvar()">💾 Salvar credenciais das 3 lojas</button>
    <script>
      function salvar() {
        const dados = {};
        ${PS_STORES.map(s =>
          `dados['${s}']={` +
          `pid:document.getElementById('${s}_pid').value.trim(),` +
          `pkey:document.getElementById('${s}_pkey').value.trim(),` +
          `at:document.getElementById('${s}_at').value.trim(),` +
          `rt:document.getElementById('${s}_rt').value.trim(),` +
          `sid:document.getElementById('${s}_sid').value.trim()};`
        ).join('')}
        document.querySelector('button').textContent = '⏳ Salvando...';
        google.script.run
          .withSuccessHandler(m => { alert(m); google.script.host.close(); })
          .withFailureHandler(e => {
            alert('Erro: ' + e.message);
            document.querySelector('button').textContent = '💾 Salvar credenciais das 3 lojas';
          })
          .ps_salvarCredenciais(dados);
      }
    </script>
  `).setWidth(540).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Importar Credenciais — Shopee');
}

function ps_salvarCredenciais(dados) {
  const props   = PropertiesService.getScriptProperties();
  const expires = String(Math.floor(Date.now() / 1000) + 14400);
  const salvos  = [];

  PS_STORES.forEach(s => {
    const d      = dados[s] || {};
    const prefix = s.toUpperCase() + '_';
    if (d.pid)  props.setProperty(prefix + 'PARTNER_ID',  d.pid);
    if (d.pkey) props.setProperty(prefix + 'PARTNER_KEY', d.pkey);
    if (d.at)   props.setProperty(prefix + 'ACCESS_TOKEN',  d.at);
    if (d.rt)   props.setProperty(prefix + 'REFRESH_TOKEN', d.rt);
    if (d.sid)  props.setProperty(prefix + 'SHOP_ID',       d.sid);
    if (d.at && d.rt && d.sid) {
      props.setProperty(prefix + 'TOKEN_EXPIRES', expires);
      salvos.push(s);
    }
  });

  if (!salvos.length) throw new Error('Nenhuma loja foi preenchida completamente (AT + RT + SHOP_ID).');
  return '✅ Credenciais salvas: ' + salvos.join(', ');
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
      .hint{color:#888;font-size:11px;margin-top:6px;line-height:1.5}
    </style>
    <p>Informe o <b>Discount ID</b> da promoção <b>ativa</b> de cada loja.</p>
    <p class="hint">Onde encontrar: Central de Marketing da Shopee → sua promoção → a URL da página contém o ID numérico.</p>
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
  `).setWidth(420).setHeight(320);
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
  const ui  = SpreadsheetApp.getUi();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(loja);
  if (!aba) throw new Error('Aba "' + loja + '" não encontrada.');

  const props      = PropertiesService.getScriptProperties();
  const discountId = parseInt(props.getProperty(loja.toUpperCase() + '_DISCOUNT_ID') || '0');
  if (!discountId) throw new Error(
    'Discount ID da loja ' + loja + ' não configurado.\nUse o menu 🎯 Configurar IDs de Promoção.'
  );

  const lastRow = aba.getLastRow();
  if (lastRow <= PS_HDR) { ui.alert(loja + ': sem dados na aba.'); return; }

  const NUM_COLS = Math.max(...Object.values(PS_COL)) + 1;
  const allData  = aba.getRange(PS_HDR + 1, 1, lastRow - PS_HDR, NUM_COLS).getValues();
  const statusCol = PS_COL.STATUS + 1;

  const pendentes = allData
    .map((row, i) => ({ i, row }))
    .filter(({ row }) => String(row[PS_COL.STATUS] || '').trim() === PS_PENDING);

  if (!pendentes.length) {
    ui.alert(loja + ': nenhum item com status "' + PS_PENDING + '".'); return;
  }

  // Agrupa model_ids pelo mesmo item_id
  const itemMap = {};
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

  const itemList = Object.entries(itemMap).map(([itemId, info]) => {
    const e = { item_id: parseInt(itemId), purchase_limit: info.purchase_limit, _rowIdxs: info.rowIdxs };
    if (info.temModel) e.model_list = info.models;
    else               e.item_promotion_price = info.preco;
    return e;
  });

  let adicionados = 0, erros = 0;

  for (let i = 0; i < itemList.length; i += 50) {
    const lote    = itemList.slice(i, i + 50);
    const rowMap  = {};
    const payload = lote.map(({ _rowIdxs, ...item }) => { rowMap[item.item_id] = _rowIdxs; return item; });

    let respData;
    try {
      respData = _ps_post(loja, '/api/v2/discount/add_discount_item', {
        discount_id: discountId, item_list: payload,
      });
    } catch (e) {
      lote.forEach(({ item_id, _rowIdxs }) =>
        _rowIdxs.forEach(r => aba.getRange(PS_HDR + 1 + r, statusCol).setValue('❌ Erro API: ' + e.message))
      );
      erros += lote.length; continue;
    }

    const erroMap = {};
    (respData.error_list || []).forEach(e => { erroMap[e.item_id] = e.failed_reason || e.msg || String(e.error_code || 'Erro'); });

    payload.forEach(item => {
      const msg = erroMap[item.item_id];
      rowMap[item.item_id].forEach(r =>
        aba.getRange(PS_HDR + 1 + r, statusCol).setValue(msg ? '❌ ' + msg : '✅ Adicionado')
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
// HTTP + AUTH POR LOJA
// ============================================================
function _ps_getToken(loja) {
  const prefix  = loja.toUpperCase() + '_';
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty(prefix + 'ACCESS_TOKEN');
  const refresh = props.getProperty(prefix + 'REFRESH_TOKEN');
  const shopId  = props.getProperty(prefix + 'SHOP_ID');
  const expires = parseInt(props.getProperty(prefix + 'TOKEN_EXPIRES') || '0');
  const now     = Math.floor(Date.now() / 1000);

  if (!token) throw new Error('Token da loja ' + loja + ' não encontrado.\nUse o menu 📋 Importar credenciais.');
  if (now >= expires - 300) return _ps_renovarToken(loja, refresh, parseInt(shopId));
  return { token, shopId };
}

function _ps_renovarToken(loja, refreshToken, shopId) {
  const pid    = _ps_partnerId(loja);
  const pkey   = _ps_partnerKey(loja);
  const prefix = loja.toUpperCase() + '_';
  const path   = '/api/v2/auth/access_token/get';
  const ts     = Math.floor(Date.now() / 1000);
  const sign   = _ps_sign(pid + path + ts, pkey);
  const resp   = UrlFetchApp.fetch(
    PS_BASE + path + '?partner_id=' + pid + '&timestamp=' + ts + '&sign=' + sign,
    { method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({ refresh_token: refreshToken, shop_id: shopId, partner_id: pid }) }
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
  const pid            = _ps_partnerId(loja);
  const pkey           = _ps_partnerKey(loja);
  const { token, shopId } = _ps_getToken(loja);
  const ts   = Math.floor(Date.now() / 1000);
  const sign = _ps_sign(pid + path + ts + token + parseInt(shopId), pkey);
  const url  = PS_BASE + path
    + '?partner_id=' + pid + '&timestamp=' + ts
    + '&access_token=' + token + '&shop_id=' + shopId + '&sign=' + sign;

  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (data.error && data.error !== '') throw new Error('[' + path + '] ' + (data.message || data.error));
  return data.response || data;
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
