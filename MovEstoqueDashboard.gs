// ============================================================
// DASHBOARD — Movimentação de Estoque (App Web autônomo)
//
// Deploy: Extensões → Apps Script → Implantar → Nova implantação
//   Tipo: App da Web
//   Executar como: Eu
//   Acesso: Qualquer pessoa na organização (ou Qualquer pessoa)
//
// Dependências: MovimentacaoEstoque.gs (funções _me_*)
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('MovEstoqueDashboard')
    .setTitle('Dashboard — Movimentação de Estoque')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Carrega dados diretamente da API BaseLinker ───────────────
function me_dash_getData() {
  try {
    const raw = _me_carregarTodos();
    const itens = raw.map(r => ({
      sku:        r.sku,
      pick:       r.picking,
      arm:        r.arm,
      chg:        r.chg,
      locF:       r.locF,
      locG:       r.locG,
      v5:  r.v5,  v7:  r.v7,  v10: r.v10, v15: r.v15,
      mov5:  Math.max(0, r.mov5),
      mov7:  Math.max(0, r.mov7),
      mov10: Math.max(0, r.mov10),
      mov15: Math.max(0, r.mov15),
      abc:        r.abc,
      pedTotal:   r.pedTotal,
      entrDireta: r.entrDireta,
      me2:        r.me2,
      retirada:   r.retirada,
      xpress:     r.xpress,
      pedMov:     r.pedMov,
    }));
    return { itens, ts: _me_dash_ts() };
  } catch(e) {
    return { error: e.message + '\n\n' + (e.stack || '') };
  }
}

function _me_dash_ts() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}
