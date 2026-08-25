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

// ── Lê dados da aba Planejada ─────────────────────────────────
function me_dash_getData() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Planejada');

  if (!aba) {
    return { error: 'Aba "Planejada" não encontrada. Clique em "▶ Gerar Planejada" primeiro.' };
  }

  const lastRow = aba.getLastRow();
  if (lastRow <= ME.HEADER_ROWS) {
    return { itens: [], ts: _me_dash_ts() };
  }

  // Colunas lidas (1-indexed → valores 0-indexed no array):
  //   A(1)=SKU  B(2)=picking  C(3)=armPad  F(6)=locF
  //   P(16)=mov5  Q(17)=mov7  R(18)=mov10  S(19)=mov15
  //   T(20)=abc   AE(31)=pedTotal
  const NCOLS   = 31;
  const dataRows = lastRow - ME.HEADER_ROWS;
  const data    = aba.getRange(ME.HEADER_ROWS + 1, 1, dataRows, NCOLS).getValues();

  const itens = data
    .filter(r => String(r[0] || '').trim())
    .map(r => ({
      sku:   String(r[0]  || '').trim(),
      pick:  Number(r[1]  || 0),
      bruto: Number(r[2]  || 0),
      locF:  String(r[5]  || ''),
      mov5:  Math.max(0, Number(r[15] || 0)),
      mov7:  Math.max(0, Number(r[16] || 0)),
      mov10: Math.max(0, Number(r[17] || 0)),
      mov15: Math.max(0, Number(r[18] || 0)),
      abc:   String(r[19] || '').trim(),
      peds:  Number(r[30] || 0),
    }));

  return { itens, ts: _me_dash_ts() };
}

// ── Gera aba Planejada sem ui.alert() (seguro para web app) ──
function me_dash_gerarPlanejada() {
  try {
    const itens = _me_carregarTodos();
    _me_escreverAba('Planejada', itens);
    return { ok: true, count: itens.length };
  } catch(e) {
    return { ok: false, error: e.message + '\n\n' + (e.stack || '') };
  }
}

function _me_dash_ts() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}
