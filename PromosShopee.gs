// ============================================================
// PROMOS SHOPEE — Confronto Promo ANTIGA × Promo VIGENTE
//
// Problema: ao clonar uma promoção no Shopee Seller Center, nem todos
// os produtos são copiados para a nova campanha. Este script confronta
// o arquivo da promo ANTIGA (referência, com os preços corretos) com o
// da promo VIGENTE e mostra o que precisa ser inserido manualmente.
//
// Chave de cruzamento: ID de variação (coluna D dos dois arquivos).
// Preço confrontado:   Preço de desconto (coluna H dos dois arquivos).
//
// Fluxo:
//   1. Menu "🏷️ Promos Shopee" → "📤 Importar e Comparar"
//   2. Selecione os 3 arquivos da pasta "Antigas" e os 3 da "Vigentes"
//   3. Gera uma aba por conta (Humble / Najumi / Sky) com TODOS os itens
//      da promo antiga + status + o preço que está hoje na vigente
//   4. (Opcional) "📥 Gerar abas de INSERIR" monta as planilhas já no
//      layout de importação do Shopee, só com o que falta corrigir
//
// OBS: se o projeto Apps Script já tiver outro onOpen(), apague o
// onOpen() deste arquivo e chame pspCriarMenu() de dentro do existente.
// ============================================================

const PSP = {
  CONTAS: ['Humble', 'Najumi', 'Sky'],

  ABA_RESUMO:     '📋 Resumo Promos',
  SUFIXO_INSERIR: ' — INSERIR',

  // Colunas do arquivo exportado pelo Shopee (1-based) — usadas como
  // fallback caso o cabeçalho não seja reconhecido pelo nome
  COL: {
    produto:      1, // A
    nome:         2, // B
    parentSku:    3, // C
    variacao:     4, // D  ← chave
    nomeVariacao: 5, // E
    sku:          6, // F
    precoOrig:    7, // G
    precoPromo:   8, // H  ← preço confrontado
    limite:       9, // I
  },

  TOLERANCIA: 0.005, // diferença de preço ignorada (arredondamento)
};

const PSP_ST = {
  FALTA: 'FALTA INSERIR',
  DIF:   'PREÇO DIFERENTE',
  OK:    'OK',
};

// Ordem de exibição: primeiro o que exige ação
const PSP_PRIORIDADE = { 'FALTA INSERIR': 0, 'PREÇO DIFERENTE': 1, 'OK': 2 };

const PSP_HEADER = [
  'ID do produto',
  'Nome do Produto',
  'Nº de Ref. Parent SKU',
  'ID de variação',
  'Variação de nome',
  'Nº de Ref. SKU',
  'Preço original (ANTIGA)',
  'Preço de desconto (ANTIGA)',
  'Limite de compra',
  'STATUS',
  'Preço de desconto (VIGENTE)',
  'Diferença (Vigente - Antiga)',
  'Preço original (VIGENTE)',
];

// Cabeçalho exato do arquivo de importação do Shopee (abas "— INSERIR")
const PSP_HEADER_SHOPEE = [
  'ID do produto',
  'Nome do Produto. (Opcional)',
  'Nº de Ref. Parent SKU. (Opcional)',
  'ID de variação',
  'Variação de nome. (Opcional)',
  'Nº de Ref. SKU. (Opcional)',
  'Preço original (opcional)',
  'Preço de desconto',
  'Limite de compra (Opcional)',
];

// Colunas do relatório (1-based) que guardam valores em dinheiro
const PSP_COLS_DINHEIRO = [7, 8, 11, 12, 13];
// Colunas do relatório que precisam ficar como TEXTO (IDs/SKUs longos)
const PSP_COLS_TEXTO = [1, 3, 4, 6];

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  pspCriarMenu();
}

function pspCriarMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🏷️ Promos Shopee')
    .addItem('📤 Importar e Comparar',      'pspImportar')
    .addSeparator()
    .addItem('📥 Gerar abas de INSERIR',    'pspGerarInserir')
    .addItem('🔄 Atualizar Resumo',         'pspAtualizarResumo')
    .addSeparator()
    .addItem('🗑️ Limpar abas geradas',      'pspLimpar')
    .addToUi();
}

// ── Diálogo de upload ─────────────────────────────────────────
function pspImportar() {
  const html = HtmlService.createHtmlOutputFromFile('PromosShopeeUpload')
    .setWidth(620)
    .setHeight(600)
    .setTitle('Importar Promos Shopee');
  SpreadsheetApp.getUi().showModalDialog(html, '🏷️ Importar e Comparar Promos');
}

/** Lista de contas para o diálogo montar os seletores. */
function pspContas() {
  return PSP.CONTAS;
}

// ============================================================
// PROCESSAMENTO DE UMA CONTA (chamado pelo diálogo)
//
// antiga / vigente: arrays de linhas já normalizadas pelo diálogo,
// cada linha com 9 posições na ordem das colunas A..I do Shopee.
// ============================================================
function pspProcessarConta(conta, antiga, vigente) {
  if (!conta) throw new Error('Conta não informada.');
  if (!antiga  || !antiga.length)  throw new Error('Arquivo da promo ANTIGA de ' + conta + ' está vazio.');
  if (!vigente || !vigente.length) throw new Error('Arquivo da promo VIGENTE de ' + conta + ' está vazio.');

  // Índice da vigente por ID de variação
  const mapaVigente = {};
  vigente.forEach(r => {
    const chave = _pspChave(r);
    if (!chave || mapaVigente[chave]) return; // 1ª ocorrência vence
    mapaVigente[chave] = {
      promo: _pspNum(r[PSP.COL.precoPromo - 1]),
      orig:  _pspNum(r[PSP.COL.precoOrig  - 1]),
    };
  });

  const vistos  = {};
  const linhas  = [];
  const cont    = { total: 0, falta: 0, dif: 0, ok: 0, semChave: 0 };

  antiga.forEach(r => {
    const chave = _pspChave(r);
    if (!chave) { cont.semChave++; return; }
    if (vistos[chave]) return;
    vistos[chave] = true;

    const promoAntiga = _pspNum(r[PSP.COL.precoPromo - 1]);
    const origAntiga  = _pspNum(r[PSP.COL.precoOrig  - 1]);
    const naVigente   = mapaVigente[chave];

    let status, promoVigente, origVigente, diferenca;

    if (!naVigente) {
      status       = PSP_ST.FALTA;
      promoVigente = '';   // não está na promo atual
      origVigente  = '';
      diferenca    = '';
      cont.falta++;
    } else {
      promoVigente = naVigente.promo === null ? '' : naVigente.promo;
      origVigente  = naVigente.orig  === null ? '' : naVigente.orig;
      diferenca    = (promoAntiga !== null && naVigente.promo !== null)
        ? Math.round((naVigente.promo - promoAntiga) * 100) / 100
        : '';
      const igual = promoAntiga !== null && naVigente.promo !== null &&
                    Math.abs(naVigente.promo - promoAntiga) <= PSP.TOLERANCIA;
      status = igual ? PSP_ST.OK : PSP_ST.DIF;
      if (igual) cont.ok++; else cont.dif++;
    }

    linhas.push([
      _pspTexto(r[PSP.COL.produto      - 1]),
      _pspTexto(r[PSP.COL.nome         - 1]),
      _pspTexto(r[PSP.COL.parentSku    - 1]),
      _pspTexto(r[PSP.COL.variacao     - 1]),
      _pspTexto(r[PSP.COL.nomeVariacao - 1]),
      _pspTexto(r[PSP.COL.sku          - 1]),
      origAntiga  === null ? '' : origAntiga,
      promoAntiga === null ? '' : promoAntiga,
      _pspTexto(r[PSP.COL.limite       - 1]),
      status,
      promoVigente,
      diferenca,
      origVigente,
    ]);
    cont.total++;
  });

  // Ação primeiro, depois nome do produto
  linhas.sort((a, b) => {
    const pa = PSP_PRIORIDADE[a[9]], pb = PSP_PRIORIDADE[b[9]];
    if (pa !== pb) return pa - pb;
    return String(a[1]).localeCompare(String(b[1]), 'pt-BR');
  });

  _pspEscreverAba(conta, linhas);

  cont.conta        = conta;
  cont.linhasAntiga = antiga.length;
  cont.linhasVigente= vigente.length;
  return cont;
}

// ── Escrita + formatação da aba da conta ──────────────────────
function _pspEscreverAba(conta, linhas) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let aba   = ss.getSheetByName(conta);
  if (!aba) aba = ss.insertSheet(conta);

  if (aba.getFilter()) aba.getFilter().remove();
  aba.clear();
  aba.clearConditionalFormatRules();

  const NC = PSP_HEADER.length;

  aba.getRange(1, 1, 1, NC)
     .setValues([PSP_HEADER])
     .setFontWeight('bold')
     .setFontColor('#ffffff')
     .setBackground('#263238')
     .setVerticalAlignment('middle')
     .setWrap(true);

  // Bloco da promo ANTIGA (A..I) e da VIGENTE (J..M) com cores distintas
  aba.getRange(1, 1,  1, 9).setBackground('#37474f');
  aba.getRange(1, 10, 1, 4).setBackground('#bf360c');

  if (!linhas.length) {
    aba.getRange(2, 1).setValue('⚠️ Nenhuma linha válida no arquivo da promo antiga.');
    return;
  }

  // IDs como texto para não virarem notação científica
  PSP_COLS_TEXTO.forEach(c => aba.getRange(2, c, linhas.length, 1).setNumberFormat('@'));

  aba.getRange(2, 1, linhas.length, NC).setValues(linhas);

  PSP_COLS_DINHEIRO.forEach(c =>
    aba.getRange(2, c, linhas.length, 1).setNumberFormat('R$ #,##0.00'));

  aba.getRange(2, 10, linhas.length, 1)
     .setHorizontalAlignment('center')
     .setFontWeight('bold');

  // Realce por status (linha inteira)
  const alvo  = aba.getRange(2, 1, linhas.length, NC);
  const regras = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2="' + PSP_ST.FALTA + '"')
      .setBackground('#ffcdd2').setFontColor('#b71c1c').setRanges([alvo]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2="' + PSP_ST.DIF + '"')
      .setBackground('#ffe0b2').setFontColor('#e65100').setRanges([alvo]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2="' + PSP_ST.OK + '"')
      .setBackground('#e8f5e9').setRanges([alvo]).build(),
  ];
  aba.setConditionalFormatRules(regras);

  aba.setFrozenRows(1);
  aba.setFrozenColumns(1);
  aba.getRange(1, 1, linhas.length + 1, NC).createFilter();

  aba.setColumnWidth(1, 110);
  aba.setColumnWidth(2, 330);
  aba.setColumnWidth(3, 150);
  aba.setColumnWidth(4, 120);
  aba.setColumnWidth(5, 140);
  aba.setColumnWidth(6, 150);
  for (let c = 7; c <= NC; c++) aba.setColumnWidth(c, 125);
  aba.setRowHeight(1, 42);
}

// ============================================================
// RESUMO — lê as abas das contas já geradas
// ============================================================
function pspAtualizarResumo() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const linhas = [];

  PSP.CONTAS.forEach(conta => {
    const aba = ss.getSheetByName(conta);
    if (!aba || aba.getLastRow() < 2) {
      linhas.push([conta, '—', '—', '—', '—', 'Aba não gerada']);
      return;
    }
    const status = aba.getRange(2, 10, aba.getLastRow() - 1, 1).getValues();
    let falta = 0, dif = 0, ok = 0;
    status.forEach(([s]) => {
      if (s === PSP_ST.FALTA)    falta++;
      else if (s === PSP_ST.DIF) dif++;
      else if (s === PSP_ST.OK)  ok++;
    });
    const total = falta + dif + ok;
    linhas.push([
      conta, total, falta, dif, ok,
      falta + dif === 0 ? '✅ Nada a fazer' : '⚠️ ' + (falta + dif) + ' item(ns) para corrigir',
    ]);
  });

  let aba = ss.getSheetByName(PSP.ABA_RESUMO);
  if (!aba) aba = ss.insertSheet(PSP.ABA_RESUMO, 0);
  aba.clear();

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  aba.getRange(1, 1).setValue('🏷️ Promos Shopee — Antiga × Vigente   (atualizado em ' + ts + ')')
     .setFontSize(13).setFontWeight('bold');

  const HEADER = ['Conta', 'Itens na promo antiga', 'Falta inserir', 'Preço diferente', 'OK', 'Situação'];
  aba.getRange(3, 1, 1, HEADER.length)
     .setValues([HEADER])
     .setFontWeight('bold').setFontColor('#ffffff').setBackground('#263238');
  aba.getRange(4, 1, linhas.length, HEADER.length).setValues(linhas);

  aba.getRange(4, 3, linhas.length, 1).setFontColor('#b71c1c').setFontWeight('bold');
  aba.getRange(4, 4, linhas.length, 1).setFontColor('#e65100').setFontWeight('bold');
  aba.setColumnWidth(1, 120);
  for (let c = 2; c <= 5; c++) aba.setColumnWidth(c, 150);
  aba.setColumnWidth(6, 240);

  return linhas;
}

// ============================================================
// ABAS "— INSERIR": layout de importação do Shopee, apenas com
// os itens que faltam ou estão com preço diferente do de referência.
// O preço usado é o da promo ANTIGA (a referência correta).
// ============================================================
function pspGerarInserir() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg  = '';

  PSP.CONTAS.forEach(conta => {
    const origem = ss.getSheetByName(conta);
    if (!origem || origem.getLastRow() < 2) {
      msg += '• ' + conta + ': aba não gerada — importe os arquivos primeiro.\n';
      return;
    }

    const dados = origem.getRange(2, 1, origem.getLastRow() - 1, PSP_HEADER.length).getValues();
    const saida = dados
      .filter(r => r[9] === PSP_ST.FALTA || r[9] === PSP_ST.DIF)
      .map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8]]);

    const nome = conta + PSP.SUFIXO_INSERIR;
    let aba = ss.getSheetByName(nome);
    if (!aba) aba = ss.insertSheet(nome);
    if (aba.getFilter()) aba.getFilter().remove();
    aba.clear();

    aba.getRange(1, 1, 1, PSP_HEADER_SHOPEE.length)
       .setValues([PSP_HEADER_SHOPEE])
       .setFontWeight('bold').setFontColor('#ffffff').setBackground('#ee4d2d').setWrap(true);

    if (saida.length) {
      [1, 3, 4, 6].forEach(c => aba.getRange(2, c, saida.length, 1).setNumberFormat('@'));
      aba.getRange(2, 1, saida.length, PSP_HEADER_SHOPEE.length).setValues(saida);
      [7, 8].forEach(c => aba.getRange(2, c, saida.length, 1).setNumberFormat('0.00'));
    }

    aba.setFrozenRows(1);
    aba.setColumnWidth(2, 330);
    msg += '• ' + conta + ': ' + saida.length + ' item(ns) para inserir/corrigir.\n';
  });

  ui.alert('📥 Abas de INSERIR geradas',
    msg + '\nCada aba está no layout de importação do Shopee, com o preço de desconto da promo ANTIGA.\n' +
    'Baixe em Arquivo → Fazer download → .xlsx antes de subir no Seller Center.',
    ui.ButtonSet.OK);
}

// ── Limpeza ───────────────────────────────────────────────────
function pspLimpar() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('🗑️ Limpar abas geradas',
    'Isso apaga as abas das contas, as abas "— INSERIR" e o resumo. Continuar?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alvos = [PSP.ABA_RESUMO];
  PSP.CONTAS.forEach(c => alvos.push(c, c + PSP.SUFIXO_INSERIR));

  let apagadas = 0;
  alvos.forEach(nome => {
    const aba = ss.getSheetByName(nome);
    if (aba && ss.getSheets().length > 1) { ss.deleteSheet(aba); apagadas++; }
  });
  ui.alert('✅ ' + apagadas + ' aba(s) removida(s).');
}

// ============================================================
// HELPERS
// ============================================================

/**
 * Chave de cruzamento: ID de variação (coluna D).
 * Produtos sem variação vêm com D em branco no export do Shopee —
 * nesse caso cai para o ID do produto, com prefixo para não colidir.
 */
function _pspChave(linha) {
  const variacao = _pspTexto(linha[PSP.COL.variacao - 1]);
  if (variacao) return variacao;
  const produto = _pspTexto(linha[PSP.COL.produto - 1]);
  return produto ? 'P:' + produto : '';
}

/** Normaliza célula para texto, removendo separador de milhar de IDs. */
function _pspTexto(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim();
  // "18.999.208.348" → "18999208348" (só quando é ID puramente numérico)
  if (/^\d{1,3}([.,]\d{3})+$/.test(s)) s = s.replace(/[.,]/g, '');
  // 1.2964231366e+11 → 129642313660
  if (typeof v === 'number' && Number.isInteger(v)) s = v.toFixed(0);
  return s;
}

/**
 * Converte texto de preço em número. Aceita "23.50", "23,50",
 * "1.234,56", "1,234.56" e "R$ 23,50". Retorna null se não for número.
 */
function _pspNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;

  let s = String(v).replace(/[^\d.,-]/g, '').trim();
  if (!s) return null;

  const temPonto  = s.indexOf('.') >= 0;
  const temVirg   = s.indexOf(',') >= 0;

  if (temPonto && temVirg) {
    // O separador que aparece por último é o decimal
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else                                          s = s.replace(/,/g, '');
  } else if (temVirg) {
    s = s.replace(',', '.');
  } else if (/^\d{1,3}\.\d{3}$/.test(s)) {
    // "1.234" isolado = separador de milhar, não decimal
    s = s.replace('.', '');
  }

  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
