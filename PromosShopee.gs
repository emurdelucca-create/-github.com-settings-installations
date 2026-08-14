// ============================================================
// PROMOS SHOPEE — Confronto Promo ANTIGA × Promo VIGENTE
//
// Problema: ao clonar uma promoção no Shopee Seller Center, nem todos
// os produtos são copiados para a nova campanha. Este script confronta
// os dois arquivos e mostra o que precisa ser inserido manualmente.
//
// REFERÊNCIA (padrão AUTO): por conta, vence o arquivo com MAIS itens
// em promoção — ele é tratado como a lista completa e correta. O outro
// é o comparado. Dá para travar a referência pelo menu, se preferir.
//
// Chave de cruzamento: ID de variação (coluna D dos dois arquivos).
// Preço confrontado:   Preço de desconto (coluna H dos dois arquivos).
//
// Fluxo:
//   1. Menu "🏷️ Promos Shopee" → "📤 Importar e Comparar"
//   2. Selecione os 3 arquivos da pasta "Antigas" e os 3 da "Vigentes"
//   3. Gera uma aba por conta (Humble / Najumi / Sky) com TODOS os itens
//      da referência + status + o preço que está no outro arquivo
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

  // Layout da aba de cada conta
  LINHA_TITULO: 1,
  LINHA_HEADER: 2,
  LINHA_DADOS:  3,

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

  // Coluna do STATUS na aba gerada
  COL_STATUS: 10,

  TOLERANCIA: 0.005, // diferença de preço ignorada (arredondamento)

  PROP_MODO: 'PSP_MODO_REFERENCIA', // AUTO | ANTIGA | VIGENTE
  PROP_REF:  'PSP_REF_',            // + nome da conta → lado usado como referência
};

const PSP_ANTIGA  = 'ANTIGA';
const PSP_VIGENTE = 'VIGENTE';

const PSP_ST = {
  DIF: 'PREÇO DIFERENTE',
  OK:  'OK',
  falta: function(ladoComparado) { return 'FALTA NA ' + ladoComparado; },
};

const PSP_HEADER_FIXO = [
  'ID do produto',
  'Nome do Produto',
  'Nº de Ref. Parent SKU',
  'ID de variação',
  'Variação de nome',
  'Nº de Ref. SKU',
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

// Colunas do relatório (1-based) com valores em dinheiro / texto
const PSP_COLS_DINHEIRO = [7, 8, 11, 12, 13];
const PSP_COLS_TEXTO    = [1, 3, 4, 6];
const PSP_NCOLS         = 13;

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
    .addSubMenu(SpreadsheetApp.getUi()
      .createMenu('⚙️ Referência')
      .addItem('Automática — a que tiver mais itens', 'pspModoAuto')
      .addItem('Sempre a ANTIGA',                     'pspModoAntiga')
      .addItem('Sempre a VIGENTE',                    'pspModoVigente')
      .addItem('Ver referência atual',                'pspVerModo'))
    .addSeparator()
    .addItem('🗑️ Limpar abas geradas',      'pspLimpar')
    .addToUi();
}

// ── Modo de referência ────────────────────────────────────────
function _pspProps()  { return PropertiesService.getDocumentProperties(); }
function _pspModo()   { return _pspProps().getProperty(PSP.PROP_MODO) || 'AUTO'; }

function pspModoAuto()    { _pspDefinirModo('AUTO'); }
function pspModoAntiga()  { _pspDefinirModo(PSP_ANTIGA); }
function pspModoVigente() { _pspDefinirModo(PSP_VIGENTE); }

function _pspDefinirModo(modo) {
  _pspProps().setProperty(PSP.PROP_MODO, modo);
  const texto = modo === 'AUTO'
    ? 'Automática: em cada conta vence o arquivo com MAIS itens em promoção.'
    : 'Travada na promo ' + modo + ' para todas as contas.';
  SpreadsheetApp.getUi().alert('⚙️ Referência definida', texto + '\n\nImporte novamente para aplicar.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function pspVerModo() {
  const ui = SpreadsheetApp.getUi();
  let msg = 'Modo: ' + _pspModo() + '\n\nÚltima importação:\n';
  PSP.CONTAS.forEach(c => {
    const ref = _pspProps().getProperty(PSP.PROP_REF + c);
    msg += '• ' + c + ': ' + (ref ? 'referência = ' + ref : 'ainda não importado') + '\n';
  });
  ui.alert('⚙️ Referência atual', msg, ui.ButtonSet.OK);
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

  const idxAntiga  = _pspIndexar(antiga);
  const idxVigente = _pspIndexar(vigente);

  // Referência = arquivo com mais itens únicos (ou o lado travado no menu)
  const modo = _pspModo();
  let ladoRef;
  if (modo === PSP_ANTIGA || modo === PSP_VIGENTE) {
    ladoRef = modo;
  } else {
    ladoRef = idxVigente.ordem.length > idxAntiga.ordem.length ? PSP_VIGENTE : PSP_ANTIGA;
  }
  const ladoComp = ladoRef === PSP_ANTIGA ? PSP_VIGENTE : PSP_ANTIGA;

  const ref  = ladoRef === PSP_ANTIGA ? idxAntiga  : idxVigente;
  const comp = ladoRef === PSP_ANTIGA ? idxVigente : idxAntiga;

  const stFalta = PSP_ST.falta(ladoComp);
  const linhas  = [];
  const cont    = { total: 0, falta: 0, dif: 0, ok: 0, semChave: idxAntiga.semChave + idxVigente.semChave };

  ref.ordem.forEach(chave => {
    const r = ref.dados[chave];
    const precoRef = _pspNum(r[PSP.COL.precoPromo - 1]);
    const origRef  = _pspNum(r[PSP.COL.precoOrig  - 1]);
    const noOutro  = comp.dados[chave];

    let status, precoComp = '', origComp = '', diferenca = '';

    if (!noOutro) {
      status = stFalta;
      cont.falta++;
    } else {
      const p = _pspNum(noOutro[PSP.COL.precoPromo - 1]);
      const o = _pspNum(noOutro[PSP.COL.precoOrig  - 1]);
      precoComp = p === null ? '' : p;
      origComp  = o === null ? '' : o;
      if (precoRef !== null && p !== null) {
        diferenca = Math.round((p - precoRef) * 100) / 100;
        const igual = Math.abs(p - precoRef) <= PSP.TOLERANCIA;
        status = igual ? PSP_ST.OK : PSP_ST.DIF;
        if (igual) cont.ok++; else cont.dif++;
      } else {
        status = PSP_ST.DIF;
        cont.dif++;
      }
    }

    linhas.push([
      _pspTexto(r[PSP.COL.produto      - 1]),
      _pspTexto(r[PSP.COL.nome         - 1]),
      _pspTexto(r[PSP.COL.parentSku    - 1]),
      _pspTexto(r[PSP.COL.variacao     - 1]),
      _pspTexto(r[PSP.COL.nomeVariacao - 1]),
      _pspTexto(r[PSP.COL.sku          - 1]),
      origRef  === null ? '' : origRef,
      precoRef === null ? '' : precoRef,
      _pspTexto(r[PSP.COL.limite       - 1]),
      status,
      precoComp,
      diferenca,
      origComp,
    ]);
    cont.total++;
  });

  // Ação primeiro, depois nome do produto
  linhas.sort((a, b) => {
    const pa = _pspPrioridade(a[PSP.COL_STATUS - 1]);
    const pb = _pspPrioridade(b[PSP.COL_STATUS - 1]);
    if (pa !== pb) return pa - pb;
    return String(a[1]).localeCompare(String(b[1]), 'pt-BR');
  });

  _pspProps().setProperty(PSP.PROP_REF + conta, ladoRef);
  _pspEscreverAba(conta, linhas, {
    ladoRef:  ladoRef,
    ladoComp: ladoComp,
    qtdRef:   ref.ordem.length,
    qtdComp:  comp.ordem.length,
    stFalta:  stFalta,
  });

  cont.conta    = conta;
  cont.ladoRef  = ladoRef;
  cont.ladoComp = ladoComp;
  cont.qtdRef   = ref.ordem.length;
  cont.qtdComp  = comp.ordem.length;
  return cont;
}

/** Indexa as linhas de um arquivo por ID de variação, sem duplicatas. */
function _pspIndexar(linhas) {
  const dados = {}, ordem = [];
  let semChave = 0;
  linhas.forEach(r => {
    const chave = _pspChave(r);
    if (!chave) { semChave++; return; }
    if (dados[chave]) return; // 1ª ocorrência vence
    dados[chave] = r;
    ordem.push(chave);
  });
  return { dados: dados, ordem: ordem, semChave: semChave };
}

function _pspPrioridade(status) {
  if (String(status).indexOf('FALTA') === 0) return 0;
  if (status === PSP_ST.DIF) return 1;
  return 2;
}

// ── Escrita + formatação da aba da conta ──────────────────────
function _pspEscreverAba(conta, linhas, ctx) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let aba   = ss.getSheetByName(conta);
  if (!aba) aba = ss.insertSheet(conta);

  if (aba.getFilter()) aba.getFilter().remove();
  aba.clear();
  aba.clearConditionalFormatRules();

  const header = PSP_HEADER_FIXO.concat([
    'Preço original (' + ctx.ladoRef + ')',
    '★ Preço de desconto (' + ctx.ladoRef + ')',
    'Limite de compra',
    'STATUS',
    'Preço de desconto (' + ctx.ladoComp + ')',
    'Diferença (' + ctx.ladoComp + ' − ' + ctx.ladoRef + ')',
    'Preço original (' + ctx.ladoComp + ')',
  ]);

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  aba.getRange(PSP.LINHA_TITULO, 1, 1, PSP_NCOLS).merge()
     .setValue('🏷️ ' + conta + ' — Referência: promo ' + ctx.ladoRef + ' (' + ctx.qtdRef +
               ' itens, a que tem mais)  ×  Comparada: promo ' + ctx.ladoComp + ' (' + ctx.qtdComp +
               ' itens)   ·   atualizado em ' + ts)
     .setFontWeight('bold').setFontSize(11)
     .setBackground('#ee4d2d').setFontColor('#ffffff')
     .setVerticalAlignment('middle');

  aba.getRange(PSP.LINHA_HEADER, 1, 1, PSP_NCOLS)
     .setValues([header])
     .setFontWeight('bold')
     .setFontColor('#ffffff')
     .setBackground('#263238')
     .setVerticalAlignment('middle')
     .setWrap(true);

  // Bloco da referência (A..I) e do arquivo comparado (J..M)
  aba.getRange(PSP.LINHA_HEADER, 1,  1, 9).setBackground('#37474f');
  aba.getRange(PSP.LINHA_HEADER, 10, 1, 4).setBackground('#bf360c');

  if (!linhas.length) {
    aba.getRange(PSP.LINHA_DADOS, 1).setValue('⚠️ Nenhuma linha válida no arquivo de referência.');
    return;
  }

  // IDs como texto para não virarem notação científica
  PSP_COLS_TEXTO.forEach(c =>
    aba.getRange(PSP.LINHA_DADOS, c, linhas.length, 1).setNumberFormat('@'));

  aba.getRange(PSP.LINHA_DADOS, 1, linhas.length, PSP_NCOLS).setValues(linhas);

  PSP_COLS_DINHEIRO.forEach(c =>
    aba.getRange(PSP.LINHA_DADOS, c, linhas.length, 1).setNumberFormat('R$ #,##0.00'));

  aba.getRange(PSP.LINHA_DADOS, PSP.COL_STATUS, linhas.length, 1)
     .setHorizontalAlignment('center')
     .setFontWeight('bold');

  // Realce por status (linha inteira)
  const alvo = aba.getRange(PSP.LINHA_DADOS, 1, linhas.length, PSP_NCOLS);
  const ref$ = '$J' + PSP.LINHA_DADOS;
  aba.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + ref$ + '="' + ctx.stFalta + '"')
      .setBackground('#ffcdd2').setFontColor('#b71c1c').setRanges([alvo]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + ref$ + '="' + PSP_ST.DIF + '"')
      .setBackground('#ffe0b2').setFontColor('#e65100').setRanges([alvo]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + ref$ + '="' + PSP_ST.OK + '"')
      .setBackground('#e8f5e9').setRanges([alvo]).build(),
  ]);

  aba.setFrozenRows(PSP.LINHA_HEADER);
  aba.setFrozenColumns(1);
  aba.getRange(PSP.LINHA_HEADER, 1, linhas.length + 1, PSP_NCOLS).createFilter();

  aba.setColumnWidth(1, 110);
  aba.setColumnWidth(2, 330);
  aba.setColumnWidth(3, 150);
  aba.setColumnWidth(4, 120);
  aba.setColumnWidth(5, 140);
  aba.setColumnWidth(6, 150);
  for (let c = 7; c <= PSP_NCOLS; c++) aba.setColumnWidth(c, 130);
  aba.setRowHeight(PSP.LINHA_TITULO, 28);
  aba.setRowHeight(PSP.LINHA_HEADER, 42);
}

// ============================================================
// RESUMO — lê as abas das contas já geradas
// ============================================================
function pspAtualizarResumo() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const linhas = [];

  PSP.CONTAS.forEach(conta => {
    const aba = ss.getSheetByName(conta);
    if (!aba || aba.getLastRow() < PSP.LINHA_DADOS) {
      linhas.push([conta, '—', '—', '—', '—', '—', 'Aba não gerada']);
      return;
    }

    const n      = aba.getLastRow() - PSP.LINHA_DADOS + 1;
    const status = aba.getRange(PSP.LINHA_DADOS, PSP.COL_STATUS, n, 1).getValues();

    let falta = 0, dif = 0, ok = 0, rotuloFalta = '';
    status.forEach(([s]) => {
      if (String(s).indexOf('FALTA') === 0) { falta++; rotuloFalta = s; }
      else if (s === PSP_ST.DIF)            { dif++; }
      else if (s === PSP_ST.OK)             { ok++; }
    });

    const ladoRef = _pspProps().getProperty(PSP.PROP_REF + conta) || '—';

    // Só há o que inserir na vigente quando a referência é a ANTIGA.
    // Com a VIGENTE como referência, os "sem par" são itens que entraram
    // agora e as divergências de preço já estão com o preço da referência.
    let situacao;
    if (ladoRef === PSP_ANTIGA) {
      const acoes = falta + dif;
      situacao = acoes === 0
        ? '✅ Nada a inserir na vigente'
        : '⚠️ ' + acoes + ' item(ns) para inserir/corrigir na vigente';
    } else {
      situacao = dif === 0
        ? '✅ Nenhuma divergência de preço'
        : '⚠️ ' + dif + ' preço(s) divergente(s) — confira qual vale';
    }

    linhas.push([
      conta,
      'promo ' + ladoRef,
      falta + dif + ok,
      falta + (rotuloFalta ? ' (' + rotuloFalta.replace('FALTA NA ', 'só na ' + ladoRef + ', fora da ') + ')' : ''),
      dif,
      ok,
      situacao,
    ]);
  });

  let aba = ss.getSheetByName(PSP.ABA_RESUMO);
  if (!aba) aba = ss.insertSheet(PSP.ABA_RESUMO, 0);
  aba.clear();

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  aba.getRange(1, 1).setValue('🏷️ Promos Shopee — Antiga × Vigente   (atualizado em ' + ts + ')')
     .setFontSize(13).setFontWeight('bold');
  aba.getRange(2, 1).setValue('Referência = ' +
    (_pspModo() === 'AUTO' ? 'automática (o arquivo com mais itens em promoção)' : 'travada na promo ' + _pspModo()))
     .setFontColor('#666666');

  const HEADER = ['Conta', 'Referência', 'Itens na referência', 'Sem par no outro arquivo',
                  'Preço diferente', 'OK', 'Situação'];
  aba.getRange(4, 1, 1, HEADER.length)
     .setValues([HEADER])
     .setFontWeight('bold').setFontColor('#ffffff').setBackground('#263238').setWrap(true);
  aba.getRange(5, 1, linhas.length, HEADER.length).setValues(linhas);

  aba.getRange(5, 4, linhas.length, 1).setFontColor('#b71c1c').setFontWeight('bold');
  aba.getRange(5, 5, linhas.length, 1).setFontColor('#e65100').setFontWeight('bold');
  aba.setColumnWidth(1, 110);
  aba.setColumnWidth(2, 120);
  for (let c = 3; c <= 6; c++) aba.setColumnWidth(c, 170);
  aba.setColumnWidth(7, 300);

  return linhas;
}

// ============================================================
// ABAS "— INSERIR": layout de importação do Shopee, apenas com
// o que precisa entrar/ser corrigido NA PROMO VIGENTE, sempre com
// o preço do arquivo de referência.
//
// Quando a referência é a própria VIGENTE, os itens "sem par" estão
// só na vigente (nada a inserir) e os de preço diferente já estão com
// o preço da referência — nesses casos a aba sai vazia, de propósito.
// ============================================================
function pspGerarInserir() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg  = '';

  PSP.CONTAS.forEach(conta => {
    const origem = ss.getSheetByName(conta);
    if (!origem || origem.getLastRow() < PSP.LINHA_DADOS) {
      msg += '• ' + conta + ': aba não gerada — importe os arquivos primeiro.\n';
      return;
    }

    const ladoRef = _pspProps().getProperty(PSP.PROP_REF + conta) || PSP_ANTIGA;
    const n     = origem.getLastRow() - PSP.LINHA_DADOS + 1;
    const dados = origem.getRange(PSP.LINHA_DADOS, 1, n, PSP_NCOLS).getValues();

    const saida = dados
      .filter(r => {
        const s = r[PSP.COL_STATUS - 1];
        if (ladoRef !== PSP_ANTIGA) return false;               // nada a inserir na vigente
        return String(s).indexOf('FALTA') === 0 || s === PSP_ST.DIF;
      })
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

    msg += '• ' + conta + ': ' + saida.length + ' item(ns) para inserir/corrigir' +
           (ladoRef === PSP_ANTIGA ? '.\n' : ' — referência é a própria VIGENTE, nada a inserir.\n');
  });

  ui.alert('📥 Abas de INSERIR geradas',
    msg + '\nCada aba está no layout de importação do Shopee, com o preço de desconto da referência.\n' +
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

  const temPonto = s.indexOf('.') >= 0;
  const temVirg  = s.indexOf(',') >= 0;

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
