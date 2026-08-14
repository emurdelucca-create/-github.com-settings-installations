/**
 * ============================================================================
 *  PRECIFICAÇÃO LUCRO REAL — MARKETPLACES (ML / SHOPEE)
 *  Gerador de planilha com ICMS/DIFAL corretos (EC 87/2015 + LC 190/2022)
 * ============================================================================
 *
 *  COMO USAR
 *  1) Abra a planilha > Extensões > Apps Script
 *  2) Cole este arquivo inteiro, salve
 *  3) Execute a função  criarPlanilha()   (autorize na primeira vez)
 *  4) Recarregue a planilha — vai aparecer o menu "⚙️ Precificação"
 *
 *  O script SÓ escreve as fórmulas nas células. Depois disso todo o cálculo
 *  é dinâmico, nativo do Sheets. Você nunca mais precisa rodar nada.
 *
 *  Ele cria/recria APENAS estas 4 abas e não encosta em nenhuma outra:
 *      Config | Tabela_ICMS | Faturamento_UF | Precificacao
 *
 *  ---------------------------------------------------------------------------
 *  LOCALE — por que existe a função F() abaixo
 *  ---------------------------------------------------------------------------
 *  Numa planilha em pt-BR o Sheets separa argumentos por ";" e não por ",".
 *  Fórmula gravada com vírgula vira #ERROR! (erro de parsing) e o erro se
 *  propaga por todas as células que dependem dela.
 *
 *  Pior: literal decimal com ponto (0.12) NÃO dá erro em pt-BR — o ponto é
 *  separador de milhar, então 0.12 é lido como 12. Número errado, sem aviso.
 *
 *  Solução aqui:
 *    • detectarSeparador() descobre em tempo de execução qual separador esta
 *      planilha aceita, testando =SUM(1,2) numa aba temporária;
 *    • F() converte as vírgulas das fórmulas para o separador certo;
 *    • NENHUMA fórmula contém literal decimal — as alíquotas de 12% e 7%
 *      moram em células da aba Config (ALQ_INTER_12 e ALQ_INTER_7).
 *
 *  INVARIANTE que faz o F() ser seguro: nenhum texto entre aspas dentro das
 *  fórmulas contém vírgula, e não há literal decimal. Se você editar as
 *  fórmulas, preserve isso — senão o replace corrompe o conteúdo.
 * ============================================================================
 */

var ABAS = ['Config', 'Tabela_ICMS', 'Faturamento_UF', 'Precificacao'];

var SEP = ',';   // definido em criarPlanilha() por detectarSeparador()

var FMT_MOEDA = 'R$ #,##0.00';
var FMT_PCT   = '0.00%';

var COR_TITULO  = '#1c3d5a';
var COR_SECAO   = '#dbe5f1';
var COR_INPUT   = '#fff2cc';
var COR_AUTO    = '#f0f0f0';
var COR_RESULT  = '#d9ead3';
var COR_ALERTA  = '#f4cccc';

/**
 * UF | Estado | Grupo (S/SE ou N/NE/CO) | Alíquota interna | FCP | Faturamento
 *
 * "Grupo" define a alíquota interestadual: S/SE -> S/SE = 12%, qualquer outro
 * cruzamento = 7%. O ES conta como N/NE/CO para essa regra (Resolução do
 * Senado 22/1989).
 *
 * A coluna "Alíquota interna" já embute o adicional/FECP dos estados que o
 * cobram dentro da alíquota cheia (ex.: RJ 20% + 2% FECP = 22%). Por isso o
 * FCP vai zerado — confira com a sua contabilidade antes de mexer, para não
 * contar o adicional duas vezes.
 */
var UFS = [
  ['AC', 'Acre',                'N/NE/CO', 0.190,  0,      5739.36],
  ['AL', 'Alagoas',             'N/NE/CO', 0.200,  0,     42020.26],
  ['AP', 'Amapá',               'N/NE/CO', 0.180,  0,      3384.78],
  ['AM', 'Amazonas',            'N/NE/CO', 0.200,  0,     20726.67],
  ['BA', 'Bahia',               'N/NE/CO', 0.205,  0,    215351.57],
  ['CE', 'Ceará',               'N/NE/CO', 0.200,  0,     97517.55],
  ['DF', 'Distrito Federal',    'N/NE/CO', 0.200,  0,     18468.01],
  ['ES', 'Espírito Santo',      'N/NE/CO', 0.170,  0,     69033.55],
  ['GO', 'Goiás',               'N/NE/CO', 0.190,  0,     81697.44],
  ['MA', 'Maranhão',            'N/NE/CO', 0.230,  0,     90050.98],
  ['MT', 'Mato Grosso',         'N/NE/CO', 0.170,  0,     44814.70],
  ['MS', 'Mato Grosso do Sul',  'N/NE/CO', 0.170,  0,     44444.39],
  ['MG', 'Minas Gerais',        'S/SE',    0.180,  0,    305715.80],
  ['PA', 'Pará',                'N/NE/CO', 0.190,  0,     68296.02],
  ['PB', 'Paraíba',             'N/NE/CO', 0.200,  0,     58681.35],
  ['PR', 'Paraná',              'S/SE',    0.195,  0,    127810.82],
  ['PE', 'Pernambuco',          'N/NE/CO', 0.205,  0,    134125.25],
  ['PI', 'Piauí',               'N/NE/CO', 0.225,  0,     38922.37],
  ['RJ', 'Rio de Janeiro',      'S/SE',    0.220,  0,    198329.22],
  ['RN', 'Rio Grande do Norte', 'N/NE/CO', 0.200,  0,     42685.22],
  ['RS', 'Rio Grande do Sul',   'S/SE',    0.170,  0,     89297.15],
  ['RO', 'Rondônia',            'N/NE/CO', 0.195,  0,     10886.11],
  ['RR', 'Roraima',             'N/NE/CO', 0.200,  0,      1529.04],
  ['SC', 'Santa Catarina',      'S/SE',    0.170,  0,     80969.88],
  ['SP', 'São Paulo',           'S/SE',    0.180,  0,    537245.38],
  ['SE', 'Sergipe',             'N/NE/CO', 0.200,  0,     34095.88],
  ['TO', 'Tocantins',           'N/NE/CO', 0.200,  0,     22475.91]
];

var N_UF    = UFS.length;         // 27
var LIN_1   = 2;                  // primeira linha de dados
var LIN_N   = LIN_1 + N_UF - 1;   // última linha de dados = 28
var LIN_TOT = LIN_N + 1;          // linha de total = 29


// ===========================================================================
//  LOCALE
// ===========================================================================

/**
 * Descobre o separador de argumentos que ESTA planilha aceita, testando o
 * comportamento real em vez de adivinhar pelo código de idioma.
 *
 *   locale en_US : =SUM(1,2) -> 3      (vírgula separa argumentos)
 *   locale pt-BR : =SUM(1,2) -> 1,2    (vírgula é decimal; 1 argumento só)
 *
 * Qualquer resultado diferente de 3 significa que a vírgula não é separador.
 */
function detectarSeparador(ss) {
  var nome = '__sep_probe__';
  var antiga = ss.getSheetByName(nome);
  if (antiga) ss.deleteSheet(antiga);

  var probe = ss.insertSheet(nome);
  var sep = ',';
  try {
    probe.getRange('A1').setFormula('=SUM(1,2)');
    SpreadsheetApp.flush();
    if (probe.getRange('A1').getValue() !== 3) sep = ';';
  } catch (e) {
    sep = ';';
  }
  ss.deleteSheet(probe);
  return sep;
}

/** Converte as vírgulas de uma fórmula para o separador desta planilha. */
function F(formula) {
  return SEP === ',' ? formula : formula.split(',').join(SEP);
}

/** Versão de F() para matrizes usadas em setFormulas(). */
function FF(matriz) {
  return matriz.map(function (linha) { return linha.map(F); });
}


// ===========================================================================
//  MENU
// ===========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Precificação')
    .addItem('Criar / recriar planilha de precificação', 'criarPlanilha')
    .addToUi();
}


// ===========================================================================
//  ENTRADA PRINCIPAL
// ===========================================================================
function criarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  SEP = detectarSeparador(ss);

  // Remove SOMENTE as abas geradas por este script (as suas ficam intactas).
  ABAS.forEach(function (nome) {
    var s = ss.getSheetByName(nome);
    if (s) ss.deleteSheet(s);
  });

  // Ordem de criação importa: Tabela_ICMS é referenciada pelas outras.
  var shTab = ss.insertSheet('Tabela_ICMS');
  var shCfg = ss.insertSheet('Config');
  var shFat = ss.insertSheet('Faturamento_UF');
  var shPre = ss.insertSheet('Precificacao');

  // 1) estrutura estática
  montarTabelaICMS_estatico(shTab);
  montarConfig_estatico(shCfg);
  montarFaturamento_estatico(shFat);
  montarPrecificacao_estatico(shPre);

  // 2) named ranges (precisam existir antes das fórmulas que os usam)
  criarNamedRanges(ss, shCfg, shFat);

  // 3) fórmulas
  montarTabelaICMS_formulas(shTab);
  montarConfig_formulas(shCfg);
  montarFaturamento_formulas(shFat);
  montarPrecificacao_formulas(shPre);

  // 4) acabamento
  formatarTudo(shCfg, shTab, shFat, shPre);

  SpreadsheetApp.flush();
  ss.setActiveSheet(shPre);

  var erros = contarErros(ss);
  SpreadsheetApp.getUi().alert(
    'Planilha criada.\n\n' +
    'Idioma da planilha: ' + ss.getSpreadsheetLocale() + '\n' +
    'Separador de argumentos detectado: "' + SEP + '"\n' +
    'Células com erro de fórmula: ' + erros + '\n\n' +
    'Amarelo = você preenche.  Cinza = calculado.  Verde = resultado.\n\n' +
    'Comece por Config, depois Faturamento_UF, depois Precificacao.'
  );
}

/** Conta células em erro nas 4 abas — verificação pós-montagem. */
function contarErros(ss) {
  var n = 0;
  ABAS.forEach(function (nome) {
    var sh = ss.getSheetByName(nome);
    if (!sh) return;
    sh.getDataRange().getValues().forEach(function (linha) {
      linha.forEach(function (v) {
        if (typeof v === 'string' && v.indexOf('#') === 0 && v.length > 2) n++;
      });
    });
  });
  return n;
}


// ===========================================================================
//  ABA: Tabela_ICMS
// ===========================================================================
function montarTabelaICMS_estatico(sh) {
  sh.getRange('A1:L1').setValues([[
    'UF', 'Estado', 'Grupo', 'Alíq. interna', 'FCP',
    'Interna total', 'Interestadual NACIONAL', 'Interestadual IMPORTADA',
    'DIFAL efetivo % NACIONAL', 'DIFAL efetivo % IMPORTADA',
    'Interestadual (cenário ativo)', 'DIFAL efetivo % (cenário ativo)'
  ]]);

  var dados = UFS.map(function (u) { return [u[0], u[1], u[2], u[3], u[4]]; });
  sh.getRange(LIN_1, 1, N_UF, 5).setValues(dados);

  sh.getRange('N1').setValue('COMO ESTA ABA FUNCIONA');
  sh.getRange('N2:N15').setValues([
    ['Alíquota interestadual (Res. Senado 22/1989 e 13/2012):'],
    ['  • S/SE → S/SE ................ 12%'],
    ['  • qualquer outro cruzamento .. 7%   (o ES entra como N/NE/CO)'],
    ['  • mercadoria importada ....... 4%   (CST de origem 1, 2, 3 ou 8)'],
    ['  • venda dentro da própria UF . alíquota interna, e DIFAL = 0'],
    ['As três primeiras são editáveis em Config, caso o Senado mude.'],
    [''],
    ['DIFAL = alíquota interna do destino − alíquota interestadual.'],
    ['É REPARTIÇÃO do mesmo ICMS, não imposto adicional: a origem fica'],
    ['com a interestadual e o destino fica com a diferença.'],
    [''],
    ['Base dupla (LC 190/22, art. 13, §6º): a base do DIFAL é remontada'],
    ['"por dentro" com a alíquota interna do destino, o que eleva o DIFAL'],
    ['efetivo (ex.: MG sai de 6,00% para 7,32% do preço).']
  ]);
}

function montarTabelaICMS_formulas(sh) {
  var f = [];
  for (var i = LIN_1; i <= LIN_N; i++) {
    f.push([
      // F - interna total (interna + FCP, se habilitado)
      '=D' + i + '+IF(USA_FCP="Sim",E' + i + ',0)',
      // G - interestadual da SAÍDA, mercadoria nacional
      '=IF($A' + i + '=UF_ORIGEM,$F' + i +
        ',IF(AND($C' + i + '="S/SE",GRUPO_ORIGEM="S/SE"),ALQ_INTER_12,ALQ_INTER_7))',
      // H - interestadual da SAÍDA, mercadoria importada
      '=IF($A' + i + '=UF_ORIGEM,$F' + i + ',ALQ_INTER_IMP)',
      // I - DIFAL efetivo sobre o preço, mercadoria nacional
      '=IF(BASE_DUPLA="Sim",(1-$G' + i + ')/(1-$F' + i + ')*$F' + i + '-$G' + i +
        ',$F' + i + '-$G' + i + ')',
      // J - DIFAL efetivo sobre o preço, mercadoria importada
      '=IF(BASE_DUPLA="Sim",(1-$H' + i + ')/(1-$F' + i + ')*$F' + i + '-$H' + i +
        ',$F' + i + '-$H' + i + ')',
      // K - interestadual do cenário ativo
      '=IF(ORIGEM_MERC="Importada",$H' + i + ',$G' + i + ')',
      // L - DIFAL efetivo do cenário ativo
      '=IF(ORIGEM_MERC="Importada",$J' + i + ',$I' + i + ')'
    ]);
  }
  sh.getRange(LIN_1, 6, N_UF, 7).setFormulas(FF(f));
}


// ===========================================================================
//  ABA: Config
// ===========================================================================
function montarConfig_estatico(sh) {
  sh.getRange('A1').setValue('CONFIGURAÇÃO — PRECIFICAÇÃO LUCRO REAL');

  var linhas = [
    // [linha, rótulo, valor, tipo]  tipo: 'sec' | 'in' | 'auto'
    [ 3, '1) OPERAÇÃO', '', 'sec'],
    [ 4, 'UF de origem (seu estabelecimento)', 'SP', 'in'],
    [ 5, 'Grupo da UF de origem', '', 'auto'],
    [ 6, 'Origem da mercadoria', 'Nacional', 'in'],
    [ 7, 'Modalidade de aquisição', 'Revenda (compra no Brasil)', 'in'],
    [ 8, 'Alíquota interestadual — S/SE para S/SE', 0.12, 'in'],
    [ 9, 'Alíquota interestadual — demais cruzamentos', 0.07, 'in'],
    [10, 'Alíquota interestadual — mercadoria importada', 0.04, 'in'],
    [11, 'Crédito de ICMS acumulado é aproveitável?', 'Sim', 'in'],
    [12, 'Considerar FCP no DIFAL?', 'Não', 'in'],

    [14, '2) REGIME E TESES (Lucro Real — não cumulativo)', '', 'sec'],
    [15, 'PIS %', 0.0165, 'in'],
    [16, 'COFINS %', 0.076, 'in'],
    [17, 'Excluir ICMS da base de PIS/COFINS? (Tema 69 / RE 574.706)', 'Sim', 'in'],
    [18, 'Excluir ICMS da base do crédito? (Lei 14.592/23)', 'Sim', 'in'],
    [19, 'DIFAL com base dupla? (LC 190/22, art. 13, §6º)', 'Sim', 'in'],

    [21, '3) AQUISIÇÃO — REVENDA (compra no Brasil)', '', 'sec'],
    [22, 'Custo unitário do produto (sem IPI)', 283.87, 'in'],
    [23, 'UF do fornecedor', 'SP', 'in'],
    [24, 'Alíquota de crédito de ICMS na entrada', '', 'auto'],
    [25, 'IPI % na compra (custo — não gera crédito na revenda)', 0.09, 'in'],
    [26, 'IPI R$ na compra', '', 'auto'],
    [27, 'Frete e despesas de compra (por unidade)', 0, 'in'],
    [28, 'Custo de aquisição bruto', '', 'auto'],
    [29, 'Crédito de ICMS', '', 'auto'],
    [30, 'Base do crédito de PIS/COFINS', '', 'auto'],
    [31, 'Crédito de PIS/COFINS', '', 'auto'],

    [33, '4) AQUISIÇÃO — IMPORTAÇÃO PRÓPRIA', '', 'sec'],
    [34, 'O custo abaixo já está LÍQUIDO de créditos?', 'Não', 'in'],
    [35, 'Custo por unidade (valor aduaneiro/CIF, ou custo líquido)', 0, 'in'],
    [36, 'Imposto de Importação %   [só no modo "Não"]', 0, 'in'],
    [37, 'IPI %   [SEMPRE — é ele que gera o débito na saída]', 0, 'in'],
    [38, 'PIS-Importação %   [só no modo "Não"]', 0.021, 'in'],
    [39, 'COFINS-Importação %   [só no modo "Não"]', 0.0965, 'in'],
    [40, 'Despesas aduaneiras (Siscomex, THC, armazenagem, desembaraço)', 0, 'in'],
    [41, 'Frete interno pós-desembaraço', 0, 'in'],
    [42, 'II R$', '', 'auto'],
    [43, 'IPI R$', '', 'auto'],
    [44, 'PIS-Importação R$', '', 'auto'],
    [45, 'COFINS-Importação R$', '', 'auto'],
    [46, 'Base de cálculo do ICMS-Importação (por dentro)', '', 'auto'],
    [47, 'ICMS-Importação R$', '', 'auto'],
    [48, 'Custo de aquisição bruto', '', 'auto'],
    [49, 'Crédito de ICMS', '', 'auto'],
    [50, 'Crédito de PIS/COFINS', '', 'auto'],
    [51, 'Crédito de IPI', '', 'auto'],

    [53, '5) AQUISIÇÃO — CENÁRIO ATIVO (alimenta a precificação)', '', 'sec'],
    [54, 'Custo de aquisição bruto', '', 'auto'],
    [55, 'Crédito de ICMS na entrada', '', 'auto'],
    [56, 'Crédito de PIS/COFINS na entrada', '', 'auto'],
    [57, 'Crédito de IPI na entrada', '', 'auto'],
    [58, 'Alíquota de IPI na saída', '', 'auto'],
    [59, 'CUSTO LÍQUIDO DE CRÉDITOS (é este que vira o CMV)', '', 'res']
  ];

  linhas.forEach(function (l) {
    sh.getRange(l[0], 1).setValue(l[1]);
    if (l[3] === 'sec') {
      sh.getRange(l[0], 1, 1, 2).setBackground(COR_SECAO).setFontWeight('bold');
    } else {
      if (l[2] !== '') sh.getRange(l[0], 2).setValue(l[2]);
      sh.getRange(l[0], 2).setBackground(
        l[3] === 'in' ? COR_INPUT : (l[3] === 'res' ? COR_RESULT : COR_AUTO));
      if (l[3] === 'res') sh.getRange(l[0], 1, 1, 2).setFontWeight('bold');
    }
  });

  // Formatos
  [8, 9, 10, 15, 16, 24, 25, 36, 37, 38, 39, 58].forEach(function (r) {
    sh.getRange(r, 2).setNumberFormat(FMT_PCT);
  });
  [22, 26, 27, 28, 29, 30, 31, 35, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51,
   54, 55, 56, 57, 59].forEach(function (r) {
    sh.getRange(r, 2).setNumberFormat(FMT_MOEDA);
  });

  // Dropdowns
  var listaUF = UFS.map(function (u) { return u[0]; });
  dropdown(sh, 'B4',  listaUF);
  dropdown(sh, 'B23', listaUF);
  dropdown(sh, 'B6',  ['Nacional', 'Importada']);
  dropdown(sh, 'B7',  ['Revenda (compra no Brasil)', 'Importação própria']);
  ['B11', 'B12', 'B17', 'B18', 'B19', 'B34'].forEach(function (c) {
    dropdown(sh, c, ['Sim', 'Não']);
  });

  // Notas
  sh.getRange('D3').setValue('OBSERVAÇÕES');
  sh.getRange('D4:D53').setValues([
    ['• "Origem da mercadoria" = CST de origem do item. Importada (1/2/3/8) → saída'],
    ['  interestadual a 4%, INDEPENDENTE de quem importou. Se você compra de um'],
    ['  distribuidor nacional que importou, ainda é 4% na sua saída.'],
    [''],
    ['• "Modalidade" muda só o IPI e o crédito de entrada:'],
    ['     Revenda ............. IPI da compra é CUSTO, sem débito na saída.'],
    ['     Importação própria .. você vira equiparado a industrial: credita o IPI'],
    ['                           da DI e DEBITA IPI na saída.'],
    [''],
    ['• Vendendo a consumidor final NÃO contribuinte, o IPI INTEGRA a base do'],
    ['  ICMS (CF art. 155, §2º, XI — a exclusão só vale entre contribuintes).'],
    ['  Isso já está embutido nas fórmulas da aba Precificacao.'],
    [''],
    ['• Crédito acumulado: mercadoria importada credita ~18% na entrada e debita'],
    ['  só 4% na saída interestadual. O DIFAL vai para OUTRO estado e não abate'],
    ['  crédito de SP, então o excedente ACUMULA. Se você não consegue monetizar'],
    ['  esse saldo, marque "Não" em B11 — a planilha passa a limitar o crédito'],
    ['  aproveitado ao valor do débito próprio e mostra o que ficou preso.'],
    [''],
    ['• Tema 69 e Lei 14.592/23 puxam para lados opostos: o primeiro reduz o'],
    ['  débito de PIS/COFINS, o segundo reduz o crédito. Ambos são lei vigente.'],
    [''],
    ['• IRPJ/CSLL NÃO entram aqui. São apurados no fechamento, sobre o lucro'],
    ['  do período — depois das despesas fixas e das compensações. A margem'],
    ['  desta planilha é margem de contribuição, antes disso.'],
    [''],
    ['• B34 escolhe COMO você digita o custo da importação:'],
    [''],
    ['    "Não"  → B35 é o valor aduaneiro puro. A planilha calcula II, IPI,'],
    ['             PIS/COFINS e o ICMS "por dentro", soma no custo e devolve'],
    ['             como crédito. Use quando quiser ver a importação aberta.'],
    [''],
    ['    "Sim"  → B35 já é o custo líquido de créditos (a cotação que o'],
    ['             despachante te passa). A planilha não calcula nem credita'],
    ['             nada — usa o número como está. II, PIS-imp e COFINS-imp'],
    ['             passam a ser ignorados.'],
    [''],
    ['  Os dois modos dão a MESMA margem quando os números são coerentes.'],
    ['  O "Sim" só evita mostrar um custo bruto inflado e neutraliza o'],
    ['  limitador de crédito do B11, que não teria o que limitar.'],
    [''],
    ['• ATENÇÃO: o IPI % (B37) é lido nos DOIS modos. No "Sim" ele não gera'],
    ['  crédito, mas continua definindo o DÉBITO de IPI na saída — como'],
    ['  importador você é equiparado a industrial e deve IPI na venda.'],
    ['  Deixar B37 zerado apaga esse débito da conta.'],
    [''],
    ['• No modo "Não", ESTRANHE se o "custo de aquisição bruto" vier maior'],
    ['  que o que você pagou: é normal. O ICMS-importação é calculado "por'],
    ['  dentro" (entra na própria base), infla o custo bruto e volta inteiro'],
    ['  como crédito. Olhe sempre o CUSTO LÍQUIDO (linha 59).']
  ]);
}

function montarConfig_formulas(sh) {
  var ALQ_INT_ORIG = 'VLOOKUP(UF_ORIGEM,Tabela_ICMS!$A$' + LIN_1 + ':$F$' + LIN_N + ',6,FALSE)';

  sh.getRange('B5').setFormula(F(
    '=IFERROR(VLOOKUP(UF_ORIGEM,Tabela_ICMS!$A$' + LIN_1 + ':$C$' + LIN_N + ',3,FALSE),"?")'));

  // --- Revenda -------------------------------------------------------------
  // Crédito de entrada: mesma UF -> interna; UF diferente -> interestadual da
  // operação fornecedor→você (4% se a mercadoria for importada).
  sh.getRange('B24').setFormula(F(
    '=IF(R_UF_FORN=UF_ORIGEM,' + ALQ_INT_ORIG + ',' +
      'IF(ORIGEM_MERC="Importada",ALQ_INTER_IMP,' +
        'IF(AND(VLOOKUP(R_UF_FORN,Tabela_ICMS!$A$' + LIN_1 + ':$C$' + LIN_N + ',3,FALSE)="S/SE",' +
           'GRUPO_ORIGEM="S/SE"),ALQ_INTER_12,ALQ_INTER_7)))'));

  sh.getRange('B26').setFormula(F('=R_CUSTO*R_IPI_PCT'));
  sh.getRange('B28').setFormula(F('=R_CUSTO+R_IPI_RS+R_FRETE'));
  // IPI não integra a base do ICMS na compra para revenda (CF 155, §2º, XI).
  sh.getRange('B29').setFormula(F('=(R_CUSTO+R_FRETE)*R_ALQ_ICMS_ENT'));
  // IPI não recuperável integra o custo e entra na base do crédito (IN 2121/22,
  // art. 171, II). O ICMS sai da base desde a Lei 14.592/23.
  sh.getRange('B30').setFormula(F('=R_TOTAL-IF(EXCLUI_ICMS_CRED="Sim",R_CRED_ICMS,0)'));
  sh.getRange('B31').setFormula(F('=R_BASE_CRED_PC*(ALQ_PIS+ALQ_COFINS)'));

  // --- Importação própria --------------------------------------------------
  // Dois modos de entrada, controlados por I_CUSTO_LIQ (B34):
  //
  //   "Não"  → B35 é o valor aduaneiro. A planilha calcula II, IPI, PIS/COFINS
  //            e o ICMS "por dentro", soma tudo no custo e devolve como crédito.
  //   "Sim"  → B35 já é o custo líquido (ex.: cotação do despachante com os
  //            créditos abatidos). Zera todos os tributos e créditos daqui,
  //            porque eles já estão embutidos no número digitado.
  //
  // Os dois modos chegam ao MESMO CMV e à MESMA margem quando os números são
  // coerentes. O modo "Sim" só evita exibir um custo bruto inflado e neutraliza
  // o limitador de crédito (B11), que não teria o que limitar.
  //
  // O IPI % (B37) é lido nos DOIS modos: no "Sim" ele não gera crédito, mas
  // continua definindo o débito de IPI na saída (equiparado a industrial).
  var LIQ = 'I_CUSTO_LIQ="Sim"';

  sh.getRange('B42').setFormula(F('=IF(' + LIQ + ',0,I_VA*I_II_PCT)'));
  sh.getRange('B43').setFormula(F('=IF(' + LIQ + ',0,(I_VA+I_II_RS)*I_IPI_PCT)'));
  sh.getRange('B44').setFormula(F('=IF(' + LIQ + ',0,I_VA*I_PIS_PCT)'));   // base = valor aduaneiro (RE 559.937)
  sh.getRange('B45').setFormula(F('=IF(' + LIQ + ',0,I_VA*I_COFINS_PCT)'));
  sh.getRange('B46').setFormula(F(
    '=IF(' + LIQ + ',0,IFERROR((I_VA+I_II_RS+I_IPI_RS+I_PIS_RS+I_COFINS_RS+I_DESP)/(1-' +
    ALQ_INT_ORIG + '),0))'));
  sh.getRange('B47').setFormula(F('=I_BASE_ICMS*' + ALQ_INT_ORIG));
  sh.getRange('B48').setFormula(F(
    '=IF(' + LIQ + ',I_VA+I_DESP+I_FRETE_INT,' +
    'I_VA+I_II_RS+I_IPI_RS+I_PIS_RS+I_COFINS_RS+I_ICMS_RS+I_DESP+I_FRETE_INT)'));
  sh.getRange('B49').setFormula(F('=I_ICMS_RS'));
  sh.getRange('B50').setFormula(F('=I_PIS_RS+I_COFINS_RS'));
  sh.getRange('B51').setFormula(F('=I_IPI_RS'));

  // --- Cenário ativo -------------------------------------------------------
  var IMP = 'MODALIDADE="Importação própria"';
  sh.getRange('B54').setFormula(F('=IF(' + IMP + ',I_CUSTO_TOTAL,R_TOTAL)'));
  sh.getRange('B55').setFormula(F('=IF(' + IMP + ',I_CRED_ICMS,R_CRED_ICMS)'));
  sh.getRange('B56').setFormula(F('=IF(' + IMP + ',I_CRED_PC,R_CRED_PC)'));
  sh.getRange('B57').setFormula(F('=IF(' + IMP + ',I_CRED_IPI,0)'));
  sh.getRange('B58').setFormula(F('=IF(' + IMP + ',I_IPI_PCT,0)'));
  // O custo bruto é intermediário: no modo "Não" o ICMS-importação é calculado
  // "por dentro", infla o custo e volta inteiro como crédito. O que vira CMV
  // é o líquido — e no modo "Sim" ele é igual ao que você digitou.
  sh.getRange('B59').setFormula(F('=B54-B55-B56-B57'));

}


// ===========================================================================
//  ABA: Faturamento_UF
// ===========================================================================
function montarFaturamento_estatico(sh) {
  sh.getRange('A1:J1').setValues([[
    'UF', 'Estado', 'Faturamento R$', '% do total',
    'Interestadual % NACIONAL', 'Interestadual % IMPORTADA',
    'DIFAL efetivo % NACIONAL', 'DIFAL efetivo % IMPORTADA',
    'ICMS próprio R$ (cenário)', 'DIFAL R$ (cenário)'
  ]]);

  var dados = UFS.map(function (u) { return [u[0], u[1], u[5]]; });
  sh.getRange(LIN_1, 1, N_UF, 3).setValues(dados);
  sh.getRange(LIN_1, 3, N_UF, 1).setBackground(COR_INPUT);

  sh.getRange(LIN_TOT, 1).setValue('TOTAL');
  sh.getRange(LIN_TOT, 1, 1, 10).setFontWeight('bold').setBackground(COR_SECAO);

  // --- bloco de transparência da base de rateio ---
  var b = LIN_TOT + 2;
  sh.getRange(b, 1).setValue('BASE DO RATEIO');
  sh.getRange(b, 1, 1, 3).setBackground(COR_SECAO).setFontWeight('bold');
  sh.getRange(b + 1, 1, 2, 1).setValues([
    ['UFs com faturamento informado'],
    ['Faturamento considerado no rateio (R$)']
  ]);
  sh.getRange(b + 1, 2).setNumberFormat('0');
  sh.getRange(b + 2, 2).setNumberFormat(FMT_MOEDA);
  sh.getRange(b + 1, 2, 2, 1).setBackground(COR_AUTO);

  var r = LIN_TOT + 6;
  sh.getRange(r, 1).setValue('MÉDIAS PONDERADAS PELO SEU MIX DE ESTADOS');
  sh.getRange(r, 1, 1, 3).setBackground(COR_SECAO).setFontWeight('bold');
  sh.getRange(r + 1, 1, 10, 1).setValues([
    ['ICMS próprio médio — se TUDO for NACIONAL'],
    ['DIFAL médio — se TUDO for NACIONAL'],
    ['ICMS total médio — se TUDO for NACIONAL'],
    ['ICMS próprio médio — se TUDO for IMPORTADO'],
    ['DIFAL médio — se TUDO for IMPORTADO'],
    ['ICMS total médio — se TUDO for IMPORTADO'],
    [''],
    ['» ICMS próprio médio — CENÁRIO ATIVO'],
    ['» DIFAL médio — CENÁRIO ATIVO'],
    ['» ICMS total médio — CENÁRIO ATIVO']
  ]);
  sh.getRange(r + 8, 1, 3, 2).setBackground(COR_RESULT).setFontWeight('bold');
  sh.getRange(r + 1, 2, 10, 1).setNumberFormat(FMT_PCT);

  sh.getRange(b, 5).setValue('COMO LER');
  sh.getRange(b + 1, 5, 12, 1).setValues([
    ['Deixe em branco (ou zero) a UF onde você não vendeu — ou cujo faturamento'],
    ['você não conseguiu identificar. Ela é simplesmente ignorada: não entra no'],
    ['numerador nem no denominador da média, e o valor é abatido do total.'],
    ['A média sai apenas sobre o faturamento efetivamente rateado por UF.'],
    [''],
    ['Os dois cenários são calculados sempre, lado a lado, para você comparar.'],
    ['O "cenário ativo" é o que a aba Precificacao usa, e segue a Origem da'],
    ['mercadoria escolhida em Config!B6.'],
    [''],
    ['Repare que na mercadoria importada o DIFAL médio sobe muito (a origem'],
    ['fica com só 4%) enquanto o ICMS próprio despenca. O ICMS TOTAL é quase'],
    ['o mesmo nos dois cenários — o que muda é para QUEM vai.']
  ]);
}

function montarFaturamento_formulas(sh) {
  var T = 'Tabela_ICMS!$A:$L';
  var f = [];
  for (var i = LIN_1; i <= LIN_N; i++) {
    f.push([
      '=IFERROR($C' + i + '/$C$' + LIN_TOT + ',0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',7,FALSE),0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',8,FALSE),0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',9,FALSE),0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',10,FALSE),0)',
      '=$C' + i + '*IFERROR(VLOOKUP($A' + i + ',' + T + ',11,FALSE),0)',
      '=$C' + i + '*IFERROR(VLOOKUP($A' + i + ',' + T + ',12,FALSE),0)'
    ]);
  }
  sh.getRange(LIN_1, 4, N_UF, 7).setFormulas(FF(f));

  sh.getRange(LIN_TOT, 3).setFormula(F('=SUM(C' + LIN_1 + ':C' + LIN_N + ')'));
  sh.getRange(LIN_TOT, 4).setFormula(F('=SUM(D' + LIN_1 + ':D' + LIN_N + ')'));
  sh.getRange(LIN_TOT, 9).setFormula(F('=SUM(I' + LIN_1 + ':I' + LIN_N + ')'));
  sh.getRange(LIN_TOT, 10).setFormula(F('=SUM(J' + LIN_1 + ':J' + LIN_N + ')'));

  // Média ponderada = SUMPRODUCT(faturamento; alíquota) / faturamento total
  function pond(col) {
    return '=IFERROR(SUMPRODUCT($C$' + LIN_1 + ':$C$' + LIN_N + ',' +
           col + '$' + LIN_1 + ':' + col + '$' + LIN_N + ')/$C$' + LIN_TOT + ',0)';
  }

  // UFs em branco/zero saem da conta: não somam no total nem no SUMPRODUCT.
  var b = LIN_TOT + 2;
  sh.getRange(b + 1, 2).setFormula(F('=COUNTIF(C' + LIN_1 + ':C' + LIN_N + ',">0")'));
  sh.getRange(b + 2, 2).setFormula(F('=C' + LIN_TOT));

  var r = LIN_TOT + 7;
  sh.getRange(r,     2).setFormula(F(pond('$E')));           // ICMS próprio nacional
  sh.getRange(r + 1, 2).setFormula(F(pond('$G')));           // DIFAL nacional
  sh.getRange(r + 2, 2).setFormula(F('=B' + r + '+B' + (r + 1)));
  sh.getRange(r + 3, 2).setFormula(F(pond('$F')));           // ICMS próprio importado
  sh.getRange(r + 4, 2).setFormula(F(pond('$H')));           // DIFAL importado
  sh.getRange(r + 5, 2).setFormula(F('=B' + (r + 3) + '+B' + (r + 4)));
  sh.getRange(r + 7, 2).setFormula(F('=IF(ORIGEM_MERC="Importada",B' + (r + 3) + ',B' + r + ')'));
  sh.getRange(r + 8, 2).setFormula(F('=IF(ORIGEM_MERC="Importada",B' + (r + 4) + ',B' + (r + 1) + ')'));
  sh.getRange(r + 9, 2).setFormula(F('=B' + (r + 7) + '+B' + (r + 8)));
}


// ===========================================================================
//  ABA: Precificacao
// ===========================================================================
function montarPrecificacao_estatico(sh) {
  sh.getRange('A1').setValue('PRECIFICAÇÃO POR CANAL — LUCRO REAL');
  sh.getRange('A3:D3').setValues([['Item', 'Mercado Livre', 'Shopee', 'Canal 3']]);
  sh.getRange('A3:D3').setBackground(COR_TITULO).setFontColor('#ffffff').setFontWeight('bold');

  var L = [
    [ 4, 'Preço de venda ao consumidor (R$)', 'in',   [649.99, '', '']],
    [ 5, 'Comissão do canal (%)', 'in',                [0.12, 0.20, 0]],
    [ 6, 'Comissão (R$)', 'auto', null],
    [ 7, 'Frete grátis / subsídio de frete (R$)', 'in', [68.95, 0, 0]],
    [ 8, 'Taxa fixa (R$)', 'in',                       [0, 4.00, 0]],
    [ 9, 'Embalagem (R$)', 'in',                       [2.00, 0.50, 0]],
    [10, 'Outros custos variáveis (R$)', 'in',         [0, 0, 0]],
    [11, '% das taxas do canal com crédito de PIS/COFINS', 'in', [1, 1, 1]],
    [12, 'Crédito de PIS/COFINS sobre taxas (R$)', 'auto', null],

    [14, 'IMPOSTOS SOBRE A VENDA', 'sec', null],
    [15, 'Alíquota de IPI na saída', 'auto', null],
    [16, 'IPI destacado na saída (R$)', 'auto', null],
    [17, 'Receita bruta (R$)  [preço − IPI]', 'auto', null],
    [18, 'Alíquota de ICMS próprio — média ponderada', 'auto', null],
    [19, 'ICMS próprio devido à UF de origem (R$)', 'auto', null],
    [20, 'DIFAL médio efetivo %', 'auto', null],
    [21, 'DIFAL devido às UFs de destino (R$)', 'auto', null],
    [22, 'ICMS TOTAL da operação (R$)', 'auto', null],
    [23, 'ICMS total / preço de venda', 'auto', null],
    [24, 'Base de PIS/COFINS', 'auto', null],
    [25, 'PIS (R$)', 'auto', null],
    [26, 'COFINS (R$)', 'auto', null],

    [28, 'CUSTO E CRÉDITOS DE ENTRADA', 'sec', null],
    [29, 'Custo de aquisição bruto (R$)', 'auto', null],
    [30, 'Crédito de ICMS disponível (R$)', 'auto', null],
    [31, 'Crédito de ICMS aproveitado (R$)', 'auto', null],
    [32, 'Crédito de ICMS que ACUMULA / se perde (R$)', 'auto', null],
    [33, 'Crédito de PIS/COFINS na entrada (R$)', 'auto', null],
    [34, 'Crédito de IPI na entrada (R$)', 'auto', null],
    [35, 'CMV líquido (R$)', 'auto', null],

    [37, 'RESULTADO', 'sec', null],
    [38, 'Margem de contribuição (R$)', 'res', null],
    [39, 'Margem de contribuição (%)', 'res', null],

    [41, 'SIMULADOR — PREÇO PARA ATINGIR UMA MARGEM', 'sec', null],
    [42, 'MC alvo (% sobre a receita bruta)', 'in', [0.15, 0.15, 0.15]],
    [43, 'Coeficiente K (parcela do preço que sobra)', 'auto', null],
    [44, 'Custos fixos por unidade (R$)', 'auto', null],
    [45, 'PREÇO NECESSÁRIO (R$)', 'res', null]
  ];

  L.forEach(function (l) {
    sh.getRange(l[0], 1).setValue(l[1]);
    if (l[2] === 'sec') {
      sh.getRange(l[0], 1, 1, 4).setBackground(COR_SECAO).setFontWeight('bold');
    } else {
      var cor = l[2] === 'in' ? COR_INPUT : (l[2] === 'res' ? COR_RESULT : COR_AUTO);
      sh.getRange(l[0], 2, 1, 3).setBackground(cor);
      if (l[2] === 'res') sh.getRange(l[0], 1, 1, 4).setFontWeight('bold');
      if (l[3]) sh.getRange(l[0], 2, 1, 3).setValues([l[3]]);
    }
  });

  [5, 11, 15, 18, 20, 23, 39, 42, 43].forEach(function (r) {
    sh.getRange(r, 2, 1, 3).setNumberFormat(FMT_PCT);
  });
  [4, 6, 7, 8, 9, 10, 12, 16, 17, 19, 21, 22, 24, 25, 26,
   29, 30, 31, 32, 33, 34, 35, 38, 44, 45].forEach(function (r) {
    sh.getRange(r, 2, 1, 3).setNumberFormat(FMT_MOEDA);
  });

  sh.getRange('F3').setValue('NOTAS DE CÁLCULO');
  sh.getRange('F4:F26').setValues([
    ['ICMS (linhas 18-22) — o ponto central:'],
    ['   A origem NÃO fica com a alíquota interna. Numa venda interestadual a'],
    ['   consumidor final não contribuinte, a origem fica com a alíquota'],
    ['   INTERESTADUAL (7%, 12% ou 4%) e o destino fica com o DIFAL.'],
    ['   ICMS total = alíquota interna do destino (ou um pouco mais, com base'],
    ['   dupla). Somar "interna da origem + DIFAL" é contar duas vezes.'],
    [''],
    ['Base do ICMS = preço cheio, IPI incluído (venda a não contribuinte).'],
    [''],
    ['Base de PIS/COFINS (linha 24) = receita bruta − ICMS destacado, quando'],
    ['o Tema 69 está ligado. O IPI não é receita, então já sai antes.'],
    [''],
    ['Crédito de ICMS (linhas 30-32): se em Config você marcou que o saldo'],
    ['acumulado NÃO é aproveitável, o crédito usado é limitado ao débito'],
    ['próprio. O DIFAL não abate crédito da origem — ele é devido a outro'],
    ['estado. A linha 32 mostra quanto de crédito fica preso.'],
    [''],
    ['IRPJ e CSLL não entram aqui: são apurados no fechamento, sobre o lucro'],
    ['do período. A linha 39 é margem de CONTRIBUIÇÃO — o que sobra da venda'],
    ['para pagar despesa fixa e, só no fim, o imposto sobre o lucro.'],
    [''],
    ['Simulador (41-45): P = Custos fixos / (K − MC alvo / (1+IPI)). O crédito'],
    ['de ICMS limitado é tratado como constante — confira a linha 39 depois.']
  ]);
}

function montarPrecificacao_formulas(sh) {
  ['B', 'C', 'D'].forEach(function (c) {
    var m = {};
    for (var r = 4; r <= 45; r++) m[r] = c + r;

    var PC = '(ALQ_PIS+ALQ_COFINS)';

    sh.getRange(c + '6').setFormula(F('=' + m[4] + '*' + m[5]));
    sh.getRange(c + '12').setFormula(F(
      '=(' + m[6] + '+' + m[7] + '+' + m[8] + ')*' + m[11] + '*' + PC));

    sh.getRange(c + '15').setFormula(F('=ALQ_IPI_SAIDA'));
    sh.getRange(c + '16').setFormula(F('=' + m[4] + '*' + m[15] + '/(1+' + m[15] + ')'));
    sh.getRange(c + '17').setFormula(F('=' + m[4] + '-' + m[16]));
    sh.getRange(c + '18').setFormula(F('=ICMS_PROP_MED'));
    sh.getRange(c + '19').setFormula(F('=' + m[4] + '*' + m[18]));
    sh.getRange(c + '20').setFormula(F('=DIFAL_MED'));
    sh.getRange(c + '21').setFormula(F('=' + m[4] + '*' + m[20]));
    sh.getRange(c + '22').setFormula(F('=' + m[19] + '+' + m[21]));
    sh.getRange(c + '23').setFormula(F('=IFERROR(' + m[22] + '/' + m[4] + ',0)'));
    sh.getRange(c + '24').setFormula(F(
      '=' + m[17] + '-IF(EXCLUI_ICMS_PC="Sim",' + m[19] + ',0)'));
    sh.getRange(c + '25').setFormula(F('=' + m[24] + '*ALQ_PIS'));
    sh.getRange(c + '26').setFormula(F('=' + m[24] + '*ALQ_COFINS'));

    sh.getRange(c + '29').setFormula(F('=CUSTO_BRUTO'));
    sh.getRange(c + '30').setFormula(F('=CRED_ICMS_ENT'));
    sh.getRange(c + '31').setFormula(F(
      '=IF(CRED_ICMS_APROV="Sim",' + m[30] + ',MIN(' + m[30] + ',' + m[19] + '))'));
    sh.getRange(c + '32').setFormula(F('=' + m[30] + '-' + m[31]));
    sh.getRange(c + '33').setFormula(F('=CRED_PC_ENT'));
    sh.getRange(c + '34').setFormula(F('=CRED_IPI_ENT'));
    sh.getRange(c + '35').setFormula(F(
      '=' + m[29] + '-' + m[31] + '-' + m[33] + '-' + m[34]));

    sh.getRange(c + '38').setFormula(F(
      '=IF(N(' + m[4] + ')=0,"",' +
      m[4] + '-' + m[6] + '-' + m[7] + '-' + m[8] + '-' + m[9] + '-' + m[10] +
      '+' + m[12] + '-' + m[35] + '-' + m[16] + '-' + m[19] + '-' + m[21] +
      '-' + m[25] + '-' + m[26] + ')'));
    sh.getRange(c + '39').setFormula(F('=IFERROR(' + m[38] + '/' + m[17] + ',"")'));

    // K = 1 − comissão − IPI/(1+IPI) − ICMS próprio − DIFAL
    //       − (PIS+COFINS)*(1/(1+IPI) − Tema69*ICMS próprio)
    //       + comissão*%crédito*(PIS+COFINS)
    sh.getRange(c + '43').setFormula(F(
      '=1-' + m[5] + '-' + m[15] + '/(1+' + m[15] + ')-' + m[18] + '-' + m[20] +
      '-' + PC + '*(1/(1+' + m[15] + ')-IF(EXCLUI_ICMS_PC="Sim",' + m[18] + ',0))' +
      '+' + m[5] + '*' + m[11] + '*' + PC));
    sh.getRange(c + '44').setFormula(F(
      '=' + m[7] + '+' + m[8] + '+' + m[9] + '+' + m[10] + '+' + m[35] +
      '-(' + m[7] + '+' + m[8] + ')*' + m[11] + '*' + PC));
    sh.getRange(c + '45').setFormula(F(
      '=IFERROR(' + m[44] + '/(' + m[43] + '-' + m[42] + '/(1+' + m[15] + ')),"")'));
  });

  // Alertas visuais
  var regra1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setBackground(COR_ALERTA)
    .setRanges([sh.getRange('B38:D39')]).build();
  var regra2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground(COR_ALERTA)
    .setRanges([sh.getRange('B32:D32')]).build();
  sh.setConditionalFormatRules([regra1, regra2]);
}


// ===========================================================================
//  NAMED RANGES
// ===========================================================================
function criarNamedRanges(ss, cfg, fat) {
  var r = LIN_TOT + 7;
  var nomes = [
    ['UF_ORIGEM',        cfg, 'B4'],
    ['GRUPO_ORIGEM',     cfg, 'B5'],
    ['ORIGEM_MERC',      cfg, 'B6'],
    ['MODALIDADE',       cfg, 'B7'],
    ['ALQ_INTER_12',     cfg, 'B8'],
    ['ALQ_INTER_7',      cfg, 'B9'],
    ['ALQ_INTER_IMP',    cfg, 'B10'],
    ['CRED_ICMS_APROV',  cfg, 'B11'],
    ['USA_FCP',          cfg, 'B12'],

    ['ALQ_PIS',          cfg, 'B15'],
    ['ALQ_COFINS',       cfg, 'B16'],
    ['EXCLUI_ICMS_PC',   cfg, 'B17'],
    ['EXCLUI_ICMS_CRED', cfg, 'B18'],
    ['BASE_DUPLA',       cfg, 'B19'],

    ['R_CUSTO',          cfg, 'B22'],
    ['R_UF_FORN',        cfg, 'B23'],
    ['R_ALQ_ICMS_ENT',   cfg, 'B24'],
    ['R_IPI_PCT',        cfg, 'B25'],
    ['R_IPI_RS',         cfg, 'B26'],
    ['R_FRETE',          cfg, 'B27'],
    ['R_TOTAL',          cfg, 'B28'],
    ['R_CRED_ICMS',      cfg, 'B29'],
    ['R_BASE_CRED_PC',   cfg, 'B30'],
    ['R_CRED_PC',        cfg, 'B31'],

    ['I_CUSTO_LIQ',      cfg, 'B34'],
    ['I_VA',             cfg, 'B35'],
    ['I_II_PCT',         cfg, 'B36'],
    ['I_IPI_PCT',        cfg, 'B37'],
    ['I_PIS_PCT',        cfg, 'B38'],
    ['I_COFINS_PCT',     cfg, 'B39'],
    ['I_DESP',           cfg, 'B40'],
    ['I_FRETE_INT',      cfg, 'B41'],
    ['I_II_RS',          cfg, 'B42'],
    ['I_IPI_RS',         cfg, 'B43'],
    ['I_PIS_RS',         cfg, 'B44'],
    ['I_COFINS_RS',      cfg, 'B45'],
    ['I_BASE_ICMS',      cfg, 'B46'],
    ['I_ICMS_RS',        cfg, 'B47'],
    ['I_CUSTO_TOTAL',    cfg, 'B48'],
    ['I_CRED_ICMS',      cfg, 'B49'],
    ['I_CRED_PC',        cfg, 'B50'],
    ['I_CRED_IPI',       cfg, 'B51'],

    ['CUSTO_BRUTO',      cfg, 'B54'],
    ['CRED_ICMS_ENT',    cfg, 'B55'],
    ['CRED_PC_ENT',      cfg, 'B56'],
    ['CRED_IPI_ENT',     cfg, 'B57'],
    ['ALQ_IPI_SAIDA',    cfg, 'B58'],
    ['CUSTO_LIQUIDO',    cfg, 'B59'],

    ['ICMS_PROP_NAC',    fat, 'B' + r],
    ['DIFAL_MED_NAC',    fat, 'B' + (r + 1)],
    ['ICMS_PROP_IMP',    fat, 'B' + (r + 3)],
    ['DIFAL_MED_IMP',    fat, 'B' + (r + 4)],
    ['ICMS_PROP_MED',    fat, 'B' + (r + 7)],
    ['DIFAL_MED',        fat, 'B' + (r + 8)]
  ];

  nomes.forEach(function (n) {
    ss.setNamedRange(n[0], n[1].getRange(n[2]));
  });
}


// ===========================================================================
//  UTILITÁRIOS
// ===========================================================================
function dropdown(sh, a1, valores) {
  sh.getRange(a1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(valores, true)
      .setAllowInvalid(false)
      .build());
}

function formatarTudo(cfg, tab, fat, pre) {
  [cfg, tab, fat, pre].forEach(function (sh) {
    sh.getRange('A1').setFontSize(13).setFontWeight('bold').setFontColor(COR_TITULO);
  });

  cfg.setColumnWidth(1, 400); cfg.setColumnWidth(2, 150); cfg.setColumnWidth(3, 25);
  cfg.setColumnWidth(4, 620);
  cfg.setFrozenRows(1);

  tab.getRange('A1:L1').setBackground(COR_TITULO).setFontColor('#ffffff')
     .setFontWeight('bold').setWrap(true);
  tab.setColumnWidth(1, 55); tab.setColumnWidth(2, 165);
  for (var i = 3; i <= 12; i++) tab.setColumnWidth(i, 105);
  tab.setColumnWidth(13, 25); tab.setColumnWidth(14, 560);
  tab.getRange(LIN_1, 4, N_UF, 9).setNumberFormat(FMT_PCT);
  tab.setFrozenRows(1); tab.setFrozenColumns(2);

  fat.getRange('A1:J1').setBackground(COR_TITULO).setFontColor('#ffffff')
     .setFontWeight('bold').setWrap(true);
  fat.setColumnWidth(1, 55); fat.setColumnWidth(2, 165);
  for (var j = 3; j <= 10; j++) fat.setColumnWidth(j, 115);
  fat.getRange(LIN_1, 3, N_UF + 1, 1).setNumberFormat(FMT_MOEDA);
  fat.getRange(LIN_1, 4, N_UF + 1, 5).setNumberFormat(FMT_PCT);
  fat.getRange(LIN_1, 9, N_UF + 1, 2).setNumberFormat(FMT_MOEDA);
  fat.getRange(LIN_TOT + 2, 1).setFontWeight('bold');
  fat.getRange(LIN_TOT + 6, 1).setFontWeight('bold');
  fat.setFrozenRows(1); fat.setFrozenColumns(2);

  pre.setColumnWidth(1, 340);
  pre.setColumnWidth(2, 130); pre.setColumnWidth(3, 130); pre.setColumnWidth(4, 130);
  pre.setColumnWidth(5, 25);  pre.setColumnWidth(6, 620);
  pre.setFrozenRows(3); pre.setFrozenColumns(1);
}
