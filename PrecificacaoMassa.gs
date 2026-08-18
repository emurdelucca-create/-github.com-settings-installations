/**
 * ============================================================================
 *  PRECIFICAÇÃO EM MASSA — LUCRO REAL — MARKETPLACES
 *  Uma linha por anúncio/SKU. ICMS e DIFAL pela EC 87/2015 + LC 190/2022.
 * ============================================================================
 *
 *  COMO USAR
 *  1) Numa planilha NOVA: Extensões > Apps Script
 *  2) Cole este arquivo inteiro, salve
 *  3) Execute  criarPlanilha()   (autorize na primeira vez)
 *  4) Recarregue — aparece o menu "⚙️ Precificação em massa"
 *
 *  O script só escreve as fórmulas. Depois disso o cálculo é nativo do Sheets
 *  e recalcula sozinho. Cria 4 abas e não encosta em nenhuma outra:
 *      Precificacao | Parametros | Tabela_ICMS | Faturamento_UF
 *
 *  ---------------------------------------------------------------------------
 *  LOCALE — por que existe a função F()
 *  ---------------------------------------------------------------------------
 *  Em pt-BR o Sheets separa argumentos por ";" e não por ",". Fórmula gravada
 *  com vírgula vira #ERROR!. E literal decimal com ponto (0.12) não dá erro —
 *  o ponto é separador de milhar, então 0.12 vira 12. Errado e silencioso.
 *
 *  Por isso: detectarSeparador() descobre o separador testando =SUM(1,2) numa
 *  aba temporária, F() converte as vírgulas, e NENHUMA fórmula contém literal
 *  decimal — as alíquotas moram em células da aba Parametros.
 *
 *  INVARIANTE que faz o F() ser seguro: nenhum texto entre aspas dentro das
 *  fórmulas contém vírgula, e não há literal decimal. Preserve isso ao editar.
 * ============================================================================
 */

var ABAS = ['Precificacao', 'Parametros', 'Tabela_ICMS', 'Faturamento_UF'];

var SEP = ',';          // definido em criarPlanilha() por detectarSeparador()
var LIN_DADOS = 5;      // primeira linha de dados da aba Precificacao
var N_LINHAS  = 500;    // quantas linhas de SKU preparar

var FMT_MOEDA = 'R$ #,##0.00';
var FMT_PCT   = '0.00%';

var COR_TITULO = '#1c3d5a';
var COR_SECAO  = '#dbe5f1';
var COR_INPUT  = '#fff2cc';
var COR_AUTO   = '#f0f0f0';
var COR_RESULT = '#d9ead3';
var COR_ALERTA = '#f4cccc';

var CANAIS = ['Mercado Livre', 'Shopee', 'TikTok'];

/**
 * UF | Estado | Grupo (S/SE ou N/NE/CO) | Alíquota interna | FCP | Faturamento
 *
 * Grupo define a alíquota interestadual: S/SE -> S/SE = 12%, qualquer outro
 * cruzamento = 7% (Resolução do Senado 22/1989). O ES conta como N/NE/CO.
 *
 * A alíquota interna já embute o adicional/FECP dos estados que o cobram
 * dentro da alíquota cheia (ex.: RJ 20% + 2% = 22%). Por isso o FCP vai
 * zerado — confira com a contabilidade antes de mexer, para não contar duas
 * vezes o mesmo adicional.
 */
var UFS = [
  ['AC', 'Acre',                'N/NE/CO', 0.190, 0,     5739.36],
  ['AL', 'Alagoas',             'N/NE/CO', 0.200, 0,    42020.26],
  ['AP', 'Amapá',               'N/NE/CO', 0.180, 0,     3384.78],
  ['AM', 'Amazonas',            'N/NE/CO', 0.200, 0,    20726.67],
  ['BA', 'Bahia',               'N/NE/CO', 0.205, 0,   215351.57],
  ['CE', 'Ceará',               'N/NE/CO', 0.200, 0,    97517.55],
  ['DF', 'Distrito Federal',    'N/NE/CO', 0.200, 0,    18468.01],
  ['ES', 'Espírito Santo',      'N/NE/CO', 0.170, 0,    69033.55],
  ['GO', 'Goiás',               'N/NE/CO', 0.190, 0,    81697.44],
  ['MA', 'Maranhão',            'N/NE/CO', 0.230, 0,    90050.98],
  ['MT', 'Mato Grosso',         'N/NE/CO', 0.170, 0,    44814.70],
  ['MS', 'Mato Grosso do Sul',  'N/NE/CO', 0.170, 0,    44444.39],
  ['MG', 'Minas Gerais',        'S/SE',    0.180, 0,   305715.80],
  ['PA', 'Pará',                'N/NE/CO', 0.190, 0,    68296.02],
  ['PB', 'Paraíba',             'N/NE/CO', 0.200, 0,    58681.35],
  ['PR', 'Paraná',              'S/SE',    0.195, 0,   127810.82],
  ['PE', 'Pernambuco',          'N/NE/CO', 0.205, 0,   134125.25],
  ['PI', 'Piauí',               'N/NE/CO', 0.225, 0,    38922.37],
  ['RJ', 'Rio de Janeiro',      'S/SE',    0.220, 0,   198329.22],
  ['RN', 'Rio Grande do Norte', 'N/NE/CO', 0.200, 0,    42685.22],
  ['RS', 'Rio Grande do Sul',   'S/SE',    0.170, 0,    89297.15],
  ['RO', 'Rondônia',            'N/NE/CO', 0.195, 0,    10886.11],
  ['RR', 'Roraima',             'N/NE/CO', 0.200, 0,     1529.04],
  ['SC', 'Santa Catarina',      'S/SE',    0.170, 0,    80969.88],
  ['SP', 'São Paulo',           'S/SE',    0.180, 0,   537245.38],
  ['SE', 'Sergipe',             'N/NE/CO', 0.200, 0,    34095.88],
  ['TO', 'Tocantins',           'N/NE/CO', 0.200, 0,    22475.91]
];

var N_UF    = UFS.length;
var LIN_1   = 2;
var LIN_N   = LIN_1 + N_UF - 1;
var LIN_TOT = LIN_N + 1;


// ===========================================================================
//  COLUNAS DA ABA Precificacao
//  [letra, bloco (linha 1), subgrupo (linha 3), título (linha 4), tipo, formato]
//  tipo: 'in' = você preenche | 'fx' = fórmula | 'res' = fórmula em destaque
// ===========================================================================
var COLS = [
  ['A',  'SKU',                    '',                   '',                              'in',  'texto'],
  ['B',  'Comissões Marketplace',  'Canal de Venda',     '',                              'in',  'texto'],
  ['C',  '',                       'Estrutura Comissão', 'Comissão %',                    'in',  'pct'  ],
  ['D',  '',                       '',                   'Taxa Fixa R$',                  'in',  'money'],
  ['E',  '',                       '',                   'Frete Grátis R$',               'in',  'money'],
  ['F',  '',                       '',                   'Outros Custos R$',              'in',  'money'],
  ['G',  '',                       '',                   'Total Comissões R$',            'fx',  'money'],
  ['H',  '',                       'Crédito Imposto',    'Crédito Pis Cofins R$',         'fx',  'money'],
  ['I',  'Venda',                  '',                   'Preço de venda',                'in',  'money'],
  ['J',  '',                       '',                   'MC R$',                         'res', 'money'],
  ['K',  '',                       '',                   'MC %',                          'res', 'pct'  ],
  ['L',  'Simulador',              '',                   'MC % Objetiva',                 'in',  'pct'  ],
  ['M',  '',                       '',                   'MC R$',                         'fx',  'money'],
  ['N',  '',                       '',                   'Valor de Venda',                'res', 'money'],
  ['O',  'Custos Operacionais',    '',                   'Embalagem',                     'in',  'money'],
  ['P',  'Compra',                 '',                   'Origem (Nacional ou Importada)','in',  'texto'],
  ['Q',  '',                       '',                   'CMV já com créditos? (Sim/Não)','in',  'texto'],
  ['R',  '',                       '',                   'UF meu estabelecimento',        'in',  'texto'],
  ['S',  '',                       '',                   'UF de compra',                  'in',  'texto'],
  ['T',  '',                       '',                   'ICMS %',                        'fx',  'pct'  ],
  ['U',  '',                       '',                   'Monofásico? (ver dropdown)',    'in',  'texto'],
  ['V',  '',                       '',                   'Pis %',                         'in',  'pct'  ],
  ['W',  '',                       '',                   'Cofins %',                      'in',  'pct'  ],
  ['X',  '',                       '',                   'ST %',                          'in',  'pct'  ],
  ['Y',  '',                       '',                   'IPI %',                         'in',  'pct'  ],
  ['Z',  '',                       '',                   'IPI só custo ou déb. e créd.?', 'in',  'texto'],
  ['AA', '',                       '',                   'Custo de compra R$',            'in',  'money'],
  ['AB', 'Créditos',               '',                   'ICMS',                          'fx',  'money'],
  ['AC', '',                       '',                   'Pis',                           'fx',  'money'],
  ['AD', '',                       '',                   'Cofins',                        'fx',  'money'],
  ['AE', '',                       '',                   'IPI R$',                        'fx',  'money'],
  ['AF', 'Custo na compra',        '',                   'IPI R$',                        'fx',  'money'],
  ['AG', '',                       '',                   'ST R$',                         'fx',  'money'],
  ['AH', '',                       '',                   'CMV líquido',                   'res', 'money'],
  ['AI', 'Débitos',                'ICMS',               'ICMS próprio R$',               'fx',  'money'],
  ['AJ', '',                       'Pis',                'Pis R$',                        'fx',  'money'],
  ['AK', '',                       'Cofins',             'Cofins R$',                     'fx',  'money'],
  ['AL', '',                       'IPI R$',             'IPI R$',                        'fx',  'money'],
  ['AM', '',                       'Difal',              'Difal Médio %',                 'fx',  'pct'  ],
  ['AN', '',                       '',                   'Difal Médio R$',                'fx',  'money'],
  ['AO', 'Diagnóstico',            '',                   'ICMS total / preço',            'fx',  'pct'  ],
  ['AP', '',                       '',                   'Crédito ICMS que acumula R$',   'fx',  'money']
];


// ===========================================================================
//  LOCALE
// ===========================================================================

/**
 * Descobre o separador de argumentos que ESTA planilha aceita, testando o
 * comportamento real em vez de adivinhar pelo código de idioma.
 *   en_US : =SUM(1,2) -> 3     (vírgula separa argumentos)
 *   pt-BR : =SUM(1,2) -> 1,2   (vírgula é decimal; um argumento só)
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

function F(formula) {
  return SEP === ',' ? formula : formula.split(',').join(SEP);
}

function FF(matriz) {
  return matriz.map(function (linha) { return linha.map(F); });
}


// ===========================================================================
//  MENU / ENTRADA
// ===========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Precificação em massa')
    .addItem('Criar / recriar planilha', 'criarPlanilha')
    .addToUi();
}

function criarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SEP = detectarSeparador(ss);

  // Uma planilha não pode ficar sem nenhuma aba. Se as 4 abas do script forem
  // as únicas do documento, apagar a última estoura e deixa o arquivo quebrado.
  // A aba-guarda segura o documento durante a exclusão e sai no fim.
  var guarda = ss.insertSheet('__guarda__');

  ABAS.forEach(function (nome) {
    var s = ss.getSheetByName(nome);
    if (s) ss.deleteSheet(s);
  });

  var shTab = ss.insertSheet('Tabela_ICMS');
  var shPar = ss.insertSheet('Parametros');
  var shFat = ss.insertSheet('Faturamento_UF');
  var shPre = ss.insertSheet('Precificacao');

  montarTabelaICMS_estatico(shTab);
  montarParametros_estatico(shPar);
  montarFaturamento_estatico(shFat);
  montarPrecificacao_estatico(shPre);

  criarNamedRanges(ss, shPar, shFat);

  montarTabelaICMS_formulas(shTab);
  montarFaturamento_formulas(shFat);
  montarPrecificacao_formulas(shPre);

  formatarTudo(shPar, shTab, shFat, shPre);

  ss.deleteSheet(guarda);
  SpreadsheetApp.flush();
  ss.setActiveSheet(shPre);

  SpreadsheetApp.getUi().alert(
    'Planilha criada.\n\n' +
    'Idioma: ' + ss.getSpreadsheetLocale() + '\n' +
    'Separador detectado: "' + SEP + '"\n' +
    'Células com erro: ' + contarErros(ss) + '\n\n' +
    'Amarelo = você preenche.  Cinza = fórmula.  Verde = resultado.\n\n' +
    '1) Confira a aba Parametros\n' +
    '2) Confira o faturamento por UF em Faturamento_UF\n' +
    '3) Preencha uma linha por SKU na aba Precificacao (a partir da linha 5)'
  );
}

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
//  ABA: Parametros
// ===========================================================================
function montarParametros_estatico(sh) {
  sh.getRange('A1').setValue('PARÂMETROS GLOBAIS — valem para TODAS as linhas');

  var linhas = [
    [ 3, '1) SEU ESTABELECIMENTO', '', 'sec'],
    [ 4, 'UF de origem (usada no rateio do DIFAL)', 'SP', 'in'],
    [ 5, 'Grupo da UF de origem', '', 'auto'],

    [ 7, '2) ALÍQUOTAS INTERESTADUAIS (Resolução do Senado)', '', 'sec'],
    [ 8, 'S/SE para S/SE', 0.12, 'in'],
    [ 9, 'Demais cruzamentos', 0.07, 'in'],
    [10, 'Mercadoria importada (Res. 13/2012)', 0.04, 'in'],

    [12, '3) REGIME E TESES (Lucro Real — não cumulativo)', '', 'sec'],
    [13, 'PIS % geral (usado no crédito sobre taxas do canal)', 0.0165, 'in'],
    [14, 'COFINS % geral (idem)', 0.076, 'in'],
    [15, '% das taxas do canal que geram crédito de PIS/COFINS', 1, 'in'],
    [16, 'Excluir ICMS da base de PIS/COFINS? (Tema 69 / RE 574.706)', 'Sim', 'in'],
    [17, 'Excluir ICMS da base do crédito? (Lei 14.592/23)', 'Sim', 'in'],
    [18, 'DIFAL com base dupla? (LC 190/22, art. 13, §6º)', 'Sim', 'in'],
    [19, 'Considerar FCP no DIFAL?', 'Não', 'in'],
    [20, 'Crédito de ICMS acumulado é aproveitável?', 'Sim', 'in'],
    [21, 'Texto da opção "revendedor" da coluna U (não altere)', 'Sim - revendedor', 'auto'],

    [23, '4) MÉDIAS DO SEU MIX DE ESTADOS (vêm de Faturamento_UF)', '', 'sec'],
    [24, 'ICMS próprio médio — mercadoria NACIONAL', '', 'auto'],
    [25, 'DIFAL médio — mercadoria NACIONAL', '', 'auto'],
    [26, 'ICMS próprio médio — mercadoria IMPORTADA', '', 'auto'],
    [27, 'DIFAL médio — mercadoria IMPORTADA', '', 'auto'],

    [29, '5) PADRÕES SUGERIDOS PARA LINHAS NOVAS (referência, não alimenta nada)', '', 'sec'],
    [30, 'PIS % do item — regime normal', 0.0165, 'in'],
    [31, 'COFINS % do item — regime normal', 0.076, 'in'],
    [32, 'PIS % — fabricante/importador de autopeça monofásica', 0.023, 'in'],
    [33, 'COFINS % — fabricante/importador de autopeça monofásica', 0.108, 'in']
  ];

  linhas.forEach(function (l) {
    sh.getRange(l[0], 1).setValue(l[1]);
    if (l[3] === 'sec') {
      sh.getRange(l[0], 1, 1, 2).setBackground(COR_SECAO).setFontWeight('bold');
    } else {
      if (l[2] !== '') sh.getRange(l[0], 2).setValue(l[2]);
      sh.getRange(l[0], 2).setBackground(l[3] === 'in' ? COR_INPUT : COR_AUTO);
    }
  });

  [8, 9, 10, 13, 14, 15, 24, 25, 26, 27, 30, 31, 32, 33].forEach(function (r) {
    sh.getRange(r, 2).setNumberFormat(FMT_PCT);
  });

  var listaUF = UFS.map(function (u) { return u[0]; });
  dropdown(sh, 'B4', listaUF);
  ['B16', 'B17', 'B18', 'B19', 'B20'].forEach(function (c) {
    dropdown(sh, c, ['Sim', 'Não']);
  });

  sh.getRange('B5').setFormula(F(
    '=IFERROR(VLOOKUP(UF_ORIGEM,Tabela_ICMS!$A$' + LIN_1 + ':$C$' + LIN_N + ',3,FALSE),"?")'));
  sh.getRange('B24').setFormula(F('=ICMS_PROP_NAC'));
  sh.getRange('B25').setFormula(F('=DIFAL_MED_NAC'));
  sh.getRange('B26').setFormula(F('=ICMS_PROP_IMP'));
  sh.getRange('B27').setFormula(F('=DIFAL_MED_IMP'));

  sh.getRange('D3').setValue('O QUE CADA COISA FAZ');
  sh.getRange('D4:D55').setValues([
    ['• Tema 69 (linha 16): o ICMS destacado não integra a base de PIS/COFINS.'],
    ['  Decisão definitiva do STF, vale para todo mundo. Deixe "Sim".'],
    [''],
    ['• Lei 14.592/23 (linha 17): o ICMS da nota de COMPRA não gera crédito de'],
    ['  PIS/COFINS. Obrigatório desde 05/2023. Puxa para o lado oposto do Tema'],
    ['  69 — um reduz o débito, o outro reduz o crédito. Deixe "Sim".'],
    [''],
    ['• Base dupla (linha 18): a base do DIFAL é remontada "por dentro" com a'],
    ['  alíquota interna do destino (LC 190/22, art. 13, §6º). Eleva o DIFAL'],
    ['  efetivo — em MG sai de 6,00% para 7,32% do preço.'],
    [''],
    ['• As médias da seção 4 saem do SEU mix de estados, na aba Faturamento_UF.'],
    ['  Cada linha da Precificacao escolhe entre a média NACIONAL e a IMPORTADA'],
    ['  conforme a coluna P daquela linha.'],
    [''],
    ['• A UF de origem (linha 4) vale para o rateio do DIFAL de todas as linhas.'],
    ['  A coluna R da Precificacao é usada só para achar a alíquota de crédito'],
    ['  da compra. Se você tiver estabelecimento em mais de uma UF, o DIFAL'],
    ['  médio não vai refletir isso — rode uma planilha por estabelecimento.'],
    [''],
    ['• IRPJ e CSLL não entram: são apurados no fechamento, sobre o lucro do'],
    ['  período. O que a planilha calcula é margem de CONTRIBUIÇÃO.'],
    [''],
    ['• Coluna U (monofásico) tem TRÊS estados, e a diferença é grande:'],
    [''],
    ['    "Não" ................... PIS/COFINS normais dos dois lados.'],
    [''],
    ['    "Sim - revendedor" ...... você COMPRA de um fabricante/importador'],
    ['                              brasileiro. Saída a ZERO e sem crédito'],
    ['                              na entrada. É o caso de quem revende.'],
    [''],
    ['    "Sim - fabricante/importador" ... você é quem IMPORTA ou fabrica.'],
    ['                              A concentração cai em VOCÊ: não há'],
    ['                              alíquota zero. Preencha V e W com as'],
    ['                              alíquotas concentradas do seu produto'],
    ['                              (autopeças da Lei 10.485: 2,3% e 10,8%).'],
    ['                              A mercadoria segue sem gerar crédito.'],
    [''],
    ['  O crédito de PIS/COFINS sobre as taxas do canal (coluna H) NÃO segue'],
    ['  a coluna U: ele vem da nota de SERVIÇO que o marketplace emite contra'],
    ['  você, e o monofásico da mercadoria não alcança esse serviço. Se quiser'],
    ['  desligar ou reduzir esse crédito, use a linha 15 desta aba.'],
    [''],
    ['  Ou seja: importar direto um item monofásico tira a alíquota zero e'],
    ['  ainda paga MAIS que o regime normal (13,1% contra 9,25%). Confirme'],
    ['  as alíquotas do seu produto no art. 3º da lei antes de usar.'],
    [''],
    ['• LIMITAÇÃO — ST: quando a coluna X tem ST, a planilha trata o ST como'],
    ['  custo e zera o crédito de ICMS da compra, mas segue debitando ICMS e'],
    ['  DIFAL normais na venda. Isso está certo para a venda interestadual a'],
    ['  consumidor final, mas ignora o direito a ressarcimento do ST retido.'],
    ['  Itens com ST vão aparecer PIOR do que realmente são.']
  ]);
}


// ===========================================================================
//  ABA: Tabela_ICMS
// ===========================================================================
function montarTabelaICMS_estatico(sh) {
  sh.getRange('A1:J1').setValues([[
    'UF', 'Estado', 'Grupo', 'Alíq. interna', 'FCP', 'Interna total',
    'Interestadual NACIONAL', 'Interestadual IMPORTADA',
    'DIFAL efetivo % NACIONAL', 'DIFAL efetivo % IMPORTADA'
  ]]);
  sh.getRange(LIN_1, 1, N_UF, 5).setValues(
    UFS.map(function (u) { return [u[0], u[1], u[2], u[3], u[4]]; }));

  sh.getRange('L1').setValue('COMO ESTA ABA FUNCIONA');
  sh.getRange('L2:L14').setValues([
    ['Alíquota interestadual (Res. Senado 22/1989 e 13/2012):'],
    ['  • S/SE → S/SE ................ 12%'],
    ['  • qualquer outro cruzamento .. 7%   (o ES entra como N/NE/CO)'],
    ['  • mercadoria importada ....... 4%   (CST de origem 1, 2, 3 ou 8)'],
    ['  • venda dentro da própria UF . alíquota interna, e DIFAL = 0'],
    [''],
    ['DIFAL = alíquota interna do destino − alíquota interestadual.'],
    ['É REPARTIÇÃO do mesmo ICMS, não imposto adicional: a origem fica com a'],
    ['interestadual e o destino fica com a diferença. Debitar a alíquota'],
    ['interna da origem E somar o DIFAL conta o mesmo imposto duas vezes.'],
    [''],
    ['Base dupla: a base do DIFAL é remontada "por dentro" com a alíquota'],
    ['interna do destino, o que eleva o DIFAL efetivo sobre o preço.']
  ]);
}

function montarTabelaICMS_formulas(sh) {
  var f = [];
  for (var i = LIN_1; i <= LIN_N; i++) {
    f.push([
      '=D' + i + '+IF(USA_FCP="Sim",E' + i + ',0)',
      '=IF($A' + i + '=UF_ORIGEM,$F' + i +
        ',IF(AND($C' + i + '="S/SE",GRUPO_ORIGEM="S/SE"),ALQ_INTER_12,ALQ_INTER_7))',
      '=IF($A' + i + '=UF_ORIGEM,$F' + i + ',ALQ_INTER_IMP)',
      '=IF(BASE_DUPLA="Sim",(1-$G' + i + ')/(1-$F' + i + ')*$F' + i + '-$G' + i +
        ',$F' + i + '-$G' + i + ')',
      '=IF(BASE_DUPLA="Sim",(1-$H' + i + ')/(1-$F' + i + ')*$F' + i + '-$H' + i +
        ',$F' + i + '-$H' + i + ')'
    ]);
  }
  sh.getRange(LIN_1, 6, N_UF, 5).setFormulas(FF(f));
}


// ===========================================================================
//  ABA: Faturamento_UF
// ===========================================================================
function montarFaturamento_estatico(sh) {
  sh.getRange('A1:H1').setValues([[
    'UF', 'Estado', 'Faturamento R$', '% do total',
    'Interestadual % NACIONAL', 'Interestadual % IMPORTADA',
    'DIFAL efetivo % NACIONAL', 'DIFAL efetivo % IMPORTADA'
  ]]);
  sh.getRange(LIN_1, 1, N_UF, 3).setValues(
    UFS.map(function (u) { return [u[0], u[1], u[5]]; }));
  sh.getRange(LIN_1, 3, N_UF, 1).setBackground(COR_INPUT);

  sh.getRange(LIN_TOT, 1).setValue('TOTAL');
  sh.getRange(LIN_TOT, 1, 1, 8).setFontWeight('bold').setBackground(COR_SECAO);

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
  sh.getRange(r + 1, 1, 6, 1).setValues([
    ['ICMS próprio médio — mercadoria NACIONAL'],
    ['DIFAL médio — mercadoria NACIONAL'],
    ['ICMS total médio — mercadoria NACIONAL'],
    ['ICMS próprio médio — mercadoria IMPORTADA'],
    ['DIFAL médio — mercadoria IMPORTADA'],
    ['ICMS total médio — mercadoria IMPORTADA']
  ]);
  sh.getRange(r + 1, 1, 6, 2).setBackground(COR_RESULT).setFontWeight('bold');
  sh.getRange(r + 1, 2, 6, 1).setNumberFormat(FMT_PCT);

  sh.getRange(b, 5).setValue('COMO LER');
  sh.getRange(b + 1, 5, 10, 1).setValues([
    ['Deixe em branco (ou zero) a UF onde você não vendeu — ou cujo faturamento'],
    ['você não conseguiu identificar. Ela é ignorada: não entra no numerador'],
    ['nem no denominador, e o valor é abatido do total. A média sai só sobre o'],
    ['faturamento efetivamente rateado por UF.'],
    [''],
    ['Os dois cenários são calculados sempre, lado a lado. Cada linha da aba'],
    ['Precificacao escolhe um deles conforme a coluna P (Origem).'],
    [''],
    ['Na mercadoria importada o DIFAL médio sobe muito (a origem fica com só'],
    ['4%) e o ICMS próprio despenca. O ICMS TOTAL quase não muda.']
  ]);
}

function montarFaturamento_formulas(sh) {
  var T = 'Tabela_ICMS!$A:$J';
  var f = [];
  for (var i = LIN_1; i <= LIN_N; i++) {
    f.push([
      '=IFERROR($C' + i + '/$C$' + LIN_TOT + ',0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',7,FALSE),0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',8,FALSE),0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',9,FALSE),0)',
      '=IFERROR(VLOOKUP($A' + i + ',' + T + ',10,FALSE),0)'
    ]);
  }
  sh.getRange(LIN_1, 4, N_UF, 5).setFormulas(FF(f));

  sh.getRange(LIN_TOT, 3).setFormula(F('=SUM(C' + LIN_1 + ':C' + LIN_N + ')'));
  sh.getRange(LIN_TOT, 4).setFormula(F('=SUM(D' + LIN_1 + ':D' + LIN_N + ')'));

  function pond(col) {
    return '=IFERROR(SUMPRODUCT($C$' + LIN_1 + ':$C$' + LIN_N + ',' +
           col + '$' + LIN_1 + ':' + col + '$' + LIN_N + ')/$C$' + LIN_TOT + ',0)';
  }

  var b = LIN_TOT + 2;
  sh.getRange(b + 1, 2).setFormula(F('=COUNTIF(C' + LIN_1 + ':C' + LIN_N + ',">0")'));
  sh.getRange(b + 2, 2).setFormula(F('=C' + LIN_TOT));

  var r = LIN_TOT + 7;
  sh.getRange(r,     2).setFormula(F(pond('$E')));
  sh.getRange(r + 1, 2).setFormula(F(pond('$G')));
  sh.getRange(r + 2, 2).setFormula(F('=B' + r + '+B' + (r + 1)));
  sh.getRange(r + 3, 2).setFormula(F(pond('$F')));
  sh.getRange(r + 4, 2).setFormula(F(pond('$H')));
  sh.getRange(r + 5, 2).setFormula(F('=B' + (r + 3) + '+B' + (r + 4)));
}


// ===========================================================================
//  ABA: Precificacao (horizontal, uma linha por SKU)
// ===========================================================================
function montarPrecificacao_estatico(sh) {
  var nc = COLS.length;

  sh.getRange(1, 1, 4, nc).setBackground(COR_TITULO).setFontColor('#ffffff')
    .setFontWeight('bold').setWrap(true).setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  COLS.forEach(function (c, i) {
    if (c[1]) sh.getRange(1, i + 1).setValue(c[1]);
    if (c[2]) sh.getRange(3, i + 1).setValue(c[2]);
    if (c[3]) sh.getRange(4, i + 1).setValue(c[3]);
  });

  // Mescla o rótulo de bloco (linha 1) sobre as colunas do bloco
  var ini = 0;
  for (var i = 1; i <= nc; i++) {
    if (i === nc || COLS[i][1]) {
      if (i - 1 > ini) sh.getRange(1, ini + 1, 1, i - ini).merge();
      ini = i;
    }
  }

  // Cores, formatos e validação vão SÓ na linha 5. O copyTo em
  // montarPrecificacao_formulas replica tudo para as demais linhas — é o que
  // mantém a execução leve o bastante para o serviço do Sheets aguentar.
  COLS.forEach(function (c, i) {
    var rg = sh.getRange(LIN_DADOS, i + 1);
    rg.setBackground(c[4] === 'in' ? COR_INPUT : (c[4] === 'res' ? COR_RESULT : COR_AUTO));
    if (c[5] === 'money') rg.setNumberFormat(FMT_MOEDA);
    else if (c[5] === 'pct') rg.setNumberFormat(FMT_PCT);
  });

  var listaUF = UFS.map(function (u) { return u[0]; });
  dropCel(sh, 'B',  CANAIS);
  dropCel(sh, 'P',  ['Nacional', 'Importada']);
  dropCel(sh, 'Q',  ['Sim', 'Não']);
  dropCel(sh, 'R',  listaUF);
  dropCel(sh, 'S',  listaUF);
  dropCel(sh, 'U',  ['Não', 'Sim - revendedor', 'Sim - fabricante/importador']);
  dropCel(sh, 'Z',  ['Só custo', 'Débito e Crédito']);
}

function montarPrecificacao_formulas(sh) {
  var n = N_LINHAS, L = LIN_DADOS;
  var TAB = 'Tabela_ICMS!$A:$J';
  // REVENDEDOR aponta para uma célula de Parametros com o texto exato da opção,
  // para as fórmulas não dependerem de digitação idêntica em 500 linhas.

  // Blocos e sua atribuição por coluna. Cada função recebe o número da linha.
  var defs = {
    // Total de custos do canal: comissão + taxa fixa + frete grátis + outros
    G:  function (r) { return '=IF($A' + r + '="","",$I' + r + '*$C' + r + '+$D' + r + '+$E' + r + '+$F' + r + ')'; },
    // Crédito de PIS/COFINS sobre as taxas do canal (zerado se monofásico)
    // O crédito aqui vem da NF de SERVIÇO que o canal emite contra você, não da
    // mercadoria. O regime monofásico do produto não alcança esse serviço, então
    // a coluna U não entra nesta conta.
    H:  function (r) { return '=IF($A' + r + '="","",$G' + r + '*(ALQ_PIS+ALQ_COFINS)*CRED_TAXAS)'; },

    // Alíquota de ICMS da COMPRA. Mesma UF -> interna; UF diferente ->
    // interestadual da operação fornecedor→você (4% se a mercadoria for
    // importada). Pode ser sobrescrita à mão: fornecedor do Simples Nacional
    // dá crédito menor, escrito nas observações da nota.
    T:  function (r) { return '=IF($A' + r + '="","",IF($S' + r + '=$R' + r +
          ',IFERROR(VLOOKUP($R' + r + ',' + TAB + ',6,FALSE),0)' +
          ',IF($P' + r + '="Importada",ALQ_INTER_IMP' +
          ',IF(AND(IFERROR(VLOOKUP($S' + r + ',' + TAB + ',3,FALSE),"")="S/SE"' +
             ',IFERROR(VLOOKUP($R' + r + ',' + TAB + ',3,FALSE),"")="S/SE")' +
             ',ALQ_INTER_12,ALQ_INTER_7))))'; },

    // --- Créditos da entrada. Todos zerados quando o CMV digitado já é
    //     líquido (coluna Q = "Sim"), porque já estão embutidos nele.
    AB: function (r) { return '=IF($A' + r + '="","",IF(OR($Q' + r + '="Sim",$X' + r + '>0),0,$AA' + r + '*$T' + r + '))'; },
    AC: function (r) { return '=IF($A' + r + '="","",IF(OR($Q' + r + '="Sim",LEFT($U' + r + ',3)="Sim"),0,' +
          '($AA' + r + '+$AF' + r + '-IF(EXCLUI_ICMS_CRED="Sim",$AB' + r + ',0))*$V' + r + '))'; },
    AD: function (r) { return '=IF($A' + r + '="","",IF(OR($Q' + r + '="Sim",LEFT($U' + r + ',3)="Sim"),0,' +
          '($AA' + r + '+$AF' + r + '-IF(EXCLUI_ICMS_CRED="Sim",$AB' + r + ',0))*$W' + r + '))'; },
    AE: function (r) { return '=IF($A' + r + '="","",IF($Z' + r + '="Débito e Crédito",$AF' + r + ',0))'; },

    // --- Custos da entrada que NÃO voltam como crédito
    AF: function (r) { return '=IF($A' + r + '="","",IF($Q' + r + '="Sim",0,$AA' + r + '*$Y' + r + '))'; },
    AG: function (r) { return '=IF($A' + r + '="","",IF($Q' + r + '="Sim",0,$AA' + r + '*$X' + r + '))'; },
    AH: function (r) { return '=IF($A' + r + '="","",$AA' + r + '+$AF' + r + '+$AG' + r +
          '-MIN($AB' + r + ',IF(CRED_ICMS_APROV="Sim",$AB' + r + ',$AI' + r + '))' +
          '-$AC' + r + '-$AD' + r + '-$AE' + r + ')'; },

    // --- Débitos da saída
    AI: function (r) { return '=IF($A' + r + '="","",$I' + r + '*IF($P' + r + '="Importada",ICMS_PROP_IMP,ICMS_PROP_NAC))'; },
    AJ: function (r) { return '=IF($A' + r + '="","",IF($U' + r + '=REVENDEDOR,0,' +
          '(($I' + r + '-$AL' + r + ')-IF(EXCLUI_ICMS_PC="Sim",$AI' + r + ',0))*$V' + r + '))'; },
    AK: function (r) { return '=IF($A' + r + '="","",IF($U' + r + '=REVENDEDOR,0,' +
          '(($I' + r + '-$AL' + r + ')-IF(EXCLUI_ICMS_PC="Sim",$AI' + r + ',0))*$W' + r + '))'; },
    // IPI da saída só existe quando você é equiparado a industrial
    AL: function (r) { return '=IF($A' + r + '="","",IF($Z' + r + '="Débito e Crédito",$I' + r + '*$Y' + r + '/(1+$Y' + r + '),0))'; },
    AM: function (r) { return '=IF($A' + r + '="","",IF($P' + r + '="Importada",DIFAL_MED_IMP,DIFAL_MED_NAC))'; },
    AN: function (r) { return '=IF($A' + r + '="","",$I' + r + '*$AM' + r + ')'; },

    // --- Resultado
    J:  function (r) { return '=IF($A' + r + '="","",$I' + r + '-$G' + r + '-$O' + r + '+$H' + r +
          '-$AH' + r + '-$AL' + r + '-$AI' + r + '-$AN' + r + '-$AJ' + r + '-$AK' + r + ')'; },
    K:  function (r) { return '=IF($A' + r + '="","",IFERROR($J' + r + '/($I' + r + '-$AL' + r + '),""))'; },

    // --- Simulador: inverte a equação da margem.
    //     MC = Preço x K - Fixos   =>   Preço = Fixos / (K - alvo/(1+IPI))
    //     K e Fixos ficam embutidos aqui para não gastar colunas.
    N:  function (r) {
          var ipi = 'IF($Z' + r + '="Débito e Crédito",$Y' + r + ',0)';
          var pcI = 'IF($U' + r + '=REVENDEDOR,0,$V' + r + '+$W' + r + ')';
          var crT = '((ALQ_PIS+ALQ_COFINS)*CRED_TAXAS)';
          var icm = 'IF($P' + r + '="Importada",ICMS_PROP_IMP,ICMS_PROP_NAC)';
          var dif = 'IF($P' + r + '="Importada",DIFAL_MED_IMP,DIFAL_MED_NAC)';
          var K = '(1-$C' + r + '-' + ipi + '/(1+' + ipi + ')-' + icm + '-' + dif +
                  '-' + pcI + '*(1/(1+' + ipi + ')-IF(EXCLUI_ICMS_PC="Sim",' + icm + ',0))' +
                  '+$C' + r + '*' + crT + ')';
          var FX = '($D' + r + '+$E' + r + '+$F' + r + '+$O' + r + '+$AH' + r +
                   '-($D' + r + '+$E' + r + '+$F' + r + ')*' + crT + ')';
          return '=IF($A' + r + '="","",IFERROR(' + FX + '/(' + K + '-$L' + r + '/(1+' + ipi + ')),""))';
        },
    M:  function (r) {
          var ipi = 'IF($Z' + r + '="Débito e Crédito",$Y' + r + ',0)';
          return '=IF($A' + r + '="","",IFERROR($L' + r + '*$N' + r + '/(1+' + ipi + '),""))';
        },

    // --- Diagnóstico
    AO: function (r) { return '=IF($A' + r + '="","",IFERROR(($AI' + r + '+$AN' + r + ')/$I' + r + ',""))'; },
    AP: function (r) { return '=IF($A' + r + '="","",IF(CRED_ICMS_APROV="Sim",0,MAX(0,$AB' + r + '-$AI' + r + ')))'; }
  };

  // Escreve as fórmulas UMA vez, na linha 5, e replica com copyTo.
  //
  // Gerar as 500 linhas uma a uma custava ~11.000 strings de fórmula numa só
  // execução, e o serviço do Sheets derrubava o script. O copyTo faz o mesmo
  // trabalho do lado do servidor e ainda leva junto cor, formato numérico e
  // validação de dados da linha modelo.
  //
  // As referências das fórmulas são do tipo $A5: coluna travada, linha solta.
  // É exatamente o que o copyTo precisa para ajustar 5 -> 6 -> 7 sozinho.
  var modelo = COLS.map(function (c) {
    return defs[c[0]] ? F(defs[c[0]](L)) : '';
  });
  sh.getRange(L, 1, 1, COLS.length).setFormulas([modelo]);
  SpreadsheetApp.flush();

  sh.getRange(L, 1, 1, COLS.length)
    .copyTo(sh.getRange(L + 1, 1, n - 1, COLS.length));

  // A linha de exemplo entra DEPOIS do copyTo, senão 'EXEMPLO-001' seria
  // replicado nas 499 linhas seguintes.
  // Só as colunas de entrada recebem valor; as de fórmula ficam intactas.
  [['A', 'EXEMPLO-001'], ['B', 'Mercado Livre'], ['C', 0.12], ['D', 0],
   ['E', 68.95], ['F', 0], ['I', 649.99], ['L', 0.15], ['O', 2.00],
   ['P', 'Nacional'], ['Q', 'Não'], ['R', 'SP'], ['S', 'SP'], ['U', 'Não'],
   ['V', 0.0165], ['W', 0.076], ['X', 0], ['Y', 0.15], ['Z', 'Só custo'],
   ['AA', 283.87]
  ].forEach(function (v) {
    sh.getRange(L, colNum(v[0])).setValue(v[1]);
  });

  // Alertas: margem negativa e crédito de ICMS empoçando
  var regras = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
      .setBackground(COR_ALERTA)
      .setRanges([sh.getRange(L, colNum('J'), n, 2)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setBackground(COR_ALERTA)
      .setRanges([sh.getRange(L, colNum('AP'), n, 1)]).build()
  ];
  sh.setConditionalFormatRules(regras);
}


// ===========================================================================
//  NAMED RANGES
// ===========================================================================
function criarNamedRanges(ss, par, fat) {
  var r = LIN_TOT + 7;
  [
    ['UF_ORIGEM',        par, 'B4'],
    ['GRUPO_ORIGEM',     par, 'B5'],
    ['ALQ_INTER_12',     par, 'B8'],
    ['ALQ_INTER_7',      par, 'B9'],
    ['ALQ_INTER_IMP',    par, 'B10'],
    ['ALQ_PIS',          par, 'B13'],
    ['ALQ_COFINS',       par, 'B14'],
    ['CRED_TAXAS',       par, 'B15'],
    ['EXCLUI_ICMS_PC',   par, 'B16'],
    ['EXCLUI_ICMS_CRED', par, 'B17'],
    ['BASE_DUPLA',       par, 'B18'],
    ['USA_FCP',          par, 'B19'],
    ['CRED_ICMS_APROV',  par, 'B20'],
    ['REVENDEDOR',       par, 'B21'],
    ['ICMS_PROP_NAC',    fat, 'B' + r],
    ['DIFAL_MED_NAC',    fat, 'B' + (r + 1)],
    ['ICMS_PROP_IMP',    fat, 'B' + (r + 3)],
    ['DIFAL_MED_IMP',    fat, 'B' + (r + 4)]
  ].forEach(function (n) {
    ss.setNamedRange(n[0], n[1].getRange(n[2]));
  });
}


// ===========================================================================
//  UTILITÁRIOS
// ===========================================================================
function colNum(letra) {
  var n = 0;
  for (var i = 0; i < letra.length; i++) n = n * 26 + (letra.charCodeAt(i) - 64);
  return n;
}

function dropdown(sh, a1, valores) {
  sh.getRange(a1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(valores, true).setAllowInvalid(false).build());
}

function dropCel(sh, letra, valores) {
  sh.getRange(LIN_DADOS, colNum(letra)).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(valores, true).setAllowInvalid(false).build());
}

function formatarTudo(par, tab, fat, pre) {
  [par, tab, fat].forEach(function (sh) {
    sh.getRange('A1').setFontSize(13).setFontWeight('bold').setFontColor(COR_TITULO);
    sh.setFrozenRows(1);
  });

  par.setColumnWidth(1, 400); par.setColumnWidth(2, 130); par.setColumnWidth(3, 25);
  par.setColumnWidth(4, 620);

  tab.getRange('A1:J1').setBackground(COR_TITULO).setFontColor('#ffffff')
     .setFontWeight('bold').setWrap(true);
  tab.setColumnWidth(1, 55); tab.setColumnWidth(2, 165);
  for (var i = 3; i <= 10; i++) tab.setColumnWidth(i, 105);
  tab.setColumnWidth(11, 25); tab.setColumnWidth(12, 560);
  tab.getRange(LIN_1, 4, N_UF, 7).setNumberFormat(FMT_PCT);
  tab.setFrozenColumns(2);

  fat.getRange('A1:H1').setBackground(COR_TITULO).setFontColor('#ffffff')
     .setFontWeight('bold').setWrap(true);
  fat.setColumnWidth(1, 55); fat.setColumnWidth(2, 165);
  for (var j = 3; j <= 8; j++) fat.setColumnWidth(j, 115);
  fat.getRange(LIN_1, 3, N_UF + 1, 1).setNumberFormat(FMT_MOEDA);
  fat.getRange(LIN_1, 4, N_UF + 1, 5).setNumberFormat(FMT_PCT);
  fat.setFrozenColumns(2);

  pre.setColumnWidth(1, 150);
  for (var k = 2; k <= COLS.length; k++) pre.setColumnWidth(k, 95);
  pre.setRowHeight(4, 46);
  pre.setFrozenRows(4);
  pre.setFrozenColumns(1);
}
