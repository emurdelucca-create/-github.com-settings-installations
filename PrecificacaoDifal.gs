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
 * ============================================================================
 */

var ABAS = ['Config', 'Tabela_ICMS', 'Faturamento_UF', 'Precificacao'];

var FMT_MOEDA = 'R$ #,##0.00';
var FMT_PCT   = '0.00%';
var FMT_PCT4  = '0.0000%';

var COR_TITULO  = '#1c3d5a';
var COR_SECAO   = '#dbe5f1';
var COR_INPUT   = '#fff2cc';
var COR_AUTO    = '#f0f0f0';
var COR_RESULT  = '#d9ead3';
var COR_ALERTA  = '#f4cccc';

/**
 * UF | Estado | Grupo (S/SE ou N/NE/CO) | Alíquota interna | FCP | Faturamento inicial
 *
 * "Grupo" define a alíquota interestadual: S/SE -> S/SE = 12%, qualquer outro
 * cruzamento = 7%. O ES conta como N/NE/CO para essa regra (é a regra do
 * Senado: Resolução 22/1989).
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

var N_UF   = UFS.length;      // 27
var LIN_1  = 2;               // primeira linha de dados
var LIN_N  = LIN_1 + N_UF - 1;// última linha de dados = 28
var LIN_TOT = LIN_N + 1;      // linha de total = 29


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

  // Remove SOMENTE as abas geradas por este script (as suas ficam intactas).
  ABAS.forEach(function (nome) {
    var s = ss.getSheetByName(nome);
    if (s) ss.deleteSheet(s);
  });

  // Ordem de criação importa: Tabela_ICMS é referenciada pelas outras.
  var shTab  = ss.insertSheet('Tabela_ICMS');
  var shCfg  = ss.insertSheet('Config');
  var shFat  = ss.insertSheet('Faturamento_UF');
  var shPre  = ss.insertSheet('Precificacao');

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

  ss.setActiveSheet(shPre);
  SpreadsheetApp.getUi().alert(
    'Planilha criada.\n\n' +
    'Amarelo = você preenche.  Cinza = calculado.  Verde = resultado.\n\n' +
    'Comece por Config, depois Faturamento_UF, depois Precificacao.'
  );
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
  sh.getRange('N2:N12').setValues([
    ['Alíquota interestadual (Res. Senado 22/1989 e 13/2012):'],
    ['  • S/SE → S/SE ................ 12%'],
    ['  • qualquer outro cruzamento .. 7%   (o ES entra como N/NE/CO)'],
    ['  • mercadoria importada ....... 4%   (CST de origem 1, 2, 3 ou 8)'],
    ['  • venda dentro da própria UF . alíquota interna, e DIFAL = 0'],
    [''],
    ['DIFAL = alíquota interna do destino − alíquota interestadual.'],
    ['É REPARTIÇÃO do mesmo ICMS, não imposto adicional: a origem fica'],
    ['com a interestadual e o destino fica com a diferença.'],
    [''],
    ['Base dupla (LC 190/22, art. 13, §6º): a base do DIFAL é remontada'],
  ]);
  sh.getRange('N13:N15').setValues([
    ['"por dentro" com a alíquota interna do destino, o que eleva o DIFAL'],
    ['efetivo (ex.: MG sai de 6,00% para 7,32% do preço). Ligue/desligue'],
    ['isso em Config.']
  ]);
}

function montarTabelaICMS_formulas(sh) {
  var f = [];
  for (var i = LIN_1; i <= LIN_N; i++) {
    f.push([
      // F - interna total (interna + FCP, se habilitado)
      '=D' + i + '+IF(USA_FCP="Sim",E' + i + ',0)',
      // G - interestadual da SAÍDA para mercadoria nacional
      '=IF($A' + i + '=UF_ORIGEM,$F' + i + ',IF(AND($C' + i + '="S/SE",GRUPO_ORIGEM="S/SE"),0.12,0.07))',
      // H - interestadual da SAÍDA para mercadoria importada
      '=IF($A' + i + '=UF_ORIGEM,$F' + i + ',ALQ_INTER_IMP)',
      // I - DIFAL efetivo sobre o preço, mercadoria nacional
      '=IF(BASE_DUPLA="Sim",(1-$G' + i + ')/(1-$F' + i + ')*$F' + i + '-$G' + i + ',$F' + i + '-$G' + i + ')',
      // J - DIFAL efetivo sobre o preço, mercadoria importada
      '=IF(BASE_DUPLA="Sim",(1-$H' + i + ')/(1-$F' + i + ')*$F' + i + '-$H' + i + ',$F' + i + '-$H' + i + ')',
      // K - interestadual do cenário ativo
      '=IF(ORIGEM_MERC="Importada",$H' + i + ',$G' + i + ')',
      // L - DIFAL efetivo do cenário ativo
      '=IF(ORIGEM_MERC="Importada",$J' + i + ',$I' + i + ')'
    ]);
  }
  sh.getRange(LIN_1, 6, N_UF, 7).setFormulas(f);
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
    [ 8, 'Alíquota interestadual — mercadoria importada', 0.04, 'in'],
    [ 9, 'Crédito de ICMS acumulado é aproveitável?', 'Sim', 'in'],
    [10, 'Considerar FCP no DIFAL?', 'Não', 'in'],

    [12, '2) REGIME E TESES (Lucro Real — não cumulativo)', '', 'sec'],
    [13, 'PIS %', 0.0165, 'in'],
    [14, 'COFINS %', 0.076, 'in'],
    [15, 'Excluir ICMS da base de PIS/COFINS? (Tema 69 / RE 574.706)', 'Sim', 'in'],
    [16, 'Excluir ICMS da base do crédito? (Lei 14.592/23)', 'Sim', 'in'],
    [17, 'DIFAL com base dupla? (LC 190/22, art. 13, §6º)', 'Sim', 'in'],

    [20, '3) AQUISIÇÃO — REVENDA (compra no Brasil)', '', 'sec'],
    [21, 'Custo unitário do produto (sem IPI)', 283.87, 'in'],
    [22, 'UF do fornecedor', 'SP', 'in'],
    [23, 'Alíquota de crédito de ICMS na entrada', '', 'auto'],
    [24, 'IPI % na compra (custo — não gera crédito na revenda)', 0.09, 'in'],
    [25, 'IPI R$ na compra', '', 'auto'],
    [26, 'Frete e despesas de compra (por unidade)', 0, 'in'],
    [27, 'Custo de aquisição bruto', '', 'auto'],
    [28, 'Crédito de ICMS', '', 'auto'],
    [29, 'Base do crédito de PIS/COFINS', '', 'auto'],
    [30, 'Crédito de PIS/COFINS', '', 'auto'],

    [32, '4) AQUISIÇÃO — IMPORTAÇÃO PRÓPRIA', '', 'sec'],
    [33, 'Valor aduaneiro / CIF (R$, por unidade)', 0, 'in'],
    [34, 'Imposto de Importação %', 0, 'in'],
    [35, 'IPI % (crédito na DI e débito na saída)', 0, 'in'],
    [36, 'PIS-Importação %', 0.021, 'in'],
    [37, 'COFINS-Importação %', 0.0965, 'in'],
    [38, 'Despesas aduaneiras (Siscomex, THC, armazenagem, desembaraço)', 0, 'in'],
    [39, 'Frete interno pós-desembaraço', 0, 'in'],
    [40, 'II R$', '', 'auto'],
    [41, 'IPI R$', '', 'auto'],
    [42, 'PIS-Importação R$', '', 'auto'],
    [43, 'COFINS-Importação R$', '', 'auto'],
    [44, 'Base de cálculo do ICMS-Importação (por dentro)', '', 'auto'],
    [45, 'ICMS-Importação R$', '', 'auto'],
    [46, 'Custo de aquisição bruto', '', 'auto'],
    [47, 'Crédito de ICMS', '', 'auto'],
    [48, 'Crédito de PIS/COFINS', '', 'auto'],
    [49, 'Crédito de IPI', '', 'auto'],

    [51, '5) AQUISIÇÃO — CENÁRIO ATIVO (alimenta a precificação)', '', 'sec'],
    [52, 'Custo de aquisição bruto', '', 'auto'],
    [53, 'Crédito de ICMS na entrada', '', 'auto'],
    [54, 'Crédito de PIS/COFINS na entrada', '', 'auto'],
    [55, 'Crédito de IPI na entrada', '', 'auto'],
    [56, 'Alíquota de IPI na saída', '', 'auto']
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

  // Formatos
  [8, 13, 14, 23, 24, 34, 35, 36, 37, 56].forEach(function (r) {
    sh.getRange(r, 2).setNumberFormat(FMT_PCT);
  });
  [21, 25, 26, 27, 28, 29, 30, 33, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49,
   52, 53, 54, 55].forEach(function (r) {
    sh.getRange(r, 2).setNumberFormat(FMT_MOEDA);
  });

  // Dropdowns
  var listaUF = UFS.map(function (u) { return u[0]; });
  dropdown(sh, 'B4',  listaUF);
  dropdown(sh, 'B22', listaUF);
  dropdown(sh, 'B6',  ['Nacional', 'Importada']);
  dropdown(sh, 'B7',  ['Revenda (compra no Brasil)', 'Importação própria']);
  ['B9', 'B10', 'B15', 'B16', 'B17'].forEach(function (c) {
    dropdown(sh, c, ['Sim', 'Não']);
  });

  // Notas
  sh.getRange('D3').setValue('OBSERVAÇÕES');
  sh.getRange('D4:D28').setValues([
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
    ['  esse saldo, marque "Não" em B9 — a planilha passa a limitar o crédito'],
    ['  aproveitado ao valor do débito próprio e mostra o que ficou preso.'],
    [''],
    ['• Tema 69 e Lei 14.592/23 puxam para lados opostos: o primeiro reduz o'],
    ['  débito de PIS/COFINS, o segundo reduz o crédito. Ambos são lei vigente.'],
    [''],
    ['• IRPJ/CSLL NÃO entram aqui. São apurados no fechamento, sobre o lucro'],
    ['  do período — depois das despesas fixas e das compensações. A margem'],
    ['  desta planilha é margem de contribuição, antes disso.']
  ]);
}

function montarConfig_formulas(sh) {
  var ALQ_INT_ORIG = 'VLOOKUP(UF_ORIGEM,Tabela_ICMS!$A$' + LIN_1 + ':$F$' + LIN_N + ',6,FALSE)';

  sh.getRange('B5').setFormula(
    '=IFERROR(VLOOKUP(UF_ORIGEM,Tabela_ICMS!$A$' + LIN_1 + ':$C$' + LIN_N + ',3,FALSE),"?")');

  // --- Revenda -------------------------------------------------------------
  // Crédito de entrada: mesma UF -> interna; UF diferente -> interestadual da
  // operação fornecedor→você (4% se a mercadoria for importada).
  sh.getRange('B23').setFormula(
    '=IF(R_UF_FORN=UF_ORIGEM,' + ALQ_INT_ORIG + ',' +
      'IF(ORIGEM_MERC="Importada",ALQ_INTER_IMP,' +
        'IF(AND(VLOOKUP(R_UF_FORN,Tabela_ICMS!$A$' + LIN_1 + ':$C$' + LIN_N + ',3,FALSE)="S/SE",' +
           'GRUPO_ORIGEM="S/SE"),0.12,0.07)))');

  sh.getRange('B25').setFormula('=R_CUSTO*R_IPI_PCT');
  sh.getRange('B27').setFormula('=R_CUSTO+R_IPI_RS+R_FRETE');
  // IPI não integra a base do ICMS na compra para revenda (CF 155, §2º, XI).
  sh.getRange('B28').setFormula('=(R_CUSTO+R_FRETE)*R_ALQ_ICMS_ENT');
  // IPI não recuperável integra o custo e entra na base do crédito (IN 2121/22,
  // art. 171, II). O ICMS sai da base desde a Lei 14.592/23.
  sh.getRange('B29').setFormula('=R_TOTAL-IF(EXCLUI_ICMS_CRED="Sim",R_CRED_ICMS,0)');
  sh.getRange('B30').setFormula('=R_BASE_CRED_PC*(ALQ_PIS+ALQ_COFINS)');

  // --- Importação própria --------------------------------------------------
  sh.getRange('B40').setFormula('=I_VA*I_II_PCT');
  sh.getRange('B41').setFormula('=(I_VA+I_II_RS)*I_IPI_PCT');
  sh.getRange('B42').setFormula('=I_VA*I_PIS_PCT');    // base = valor aduaneiro (RE 559.937)
  sh.getRange('B43').setFormula('=I_VA*I_COFINS_PCT');
  sh.getRange('B44').setFormula(
    '=IFERROR((I_VA+I_II_RS+I_IPI_RS+I_PIS_RS+I_COFINS_RS+I_DESP)/(1-' + ALQ_INT_ORIG + '),0)');
  sh.getRange('B45').setFormula('=I_BASE_ICMS*' + ALQ_INT_ORIG);
  sh.getRange('B46').setFormula(
    '=I_VA+I_II_RS+I_IPI_RS+I_PIS_RS+I_COFINS_RS+I_ICMS_RS+I_DESP+I_FRETE_INT');
  sh.getRange('B47').setFormula('=I_ICMS_RS');
  sh.getRange('B48').setFormula('=I_PIS_RS+I_COFINS_RS');
  sh.getRange('B49').setFormula('=I_IPI_RS');

  // --- Cenário ativo -------------------------------------------------------
  var IMP = 'MODALIDADE="Importação própria"';
  sh.getRange('B52').setFormula('=IF(' + IMP + ',I_CUSTO_TOTAL,R_TOTAL)');
  sh.getRange('B53').setFormula('=IF(' + IMP + ',I_CRED_ICMS,R_CRED_ICMS)');
  sh.getRange('B54').setFormula('=IF(' + IMP + ',I_CRED_PC,R_CRED_PC)');
  sh.getRange('B55').setFormula('=IF(' + IMP + ',I_CRED_IPI,0)');
  sh.getRange('B56').setFormula('=IF(' + IMP + ',I_IPI_PCT,0)');
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
  sh.getRange(LIN_1, 4, N_UF, 7).setFormulas(f);

  sh.getRange(LIN_TOT, 3).setFormula('=SUM(C' + LIN_1 + ':C' + LIN_N + ')');
  sh.getRange(LIN_TOT, 4).setFormula('=SUM(D' + LIN_1 + ':D' + LIN_N + ')');
  sh.getRange(LIN_TOT, 9).setFormula('=SUM(I' + LIN_1 + ':I' + LIN_N + ')');
  sh.getRange(LIN_TOT, 10).setFormula('=SUM(J' + LIN_1 + ':J' + LIN_N + ')');

  // Média ponderada = SUMPRODUCT(faturamento; alíquota) / faturamento total
  function pond(col) {
    return '=IFERROR(SUMPRODUCT($C$' + LIN_1 + ':$C$' + LIN_N + ',' +
           col + '$' + LIN_1 + ':' + col + '$' + LIN_N + ')/$C$' + LIN_TOT + ',0)';
  }

  // UFs em branco/zero saem da conta: não somam no total nem no SUMPRODUCT.
  var b = LIN_TOT + 2;
  sh.getRange(b + 1, 2).setFormula('=COUNTIF(C' + LIN_1 + ':C' + LIN_N + ',">0")');
  sh.getRange(b + 2, 2).setFormula('=C' + LIN_TOT);

  var r = LIN_TOT + 7;
  sh.getRange(r,     2).setFormula(pond('$E'));           // ICMS próprio nacional
  sh.getRange(r + 1, 2).setFormula(pond('$G'));           // DIFAL nacional
  sh.getRange(r + 2, 2).setFormula('=B' + r + '+B' + (r + 1));
  sh.getRange(r + 3, 2).setFormula(pond('$F'));           // ICMS próprio importado
  sh.getRange(r + 4, 2).setFormula(pond('$H'));           // DIFAL importado
  sh.getRange(r + 5, 2).setFormula('=B' + (r + 3) + '+B' + (r + 4));
  sh.getRange(r + 7, 2).setFormula('=IF(ORIGEM_MERC="Importada",B' + (r + 3) + ',B' + r + ')');
  sh.getRange(r + 8, 2).setFormula('=IF(ORIGEM_MERC="Importada",B' + (r + 4) + ',B' + (r + 1) + ')');
  sh.getRange(r + 9, 2).setFormula('=B' + (r + 7) + '+B' + (r + 8));
}


// ===========================================================================
//  ABA: Precificacao
// ===========================================================================
function montarPrecificacao_estatico(sh) {
  sh.getRange('A1').setValue('PRECIFICAÇÃO POR CANAL — LUCRO REAL');
  sh.getRange('A3:E3').setValues([['Item', 'Mercado Livre', 'Shopee', 'Canal 3', '']]);
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
  var cols = ['B', 'C', 'D'];
  cols.forEach(function (c) {
    var m = {};
    for (var r = 4; r <= 45; r++) m[r] = c + r;

    var PC = '(ALQ_PIS+ALQ_COFINS)';

    sh.getRange(c + '6').setFormula('=' + m[4] + '*' + m[5]);
    sh.getRange(c + '12').setFormula('=(' + m[6] + '+' + m[7] + '+' + m[8] + ')*' + m[11] + '*' + PC);

    sh.getRange(c + '15').setFormula('=ALQ_IPI_SAIDA');
    sh.getRange(c + '16').setFormula('=' + m[4] + '*' + m[15] + '/(1+' + m[15] + ')');
    sh.getRange(c + '17').setFormula('=' + m[4] + '-' + m[16]);
    sh.getRange(c + '18').setFormula('=ICMS_PROP_MED');
    sh.getRange(c + '19').setFormula('=' + m[4] + '*' + m[18]);
    sh.getRange(c + '20').setFormula('=DIFAL_MED');
    sh.getRange(c + '21').setFormula('=' + m[4] + '*' + m[20]);
    sh.getRange(c + '22').setFormula('=' + m[19] + '+' + m[21]);
    sh.getRange(c + '23').setFormula('=IFERROR(' + m[22] + '/' + m[4] + ',0)');
    sh.getRange(c + '24').setFormula('=' + m[17] + '-IF(EXCLUI_ICMS_PC="Sim",' + m[19] + ',0)');
    sh.getRange(c + '25').setFormula('=' + m[24] + '*ALQ_PIS');
    sh.getRange(c + '26').setFormula('=' + m[24] + '*ALQ_COFINS');

    sh.getRange(c + '29').setFormula('=CUSTO_BRUTO');
    sh.getRange(c + '30').setFormula('=CRED_ICMS_ENT');
    sh.getRange(c + '31').setFormula('=IF(CRED_ICMS_APROV="Sim",' + m[30] + ',MIN(' + m[30] + ',' + m[19] + '))');
    sh.getRange(c + '32').setFormula('=' + m[30] + '-' + m[31]);
    sh.getRange(c + '33').setFormula('=CRED_PC_ENT');
    sh.getRange(c + '34').setFormula('=CRED_IPI_ENT');
    sh.getRange(c + '35').setFormula('=' + m[29] + '-' + m[31] + '-' + m[33] + '-' + m[34]);

    sh.getRange(c + '38').setFormula(
      '=IF(N(' + m[4] + ')=0,"",' +
      m[4] + '-' + m[6] + '-' + m[7] + '-' + m[8] + '-' + m[9] + '-' + m[10] +
      '+' + m[12] + '-' + m[35] + '-' + m[16] + '-' + m[19] + '-' + m[21] +
      '-' + m[25] + '-' + m[26] + ')');
    sh.getRange(c + '39').setFormula('=IFERROR(' + m[38] + '/' + m[17] + ',"")');

    // K = 1 − comissão − IPI/(1+IPI) − ICMS próprio − DIFAL
    //       − (PIS+COFINS)*(1/(1+IPI) − Tema69*ICMS próprio)
    //       + comissão*%crédito*(PIS+COFINS)
    sh.getRange(c + '43').setFormula(
      '=1-' + m[5] + '-' + m[15] + '/(1+' + m[15] + ')-' + m[18] + '-' + m[20] +
      '-' + PC + '*(1/(1+' + m[15] + ')-IF(EXCLUI_ICMS_PC="Sim",' + m[18] + ',0))' +
      '+' + m[5] + '*' + m[11] + '*' + PC);
    sh.getRange(c + '44').setFormula(
      '=' + m[7] + '+' + m[8] + '+' + m[9] + '+' + m[10] + '+' + m[35] +
      '-(' + m[7] + '+' + m[8] + ')*' + m[11] + '*' + PC);
    sh.getRange(c + '45').setFormula(
      '=IFERROR(' + m[44] + '/(' + m[43] + '-' + m[42] + '/(1+' + m[15] + ')),"")');
  });

  // Alertas visuais
  var reg = sh.getRange('B38:D39');
  var regra = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setBackground(COR_ALERTA).setRanges([reg]).build();
  var reg2 = sh.getRange('B32:D32');
  var regra2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground(COR_ALERTA).setRanges([reg2]).build();
  sh.setConditionalFormatRules([regra, regra2]);
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
    ['ALQ_INTER_IMP',    cfg, 'B8'],
    ['CRED_ICMS_APROV',  cfg, 'B9'],
    ['USA_FCP',          cfg, 'B10'],

    ['ALQ_PIS',          cfg, 'B13'],
    ['ALQ_COFINS',       cfg, 'B14'],
    ['EXCLUI_ICMS_PC',   cfg, 'B15'],
    ['EXCLUI_ICMS_CRED', cfg, 'B16'],
    ['BASE_DUPLA',       cfg, 'B17'],

    ['R_CUSTO',          cfg, 'B21'],
    ['R_UF_FORN',        cfg, 'B22'],
    ['R_ALQ_ICMS_ENT',   cfg, 'B23'],
    ['R_IPI_PCT',        cfg, 'B24'],
    ['R_IPI_RS',         cfg, 'B25'],
    ['R_FRETE',          cfg, 'B26'],
    ['R_TOTAL',          cfg, 'B27'],
    ['R_CRED_ICMS',      cfg, 'B28'],
    ['R_BASE_CRED_PC',   cfg, 'B29'],
    ['R_CRED_PC',        cfg, 'B30'],

    ['I_VA',             cfg, 'B33'],
    ['I_II_PCT',         cfg, 'B34'],
    ['I_IPI_PCT',        cfg, 'B35'],
    ['I_PIS_PCT',        cfg, 'B36'],
    ['I_COFINS_PCT',     cfg, 'B37'],
    ['I_DESP',           cfg, 'B38'],
    ['I_FRETE_INT',      cfg, 'B39'],
    ['I_II_RS',          cfg, 'B40'],
    ['I_IPI_RS',         cfg, 'B41'],
    ['I_PIS_RS',         cfg, 'B42'],
    ['I_COFINS_RS',      cfg, 'B43'],
    ['I_BASE_ICMS',      cfg, 'B44'],
    ['I_ICMS_RS',        cfg, 'B45'],
    ['I_CUSTO_TOTAL',    cfg, 'B46'],
    ['I_CRED_ICMS',      cfg, 'B47'],
    ['I_CRED_PC',        cfg, 'B48'],
    ['I_CRED_IPI',       cfg, 'B49'],

    ['CUSTO_BRUTO',      cfg, 'B52'],
    ['CRED_ICMS_ENT',    cfg, 'B53'],
    ['CRED_PC_ENT',      cfg, 'B54'],
    ['CRED_IPI_ENT',     cfg, 'B55'],
    ['ALQ_IPI_SAIDA',    cfg, 'B56'],

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
    sh.setFrozenRows(1);
  });

  cfg.setColumnWidth(1, 400); cfg.setColumnWidth(2, 150); cfg.setColumnWidth(3, 25);
  cfg.setColumnWidth(4, 620);

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
  fat.getRange(LIN_TOT + 2, 1, 1, 1).setFontWeight('bold');
  fat.getRange(LIN_TOT + 6, 1, 1, 1).setFontWeight('bold');
  fat.setFrozenRows(1); fat.setFrozenColumns(2);

  pre.setColumnWidth(1, 340);
  pre.setColumnWidth(2, 130); pre.setColumnWidth(3, 130); pre.setColumnWidth(4, 130);
  pre.setColumnWidth(5, 25);  pre.setColumnWidth(6, 620);
  pre.setFrozenRows(3); pre.setFrozenColumns(1);
}
