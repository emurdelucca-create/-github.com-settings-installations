// ============================================================
// COMPARAÇÃO VENDAS 2026 — GE Finance + BaseLinker
// Análise de faturamento por canal, sem desmembramento
// ============================================================

const CV = {
  ID_GE:  '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo',
  ABA_GE: 'Dados Completos',
  ID_BL:  '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik',
  ABA_BL: 'estoque',

  HEADER_ROWS: 3,
  NCOLS:       30, // A(1) a AD(30)

  GE_SKU:    1,
  GE_QTD:    2,
  GE_DATA:   3,
  GE_CANAL:  7,
  GE_FAT:    21,
  GE_MARGEM: 45,

  BL_SKU:       1,
  BL_SALDO_PAD: 4,
  BL_SALDO_ARM: 5,
  BL_SALDO_CHG: 6,
  BL_TIPO:      7,
};

const CV_ML     = ['ML Edmotos', 'ML Humble', 'ML Najumi', 'ML Sky', 'ML Moto Vibe'];
const CV_SHOPEE = ['Shopee Humble', 'Shopee Najumi', 'Shopee Sky'];

const CV_PER_SEM = [
  {de:0,ate:7},{de:7,ate:15},{de:15,ate:30},
  {de:30,ate:45},{de:45,ate:60},{de:60,ate:75},{de:75,ate:90},
];
const CV_PER_CON = [
  {nome:'0-30',  de:0,  ate:30},
  {nome:'30-60', de:30, ate:60},
  {nome:'60-90', de:60, ate:90},
];

// ============================================================
// HELPER
// ============================================================
function _cvGetAba(ss, nome) {
  let aba = ss.getSheetByName(nome);
  if (aba) return aba;
  const norm = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const t = norm(nome);
  for (const a of ss.getSheets()) if (norm(a.getName()) === t) return a;
  return null;
}

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Comparação Vendas')
    .addItem('ML Humble',              'cvGerarML_Humble')
    .addItem('ML Najumi',              'cvGerarML_Najumi')
    .addItem('ML Sky',                 'cvGerarML_Sky')
    .addItem('ML Edmotos',             'cvGerarML_Edmotos')
    .addItem('ML Moto Vibe',           'cvGerarML_MotoVibe')
    .addSeparator()
    .addItem('Shopee Humble',          'cvGerarShopee_Humble')
    .addItem('Shopee Najumi',          'cvGerarShopee_Najumi')
    .addItem('Shopee Sky',             'cvGerarShopee_Sky')
    .addSeparator()
    .addItem('Todos ML',               'cvGerarTodosML')
    .addItem('Todos Shopee',           'cvGerarTodosShopee')
    .addSeparator()
    .addItem('🚀 Gerar Todos os Canais', 'cvGerarTodos')
    .addSeparator()
    .addItem('🔍 Diagnóstico',         'cvDiagnostico')
    .addToUi();
}

// ============================================================
// FUNÇÕES POR CANAL (chamadas pelo menu)
// ============================================================
function cvGerarML_Humble()     { _cvGerar('ML Humble',     ['ML Humble']);     }
function cvGerarML_Najumi()     { _cvGerar('ML Najumi',     ['ML Najumi']);     }
function cvGerarML_Sky()        { _cvGerar('ML Sky',        ['ML Sky']);        }
function cvGerarML_Edmotos()    { _cvGerar('ML Edmotos',    ['ML Edmotos']);    }
function cvGerarML_MotoVibe()   { _cvGerar('ML Moto Vibe',  ['ML Moto Vibe']);  }
function cvGerarShopee_Humble() { _cvGerar('Shopee Humble', ['Shopee Humble']); }
function cvGerarShopee_Najumi() { _cvGerar('Shopee Najumi', ['Shopee Najumi']); }
function cvGerarShopee_Sky()    { _cvGerar('Shopee Sky',    ['Shopee Sky']);    }
function cvGerarTodosML()       { _cvGerar('Todos ML',      CV_ML);             }
function cvGerarTodosShopee()   { _cvGerar('Todos Shopee',  CV_SHOPEE);         }

function cvGerarTodos() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Gerando todos os canais... Aguarde alguns minutos.');
  try {
    const ge = _cvCarregarGE();
    const bl = _cvCarregarBL();
    const canais = [
      ['ML Humble',     ['ML Humble']],
      ['ML Najumi',     ['ML Najumi']],
      ['ML Sky',        ['ML Sky']],
      ['ML Edmotos',    ['ML Edmotos']],
      ['ML Moto Vibe',  ['ML Moto Vibe']],
      ['Shopee Humble', ['Shopee Humble']],
      ['Shopee Najumi', ['Shopee Najumi']],
      ['Shopee Sky',    ['Shopee Sky']],
      ['Todos ML',      CV_ML],
      ['Todos Shopee',  CV_SHOPEE],
    ];
    canais.forEach(([nome, lista]) => _cvGerar(nome, lista, ge, bl, true));
    ui.alert('✅ Todos os 10 canais gerados com sucesso!');
  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ============================================================
// DIAGNÓSTICO
// ============================================================
function cvDiagnostico() {
  const ui = SpreadsheetApp.getUi();
  let msg  = '';
  [[CV.ID_GE, CV.ABA_GE, 'GE Finance'], [CV.ID_BL, CV.ABA_BL, 'BaseLinker']].forEach(([id, aba, label]) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const a  = _cvGetAba(ss, aba);
      const ok = a ? '✅ ' + a.getLastRow() + ' linhas' : '❌ aba não encontrada';
      msg += label + ' → "' + aba + '": ' + ok + '\n';
      msg += '  Abas: ' + ss.getSheets().map(s => '"' + s.getName() + '"').join(', ') + '\n\n';
    } catch(e) {
      msg += label + ' → ERRO: ' + e.message + '\n\n';
    }
  });
  ui.alert('🔍 Diagnóstico', msg, ui.ButtonSet.OK);
}

// ============================================================
// CARREGAR GE FINANCE
// Retorna: { dados[], dataMax }
// ============================================================
function _cvCarregarGE() {
  const ss  = SpreadsheetApp.openById(CV.ID_GE);
  const aba = _cvGetAba(ss, CV.ABA_GE);
  if (!aba) throw new Error('Aba GE Finance "' + CV.ABA_GE + '" não encontrada.');
  const dados = aba.getDataRange().getValues();

  let dataMax = new Date(0);
  for (let i = 1; i < dados.length; i++) {
    const d = new Date(dados[i][CV.GE_DATA]);
    if (!isNaN(d) && d > dataMax) dataMax = d;
  }
  dataMax.setHours(23, 59, 59, 999);

  return { dados, dataMax };
}

// ============================================================
// CARREGAR BASELINKER ESTOQUE
// Retorna: { mapaTipo{sku→'Simples'|'Composto'}, mapaEstoque{sku→saldo} }
// ============================================================
function _cvCarregarBL() {
  const ss  = SpreadsheetApp.openById(CV.ID_BL);
  const aba = _cvGetAba(ss, CV.ABA_BL);
  if (!aba) throw new Error('Aba BaseLinker "' + CV.ABA_BL + '" não encontrada.');
  const dados = aba.getDataRange().getValues();

  const mapaTipo    = {};
  const mapaEstoque = {};

  for (let i = 1; i < dados.length; i++) {
    const row  = dados[i];
    const sku  = String(row[CV.BL_SKU]  || '').trim();
    const tipo = String(row[CV.BL_TIPO] || '').trim();
    if (!sku) continue;

    const saldo = Number(row[CV.BL_SALDO_PAD] || 0)
                + Number(row[CV.BL_SALDO_ARM] || 0)
                + Number(row[CV.BL_SALDO_CHG] || 0);

    if (tipo === 'Simples') {
      mapaTipo[sku]    = 'Simples';
      mapaEstoque[sku] = saldo;
    } else if (tipo === 'PAI') {
      // PAI = linha pai de produto composto → estoque total disponível
      mapaTipo[sku]    = 'Composto';
      mapaEstoque[sku] = saldo;
    }
  }

  return { mapaTipo, mapaEstoque };
}

// ============================================================
// GERAR AS DUAS ABAS DE UM CANAL
// ge / bl opcionais: passados por cvGerarTodos para evitar reload
// silencioso: true → sem alert individual (usado em cvGerarTodos)
// ============================================================
function _cvGerar(nomeCanal, listaCanais, ge, bl, silencioso) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    if (!ge) ge = _cvCarregarGE();
    if (!bl) bl = _cvCarregarBL();

    const { dados, dataMax }        = ge;
    const { mapaTipo, mapaEstoque } = bl;

    const canaisSet = new Set(listaCanais.map(c => c.toLowerCase()));

    // ── Agregar vendas por SKU ─────────────────────────────────
    const skuMap = {};

    for (let i = 1; i < dados.length; i++) {
      const row   = dados[i];
      const canal = String(row[CV.GE_CANAL] || '').trim();
      if (!canaisSet.has(canal.toLowerCase())) continue;

      const d = new Date(row[CV.GE_DATA]);
      if (isNaN(d)) continue;
      const dias = Math.floor((dataMax - d) / 86400000);
      if (dias < 0 || dias >= 90) continue;

      const sku    = String(row[CV.GE_SKU]    || '').trim();
      const qtd    = Number(row[CV.GE_QTD]    || 0);
      const fat    = Number(row[CV.GE_FAT]    || 0);
      const margem = Number(row[CV.GE_MARGEM] || 0);
      if (!sku) continue;

      if (!skuMap[sku]) {
        skuMap[sku] = {
          qtdSem: [0, 0, 0, 0, 0, 0, 0],
          cons: {
            '0-30':  { qtd: 0, fat: 0, margem: 0 },
            '30-60': { qtd: 0, fat: 0, margem: 0 },
            '60-90': { qtd: 0, fat: 0, margem: 0 },
          },
        };
      }
      const d_ = skuMap[sku];

      CV_PER_SEM.forEach((p, idx) => {
        if (dias >= p.de && dias < p.ate) d_.qtdSem[idx] += qtd;
      });
      CV_PER_CON.forEach(p => {
        if (dias < p.de || dias >= p.ate) return;
        d_.cons[p.nome].qtd    += qtd;
        d_.cons[p.nome].fat    += fat;
        d_.cons[p.nome].margem += margem;
      });
    }

    if (Object.keys(skuMap).length === 0) {
      if (!silencioso) ui.alert('ℹ️ Nenhuma venda em "' + nomeCanal + '" nos últimos 90 dias.');
      return;
    }

    // ── Montar itens ──────────────────────────────────────────
    const itens = Object.entries(skuMap).map(([sku, d]) => {
      const c030  = d.cons['0-30'];
      const c3060 = d.cons['30-60'];
      const c6090 = d.cons['60-90'];
      const fat90 = c030.fat + c3060.fat + c6090.fat;

      const mg030  = c030.fat  > 0 ? c030.margem  / c030.fat  : 0;
      const mg3060 = c3060.fat > 0 ? c3060.margem / c3060.fat : 0;
      const mg6090 = c6090.fat > 0 ? c6090.margem / c6090.fat : 0;

      const tm030  = c030.qtd  > 0 ? c030.fat  / c030.qtd  : 0;
      const tm3060 = c3060.qtd > 0 ? c3060.fat / c3060.qtd : 0;
      const tm6090 = c6090.qtd > 0 ? c6090.fat / c6090.qtd : 0;

      const difZ  = c030.fat - c3060.fat;
      const difAA = c3060.fat !== 0 ? difZ  / Math.abs(c3060.fat) : (c030.fat > 0 ? 1 : 0);
      const difAB = c030.fat - c6090.fat;
      const difAC = c6090.fat !== 0 ? difAB / Math.abs(c6090.fat) : (c030.fat > 0 ? 1 : 0);

      return {
        sku,
        tipo:    mapaTipo[sku]                               || '',
        estoque: mapaEstoque[sku] !== undefined ? mapaEstoque[sku] : 0,
        qtdSem:  d.qtdSem,
        c030, c3060, c6090,
        mg030, mg3060, mg6090,
        tm030, tm3060, tm6090,
        difZ, difAA, difAB, difAC,
        fat90,
      };
    });

    // Aba "+": maior fat 0-90 primeiro
    const itensPlus  = [...itens].sort((a, b) => b.fat90 - a.fat90);
    // Aba "−": maior queda fat 0-30 vs 60-90 primeiro (mais negativo = mais caiu)
    const itensMinus = [...itens].sort((a, b) => a.difAB - b.difAB);

    _cvEscreverAba(ss, nomeCanal + ' +', itensPlus);
    _cvEscreverAba(ss, nomeCanal + ' -', itensMinus);

    if (!silencioso) {
      ui.alert('✅ ' + nomeCanal + ': ' + itens.length + ' SKUs | abas "' +
               nomeCanal + ' +" e "' + nomeCanal + ' -" geradas.');
    }

  } catch(e) {
    if (!silencioso) ui.alert('❌ Erro [' + nomeCanal + ']: ' + e.message);
    else throw e;
  }
}

// ============================================================
// ESCREVER ABA
// Colunas A-AD (30 colunas), 3 linhas de cabeçalho
//
// A(1):  SKU
// B(2):  Tipo
// C(3):  Estoque
// D-J(4-10): Qtd semanal 0-7, 7-15, 15-30, 30-45, 45-60, 60-75, 75-90
// K-O(11-15): 0-30  → Qtd.Vend., Fat., Margem R$, Margem %, Ticket Médio
// P-T(16-20): 30-60 → idem
// U-Y(21-25): 60-90 → idem
// Z(26):  Fat 0-30 − Fat 30-60
// AA(27): % 0-30 / 30-60
// AB(28): Fat 0-30 − Fat 60-90
// AC(29): % 0-30 / 60-90
// AD(30): Fat R$ 0-90
// ============================================================
function _cvEscreverAba(ss, nomeAba, itens) {
  let aba = _cvGetAba(ss, nomeAba);
  if (aba) {
    aba.getRange(1, 1, aba.getMaxRows(), aba.getMaxColumns()).breakApart();
    aba.clearContents();
    aba.clearFormats();
  } else {
    aba = ss.insertSheet(nomeAba);
  }

  // ── Linha 1: seções principais ─────────────────────────────
  const row1 = new Array(CV.NCOLS).fill('');
  row1[3]  = 'Qtd. Vend.';  // D (0-idx 3)  → merge D-J
  row1[10] = 'Vendas';       // K (0-idx 10) → merge K-Y
  row1[25] = 'Diferenças';   // Z (0-idx 25) → merge Z-AC
  row1[29] = 'Fat. R$';      // AD (0-idx 29)

  // ── Linha 2: sub-seções de período ─────────────────────────
  const row2 = new Array(CV.NCOLS).fill('');
  row2[10] = '0-30';   // K → merge K-O
  row2[15] = '30-60';  // P → merge P-T
  row2[20] = '60-90';  // U → merge U-Y

  // ── Linha 3: nomes das colunas ─────────────────────────────
  const SUB5 = ['Qtd. Vend.', 'Fat.', 'Margem (R$)', 'Margem (%)', 'Ticket Médio'];
  const row3 = [
    'SKU', 'Tipo (Simples ou Composto)', 'Estoque',
    '0-7', '7-15', '15-30', '30-45', '45-60', '60-75', '75-90',
    ...SUB5, ...SUB5, ...SUB5,
    '0-30 − 30-60 (R$)', '% 0-30 / 30-60',
    '0-30 − 60-90 (R$)', '% 0-30 / 60-90',
    'Fat. R$ 0-90',
  ];

  aba.getRange(1, 1, 1, CV.NCOLS).setValues([row1]);
  aba.getRange(2, 1, 1, CV.NCOLS).setValues([row2]);
  aba.getRange(3, 1, 1, CV.NCOLS).setValues([row3]);

  // Merges linha 1
  aba.getRange(1, 4,  1, 7).merge();   // D-J  (7 cols): Qtd. Vend.
  aba.getRange(1, 11, 1, 15).merge();  // K-Y  (15 cols): Vendas
  aba.getRange(1, 26, 1, 4).merge();   // Z-AC (4 cols):  Diferenças

  // Merges linha 2
  aba.getRange(2, 11, 1, 5).merge();  // K-O: 0-30
  aba.getRange(2, 16, 1, 5).merge();  // P-T: 30-60
  aba.getRange(2, 21, 1, 5).merge();  // U-Y: 60-90

  if (itens.length === 0) return;

  // ── Dados ──────────────────────────────────────────────────
  const PRIMA = CV.HEADER_ROWS + 1; // linha 4

  const rows = itens.map(item => [
    item.sku,           // A  (1)
    item.tipo,          // B  (2)
    item.estoque,       // C  (3)
    item.qtdSem[0],     // D  (4)  0-7
    item.qtdSem[1],     // E  (5)  7-15
    item.qtdSem[2],     // F  (6)  15-30
    item.qtdSem[3],     // G  (7)  30-45
    item.qtdSem[4],     // H  (8)  45-60
    item.qtdSem[5],     // I  (9)  60-75
    item.qtdSem[6],     // J  (10) 75-90
    item.c030.qtd,      // K  (11)
    item.c030.fat,      // L  (12)
    item.c030.margem,   // M  (13)
    item.mg030,         // N  (14) Margem %
    item.tm030,         // O  (15) Ticket Médio
    item.c3060.qtd,     // P  (16)
    item.c3060.fat,     // Q  (17)
    item.c3060.margem,  // R  (18)
    item.mg3060,        // S  (19) Margem %
    item.tm3060,        // T  (20) Ticket Médio
    item.c6090.qtd,     // U  (21)
    item.c6090.fat,     // V  (22)
    item.c6090.margem,  // W  (23)
    item.mg6090,        // X  (24) Margem %
    item.tm6090,        // Y  (25) Ticket Médio
    item.difZ,          // Z  (26) Fat 0-30 − Fat 30-60
    item.difAA,         // AA (27) %
    item.difAB,         // AB (28) Fat 0-30 − Fat 60-90
    item.difAC,         // AC (29) %
    item.fat90,         // AD (30)
  ]);

  aba.getRange(PRIMA, 1, rows.length, CV.NCOLS).setValues(rows);

  // SKU como texto
  aba.getRange(PRIMA, 1, rows.length, 1).setNumberFormat('@');

  // Margem % e variação %
  [14, 19, 24, 27, 29].forEach(col => {
    aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('0.0%');
  });

  // Faturamento e Margem R$
  [12, 13, 17, 18, 22, 23, 26, 28, 30].forEach(col => {
    aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('R$ #,##0.00');
  });

  // Ticket Médio
  [15, 20, 25].forEach(col => {
    aba.getRange(PRIMA, col, rows.length, 1).setNumberFormat('R$ #,##0.00');
  });

  // Colorir colunas de diferença
  _cvColorirDif(aba, rows, PRIMA, 26); // Z:  0-30 vs 30-60
  _cvColorirDif(aba, rows, PRIMA, 28); // AB: 0-30 vs 60-90

  SpreadsheetApp.flush();
}

// ============================================================
// COLORIR COLUNA DE DIFERENÇA
// positivo (cresceu) → verde | negativo (caiu) → vermelho | zero → amarelo
// ============================================================
function _cvColorirDif(aba, rows, prima, col) {
  const verde = [], vermelho = [], amarelo = [];
  rows.forEach((row, i) => {
    const v = row[col - 1]; // converte col 1-based para índice 0-based
    if      (v > 0) verde.push(aba.getRange(prima + i, col).getA1Notation());
    else if (v < 0) vermelho.push(aba.getRange(prima + i, col).getA1Notation());
    else            amarelo.push(aba.getRange(prima + i, col).getA1Notation());
  });
  if (verde.length)    aba.getRangeList(verde).setBackground('#c8e6c9');
  if (vermelho.length) aba.getRangeList(vermelho).setBackground('#ef9a9a');
  if (amarelo.length)  aba.getRangeList(amarelo).setBackground('#fff9c4');
}
