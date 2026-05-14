// ============================================================
// FULL ML — v2 (inclui aba Reposição Full)
// ============================================================

const ABA_DADOS      = 'Full ML';
const ABA_UPLOAD     = 'Uploads';
const ABA_REPOSICAO  = 'Reposição Full';

// # Anúncio adicionado ao final (índice 12) para não quebrar
// código existente que lê colunas por índice 0-11
const CABECALHOS = [
  'Conta', 'Inventory ID', 'SKU', 'Produto', 'Variação',
  'Vendas Brutas 30 dias (R$)', 'Unidades Vendidas 30 dias',
  'Estoque Apto (CD)', 'Estoque A Caminho',
  'Nível de Estoque', 'Observações', 'Última Atualização',
  '# Anúncio',
];

const CONTAS = [
  'ML Humble', 'ML Edmotos', 'ML Sky', 'ML Moto Vibe', 'ML Najumi'
];

// IDs das fontes externas
const RP_GE_ID  = '1OedjVQcNUoqmoPzeRs9TKsoC4LiDtemYs3cZ0--OTUo';
const RP_GE_ABA = 'Dados Completos';
const RP_BL_ID  = '1wy-tJoDxGDjfnd0AXQfdQ0bw9qmTtxz7wV4UKDlQrik';
const RP_BL_ABA = 'estoque';

// Pesos ponderados para taxa diária de vendas (soma = 1.00)
const RP_PESOS = [
  { de: 0,  ate: 7,  dias: 7,  peso: 0.25 },
  { de: 7,  ate: 15, dias: 8,  peso: 0.20 },
  { de: 15, ate: 30, dias: 15, peso: 0.18 },
  { de: 30, ate: 45, dias: 15, peso: 0.15 },
  { de: 45, ate: 60, dias: 15, peso: 0.10 },
  { de: 60, ate: 75, dias: 15, peso: 0.07 },
  { de: 75, ate: 90, dias: 15, peso: 0.05 },
];

const RP_PER_CON = [
  { nome: '0-30',  de: 0,  ate: 30 },
  { nome: '30-60', de: 30, ate: 60 },
  { nome: '60-90', de: 60, ate: 90 },
];

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Full ML')
    .addItem('📤 Upload Relatório ML',              'abrirUploadRelatorio')
    .addItem('🗑️ Limpar Dados de uma Conta',        'abrirLimparConta')
    .addItem('📋 Ver Resumo por Conta',              'verResumo')
    .addItem('📈 Gerar Curva ABC',                  'gerarCurvaABC')
    .addItem('📈 Gerar Curva ABC 2 (por anúncio)',  'gerarCurvaABC2')
    .addSeparator()
    .addItem('🔄 Gerar Reposição Full',             'gerarReposicaoFull')
    .addItem('⚙️ Configurar Dias por Curva',        'abrirConfigDias')
    .addToUi();
}

// ============================================================
// UPLOAD DE RELATÓRIO
// ============================================================
function abrirUploadRelatorio() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { box-sizing: border-box; }
        body { font-family: Arial, sans-serif; padding: 20px; margin: 0; background: #f8f9fa; }
        .card { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        h3 { color: #1a73e8; margin: 0 0 16px 0; }
        label { display: block; font-size: 13px; color: #5f6368; margin: 12px 0 4px 0; font-weight: 500; }
        select, input[type="file"] { width: 100%; padding: 8px 12px; border: 1px solid #dadce0; border-radius: 4px; font-size: 14px; background: white; }
        .btn { background: #1a73e8; color: white; border: none; padding: 10px 24px; border-radius: 4px; cursor: pointer; width: 100%; margin-top: 16px; font-size: 14px; font-weight: 500; }
        .btn:hover { background: #1557b0; }
        .btn:disabled { background: #9aa0a6; cursor: not-allowed; }
        #status { margin-top: 12px; padding: 12px; border-radius: 4px; display: none; font-size: 13px; }
        .sucesso  { background: #e6f4ea; color: #137333; border: 1px solid #ceead6; }
        .erro     { background: #fce8e6; color: #c5221f; border: 1px solid #f5c6c6; }
        .carregando { background: #e8f0fe; color: #1967d2; border: 1px solid #c5d8f6; }
        .info { background: #fef7e0; color: #b45309; border: 1px solid #fde293; font-size: 12px; margin-top: 8px; padding: 8px; border-radius: 4px; }
      </style>
    </head>
    <body>
      <div class="card">
        <h3>📤 Upload Relatório ML</h3>
        <div class="info">ℹ️ Baixe em: <b>Seller Center → Full → Planejamento de envios → Baixar relatório</b></div>
        <label>Conta do Mercado Livre:</label>
        <select id="conta">
          <option value="">-- Selecione a conta --</option>
          <option value="ML Humble">ML Humble</option>
          <option value="ML Edmotos">ML Edmotos</option>
          <option value="ML Sky">ML Sky</option>
          <option value="ML Moto Vibe">ML Moto Vibe</option>
          <option value="ML Najumi">ML Najumi</option>
        </select>
        <label>Arquivo Excel (.xlsx):</label>
        <input type="file" id="arquivo" accept=".xlsx,.xls" />
        <button class="btn" id="btnEnviar" onclick="enviarArquivo()">📊 Importar Relatório</button>
        <div id="status"></div>
      </div>
      <script>
        function enviarArquivo() {
          const arquivo = document.getElementById('arquivo').files[0];
          const conta   = document.getElementById('conta').value;
          const status  = document.getElementById('status');
          const btn     = document.getElementById('btnEnviar');
          if (!conta)   { status.style.display='block'; status.className='erro'; status.textContent='❌ Selecione a conta primeiro!'; return; }
          if (!arquivo) { status.style.display='block'; status.className='erro'; status.textContent='❌ Selecione um arquivo!'; return; }
          btn.disabled = true;
          status.style.display='block'; status.className='carregando'; status.textContent='⏳ Processando... aguarde.';
          const reader = new FileReader();
          reader.onload = function(e) {
            const base64 = e.target.result.split(',')[1];
            google.script.run
              .withSuccessHandler(r => { btn.disabled=false; status.className='sucesso'; status.textContent='✅ '+r; })
              .withFailureHandler(e => { btn.disabled=false; status.className='erro'; status.textContent='❌ '+e.message; })
              .processarRelatorio(base64, conta, arquivo.name);
          };
          reader.readAsDataURL(arquivo);
        }
      </script>
    </body>
    </html>
  `).setTitle('Upload Relatório ML').setWidth(400).setHeight(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// PROCESSAR RELATÓRIO (atualizado: captura # Anúncio)
// ============================================================
function processarRelatorio(base64, conta, nomeArquivo) {
  try {
    const blob      = Utilities.newBlob(Utilities.base64Decode(base64),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', nomeArquivo);
    const arquivo   = DriveApp.createFile(blob);
    const convertido = Drive.Files.copy(
      { title: '_temp_full_ml', mimeType: 'application/vnd.google-apps.spreadsheet' },
      arquivo.getId(), { convert: true }
    );
    const planilhaTemp = SpreadsheetApp.openById(convertido.id);
    const dados        = planilhaTemp.getSheets()[0].getDataRange().getValues();
    DriveApp.getFileById(arquivo.getId()).setTrashed(true);
    DriveApp.getFileById(convertido.id).setTrashed(true);

    let linhaHeader = -1;
    let colunas     = {};

    for (let i = 0; i < dados.length; i++) {
      const linha    = dados[i].map(c => String(c).trim());
      const linhaStr = linha.join('|').toLowerCase();
      if (linhaStr.includes('código ml') || linhaStr.includes('codigo ml')) {
        linhaHeader = i;
        linha.forEach((col, j) => {
          const c = col.toLowerCase().replace(/\n/g, ' ').trim();
          if (c === 'código ml' || c === 'codigo ml')                            colunas.inventoryId    = j;
          if (c === 'sku')                                                        colunas.sku            = j;
          if (c === 'produto')                                                    colunas.produto        = j;
          if (c.includes('variaçao') || c.includes('variação') || c === 'variacao') colunas.variacao    = j;
          if (c.includes('vendas brutas'))                                        colunas.vendasBrutas   = j;
          if (c.includes('unidades vendidas'))                                    colunas.unidadesVend   = j;
          if (c.includes('unidades aptas'))                                       colunas.aptas          = j;
          if (c.includes('a caminho'))                                            colunas.caminho        = j;
          if (c.includes('nível de estoque') || c.includes('nivel de estoque'))  colunas.nivel          = j;
          if (c.includes('observações') || c.includes('observacoes'))            colunas.obs            = j;
          if (c === '# anúncio' || c === '# anuncio' || c === 'anúncio' || c === 'anuncio') colunas.anuncio = j;
        });
        break;
      }
    }

    if (linhaHeader < 0) throw new Error('Formato inválido. Verifique se é o relatório de Planejamento de envios.');
    if (colunas.inventoryId === undefined) throw new Error('Coluna "Código ML" não encontrada.');

    const agora              = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    const linhasProcessadas  = [];

    for (let i = linhaHeader + 1; i < dados.length; i++) {
      const linha      = dados[i];
      const inventoryId = String(linha[colunas.inventoryId] || '').trim();
      if (!inventoryId || inventoryId.length < 4 || inventoryId === 'undefined') continue;

      let vendasBrutas = linha[colunas.vendasBrutas] || 0;
      if (typeof vendasBrutas === 'string') {
        vendasBrutas = parseFloat(vendasBrutas.replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.')) || 0;
      }

      linhasProcessadas.push([
        conta,
        inventoryId,
        String(colunas.sku      !== undefined ? linha[colunas.sku]      : ''),
        String(colunas.produto  !== undefined ? linha[colunas.produto]  : ''),
        String(colunas.variacao !== undefined ? linha[colunas.variacao] : ''),
        vendasBrutas,
        Number(colunas.unidadesVend !== undefined ? linha[colunas.unidadesVend] : 0),
        Number(colunas.aptas        !== undefined ? linha[colunas.aptas]        : 0),
        Number(colunas.caminho      !== undefined ? linha[colunas.caminho]      : 0),
        String(colunas.nivel        !== undefined ? linha[colunas.nivel]        : ''),
        String(colunas.obs          !== undefined ? linha[colunas.obs]          : ''),
        agora,
        String(colunas.anuncio      !== undefined ? linha[colunas.anuncio]      : ''), // # Anúncio
      ]);
    }

    if (linhasProcessadas.length === 0) throw new Error('Nenhum produto encontrado no arquivo.');

    const aba               = obterOuCriarAba();
    const dadosExistentes   = aba.getDataRange().getValues();
    const linhasOutras      = dadosExistentes.filter((l, i) => i === 0 || String(l[0]).trim() !== conta);

    aba.clearContents();
    aba.appendRow(CABECALHOS);
    formatarCabecalho(aba);

    if (linhasOutras.length > 1) {
      const outras = linhasOutras.slice(1);
      if (outras.length > 0) aba.getRange(2, 1, outras.length, CABECALHOS.length).setValues(outras);
    }

    const ultimaLinha = aba.getLastRow() + 1;
    aba.getRange(ultimaLinha, 1, linhasProcessadas.length, CABECALHOS.length).setValues(linhasProcessadas);

    const totalLinhas = aba.getLastRow() - 1;
    if (totalLinhas > 0) {
      aba.getRange(2, 6, totalLinhas, 1).setNumberFormat('R$ #,##0.00');
      aba.getRange(2, 2, totalLinhas, 1).setNumberFormat('@');
      aba.getRange(2, 3, totalLinhas, 1).setNumberFormat('@');
    }

    aba.autoResizeColumns(1, CABECALHOS.length);
    aba.setFrozenRows(1);

    gerarCurvaABC();
    gerarCurvaABC2();
    return linhasProcessadas.length + ' produtos importados da conta ' + conta + '!';

  } catch(e) {
    Logger.log('Erro: ' + e.toString());
    throw new Error(e.toString());
  }
}

// ============================================================
// CURVA ABC (original, sem alteração)
// ============================================================
function gerarCurvaABC() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ss.getSheetByName(ABA_DADOS);
  if (!abaDados) return;

  const dados = abaDados.getDataRange().getValues();
  if (dados.length <= 1) return;

  const mapaSKU = {};

  for (let i = 1; i < dados.length; i++) {
    const linha        = dados[i];
    const conta        = String(linha[0] || '').trim();
    const inventoryId  = String(linha[1] || '').trim();
    const sku          = String(linha[2] || '').trim();
    const produto      = String(linha[3] || '').trim();
    const variacao     = String(linha[4] || '').trim();
    const faturamento  = parseFloat(linha[5]) || 0;
    const unidVendidas = parseInt(linha[6]) || 0;
    const estoqueApto  = parseInt(linha[7]) || 0;
    const estoqueCam   = parseInt(linha[8]) || 0;

    if (!sku && !inventoryId) continue;
    const chave = sku || inventoryId;

    if (!mapaSKU[chave]) {
      mapaSKU[chave] = { sku, produto, variacao, faturamento: 0, unidVendidas: 0,
                         estoqueApto: 0, estoqueCaminho: 0, contas: new Set() };
    }
    mapaSKU[chave].faturamento    += faturamento;
    mapaSKU[chave].unidVendidas   += unidVendidas;
    mapaSKU[chave].estoqueApto    += estoqueApto;
    mapaSKU[chave].estoqueCaminho += estoqueCam;
    if (conta) mapaSKU[chave].contas.add(conta);
  }

  const itens    = Object.values(mapaSKU).filter(i => i.faturamento > 0 || i.unidVendidas > 0)
                         .sort((a, b) => b.faturamento - a.faturamento);
  const fatTotal = itens.reduce((s, i) => s + i.faturamento, 0);
  let fatAcum    = 0;
  const linhasABC = [];

  for (let i = 0; i < itens.length; i++) {
    const item     = itens[i];
    fatAcum       += item.faturamento;
    const percAcum = fatTotal > 0 ? (fatAcum / fatTotal) * 100 : 0;
    const curva    = percAcum <= 80 ? 'A' : percAcum <= 95 ? 'B' : 'C';

    const ticketMedio    = item.unidVendidas > 0 ? item.faturamento / item.unidVendidas : 0;
    const estoqueTotal   = item.estoqueApto + item.estoqueCaminho;
    const recomendacao   = Math.max(0, item.unidVendidas - estoqueTotal);

    linhasABC.push([
      i + 1, curva, item.sku,
      item.produto + (item.variacao && item.variacao !== 'N/A' ? ' - ' + item.variacao : ''),
      item.faturamento, percAcum.toFixed(1) + '%',
      item.unidVendidas, ticketMedio,
      Array.from(item.contas).join(', '),
      item.estoqueApto, item.estoqueCaminho, estoqueTotal, recomendacao
    ]);
  }

  let abaABC = ss.getSheetByName('Curva ABC');
  if (!abaABC) abaABC = ss.insertSheet('Curva ABC');
  else { abaABC.clearContents(); abaABC.clearFormats(); }

  const cabABC = ['Ranking','Curva','SKU','Produto','Faturamento 30d (R$)','% Acumulado',
                  'Qtd. Vendida','Ticket Médio (R$)','Conta(s) ML',
                  'Estoque Apto (CD)','Estoque A Caminho','Estoque Total','Recomendação Reposição'];

  abaABC.appendRow(cabABC);
  abaABC.getRange(1,1,1,cabABC.length).setBackground('#1a73e8').setFontColor('#ffffff')
        .setFontWeight('bold').setHorizontalAlignment('center');
  abaABC.setFrozenRows(1);

  if (linhasABC.length > 0) {
    abaABC.getRange(2, 1, linhasABC.length, cabABC.length).setValues(linhasABC);
    abaABC.getRange(2, 5, linhasABC.length, 1).setNumberFormat('R$ #,##0.00');
    abaABC.getRange(2, 8, linhasABC.length, 1).setNumberFormat('R$ #,##0.00');
    abaABC.getRange(2, 3, linhasABC.length, 1).setNumberFormat('@');

    for (let i = 0; i < linhasABC.length; i++) {
      const cor = linhasABC[i][1] === 'A' ? '#e6f4ea' : linhasABC[i][1] === 'B' ? '#fef7e0' : '#fce8e6';
      abaABC.getRange(i + 2, 1, 1, cabABC.length).setBackground(cor);
      if (linhasABC[i][12] > 0) abaABC.getRange(i + 2, 13).setBackground('#ff6d00').setFontColor('#ffffff').setFontWeight('bold');
    }
  }

  abaABC.autoResizeColumns(1, cabABC.length);
  const totalRow = abaABC.getLastRow() + 2;
  abaABC.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold');
  abaABC.getRange(totalRow, 5).setFormula('=SUM(E2:E'+(linhasABC.length+1)+')').setNumberFormat('R$ #,##0.00').setFontWeight('bold');
  abaABC.getRange(totalRow, 7).setFormula('=SUM(G2:G'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC.getRange(totalRow, 10).setFormula('=SUM(J2:J'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC.getRange(totalRow, 11).setFormula('=SUM(K2:K'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC.getRange(totalRow, 12).setFormula('=SUM(L2:L'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC.getRange(totalRow, 13).setFormula('=SUM(M2:M'+(linhasABC.length+1)+')').setFontWeight('bold');

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Curva ABC gerada!', '📊 Full ML', 5);
}

// ============================================================
// CURVA ABC 2 (original, sem alteração)
// ============================================================
function gerarCurvaABC2() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ss.getSheetByName(ABA_DADOS);
  if (!abaDados) return;

  const dados = abaDados.getDataRange().getValues();
  if (dados.length <= 1) return;

  const itens = [];
  for (let i = 1; i < dados.length; i++) {
    const linha       = dados[i];
    const conta       = String(linha[0] || '').trim();
    const inventoryId = String(linha[1] || '').trim();
    const sku         = String(linha[2] || '').trim();
    const produto     = String(linha[3] || '').trim();
    const variacao    = String(linha[4] || '').trim();
    const faturamento = parseFloat(linha[5]) || 0;
    const unidVend    = parseInt(linha[6]) || 0;
    const estoqueApto = parseInt(linha[7]) || 0;
    const estoqueCam  = parseInt(linha[8]) || 0;
    const nivel       = String(linha[9] || '').trim();
    if (!inventoryId || inventoryId.length < 4) continue;
    itens.push({ conta, inventoryId, sku, produto, variacao, faturamento, unidVend, estoqueApto, estoqueCam, nivel });
  }

  itens.sort((a, b) => b.faturamento - a.faturamento);
  const fatTotal = itens.reduce((s, i) => s + i.faturamento, 0);
  let fatAcum    = 0;
  const linhasABC = [];

  for (let i = 0; i < itens.length; i++) {
    const item     = itens[i];
    fatAcum       += item.faturamento;
    const percAcum = fatTotal > 0 ? (fatAcum / fatTotal) * 100 : 0;
    const curva    = percAcum <= 80 ? 'A' : percAcum <= 95 ? 'B' : 'C';
    const tm       = item.unidVend > 0 ? item.faturamento / item.unidVend : 0;
    const estTotal = item.estoqueApto + item.estoqueCam;
    const rec      = Math.max(0, item.unidVend - estTotal);

    linhasABC.push([
      i + 1, curva, item.conta, item.inventoryId, item.sku,
      item.produto + (item.variacao && item.variacao !== 'N/A' && item.variacao !== '-' ? ' - ' + item.variacao : ''),
      item.faturamento, percAcum.toFixed(1) + '%', item.unidVend, tm,
      item.estoqueApto, item.estoqueCam, estTotal, item.nivel, rec
    ]);
  }

  let abaABC2 = ss.getSheetByName('Curva ABC 2');
  if (!abaABC2) abaABC2 = ss.insertSheet('Curva ABC 2');
  else { abaABC2.clearContents(); abaABC2.clearFormats(); }

  const cabABC2 = ['Ranking','Curva','Conta ML','Inventory ID','SKU','Produto',
                   'Faturamento 30d (R$)','% Acumulado','Qtd. Vendida','Ticket Médio (R$)',
                   'Estoque Apto (CD)','Estoque A Caminho','Estoque Total','Nível de Estoque','Recomendação Reposição'];

  abaABC2.appendRow(cabABC2);
  abaABC2.getRange(1,1,1,cabABC2.length).setBackground('#0f9d58').setFontColor('#ffffff')
         .setFontWeight('bold').setHorizontalAlignment('center');
  abaABC2.setFrozenRows(1);

  if (linhasABC.length > 0) {
    abaABC2.getRange(2,1,linhasABC.length,cabABC2.length).setValues(linhasABC);
    abaABC2.getRange(2,7,linhasABC.length,1).setNumberFormat('R$ #,##0.00');
    abaABC2.getRange(2,10,linhasABC.length,1).setNumberFormat('R$ #,##0.00');
    abaABC2.getRange(2,4,linhasABC.length,2).setNumberFormat('@');
    for (let i = 0; i < linhasABC.length; i++) {
      const cor = linhasABC[i][1] === 'A' ? '#e6f4ea' : linhasABC[i][1] === 'B' ? '#fef7e0' : '#fce8e6';
      abaABC2.getRange(i+2,1,1,cabABC2.length).setBackground(cor);
      if (linhasABC[i][14] > 0) abaABC2.getRange(i+2,15).setBackground('#ff6d00').setFontColor('#ffffff').setFontWeight('bold');
    }
  }

  abaABC2.autoResizeColumns(1, cabABC2.length);
  const totalRow = abaABC2.getLastRow() + 2;
  abaABC2.getRange(totalRow,1).setValue('TOTAL').setFontWeight('bold');
  abaABC2.getRange(totalRow,7).setFormula('=SUM(G2:G'+(linhasABC.length+1)+')').setNumberFormat('R$ #,##0.00').setFontWeight('bold');
  abaABC2.getRange(totalRow,9).setFormula('=SUM(I2:I'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC2.getRange(totalRow,11).setFormula('=SUM(K2:K'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC2.getRange(totalRow,12).setFormula('=SUM(L2:L'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC2.getRange(totalRow,13).setFormula('=SUM(M2:M'+(linhasABC.length+1)+')').setFontWeight('bold');
  abaABC2.getRange(totalRow,15).setFormula('=SUM(O2:O'+(linhasABC.length+1)+')').setFontWeight('bold');

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Curva ABC 2 gerada!', '📊 Full ML', 5);
}

// ============================================================
// LIMPAR CONTA / RESUMO / HELPERS (originais)
// ============================================================
function abrirLimparConta() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt('Limpar dados de qual conta?', 'Digite o nome da conta:\n' + CONTAS.join('\n'), ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const conta = resp.getResponseText().trim();
  if (!CONTAS.includes(conta)) { ui.alert('Conta não encontrada: ' + conta); return; }
  const aba   = obterOuCriarAba();
  const dados = aba.getDataRange().getValues();
  const linhasFiltradas = dados.filter((l, i) => i === 0 || String(l[0]).trim() !== conta);
  aba.clearContents();
  aba.getRange(1, 1, linhasFiltradas.length, CABECALHOS.length).setValues(linhasFiltradas);
  formatarCabecalho(aba);
  ui.alert('✅ Dados da conta ' + conta + ' removidos!');
}

function verResumo() {
  const aba   = obterOuCriarAba();
  const dados = aba.getDataRange().getValues();
  const resumo = {};
  for (const conta of CONTAS) resumo[conta] = { produtos: 0, vendas: 0, aptas: 0, caminho: 0 };
  for (let i = 1; i < dados.length; i++) {
    const conta = String(dados[i][0]).trim();
    if (!resumo[conta]) continue;
    resumo[conta].produtos++;
    resumo[conta].vendas  += Number(dados[i][6]) || 0;
    resumo[conta].aptas   += Number(dados[i][7]) || 0;
    resumo[conta].caminho += Number(dados[i][8]) || 0;
  }
  let texto = '📊 RESUMO POR CONTA\n\n';
  for (const [conta, r] of Object.entries(resumo)) {
    texto += '▸ ' + conta + (r.produtos > 0
      ? '\n  Produtos: ' + r.produtos + '\n  Vendas 30d: ' + r.vendas + ' un.\n  Estoque CD: ' + r.aptas + ' un.\n  A caminho: ' + r.caminho + ' un.\n\n'
      : ': sem dados\n\n');
  }
  SpreadsheetApp.getUi().alert(texto);
}

function obterOuCriarAba() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let aba   = ss.getSheetByName(ABA_DADOS);
  if (!aba) { aba = ss.insertSheet(ABA_DADOS); aba.appendRow(CABECALHOS); formatarCabecalho(aba); }
  return aba;
}

function formatarCabecalho(aba) {
  aba.getRange(1, 1, 1, CABECALHOS.length)
     .setBackground('#1a73e8').setFontColor('#ffffff')
     .setFontWeight('bold').setHorizontalAlignment('center');
  aba.setFrozenRows(1);
}

// ============================================================
// SIDEBAR: CONFIGURAR DIAS POR CURVA
// ============================================================
function abrirConfigDias() {
  const props = PropertiesService.getScriptProperties();
  const diasA = props.getProperty('DIAS_CURVA_A') || '';
  const diasB = props.getProperty('DIAS_CURVA_B') || '';
  const diasC = props.getProperty('DIAS_CURVA_C') || '';

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { box-sizing: border-box; }
        body { font-family: Arial, sans-serif; padding: 20px; margin: 0; background: #f8f9fa; }
        .card { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        h3 { color: #1a73e8; margin: 0 0 12px 0; }
        label { display: block; font-size: 13px; color: #5f6368; margin: 12px 0 4px 0; font-weight: 600; }
        input[type="number"] { width: 100%; padding: 8px 12px; border: 1px solid #dadce0; border-radius: 4px; font-size: 14px; }
        .btn { color: white; border: none; padding: 10px 24px; border-radius: 4px; cursor: pointer; width: 100%; margin-top: 10px; font-size: 14px; font-weight: 500; }
        .btn-salvar { background: #1a73e8; }
        .btn-gerar  { background: #0f9d58; }
        .btn:hover  { opacity: 0.88; }
        .btn:disabled { background: #9aa0a6; cursor: not-allowed; }
        #status { margin-top: 12px; padding: 12px; border-radius: 4px; display: none; font-size: 13px; }
        .sucesso    { background: #e6f4ea; color: #137333; border: 1px solid #ceead6; }
        .erro       { background: #fce8e6; color: #c5221f; border: 1px solid #f5c6c6; }
        .carregando { background: #e8f0fe; color: #1967d2; border: 1px solid #c5d8f6; }
        .info { background: #fef7e0; color: #b45309; border: 1px solid #fde293; font-size: 12px; padding: 8px; border-radius: 4px; margin-bottom: 14px; }
        .bloco { border-left: 4px solid #ccc; padding-left: 10px; margin-top: 10px; }
        .bloco-a { border-color: #137333; }
        .bloco-b { border-color: #b45309; }
        .bloco-c { border-color: #c5221f; }
      </style>
    </head>
    <body>
      <div class="card">
        <h3>⚙️ Dias de Estoque por Curva</h3>
        <div class="info">ℹ️ Defina quantos dias de cobertura de estoque quer para cada curva. Ex: A = 45 dias, B = 30 dias, C = 15 dias.</div>

        <div class="bloco bloco-a">
          <label>🟢 Curva A — Dias de cobertura:</label>
          <input type="number" id="diasA" min="1" max="365" value="${diasA}" placeholder="Ex: 45">
        </div>
        <div class="bloco bloco-b">
          <label>🟡 Curva B — Dias de cobertura:</label>
          <input type="number" id="diasB" min="1" max="365" value="${diasB}" placeholder="Ex: 30">
        </div>
        <div class="bloco bloco-c">
          <label>🔴 Curva C — Dias de cobertura:</label>
          <input type="number" id="diasC" min="1" max="365" value="${diasC}" placeholder="Ex: 15">
        </div>

        <button class="btn btn-salvar" id="btnSalvar" onclick="salvar()">💾 Salvar Configuração</button>
        <button class="btn btn-gerar"  id="btnGerar"  onclick="gerar()">🔄 Salvar e Gerar Reposição</button>
        <div id="status"></div>
      </div>
      <script>
        function validar() {
          const a = parseInt(document.getElementById('diasA').value);
          const b = parseInt(document.getElementById('diasB').value);
          const c = parseInt(document.getElementById('diasC').value);
          const erros = [];
          if (!a || a <= 0) erros.push('Curva A');
          if (!b || b <= 0) erros.push('Curva B');
          if (!c || c <= 0) erros.push('Curva C');
          return erros;
        }
        function mostrar(msg, tipo) {
          const s = document.getElementById('status');
          s.style.display = 'block'; s.className = tipo; s.textContent = msg;
        }
        function getDias() {
          return [parseInt(document.getElementById('diasA').value),
                  parseInt(document.getElementById('diasB').value),
                  parseInt(document.getElementById('diasC').value)];
        }
        function salvar() {
          const erros = validar();
          if (erros.length) { mostrar('❌ Preencha os dias para: ' + erros.join(', '), 'erro'); return; }
          const btn = document.getElementById('btnSalvar');
          btn.disabled = true; mostrar('⏳ Salvando...', 'carregando');
          const [a,b,c] = getDias();
          google.script.run
            .withSuccessHandler(r => { btn.disabled=false; mostrar('✅ '+r, 'sucesso'); })
            .withFailureHandler(e => { btn.disabled=false; mostrar('❌ '+e.message, 'erro'); })
            .salvarConfigDias(a, b, c);
        }
        function gerar() {
          const erros = validar();
          if (erros.length) { mostrar('❌ Preencha os dias para: ' + erros.join(', '), 'erro'); return; }
          const btn = document.getElementById('btnGerar');
          btn.disabled = true; mostrar('⏳ Gerando aba Reposição Full... aguarde alguns minutos.', 'carregando');
          const [a,b,c] = getDias();
          google.script.run
            .withSuccessHandler(r => { btn.disabled=false; mostrar('✅ '+r, 'sucesso'); })
            .withFailureHandler(e => { btn.disabled=false; mostrar('❌ '+e.message, 'erro'); })
            .salvarEGerarReposicao(a, b, c);
        }
      </script>
    </body>
    </html>
  `).setTitle('Configurar Dias por Curva').setWidth(360).setHeight(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function salvarConfigDias(diasA, diasB, diasC) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('DIAS_CURVA_A', String(diasA));
  p.setProperty('DIAS_CURVA_B', String(diasB));
  p.setProperty('DIAS_CURVA_C', String(diasC));
  return 'Configuração salva! A=' + diasA + 'd | B=' + diasB + 'd | C=' + diasC + 'd';
}

function salvarEGerarReposicao(diasA, diasB, diasC) {
  salvarConfigDias(diasA, diasB, diasC);
  gerarReposicaoFull();
  return 'Aba "' + ABA_REPOSICAO + '" gerada com sucesso!';
}

// ============================================================
// GERAR REPOSIÇÃO FULL — ponto de entrada
// ============================================================
function gerarReposicaoFull() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const diasA = parseInt(props.getProperty('DIAS_CURVA_A') || '0');
  const diasB = parseInt(props.getProperty('DIAS_CURVA_B') || '0');
  const diasC = parseInt(props.getProperty('DIAS_CURVA_C') || '0');

  if (!diasA || !diasB || !diasC) {
    ui.alert('❌ Configure os dias por curva primeiro!\nUse o menu: ⚙️ Configurar Dias por Curva');
    return;
  }

  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const abaDados = ss.getSheetByName(ABA_DADOS);
    if (!abaDados) throw new Error('Aba "' + ABA_DADOS + '" não encontrada.');

    const dadosML = abaDados.getDataRange().getValues();
    const geData  = _rpCarregarGE();
    const blData  = _rpCarregarBL();

    _rpEscreverAba(ss, dadosML, geData, blData, { A: diasA, B: diasB, C: diasC });

    ui.alert('✅ Aba "' + ABA_REPOSICAO + '" gerada com sucesso!');
  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
    Logger.log(e.stack);
  }
}

// ============================================================
// CARREGAR GE FINANCE
// Retorna: skuCanalMap[sku][canal] = { semanas:[7], cons:{} }
// ============================================================
function _rpCarregarGE() {
  const ss  = SpreadsheetApp.openById(RP_GE_ID);
  let   aba = ss.getSheetByName(RP_GE_ABA);
  if (!aba) {
    const norm = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
    const t    = norm(RP_GE_ABA);
    for (const a of ss.getSheets()) if (norm(a.getName()) === t) { aba = a; break; }
  }
  if (!aba) throw new Error('Aba GE Finance "' + RP_GE_ABA + '" não encontrada.');

  const dados = aba.getDataRange().getValues();

  let dataMax = new Date(0);
  for (let i = 1; i < dados.length; i++) {
    const d = new Date(dados[i][3]);
    if (!isNaN(d) && d > dataMax) dataMax = d;
  }
  dataMax.setHours(23, 59, 59, 999);

  const skuCanalMap = {};

  for (let i = 1; i < dados.length; i++) {
    const row   = dados[i];
    const sku   = String(row[1] || '').trim();
    const canal = String(row[7] || '').trim();
    const qtd   = Number(row[2] || 0);
    if (!sku || !canal || qtd <= 0) continue;

    const d = new Date(row[3]);
    if (isNaN(d)) continue;
    const dias = Math.floor((dataMax - d) / 86400000);
    if (dias < 0 || dias >= 90) continue;

    if (!skuCanalMap[sku])        skuCanalMap[sku] = {};
    if (!skuCanalMap[sku][canal]) skuCanalMap[sku][canal] = {
      semanas: [0,0,0,0,0,0,0],
      cons:    { '0-30': 0, '30-60': 0, '60-90': 0 },
    };

    const d_ = skuCanalMap[sku][canal];
    RP_PESOS.forEach((p, idx) => { if (dias >= p.de && dias < p.ate) d_.semanas[idx] += qtd; });
    RP_PER_CON.forEach(p => { if (dias >= p.de && dias < p.ate) d_.cons[p.nome] += qtd; });
  }

  return skuCanalMap;
}

// ============================================================
// CARREGAR BASELINKER
// ============================================================
function _rpCarregarBL() {
  const ss  = SpreadsheetApp.openById(RP_BL_ID);
  let   aba = ss.getSheetByName(RP_BL_ABA);
  if (!aba) {
    const norm = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
    const t    = norm(RP_BL_ABA);
    for (const a of ss.getSheets()) if (norm(a.getName()) === t) { aba = a; break; }
  }
  if (!aba) throw new Error('Aba BaseLinker "' + RP_BL_ABA + '" não encontrada.');

  const dados           = aba.getDataRange().getValues();
  const mapaTipo        = {}; // sku → 'Simples' | 'Composto'
  const mapaEstoque     = {}; // sku → saldo total
  const mapaComponentes = {}; // skuPai → [{codEst, qtde}]

  for (let i = 1; i < dados.length; i++) {
    const row  = dados[i];
    const sku  = String(row[1] || '').trim();
    const tipo = String(row[7] || '').trim();
    if (!sku) continue;
    const saldo = Number(row[4]||0) + Number(row[5]||0) + Number(row[6]||0);
    if (tipo === 'Simples') { mapaTipo[sku] = 'Simples'; mapaEstoque[sku] = saldo; }
    else if (tipo === 'PAI')    { mapaTipo[sku] = 'Composto'; mapaEstoque[sku] = saldo; }
  }

  for (let i = 1; i < dados.length; i++) {
    const row    = dados[i];
    const sku    = String(row[1] || '').trim();
    const tipo   = String(row[7] || '').trim();
    const codEst = String(row[2] || '').trim();
    const qtde   = Number(row[3] || 0);
    if (!sku || tipo !== 'Composto' || !codEst) continue;
    if (!mapaComponentes[sku]) mapaComponentes[sku] = [];
    mapaComponentes[sku].push({ codEst, qtde });
  }

  return { mapaTipo, mapaEstoque, mapaComponentes };
}

// ============================================================
// ESCREVER ABA REPOSIÇÃO FULL
// Layout: 40 colunas (A=1 … AN=40)
//
// A-O  (1-15):  Curva ABC (idêntico ABC 2)
// P    (16):    # Anúncio
// Q-W  (17-23): Vendas semanais GE Finance (por conta)
// X-Z  (24-26): Vendas 0-30, 30-60, 60-90
// AA   (27):    Sugestão reposição (cálculo ponderado)
// AB   (28):    Override manual
// AC   (29):    Reposição final — fórmula =IF(AB<>"",AB,AA)
// AD   (30):    Estoque BaseLinker
// AE   (31):    Saldo — fórmula =MAX(0,AD-AC)
// AF   (32):    Coluna separadora (vazia)
// AG   (33):    SKU simples (desmembrado)
// AH   (34):    Qtd. total necessária
// AI   (35):    BL − Necessário (fórmula)
// AJ   (36):    Alerta (vermelho se falta)
// AK-AN(37-40): Produtos afetados pela falta
// ============================================================
function _rpEscreverAba(ss, dadosML, geData, blData, diasPorCurva) {
  const { mapaTipo, mapaEstoque, mapaComponentes } = blData;
  const NCOLS = 40;
  const PRIMA = 3; // dados começam na linha 3

  // ── Preservar overrides manuais (AB) de runs anteriores ──
  const mapaOverride = {};
  const abaExist     = ss.getSheetByName(ABA_REPOSICAO);
  if (abaExist) {
    const dadosAtual = abaExist.getDataRange().getValues();
    for (let i = PRIMA - 1; i < dadosAtual.length; i++) {
      const invId = String(dadosAtual[i][3] || '').trim(); // col D
      const conta = String(dadosAtual[i][2] || '').trim(); // col C
      const abVal = dadosAtual[i][27];                     // col AB (índice 27)
      if (invId && conta && abVal !== '' && abVal !== null && abVal !== undefined) {
        mapaOverride[invId + '|' + conta] = abVal;
      }
    }
  }

  // ── Montar itens (mesma lógica ABC 2) ────────────────────
  const itens = [];
  for (let i = 1; i < dadosML.length; i++) {
    const l = dadosML[i];
    const inventoryId = String(l[1] || '').trim();
    if (!inventoryId || inventoryId.length < 4) continue;
    itens.push({
      conta:        String(l[0] || '').trim(),
      inventoryId,
      sku:          String(l[2] || '').trim(),
      produto:      String(l[3] || '').trim(),
      variacao:     String(l[4] || '').trim(),
      faturamento:  parseFloat(l[5]) || 0,
      unidVend:     parseInt(l[6]) || 0,
      estoqueApto:  parseInt(l[7]) || 0,
      estoqueCam:   parseInt(l[8]) || 0,
      nivel:        String(l[9] || '').trim(),
      anuncioId:    String(l[12] || '').trim(), // # Anúncio (índice 12)
    });
  }

  itens.sort((a, b) => b.faturamento - a.faturamento);
  const fatTotal = itens.reduce((s, x) => s + x.faturamento, 0);
  let   fatAcum  = 0;

  const linhasMain = [];

  itens.forEach((item, idx) => {
    fatAcum += item.faturamento;
    const percAcum = fatTotal > 0 ? (fatAcum / fatTotal) * 100 : 0;
    const curva    = percAcum <= 80 ? 'A' : percAcum <= 95 ? 'B' : 'C';

    const tm         = item.unidVend > 0 ? item.faturamento / item.unidVend : 0;
    const estTotal   = item.estoqueApto + item.estoqueCam;
    const recML      = Math.max(0, item.unidVend - estTotal);

    const ge         = geData[item.sku] && geData[item.sku][item.conta] || null;
    const semanas    = ge ? ge.semanas : [0,0,0,0,0,0,0];
    const con030     = ge ? ge.cons['0-30']  : 0;
    const con3060    = ge ? ge.cons['30-60'] : 0;
    const con6090    = ge ? ge.cons['60-90'] : 0;

    // Taxa diária ponderada
    let taxa = 0;
    RP_PESOS.forEach((p, i) => { taxa += (semanas[i] / p.dias) * p.peso; });
    const sugestao = Math.round(taxa * (diasPorCurva[curva] || 30));

    const estBL = mapaEstoque[item.sku] !== undefined ? mapaEstoque[item.sku] : 0;

    linhasMain.push({
      // A-O
      ranking: idx + 1, curva, conta: item.conta,
      inventoryId: item.inventoryId, sku: item.sku,
      produto: item.produto + (item.variacao && item.variacao !== 'N/A' && item.variacao !== '-' ? ' - ' + item.variacao : ''),
      faturamento: item.faturamento, percAcum: percAcum.toFixed(1) + '%',
      unidVend: item.unidVend, tm,
      estoqueApto: item.estoqueApto, estoqueCam: item.estoqueCam, estTotal,
      nivel: item.nivel, recML,
      // P, Q-W, X-Z
      anuncioId: item.anuncioId, semanas, con030, con3060, con6090,
      // AA, AD
      sugestao, estBL,
    });
  });

  // ── Tabela de componentes (AG-AN) ─────────────────────────
  // Usa AC = override se existir, senão sugestao
  const skuNec = {}; // skuSimples → { qtd, afetados:Set }

  linhasMain.forEach(r => {
    const key       = r.inventoryId + '|' + r.conta;
    const override  = mapaOverride[key];
    const reposicao = (override !== undefined && override !== '') ? Number(override) : r.sugestao;
    if (reposicao <= 0) return;

    const sku = r.sku;
    if (mapaComponentes[sku]) {
      mapaComponentes[sku].forEach(comp => {
        if (!skuNec[comp.codEst]) skuNec[comp.codEst] = { qtd: 0, afetados: new Set() };
        skuNec[comp.codEst].qtd += reposicao * comp.qtde;
        skuNec[comp.codEst].afetados.add(sku);
      });
    } else {
      if (!skuNec[sku]) skuNec[sku] = { qtd: 0, afetados: new Set() };
      skuNec[sku].qtd += reposicao;
    }
  });

  const linhasComp = Object.entries(skuNec)
    .sort((a, b) => b[1].qtd - a[1].qtd)
    .map(([skuS, data]) => {
      const qtdNec  = Math.ceil(data.qtd);
      const estB    = mapaEstoque[skuS] !== undefined ? mapaEstoque[skuS] : 0;
      const saldo   = estB - qtdNec;
      const alerta  = saldo < 0 ? 'Falta: ' + Math.abs(saldo) : '';
      const af      = Array.from(data.afetados).slice(0, 4);
      return [ skuS, qtdNec, saldo, alerta, af[0]||'', af[1]||'', af[2]||'', af[3]||'' ];
    });

  // ── Criar/limpar aba ──────────────────────────────────────
  let aba = ss.getSheetByName(ABA_REPOSICAO);
  if (aba) {
    aba.getRange(1, 1, aba.getMaxRows(), aba.getMaxColumns()).breakApart();
    aba.clearContents();
    aba.clearFormats();
  } else {
    aba = ss.insertSheet(ABA_REPOSICAO);
  }

  // ── Headers ───────────────────────────────────────────────
  const h1 = new Array(NCOLS).fill('');
  h1[0]  = 'Curva ABC';
  h1[15] = '# Anúncio';
  h1[16] = 'Vendas Semanais — GE Finance (por conta)';
  h1[23] = 'Vendas por Período';
  h1[26] = 'Reposição';
  h1[29] = 'Estoque';
  h1[32] = 'Análise de Componentes (Desmembrado)';

  const h2 = [
    'Ranking','Curva','Conta ML','Inventory ID','SKU','Produto',
    'Fat. 30d (R$)','% Acum.','Qtd. Vendida','Ticket Médio (R$)',
    'Est. Apto (CD)','Est. A Caminho','Est. Total (CD)','Nível Estoque','Recomend. ML',
    '# Anúncio',
    '0-7','7-15','15-30','30-45','45-60','60-75','75-90',
    '0-30','30-60','60-90',
    'Sugestão Reposição','Override Manual','Reposição Final',
    'Estoque BL','Saldo (BL−Rep.)',
    '',
    'SKU Simples','Qtd. Necessária','BL − Necessário','Alerta',
    'Afetado 1','Afetado 2','Afetado 3','Afetado 4',
  ];

  aba.getRange(1, 1, 1, NCOLS).setValues([h1]);
  aba.getRange(2, 1, 1, NCOLS).setValues([h2]);

  // Merges linha 1
  aba.getRange(1, 1,  1, 15).merge(); // A-O
  aba.getRange(1, 17, 1,  7).merge(); // Q-W
  aba.getRange(1, 24, 1,  3).merge(); // X-Z
  aba.getRange(1, 27, 1,  3).merge(); // AA-AC
  aba.getRange(1, 30, 1,  2).merge(); // AD-AE
  aba.getRange(1, 33, 1,  8).merge(); // AG-AN

  // Cores dos headers linha 1
  const secoes = [
    [1,  15, '#1a73e8'], // Curva ABC
    [16,  1, '#5f6368'], // Anúncio
    [17,  7, '#0f9d58'], // Vendas Semanais
    [24,  3, '#137333'], // Vendas Período
    [27,  3, '#e37400'], // Reposição
    [30,  2, '#7b1fa2'], // Estoque
    [32,  1, '#ffffff'], // Separador
    [33,  8, '#b71c1c'], // Componentes
  ];
  secoes.forEach(([col, qtd, bg]) => {
    const fg = bg === '#ffffff' ? '#ffffff' : '#ffffff';
    aba.getRange(1, col, 1, qtd)
       .setBackground(bg).setFontColor(fg)
       .setFontWeight('bold').setHorizontalAlignment('center');
  });

  // Header linha 2
  aba.getRange(2, 1, 1, NCOLS)
     .setBackground('#eceff1').setFontWeight('bold')
     .setHorizontalAlignment('center')
     .setBorder(true, true, true, true, true, true, '#90a4ae',
                SpreadsheetApp.BorderStyle.SOLID);

  if (linhasMain.length === 0) { SpreadsheetApp.flush(); return; }

  // ── Montar e escrever linhas ──────────────────────────────
  const rows = linhasMain.map(r => {
    const row = new Array(NCOLS).fill('');
    row[0]  = r.ranking;
    row[1]  = r.curva;
    row[2]  = r.conta;
    row[3]  = r.inventoryId;
    row[4]  = r.sku;
    row[5]  = r.produto;
    row[6]  = r.faturamento;
    row[7]  = r.percAcum;
    row[8]  = r.unidVend;
    row[9]  = r.tm;
    row[10] = r.estoqueApto;
    row[11] = r.estoqueCam;
    row[12] = r.estTotal;
    row[13] = r.nivel;
    row[14] = r.recML;
    row[15] = r.anuncioId;
    r.semanas.forEach((v, i) => { row[16 + i] = v; });
    row[23] = r.con030;
    row[24] = r.con3060;
    row[25] = r.con6090;
    row[26] = r.sugestao;   // AA
    row[27] = '';            // AB (override manual — preenchido depois)
    // row[28] AC = fórmula
    row[29] = r.estBL;       // AD
    // row[30] AE = fórmula
    return row;
  });

  aba.getRange(PRIMA, 1, rows.length, NCOLS).setValues(rows);

  // Restaurar overrides e escrever fórmulas AC + AE
  for (let i = 0; i < linhasMain.length; i++) {
    const ln  = PRIMA + i;
    const key = linhasMain[i].inventoryId + '|' + linhasMain[i].conta;
    if (mapaOverride[key] !== undefined && mapaOverride[key] !== '') {
      aba.getRange(ln, 28).setValue(mapaOverride[key]);
    }
    aba.getRange(ln, 29).setFormula(`=IF(AB${ln}<>"",VALUE(AB${ln}),AA${ln})`); // AC
    aba.getRange(ln, 31).setFormula(`=MAX(0,AD${ln}-AC${ln})`);                  // AE
  }

  // ── Formatação das linhas de dados ───────────────────────
  const COR_CURVA = { A: '#e8f5e9', B: '#fffde7', C: '#fce4ec' };
  const grupoCurva = { A: [], B: [], C: [] };
  linhasMain.forEach((r, i) => { grupoCurva[r.curva].push(PRIMA + i); });

  Object.entries(grupoCurva).forEach(([curva, linhas]) => {
    linhas.forEach(ln => {
      aba.getRange(ln, 1, 1, 15).setBackground(COR_CURVA[curva]);
    });
  });

  // Faixas com cores de seção
  aba.getRange(PRIMA, 16, rows.length, 1).setBackground('#f5f5f5');          // P: Anúncio
  aba.getRange(PRIMA, 17, rows.length, 10).setBackground('#f1f8e9');          // Q-Z: Vendas GE
  aba.getRange(PRIMA, 27, rows.length, 1).setBackground('#fff8e1');           // AA: Sugestão
  aba.getRange(PRIMA, 28, rows.length, 1).setBackground('#fff3e0');           // AB: Override
  aba.getRange(PRIMA, 29, rows.length, 1).setBackground('#ffe0b2').setFontWeight('bold'); // AC: Final
  aba.getRange(PRIMA, 30, rows.length, 1).setBackground('#f3e5f5');           // AD: BL
  aba.getRange(PRIMA, 31, rows.length, 1).setBackground('#ede7f6');           // AE: Saldo

  // Formatos numéricos
  aba.getRange(PRIMA, 4,  rows.length, 2).setNumberFormat('@');               // Inventory ID + SKU
  aba.getRange(PRIMA, 7,  rows.length, 1).setNumberFormat('R$ #,##0.00');     // Faturamento
  aba.getRange(PRIMA, 10, rows.length, 1).setNumberFormat('R$ #,##0.00');     // Ticket Médio
  aba.getRange(PRIMA, 30, rows.length, 1).setNumberFormat('#,##0');           // BL
  aba.getRange(PRIMA, 31, rows.length, 1).setNumberFormat('#,##0');           // Saldo

  // Destacar AE = 0 quando AC > 0 (estoque não cobre reposição)
  const semEstoqueCells = [];
  linhasMain.forEach((r, i) => {
    if (r.estBL === 0 && r.sugestao > 0) {
      semEstoqueCells.push(aba.getRange(PRIMA + i, 31).getA1Notation());
    }
  });
  if (semEstoqueCells.length) aba.getRangeList(semEstoqueCells).setBackground('#ef9a9a');

  // Congelar cabeçalhos e colunas iniciais
  aba.setFrozenRows(2);
  aba.setFrozenColumns(5);

  // ── Tabela de componentes (AG-AN) ─────────────────────────
  const COL_AG = 33;

  if (linhasComp.length > 0) {
    aba.getRange(PRIMA, COL_AG, linhasComp.length, 8).setValues(linhasComp);
    aba.getRange(PRIMA, COL_AG, linhasComp.length, 1).setNumberFormat('@');     // SKU texto
    aba.getRange(PRIMA, COL_AG + 4, linhasComp.length, 4).setNumberFormat('@'); // Afetados texto

    linhasComp.forEach((comp, i) => {
      const ln     = PRIMA + i;
      const alerta = comp[3];
      const bg     = alerta ? '#ffcdd2' : '#e8f5e9';
      aba.getRange(ln, COL_AG, 1, 8).setBackground(bg);
      if (alerta) {
        aba.getRange(ln, COL_AG + 3, 1, 1)
           .setBackground('#c62828').setFontColor('#ffffff').setFontWeight('bold');
      }
    });

    // Fórmula dinâmica para AI (coluna 35) = BL daquele SKU simples - AH
    // Já escrevemos o valor calculado em linhasComp[2] (AI = saldo)
    // mas se quiser formula: aba.getRange(PRIMA+i, 35).setFormula(...)
  }

  // Header da tabela componentes (linha 2, colunas AG-AN)
  aba.getRange(2, COL_AG, 1, 8)
     .setBackground('#c62828').setFontColor('#ffffff')
     .setFontWeight('bold').setHorizontalAlignment('center');

  aba.autoResizeColumns(1, NCOLS);
  SpreadsheetApp.flush();
}
