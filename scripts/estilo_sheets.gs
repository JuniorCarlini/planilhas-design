/**
 * planilhas-design — estilo padrão para Google Sheets (Apps Script).
 *
 * Como usar: Extensões → Apps Script → cole este arquivo → salve →
 * selecione a função aplicarEstiloPadrao → Executar.
 * Também cria o menu "Estilo" na planilha (após recarregar).
 *
 * Layout esperado: título A1, subtítulo A2, linha 3 vazia, cabeçalho na linha 4.
 * Formatos de número são detectados pelo nome do cabeçalho (veja DETECTAR).
 */
const COR = {
  primaria: '#12424D', destaque: '#FFB745', primariaClara: '#E7EFF0', zebra: '#F5F8F8', borda: '#D5E0E2',
  texto: '#1E2B2E', textoSuave: '#5F7378', branco: '#FFFFFF',
  positivo: '#2E7D32', negativo: '#C62828', alerta: '#B26A00',
  entrada: '#FFF1D6', entradaTexto: '#12424D', textoFaixa: '#B7CDD2',
};
const CORES_GRAFICO = ['#12424D', '#FFB745', '#3E8A96', '#8CC2C9', '#D9822B', '#A7B4B7'];
const FONTE = 'Poppins';      // texto simples
const FONTE_TITULO = 'Sora';  // títulos (A1, cabeçalho, gráficos)
const LINHA_CABECALHO = 4;
const FORMATOS = {
  moeda: '"R$" #,##0.00;-"R$" #,##0.00',
  inteiro: '#,##0', decimal: '#,##0.00', percentual: '0.0%',
  variacao: '+0.0%;-0.0%;0.0%', data: 'dd/mm/yyyy', mes: 'mmm/yyyy',
};
// Palavras no cabeçalho → tipo de coluna
const DETECTAR = [
  [/valor|pre[cç]o|custo|total|saldo|receita|despesa|r\$/i, 'moeda'],
  [/varia[cç][aã]o|var\.?\s?%/i, 'variacao'],
  [/%|percent|taxa/i, 'percentual'],
  [/data|vencimento|prazo|in[ií]cio|fim/i, 'data'],
  [/m[eê]s|compet[eê]ncia/i, 'mes'],
  [/qtd|quantidade|n[uú]mero|qtde/i, 'inteiro'],
  [/status|situa[cç][aã]o/i, 'status'],
];
const STATUS = {
  positivo: ['OK', 'Concluído', 'Pago', 'Ativo', 'Em dia', 'Aprovado', 'Novo'],
  alerta: ['Pendente', 'Atenção', 'Em andamento', 'Em análise', 'Em renovação'],
  negativo: ['Atrasado', 'Erro', 'Vencido', 'Reprovado'],
  textoSuave: ['Cancelado', 'Inativo', 'Encerrado', 'Arquivado'], // neutro
};
const NUMERICOS = ['moeda', 'inteiro', 'decimal', 'percentual', 'variacao'];
const CENTRO = ['data', 'mes', 'status'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Estilo')
    .addItem('Aplicar estilo padrão', 'aplicarEstiloPadrao')
    .addItem('Marcar seleção como entrada', 'marcarEntrada')
    .addItem('Colorir gráficos da aba', 'estilizarGraficos')
    .addToUi();
}

function aplicarEstiloPadrao() {
  const sh = SpreadsheetApp.getActiveSheet();
  const h = LINHA_CABECALHO;
  const ultimaCol = sh.getLastColumn();
  const ultimaLin = Math.max(sh.getLastRow(), h);
  if (ultimaCol < 1) return;

  sh.getRange(1, 1, ultimaLin, ultimaCol).setFontFamily(FONTE).setFontSize(10).setFontColor(COR.texto);
  // Faixa do topo: linhas 1–2 em petróleo na largura da tabela
  sh.getRange(1, 1, 2, ultimaCol).setBackground(COR.primaria);
  sh.getRange('A1').setFontFamily(FONTE_TITULO).setFontSize(16).setFontWeight('bold')
    .setFontColor(COR.branco).setVerticalAlignment('bottom');
  sh.getRange('A2').setFontStyle('italic').setFontColor(COR.textoFaixa).setVerticalAlignment('top');
  sh.getRange('A1:A2').setHorizontalAlignment('left');
  sh.setRowHeight(1, 45);
  sh.setRowHeight(2, 32);

  const cab = sh.getRange(h, 1, 1, ultimaCol);
  cab.setFontFamily(FONTE_TITULO).setBackground(COR.primaria).setFontColor(COR.branco).setFontWeight('bold')
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(null, null, true, null, null, null, COR.destaque, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(h, 24);

  const nomes = cab.getValues()[0];
  const nDados = ultimaLin - h;
  if (nDados > 0) {
    const corpo = sh.getRange(h + 1, 1, nDados, ultimaCol);
    const fundos = [];
    for (let i = 0; i < nDados; i++) {
      fundos.push(new Array(ultimaCol).fill(i % 2 === 1 ? COR.zebra : null));
    }
    corpo.setBackgrounds(fundos).setVerticalAlignment('middle');
    corpo.setBorder(null, null, true, null, null, true, COR.borda, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeights(h + 1, nDados, 20);
  }

  nomes.forEach((nome, i) => {
    const col = i + 1;
    let tipo = 'texto';
    for (const [re, t] of DETECTAR) { if (re.test(String(nome))) { tipo = t; break; } }
    const horiz = NUMERICOS.includes(tipo) ? 'right' : CENTRO.includes(tipo) ? 'center' : 'left';
    sh.getRange(h, col, Math.max(nDados + 1, 1), 1).setHorizontalAlignment(horiz);
    if (nDados > 0 && FORMATOS[tipo]) sh.getRange(h + 1, col, nDados, 1).setNumberFormat(FORMATOS[tipo]);
    if (nDados > 0 && NUMERICOS.includes(tipo)) negativosEmVermelho_(sh, sh.getRange(h + 1, col, nDados, 1));
    if (nDados > 0 && tipo === 'status') aplicarStatus_(sh, sh.getRange(h + 1, col, nDados, 1));
    sh.autoResizeColumn(col);
    const w = sh.getColumnWidth(col);
    // folga para Poppins (~15% mais larga) e para o botão do filtro
    sh.setColumnWidth(col, Math.min(Math.max(Math.round(w * 1.15) + 24, 80), 360));
  });

  sh.setHiddenGridlines(true);
  sh.setFrozenRows(h);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(h, 1, Math.max(nDados + 1, 1), ultimaCol).createFilter();

  // Linha de total: se a última linha começa com "Total", estiliza como total
  if (nDados > 0 && String(sh.getRange(ultimaLin, 1).getValue()).toLowerCase().startsWith('total')) {
    sh.getRange(ultimaLin, 1, 1, ultimaCol).setBackground(COR.primariaClara)
      .setFontWeight('bold').setFontColor(COR.primaria)
      .setBorder(true, null, null, null, null, null, COR.primaria, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

function negativosEmVermelho_(sh, faixa) {
  const regras = sh.getConditionalFormatRules();
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
    .setFontColor(COR.negativo).setRanges([faixa]).build());
  sh.setConditionalFormatRules(regras);
}

function aplicarStatus_(sh, faixa) {
  const regras = sh.getConditionalFormatRules();
  Object.entries(STATUS).forEach(([cor, palavras]) => {
    palavras.forEach(p => regras.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(p)
        .setFontColor(COR[cor]).setBold(true).setRanges([faixa]).build()));
  });
  sh.setConditionalFormatRules(regras);
}

/** Marca as células selecionadas como "entrada" (editáveis pelo usuário). */
function marcarEntrada() {
  SpreadsheetApp.getActiveRange().setBackground(COR.entrada).setFontColor(COR.entradaTexto);
}

/** Aplica a paleta padrão a todos os gráficos da aba ativa. */
function estilizarGraficos() {
  const sh = SpreadsheetApp.getActiveSheet();
  sh.getCharts().forEach(g => {
    const novo = g.modify().setOption('colors', CORES_GRAFICO)
      .setOption('fontName', FONTE)
      .setOption('titleTextStyle', { color: COR.primaria, fontSize: 14, bold: true, fontName: FONTE_TITULO })
      .setOption('legend', { position: 'bottom', textStyle: { color: COR.textoSuave } })
      .build();
    sh.updateChart(novo);
  });
}
