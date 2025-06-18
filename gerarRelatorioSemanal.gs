function enviarResumoSemanalLeads() {
  const emailDestino = 'jampalife.business@gmail.com, kleber.jampalife@gmail.com';
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ['Zap Imóveis', 'Imovelweb', 'Chaves na Mão'];
  const hoje = new Date();
  const diaSemana = hoje.getDay();
  const domingoPassado = new Date(hoje);
  domingoPassado.setDate(domingoPassado.getDate() - diaSemana - 7);
  const sabadoPassado = new Date(domingoPassado);
  sabadoPassado.setDate(domingoPassado.getDate() + 6);

  const formatarData = data => Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const pasta = DriveApp.createFolder(`Resumo Leads (${formatarData(domingoPassado)} a ${formatarData(sabadoPassado)})`);
  const dadosPorCorretor = {};
  
const emailsUnicos = new Set();


  abas.forEach(origem => {
    const aba = planilha.getSheetByName(origem);
    if (!aba) return;

    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const data = new Date(linha[7]); // Coluna H (índice 7)
      if (isNaN(data.getTime())) continue;

      if (data >= domingoPassado && data <= sabadoPassado) {
        const corretor = linha[8] || 'Sem Corretor'; // Coluna I (índice 8)
        const email = linha[2] || ''; // Coluna C (índice 2)
        const codigo = linha[4] || ''; // Coluna E (índice 4)

        if (email && !emailsUnicos.has(email)) {
  emailsUnicos.add(email);
  
  if (!dadosPorCorretor[corretor]) dadosPorCorretor[corretor] = [];
  dadosPorCorretor[corretor].push([
    formatarData(data),
    origem,
    codigo,
    email
  ]);
}
      }
    }
  });

  // Criar planilha de resumo
  const planilhaResumo = SpreadsheetApp.create('Resumo Leads - Corretores');

  Object.entries(dadosPorCorretor).forEach(([corretor, dados]) => {
    // Ordenar os leads por data
    dados.sort((a, b) => {
      const dataA = new Date(a[0].split('/').reverse().join('-'));
      const dataB = new Date(b[0].split('/').reverse().join('-'));
      return dataA - dataB;
    });

    const aba = planilhaResumo.insertSheet(corretor);
    aba.appendRow([corretor, "", "", ""]);
    aba.getRange(1, 1, 1, 4).merge();
    aba.getRange(1, 1).setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
    aba.appendRow(['Data', 'Origem do lead', 'Código do anúncio', 'E-mail do cliente']);
    dados.forEach(d => aba.appendRow(d));
    aba.autoResizeColumns(1, 4);
  });

  const padrao = planilhaResumo.getSheetByName('Sheet1');
  if (padrao) planilhaResumo.deleteSheet(padrao);

  const pdfLeads = planilhaResumo.getBlob().getAs('application/pdf').setName('Leads por Corretor.pdf');
  pasta.createFile(pdfLeads);
  DriveApp.getFileById(planilhaResumo.getId()).setTrashed(true);

  // Criar gráfico geral
  const planilhaGrafico = SpreadsheetApp.create('Resumo Gráfico Leads');
  const abaGrafico = planilhaGrafico.getActiveSheet();
  abaGrafico.setName('Resumo');
  abaGrafico.appendRow(['Corretor', 'Total de Leads']);

  Object.entries(dadosPorCorretor).forEach(([corretor, dados]) => {
    abaGrafico.appendRow([corretor, dados.length]);
  });

  const totalLinhas = abaGrafico.getLastRow();
  if (totalLinhas > 1) {
    const range = abaGrafico.getRange(2, 1, totalLinhas - 1, 2);
    const chart = abaGrafico.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(range)
      .setPosition(totalLinhas + 2, 1, 0, 0)
      .setOption('title', 'Distribuição de Leads por Corretor')
      .build();
    abaGrafico.insertChart(chart);
  }

  const pdfGrafico = planilhaGrafico.getBlob().getAs('application/pdf').setName('Gráfico Geral de Leads.pdf');
  pasta.createFile(pdfGrafico);
  DriveApp.getFileById(planilhaGrafico.getId()).setTrashed(true);

  MailApp.sendEmail({
    to: emailDestino,
    subject: `Resumo de Leads (${formatarData(domingoPassado)} a ${formatarData(sabadoPassado)})`,
    body: 'Segue em anexo o resumo dos leads da semana passada.',
    attachments: [pdfLeads, pdfGrafico]
  });
}
