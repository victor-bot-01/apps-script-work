/**
 * Relatórios automáticos em PDF: resumo semanal, resumo mensal, produtos
 * críticos, produtos para produzir, produtos sem etiqueta, estoque.
 *
 * Cada relatório vira uma aba temporária formatada, exportada como PDF
 * (via a URL de export nativa do Sheets) e salva numa pasta do Drive —
 * o prompt não diz se deveria ser enviado por e-mail também, então por
 * ora só salva no Drive (fácil de acrescentar envio por e-mail depois,
 * igual ao alerta de falha da importação).
 *
 * Cadência: roda dentro de `executarRelatoriosAgendados`, pensado pra um
 * gatilho diário — toda segunda-feira gera o pacote semanal, todo dia 1
 * do mês gera o resumo mensal (ver `configurarGatilhosRelatorios` em
 * Gatilhos.gs).
 *
 * Se `UrlFetchApp`/`DriveApp` pedirem autorização de escopo extra na
 * primeira execução, é esperado — aceite a permissão.
 */

const NOME_PASTA_RELATORIOS = 'Relatórios Essência do Brasil';

function executarRelatoriosAgendados() {
  const hoje = new Date();
  const relatoriosDeHoje = [];

  if (hoje.getDay() === 1) {
    relatoriosDeHoje.push(
      { nome: 'Resumo Semanal', funcao: gerarRelatorioResumoSemanal },
      { nome: 'Produtos Críticos', funcao: gerarRelatorioProdutosCriticos },
      { nome: 'Produtos para Produzir', funcao: gerarRelatorioProdutosParaProduzir },
      { nome: 'Produtos Sem Etiqueta', funcao: gerarRelatorioProdutosSemEtiqueta },
      { nome: 'Estoque', funcao: gerarRelatorioEstoque }
    );
  }
  if (hoje.getDate() === 1) {
    relatoriosDeHoje.push({ nome: 'Resumo Mensal', funcao: gerarRelatorioResumoMensal });
  }

  if (relatoriosDeHoje.length === 0) return;

  const inicio = new Date();
  const gerados = [];
  const erros = [];

  relatoriosDeHoje.forEach(function (relatorio) {
    try {
      relatorio.funcao();
      gerados.push(relatorio.nome);
    } catch (erro) {
      erros.push(relatorio.nome + ': ' + (erro && erro.message ? erro.message : erro));
    }
  });

  registrarLog_({
    inicio: inicio,
    fim: new Date(),
    pedidosImportados: 0,
    erros: erros.join(' | '),
    avisos: 'Relatórios gerados: ' + (gerados.join(', ') || 'nenhum'),
  });

  if (erros.length > 0) {
    enviarAlertaFalha_('Falha ao gerar relatórios em PDF: ' + erros.join(' | '));
  }
}

function gerarTodosRelatoriosSemanais() {
  gerarRelatorioResumoSemanal();
  gerarRelatorioProdutosCriticos();
  gerarRelatorioProdutosParaProduzir();
  gerarRelatorioProdutosSemEtiqueta();
  gerarRelatorioEstoque();
}

function gerarRelatorioResumoSemanal() {
  return gerarRelatorioPdf_(
    'Resumo_Semanal',
    'Essência do Brasil — Resumo Semanal',
    ['Indicador', 'Valor'],
    kpisParaLinhas_(calcularKPIs_(7))
  );
}

function gerarRelatorioResumoMensal() {
  return gerarRelatorioPdf_(
    'Resumo_Mensal',
    'Essência do Brasil — Resumo Mensal',
    ['Indicador', 'Valor'],
    kpisParaLinhas_(calcularKPIs_(30))
  );
}

function kpisParaLinhas_(kpis) {
  return [
    ['Pedidos no período', kpis.pedidosNoPeriodo],
    ['Itens vendidos', kpis.itensVendidos],
    ['Produtos ativos', kpis.produtosAtivos],
    ['Produtos sem venda', kpis.produtosSemVenda],
    ['Estoque total', kpis.estoqueTotal],
    ['Produtos em risco de ruptura', kpis.produtosEmRiscoDeRuptura],
    ['Kits vendidos', kpis.kits],
    ['Individuais vendidos', kpis.individuais],
    ['Produto campeão', kpis.produtoCampeao],
    ['Marketplace campeão', kpis.marketplaceCampeao],
  ];
}

function gerarRelatorioProdutosCriticos() {
  const linhas = calcularEstoqueInteligente()
    .filter(function (linha) {
      return linha.status !== 'Normal';
    })
    .map(function (linha) {
      return [
        linha.produto,
        linha.estoqueAtual,
        linha.nivelMinimo,
        linha.diasRestantes !== null ? Math.round(linha.diasRestantes) : '—',
        linha.status,
      ];
    });
  return gerarRelatorioPdf_(
    'Produtos_Criticos',
    'Essência do Brasil — Produtos Críticos',
    ['Produto', 'Estoque', 'Nível Mínimo', 'Dias Restantes', 'Status'],
    linhas
  );
}

function gerarRelatorioProdutosParaProduzir() {
  const linhas = calcularPrevisaoDemanda()
    .filter(function (linha) {
      return linha.sugestaoProducao > 0;
    })
    .map(function (linha) {
      return [linha.produto, linha.sugestaoProducao];
    });
  return gerarRelatorioPdf_(
    'Produtos_Para_Produzir',
    'Essência do Brasil — Produtos para Produzir',
    ['Produto', 'Sugestão de Produção (30d)'],
    linhas
  );
}

function gerarRelatorioProdutosSemEtiqueta() {
  const linhas = obterProdutosSemEtiqueta_().map(function (produto) {
    return [produto];
  });
  return gerarRelatorioPdf_(
    'Produtos_Sem_Etiqueta',
    'Essência do Brasil — Produtos Sem Etiqueta',
    ['Produto'],
    linhas
  );
}

function gerarRelatorioEstoque() {
  const estoque = lerEstoque_();
  const linhas = Array.from(estoque.entries())
    .map(function (par) {
      return [par[0], par[1]];
    })
    .sort(function (a, b) {
      return a[0].localeCompare(b[0]);
    });
  return gerarRelatorioPdf_('Estoque', 'Essência do Brasil — Estoque Atual', ['Produto', 'Quantidade'], linhas);
}

function gerarRelatorioPdf_(nomeArquivo, titulo, cabecalho, linhas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAbaTemp = 'PDF_TEMP';

  let abaTemp = planilha.getSheetByName(nomeAbaTemp);
  if (abaTemp) planilha.deleteSheet(abaTemp);
  abaTemp = planilha.insertSheet(nomeAbaTemp);

  abaTemp
    .getRange(1, 1, 1, cabecalho.length)
    .merge()
    .setValue(titulo)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0f172a')
    .setHorizontalAlignment('center');

  abaTemp
    .getRange(2, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');

  if (linhas.length > 0) {
    abaTemp.getRange(3, 1, linhas.length, cabecalho.length).setValues(linhas);
  }
  abaTemp.autoResizeColumns(1, cabecalho.length);

  SpreadsheetApp.flush();
  const blobPdf = exportarAbaComoPdf_(planilha, abaTemp);
  planilha.deleteSheet(abaTemp);

  const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  blobPdf.setName(nomeArquivo + '_' + carimbo + '.pdf');

  return obterOuCriarPastaRelatorios_().createFile(blobPdf);
}

function exportarAbaComoPdf_(planilha, aba) {
  const url =
    planilha.getUrl().replace(/edit.*$/, '') +
    'export?format=pdf&gid=' +
    aba.getSheetId() +
    '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false';
  const token = ScriptApp.getOAuthToken();
  const resposta = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  return resposta.getBlob();
}

function obterOuCriarPastaRelatorios_() {
  const pastas = DriveApp.getFoldersByName(NOME_PASTA_RELATORIOS);
  return pastas.hasNext() ? pastas.next() : DriveApp.createFolder(NOME_PASTA_RELATORIOS);
}
