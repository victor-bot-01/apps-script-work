/**
 * Dashboard executivo: cards de KPI + gráficos, calculados a partir de
 * Pedidos + Histórico + Base_Dados (matching) + Estoque + Ficha Técnica.
 *
 * `atualizarDashboard` é chamado automaticamente ao fim de cada importação
 * bem-sucedida (ver ImportacaoPedidos.gs) e também pelo menu manual.
 *
 * "Produtos em risco de ruptura" usa o cálculo de níveis do módulo
 * Estoque Inteligente (ver EstoqueInteligente.gs), não um cálculo próprio.
 *
 * O período (dias corridos ou mês fechado) é resolvido por
 * `lerPedidosNoPeriodo_` (ver LeituraVendas.gs), que já combina "Pedidos"
 * + "Histórico" — o período pedido (ex.: 30 dias, ou "mês anterior") pode
 * cruzar a virada de mês, especialmente nos primeiros dias de cada mês,
 * logo depois do arquivamento.
 */

const NOME_ABA_DASHBOARD = 'Dashboard';
const NOME_ABA_ESTOQUE = 'Estoque';
const NOME_ABA_FICHA_TECNICA = 'Ficha Técnica';
const DIAS_PERIODO_KPI = 30;

function atualizarDashboard() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaDashboard = obterOuCriarAbaDashboard_(planilha);
  const kpis = calcularKPIs_();

  escreverCardsKPI_(abaDashboard, kpis);
  const rangesGraficos = escreverDadosGraficos_(abaDashboard, kpis);
  criarOuAtualizarGraficos_(abaDashboard, rangesGraficos);
}

function calcularKPIs_(periodo) {
  periodo = periodo || DIAS_PERIODO_KPI;
  const pedidos = lerPedidosNoPeriodo_(periodo);
  const indiceBaseDados = construirIndiceBaseDados_();

  const pedidosUnicos = new Set();
  let itensVendidos = 0;
  let kits = 0;
  let individuais = 0;
  const quantidadePorProduto = new Map();
  const quantidadePorLoja = new Map();
  const produtosVendidos = new Set();

  pedidos.forEach(function (pedido) {
    pedidosUnicos.add(String(pedido.pedido).trim().toLowerCase());
    itensVendidos += pedido.quantidade;
    quantidadePorLoja.set(pedido.loja, (quantidadePorLoja.get(pedido.loja) || 0) + pedido.quantidade);

    const chave = normalizarTitulo_(pedido.descricao);
    const produtos = indiceBaseDados.has(chave) ? indiceBaseDados.get(chave) : [];
    if (produtos.length === 0) return; // sem correspondência ou fora de escopo (ex.: joia)

    if (produtos.length > 1) kits += pedido.quantidade;
    else individuais += pedido.quantidade;

    produtos.forEach(function (produto) {
      produtosVendidos.add(produto);
      quantidadePorProduto.set(produto, (quantidadePorProduto.get(produto) || 0) + pedido.quantidade);
    });
  });

  const catalogo = obterCatalogoDeProdutosCanonicos_(indiceBaseDados);
  const produtosSemVenda = catalogo.filter(function (produto) {
    return !produtosVendidos.has(produto);
  });

  const estoque = lerEstoque_();
  const estoqueTotal = Array.from(estoque.values()).reduce(function (soma, quantidade) {
    return soma + quantidade;
  }, 0);
  const produtosEmRiscoDeRuptura = obterProdutosEmRiscoDeRuptura_();

  const rankingProdutos = Array.from(quantidadePorProduto.entries()).sort(function (a, b) {
    return b[1] - a[1];
  });
  const rankingLojas = Array.from(quantidadePorLoja.entries()).sort(function (a, b) {
    return b[1] - a[1];
  });

  return {
    rotuloPeriodo: rotuloPeriodo_(periodo),
    pedidosNoPeriodo: pedidosUnicos.size,
    itensVendidos: itensVendidos,
    produtosAtivos: produtosVendidos.size,
    produtosSemVenda: produtosSemVenda.length,
    estoqueTotal: estoqueTotal,
    produtosEmRiscoDeRuptura: produtosEmRiscoDeRuptura.length,
    kits: kits,
    individuais: individuais,
    produtoCampeao: obterChaveComMaiorValor_(quantidadePorProduto) || '—',
    marketplaceCampeao: obterChaveComMaiorValor_(quantidadePorLoja) || '—',
    rankingProdutos: rankingProdutos.slice(0, 10),
    rankingLojas: rankingLojas,
  };
}

// Só produtos individuais entram no total (kits ficam de fora, ver
// Producao.gs) — a coluna "Tipo" pode nem existir ainda (migração feita
// sob demanda por Producao.gs), então ausência de "Tipo" também conta
// como Individual, igual ao comportamento anterior a essa coluna existir.
function lerEstoque_() {
  const mapa = new Map();
  const abaEstoque = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_ESTOQUE);
  if (!abaEstoque) return mapa;

  const ultimaLinha = abaEstoque.getLastRow();
  if (ultimaLinha < 2) return mapa;

  const cabecalho = abaEstoque.getRange(1, 1, 1, abaEstoque.getLastColumn()).getValues()[0];
  const colProduto = obterIndiceColuna_(cabecalho, 'Produto (canônico)', NOME_ABA_ESTOQUE);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_ESTOQUE);
  const colTipo = cabecalho.indexOf('Tipo');

  abaEstoque
    .getRange(2, 1, ultimaLinha - 1, abaEstoque.getLastColumn())
    .getValues()
    .forEach(function (linha) {
      const produto = linha[colProduto];
      if (!produto) return;
      const tipo = colTipo !== -1 && linha[colTipo] ? linha[colTipo] : 'Individual';
      if (tipo === 'Kit') return;
      mapa.set(produto, Number(linha[colQuantidade]) || 0);
    });

  return mapa;
}

// Prioriza a Ficha Técnica como catálogo mestre (é a intenção do prompt);
// cai para os produtos distintos do Base_Dados enquanto a Ficha Técnica
// ainda não estiver preenchida.
function obterCatalogoDeProdutosCanonicos_(indiceBaseDados) {
  const abaFicha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_FICHA_TECNICA);
  if (abaFicha && abaFicha.getLastRow() >= 2) {
    const cabecalho = abaFicha.getRange(1, 1, 1, abaFicha.getLastColumn()).getValues()[0];
    const colProduto = obterIndiceColuna_(cabecalho, 'Produto (canônico)', NOME_ABA_FICHA_TECNICA);
    const produtos = abaFicha
      .getRange(2, 1, abaFicha.getLastRow() - 1, abaFicha.getLastColumn())
      .getValues()
      .map(function (linha) {
        return linha[colProduto];
      })
      .filter(function (produto) {
        return produto;
      });
    if (produtos.length > 0) return Array.from(new Set(produtos));
  }

  const catalogo = new Set();
  indiceBaseDados.forEach(function (produtos) {
    produtos.forEach(function (produto) {
      catalogo.add(produto);
    });
  });
  return Array.from(catalogo);
}

function obterChaveComMaiorValor_(mapa) {
  let melhorChave = null;
  let melhorValor = -Infinity;
  mapa.forEach(function (valor, chave) {
    if (valor > melhorValor) {
      melhorValor = valor;
      melhorChave = chave;
    }
  });
  return melhorChave;
}

function obterOuCriarAbaDashboard_(planilha) {
  let aba = planilha.getSheetByName(NOME_ABA_DASHBOARD);
  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA_DASHBOARD, 0);
  }
  aba.setTabColor('#0f172a');
  return aba;
}

function escreverCardsKPI_(aba, kpis) {
  aba.clear();

  aba
    .getRange(1, 1, 1, 12)
    .merge()
    .setValue('Essência do Brasil — Painel de Vendas (' + kpis.rotuloPeriodo + ')')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0f172a')
    .setHorizontalAlignment('center');
  aba.setRowHeight(1, 40);

  const cards = [
    { titulo: 'Pedidos no período', valor: kpis.pedidosNoPeriodo, cor: '#1d4ed8' },
    { titulo: 'Itens vendidos', valor: kpis.itensVendidos, cor: '#1d4ed8' },
    { titulo: 'Produtos ativos / sem venda', valor: kpis.produtosAtivos + ' / ' + kpis.produtosSemVenda, cor: '#15803d' },
    { titulo: 'Estoque total', valor: kpis.estoqueTotal, cor: '#15803d' },
    { titulo: 'Produtos em risco de ruptura', valor: kpis.produtosEmRiscoDeRuptura, cor: '#c2410c' },
    { titulo: 'Kits vs. individuais', valor: kpis.kits + ' / ' + kpis.individuais, cor: '#1d4ed8' },
    { titulo: 'Produto campeão', valor: kpis.produtoCampeao, cor: '#15803d' },
    { titulo: 'Marketplace campeão', valor: kpis.marketplaceCampeao, cor: '#c2410c' },
  ];

  const linhaInicial = 3;
  const alturaCard = 2;
  const larguraCard = 3;
  const espacoEntreLinhas = 1;

  cards.forEach(function (card, indice) {
    const linha = linhaInicial + Math.floor(indice / 4) * (alturaCard + espacoEntreLinhas);
    const coluna = 1 + (indice % 4) * larguraCard;
    escreverCard_(aba, linha, coluna, larguraCard, card.titulo, card.valor, card.cor);
  });

  aba.setColumnWidths(1, 12, 110);
}

function escreverCard_(aba, linha, coluna, largura, titulo, valor, cor) {
  aba
    .getRange(linha, coluna, 1, largura)
    .merge()
    .setValue(titulo)
    .setFontSize(9)
    .setFontColor('#e2e8f0')
    .setBackground(cor)
    .setHorizontalAlignment('center')
    .setWrap(true);
  aba.setRowHeight(linha, 24);

  aba
    .getRange(linha + 1, coluna, 1, largura)
    .merge()
    .setValue(valor)
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground(cor)
    .setHorizontalAlignment('center');
  aba.setRowHeight(linha + 1, 40);
}

function escreverDadosGraficos_(aba, kpis) {
  const linhaProdutos = 10;
  aba.getRange(linhaProdutos, 1).setValue('Top 10 produtos (dados do gráfico)').setFontWeight('bold');
  const dadosProdutos = [['Produto', 'Quantidade']].concat(kpis.rankingProdutos);
  aba.getRange(linhaProdutos + 1, 1, dadosProdutos.length, 2).setValues(dadosProdutos);

  const linhaLojas = linhaProdutos + dadosProdutos.length + 2;
  aba.getRange(linhaLojas, 1).setValue('Vendas por marketplace (dados do gráfico)').setFontWeight('bold');
  const dadosLojas = [['Loja', 'Quantidade']].concat(kpis.rankingLojas);
  aba.getRange(linhaLojas + 1, 1, dadosLojas.length, 2).setValues(dadosLojas);

  return {
    produtos: aba.getRange(linhaProdutos + 1, 1, dadosProdutos.length, 2),
    lojas: aba.getRange(linhaLojas + 1, 1, dadosLojas.length, 2),
  };
}

function criarOuAtualizarGraficos_(aba, ranges) {
  aba.getCharts().forEach(function (grafico) {
    aba.removeChart(grafico);
  });

  aba.insertChart(
    aba
      .newChart()
      .asColumnChart()
      .addRange(ranges.produtos)
      .setNumHeaders(1)
      .setPosition(3, 14, 0, 0)
      .setOption('title', 'Top 10 produtos (unidades)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#1d4ed8'])
      .build()
  );

  aba.insertChart(
    aba
      .newChart()
      .asPieChart()
      .addRange(ranges.lojas)
      .setNumHeaders(1)
      .setPosition(20, 14, 0, 0)
      .setOption('title', 'Vendas por marketplace')
      .build()
  );
}
