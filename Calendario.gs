/**
 * Calendário/sazonalidade: heatmap de dias com mais vendas (últimos 90
 * dias), vendas por dia da semana, classificação de produtos (em
 * crescimento/queda/voltou a vender/sem vendas recentes) e contagem de
 * produtos sem venda há 30/60/90/180 dias.
 *
 * Não implementei uma classificação de "produto sazonal" de verdade —
 * isso exige pelo menos uns 12 meses de histórico ano a ano pra não virar
 * chute, e o sistema acabou de começar a acumular dados agora. Quando
 * houver histórico suficiente, dá pra evoluir isso.
 */

const NOME_ABA_CALENDARIO = 'Calendário';
const DIAS_HEATMAP = 90;
const LIMIAR_CRESCIMENTO_PRODUTO = 20;
const LIMIAR_QUEDA_PRODUTO = -20;
const DIAS_SEM_VENDA_LIMIARES = [30, 60, 90, 180];

// Agregador puro usado pelo Web App (WebAppApi.gs).
function calcularCalendario() {
  const pedidosBrutos = lerTodosPedidosBrutos_();
  const eventosPorProduto = obterVendasPorProduto_(400);
  const classificacao = calcularClassificacaoDeProdutos_(eventosPorProduto);

  return {
    heatmapDias: calcularHeatmapDeDias_(pedidosBrutos),
    porDiaDaSemana: calcularVendasPorDiaDaSemana_(pedidosBrutos),
    classificacaoProdutos: classificacao,
    contagemSemVenda: calcularContagemSemVenda_(classificacao),
  };
}

// Utilitário manual (menu "Essência do Brasil"): grava um retrato do
// calendário numa aba. Não é chamado automaticamente pela importação
// diária — a planilha "Análise e Controle" é só banco de dados (ver
// prompt.md); a visão viva é o dashboard (Web App).
function atualizarCalendario() {
  const pedidosBrutos = lerTodosPedidosBrutos_();
  const eventosPorProduto = obterVendasPorProduto_(400);

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_CALENDARIO);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_CALENDARIO);
  aba.clear();
  aba.clearConditionalFormatRules();

  let linha = 1;
  const linhaInicioHeatmap = linha;
  const dadosHeatmap = calcularHeatmapDeDias_(pedidosBrutos);
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Vendas por Dia (últimos ' + DIAS_HEATMAP + ' dias)',
    ['Data', 'Quantidade'],
    dadosHeatmap
  );
  if (dadosHeatmap.length > 0) {
    aplicarEscalaDeCor_(aba, linhaInicioHeatmap + 2, dadosHeatmap.length);
  }

  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Vendas por Dia da Semana',
    ['Dia da Semana', 'Quantidade'],
    calcularVendasPorDiaDaSemana_(pedidosBrutos)
  );

  const classificacao = calcularClassificacaoDeProdutos_(eventosPorProduto);
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Classificação de Produtos (últimos 30d vs. 30-60 dias atrás)',
    ['Produto', 'Últimos 30d', '30-60d Atrás', 'Variação %', 'Última Venda', 'Dias Sem Venda', 'Classificação'],
    classificacao
  );

  escreverSecaoTabela_(
    aba,
    linha,
    'Produtos Sem Venda por Limiar',
    ['Limiar', 'Quantidade de Produtos'],
    calcularContagemSemVenda_(classificacao)
  );

  aba.autoResizeColumns(1, 7);
}

function calcularHeatmapDeDias_(pedidosBrutos) {
  const hoje = new Date();
  const inicio = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - DIAS_HEATMAP + 1);

  const totaisPorDia = new Map();
  for (let i = 0; i < DIAS_HEATMAP; i++) {
    const dia = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate() + i);
    totaisPorDia.set(formatarDataISO_(dia), 0);
  }

  pedidosBrutos.forEach(function (item) {
    const chave = formatarDataISO_(item.data);
    if (totaisPorDia.has(chave)) totaisPorDia.set(chave, totaisPorDia.get(chave) + item.quantidade);
  });

  return Array.from(totaisPorDia.entries());
}

function aplicarEscalaDeCor_(aba, linhaInicioDados, quantidadeLinhas) {
  const faixa = aba.getRange(linhaInicioDados, 2, quantidadeLinhas, 1);
  const regra = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint('#f1f5f9')
    .setGradientMaxpoint('#1d4ed8')
    .setRanges([faixa])
    .build();
  aba.setConditionalFormatRules(aba.getConditionalFormatRules().concat([regra]));
}

function calcularVendasPorDiaDaSemana_(pedidosBrutos) {
  const nomes = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
  const totais = [0, 0, 0, 0, 0, 0, 0];
  pedidosBrutos.forEach(function (item) {
    totais[item.data.getDay()] += item.quantidade;
  });
  return nomes
    .map(function (nome, indice) {
      return [nome, totais[indice]];
    })
    .sort(function (a, b) {
      return b[1] - a[1];
    });
}

function calcularClassificacaoDeProdutos_(eventosPorProduto) {
  const linhas = [];

  eventosPorProduto.forEach(function (eventos, produto) {
    const vendidoUltimos30 = somarQuantidadeDesde_(eventos, 30);
    const vendidoAnterior30 = somarQuantidadeEntreDias_(eventos, 30, 60);

    let ultimaVenda = null;
    eventos.forEach(function (evento) {
      if (evento.quantidade > 0 && (ultimaVenda === null || evento.data > ultimaVenda)) ultimaVenda = evento.data;
    });
    const diasSemVenda = ultimaVenda ? Math.floor((new Date() - ultimaVenda) / 86400000) : null;

    let classificacao;
    let variacao = null;
    if (vendidoUltimos30 === 0 && vendidoAnterior30 === 0) {
      classificacao = 'Sem Vendas Recentes';
    } else if (vendidoAnterior30 === 0 && vendidoUltimos30 > 0) {
      classificacao = 'Voltou a Vender';
    } else {
      variacao = ((vendidoUltimos30 - vendidoAnterior30) / vendidoAnterior30) * 100;
      if (variacao >= LIMIAR_CRESCIMENTO_PRODUTO) classificacao = 'Em Crescimento';
      else if (variacao <= LIMIAR_QUEDA_PRODUTO) classificacao = 'Em Queda';
      else classificacao = 'Estável';
    }

    linhas.push([
      produto,
      vendidoUltimos30,
      vendidoAnterior30,
      variacao !== null ? Math.round(variacao) + '%' : '—',
      ultimaVenda ? formatarDataISO_(ultimaVenda) : '—',
      diasSemVenda !== null ? diasSemVenda : '—',
      classificacao,
    ]);
  });

  return linhas.sort(function (a, b) {
    return String(a[0]).localeCompare(String(b[0]));
  });
}

function calcularContagemSemVenda_(classificacao) {
  return DIAS_SEM_VENDA_LIMIARES.map(function (limiar) {
    const quantidade = classificacao.filter(function (linha) {
      return typeof linha[5] === 'number' && linha[5] >= limiar;
    }).length;
    return [limiar + '+ dias sem venda', quantidade];
  });
}
