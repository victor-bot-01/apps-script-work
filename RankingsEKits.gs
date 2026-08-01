/**
 * Rankings e análise kit vs. individual.
 *
 * Ao contrário do Dashboard (30 dias) e do Estoque Inteligente (12
 * semanas), aqui os rankings de produto/kit usam TODO o histórico da aba
 * Pedidos — "kit mais vendido" ou "produto que mais participa de kits" só
 * fazem sentido olhando pra tudo que já foi vendido, não uma janela curta.
 *
 * "Crescimento vs. período anterior" já é period-based por natureza: usa
 * janelas móveis (últimos N dias vs. os N dias antes disso) pra semana,
 * mês e últimos 12 meses — assim o número é comparável não importa o dia
 * em que a planilha é aberta. A exceção é "mesmo mês, ano anterior", que
 * compara do dia 1 do mês até hoje, este ano vs. mesmo intervalo ano
 * passado (comparação de calendário, não janela móvel).
 */

const NOME_ABA_RANKINGS = 'Rankings';
const TAMANHO_TOP_RANKING = 10;

function atualizarRankings() {
  const vendas = obterVendasResolvidas_();
  const dadosKitIndividual = agruparKitIndividualPorProduto_(vendas);

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_RANKINGS);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_RANKINGS);
  aba.clear();

  let linha = 1;
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Top ' + TAMANHO_TOP_RANKING + ' Produtos',
    ['Produto', 'Quantidade'],
    calcularTopProdutos_(vendas)
  );
  linha = escreverSecaoTabela_(aba, linha, 'Ranking de Kits', ['Kit', 'Quantidade'], calcularRankingDeKits_(vendas));
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Crescimento vs. Período Anterior',
    ['Período', 'Atual', 'Anterior', 'Variação'],
    calcularCrescimentoPeriodos_()
  );
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Semanas do Mês com Mais Vendas (mês atual)',
    ['Semana do Mês', 'Quantidade'],
    calcularSemanasDoMes_()
  );
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Kit vs. Individual por Produto',
    ['Produto', 'Vendido em Kit', 'Vendido Individual', 'Predominância'],
    calcularKitVsIndividualPorProduto_(dadosKitIndividual)
  );
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Produtos Só-Kit / Só-Individual / Ambos',
    ['Produto', 'Classificação'],
    calcularClassificacaoKitIndividual_(dadosKitIndividual)
  );
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Top 5 Produtos que Mais Participam de Kits',
    ['Produto', 'Quantidade em Kits'],
    calcularTopParticipacaoEmKits_(dadosKitIndividual)
  );
  escreverSecaoTabela_(
    aba,
    linha,
    'Kit Mais e Menos Vendido',
    ['Destaque', 'Kit', 'Quantidade'],
    calcularExtremosDeKits_(vendas)
  );

  aba.autoResizeColumns(1, 4);
}

// Cada linha de venda resolvida a produto(s) canônico(s) via Base_Dados;
// linhas sem correspondência ou fora de escopo (ex.: joia) já saem daqui.
function obterVendasResolvidas_() {
  const abaPedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_PEDIDOS);
  if (!abaPedidos) throw new Error('Aba "' + NOME_ABA_PEDIDOS + '" não encontrada.');

  const ultimaLinha = abaPedidos.getLastRow();
  if (ultimaLinha < 2) return [];

  const cabecalho = abaPedidos.getRange(1, 1, 1, abaPedidos.getLastColumn()).getValues()[0];
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_PEDIDOS);
  const colDescricao = obterIndiceColuna_(cabecalho, 'Descrição', NOME_ABA_PEDIDOS);

  const indiceBaseDados = construirIndiceBaseDados_();

  return abaPedidos
    .getRange(2, 1, ultimaLinha - 1, abaPedidos.getLastColumn())
    .getValues()
    .map(function (linha) {
      const descricao = linha[colDescricao];
      const chave = normalizarTitulo_(descricao);
      const produtos = indiceBaseDados.has(chave) ? indiceBaseDados.get(chave) : [];
      return {
        quantidade: Number(linha[colQuantidade]) || 0,
        chaveTitulo: chave,
        produtos: produtos,
        isKit: produtos.length > 1,
      };
    })
    .filter(function (venda) {
      return venda.produtos.length > 0;
    });
}

function calcularTopProdutos_(vendas) {
  const quantidadePorProduto = new Map();
  vendas.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      quantidadePorProduto.set(produto, (quantidadePorProduto.get(produto) || 0) + venda.quantidade);
    });
  });

  return Array.from(quantidadePorProduto.entries())
    .sort(function (a, b) {
      return b[1] - a[1];
    })
    .slice(0, TAMANHO_TOP_RANKING);
}

function calcularRankingDeKits_(vendas) {
  const porKit = new Map();
  vendas
    .filter(function (venda) {
      return venda.isKit;
    })
    .forEach(function (venda) {
      if (!porKit.has(venda.chaveTitulo)) {
        porKit.set(venda.chaveTitulo, { rotulo: venda.produtos.join(' + '), quantidade: 0 });
      }
      porKit.get(venda.chaveTitulo).quantidade += venda.quantidade;
    });

  return Array.from(porKit.values())
    .sort(function (a, b) {
      return b.quantidade - a.quantidade;
    })
    .map(function (kit) {
      return [kit.rotulo, kit.quantidade];
    });
}

function agruparKitIndividualPorProduto_(vendas) {
  const dados = new Map();
  vendas.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      if (!dados.has(produto)) dados.set(produto, { kit: 0, individual: 0 });
      const registro = dados.get(produto);
      if (venda.isKit) registro.kit += venda.quantidade;
      else registro.individual += venda.quantidade;
    });
  });
  return dados;
}

function calcularKitVsIndividualPorProduto_(dadosPorProduto) {
  return Array.from(dadosPorProduto.entries())
    .map(function (par) {
      const produto = par[0];
      const registro = par[1];
      return [produto, registro.kit, registro.individual, calcularPredominancia_(registro)];
    })
    .sort(function (a, b) {
      return b[1] + b[2] - (a[1] + a[2]);
    });
}

function calcularPredominancia_(registro) {
  if (registro.kit > 0 && registro.individual === 0) return 'Só Kit';
  if (registro.individual > 0 && registro.kit === 0) return 'Só Individual';
  if (registro.kit === registro.individual) return 'Empate';
  return registro.kit > registro.individual ? 'Mais em Kit' : 'Mais Individual';
}

function calcularClassificacaoKitIndividual_(dadosPorProduto) {
  return Array.from(dadosPorProduto.entries())
    .map(function (par) {
      const produto = par[0];
      const registro = par[1];
      let classificacao;
      if (registro.kit > 0 && registro.individual === 0) classificacao = 'Só Kit';
      else if (registro.individual > 0 && registro.kit === 0) classificacao = 'Só Individual';
      else classificacao = 'Ambos';
      return [produto, classificacao];
    })
    .sort(function (a, b) {
      return a[0].localeCompare(b[0]);
    });
}

function calcularTopParticipacaoEmKits_(dadosPorProduto) {
  return Array.from(dadosPorProduto.entries())
    .filter(function (par) {
      return par[1].kit > 0;
    })
    .map(function (par) {
      return [par[0], par[1].kit];
    })
    .sort(function (a, b) {
      return b[1] - a[1];
    })
    .slice(0, 5);
}

function calcularExtremosDeKits_(vendas) {
  const ranking = calcularRankingDeKits_(vendas);
  if (ranking.length === 0) return [];

  const maisVendido = ranking[0];
  const menosVendido = ranking[ranking.length - 1];
  return [
    ['Mais vendido', maisVendido[0], maisVendido[1]],
    ['Menos vendido', menosVendido[0], menosVendido[1]],
  ];
}

function lerTodosPedidosBrutos_() {
  const abaPedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_PEDIDOS);
  if (!abaPedidos) return [];

  const ultimaLinha = abaPedidos.getLastRow();
  if (ultimaLinha < 2) return [];

  const cabecalho = abaPedidos.getRange(1, 1, 1, abaPedidos.getLastColumn()).getValues()[0];
  const colData = obterIndiceColuna_(cabecalho, 'Data', NOME_ABA_PEDIDOS);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_PEDIDOS);

  return abaPedidos
    .getRange(2, 1, ultimaLinha - 1, abaPedidos.getLastColumn())
    .getValues()
    .map(function (linha) {
      return { data: linha[colData], quantidade: Number(linha[colQuantidade]) || 0 };
    })
    .filter(function (item) {
      return item.data instanceof Date;
    });
}

function somarQuantidadeNoIntervalo_(pedidosBrutos, dataInicio, dataFimExclusiva) {
  return pedidosBrutos.reduce(function (soma, item) {
    return item.data >= dataInicio && item.data < dataFimExclusiva ? soma + item.quantidade : soma;
  }, 0);
}

function subtrairDias_(data, dias) {
  return new Date(data.getFullYear(), data.getMonth(), data.getDate() - dias);
}

function calcularCrescimentoPeriodos_() {
  const pedidosBrutos = lerTodosPedidosBrutos_();
  const hoje = new Date();
  const diaDeHoje = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());

  const linhas = [];

  linhas.push(
    construirLinhaCrescimento_(
      pedidosBrutos,
      'Semana (últimos 7 dias)',
      subtrairDias_(diaDeHoje, 7),
      diaDeHoje,
      subtrairDias_(diaDeHoje, 14),
      subtrairDias_(diaDeHoje, 7)
    )
  );
  linhas.push(
    construirLinhaCrescimento_(
      pedidosBrutos,
      'Mês (últimos 30 dias)',
      subtrairDias_(diaDeHoje, 30),
      diaDeHoje,
      subtrairDias_(diaDeHoje, 60),
      subtrairDias_(diaDeHoje, 30)
    )
  );

  const inicioMesAtual = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const inicioMesAnoAnterior = new Date(hoje.getFullYear() - 1, hoje.getMonth(), 1);
  const fimMesAnoAnterior = new Date(hoje.getFullYear() - 1, hoje.getMonth(), hoje.getDate());
  linhas.push(
    construirLinhaCrescimento_(
      pedidosBrutos,
      'Mesmo mês, ano anterior',
      inicioMesAtual,
      diaDeHoje,
      inicioMesAnoAnterior,
      fimMesAnoAnterior
    )
  );

  linhas.push(
    construirLinhaCrescimento_(
      pedidosBrutos,
      'Últimos 12 meses',
      subtrairDias_(diaDeHoje, 365),
      diaDeHoje,
      subtrairDias_(diaDeHoje, 730),
      subtrairDias_(diaDeHoje, 365)
    )
  );

  return linhas;
}

function construirLinhaCrescimento_(pedidosBrutos, rotulo, inicioAtual, fimAtual, inicioAnterior, fimAnterior) {
  const atual = somarQuantidadeNoIntervalo_(pedidosBrutos, inicioAtual, fimAtual);
  const anterior = somarQuantidadeNoIntervalo_(pedidosBrutos, inicioAnterior, fimAnterior);
  return [rotulo, atual, anterior, formatarVariacao_(atual, anterior)];
}

function formatarVariacao_(atual, anterior) {
  if (anterior === 0) return atual > 0 ? 'Novo (sem base anterior)' : '0%';
  const variacao = ((atual - anterior) / anterior) * 100;
  return (variacao >= 0 ? '+' : '') + Math.round(variacao * 10) / 10 + '%';
}

function calcularSemanasDoMes_() {
  const pedidosBrutos = lerTodosPedidosBrutos_();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const fimMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 1);

  const totaisPorSemana = [0, 0, 0, 0, 0];
  pedidosBrutos.forEach(function (item) {
    if (item.data < inicioMes || item.data >= fimMes) return;
    const indiceSemana = Math.min(Math.floor((item.data.getDate() - 1) / 7), 4);
    totaisPorSemana[indiceSemana] += item.quantidade;
  });

  return totaisPorSemana
    .map(function (quantidade, indice) {
      return ['Semana ' + (indice + 1), quantidade];
    })
    .sort(function (a, b) {
      return b[1] - a[1];
    });
}

function escreverSecaoTabela_(aba, linhaInicial, titulo, cabecalhos, linhas) {
  aba
    .getRange(linhaInicial, 1, 1, cabecalhos.length)
    .merge()
    .setValue(titulo)
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0f172a');

  aba
    .getRange(linhaInicial + 1, 1, 1, cabecalhos.length)
    .setValues([cabecalhos])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');

  if (linhas.length > 0) {
    aba.getRange(linhaInicial + 2, 1, linhas.length, cabecalhos.length).setValues(linhas);
  }

  return linhaInicial + 2 + linhas.length + 1;
}
