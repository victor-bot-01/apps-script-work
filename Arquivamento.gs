/**
 * Arquivamento de "Pedidos": no início de cada execução diária, antes de
 * importar pedidos novos (ver ImportacaoPedidos.gs), verifica se a aba
 * "Pedidos" tem linhas de um mês já fechado e as move para "Histórico".
 * Também expurga de "Histórico" linhas com mais de ANOS_RETENCAO_HISTORICO
 * anos, resumindo-as antes (por Produto canônico + Ano-Mês + Loja) na aba
 * "Resumo Histórico" — assim não se perde a noção de tendência de longo
 * prazo mesmo depois de descartar o pedido detalhado.
 *
 * "Pedidos" fica enxuta (só mês corrente); "Histórico" é a memória de
 * vendas de até 2 anos que alimenta os módulos de análise via
 * lerPedidosCombinados_ (ver LeituraVendas.gs); "Resumo Histórico" é a
 * memória de longuíssimo prazo, sem granularidade de pedido/kit/loja por
 * linha — só Produto + Ano-Mês + Loja com a soma vendida.
 *
 * Ambas as etapas são baratas de rodar todo dia (viram no-op na maioria
 * dos dias) e não precisam de gatilho separado — ver a chamada em
 * executarImportacaoComTratamento_ (ImportacaoPedidos.gs).
 */

const NOME_ABA_HISTORICO = 'Histórico';
const NOME_ABA_RESUMO_HISTORICO = 'Resumo Histórico';
const ANOS_RETENCAO_HISTORICO = 2;

// Chamado pelo pipeline de importação diária. Retorna uma lista de avisos
// (string) prontos pra entrar no Log de Execução — vazia se não havia
// nada pra arquivar/expurgar hoje.
function executarArquivamento_() {
  const avisos = [];

  const totalArquivado = arquivarPedidosDoMesFechado_();
  if (totalArquivado > 0) {
    avisos.push(
      'Arquivamento: movida(s) ' + totalArquivado + ' linha(s) de mês(es) fechado(s) de "Pedidos" para "Histórico".'
    );
  }

  const resultadoExpurgo = expurgarHistoricoAntigo_();
  if (resultadoExpurgo.linhasExpurgadas > 0) {
    avisos.push(
      'Expurgo: removida(s) ' +
        resultadoExpurgo.linhasExpurgadas +
        ' linha(s) de "Histórico" com mais de ' +
        ANOS_RETENCAO_HISTORICO +
        ' ano(s), resumida(s) em "Resumo Histórico" (' +
        resultadoExpurgo.chavesAtualizadas +
        ' chave(s) produto/mês/loja atualizada(s)).'
    );
  }

  return avisos;
}

// Move para "Histórico" toda linha de "Pedidos" cujo mês (coluna Data)
// não é o mês corrente. Linha sem data válida é mantida em "Pedidos" por
// segurança (evita perder um registro num "pra onde vai" ambíguo).
// Retorna a quantidade de linhas movidas (0 = nada a fazer hoje).
function arquivarPedidosDoMesFechado_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaPedidos = planilha.getSheetByName(NOME_ABA_PEDIDOS);
  const abaHistorico = planilha.getSheetByName(NOME_ABA_HISTORICO);
  if (!abaPedidos || !abaHistorico) {
    throw new Error('Abas "Pedidos"/"Histórico" não encontradas. Rode setupEstrutura() primeiro.');
  }

  const ultimaLinha = abaPedidos.getLastRow();
  if (ultimaLinha < 2) return 0;

  const cabecalho = abaPedidos.getRange(1, 1, 1, abaPedidos.getLastColumn()).getValues()[0];
  const colData = obterIndiceColuna_(cabecalho, 'Data', NOME_ABA_PEDIDOS);
  const mesAtualChave = formatarAnoMes_(new Date());

  const todasAsLinhas = abaPedidos.getRange(2, 1, ultimaLinha - 1, abaPedidos.getLastColumn()).getValues();
  const linhasParaManter = [];
  const linhasParaArquivar = [];

  todasAsLinhas.forEach(function (linha) {
    const data = linha[colData];
    if (!(data instanceof Date) || formatarAnoMes_(data) === mesAtualChave) {
      linhasParaManter.push(linha);
    } else {
      linhasParaArquivar.push(linha);
    }
  });

  if (linhasParaArquivar.length === 0) return 0;

  abaHistorico
    .getRange(abaHistorico.getLastRow() + 1, 1, linhasParaArquivar.length, linhasParaArquivar[0].length)
    .setValues(linhasParaArquivar);

  reescreverCorpoDaAba_(abaPedidos, cabecalho.length, linhasParaManter);

  return linhasParaArquivar.length;
}

// Remove de "Histórico" toda linha com mais de ANOS_RETENCAO_HISTORICO
// anos (janela móvel), somando antes por Produto canônico + Ano-Mês +
// Loja na aba "Resumo Histórico". Um título sem correspondência no
// Base_Dados (ou fora de escopo, ex.: joia) ainda é expurgado de
// "Histórico" — a política de retenção é por idade, não por matching —
// só não contribui pro resumo, já que não há produto canônico pra somar.
function expurgarHistoricoAntigo_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaHistorico = planilha.getSheetByName(NOME_ABA_HISTORICO);
  const abaResumo = planilha.getSheetByName(NOME_ABA_RESUMO_HISTORICO);
  if (!abaHistorico || !abaResumo) {
    throw new Error('Abas "Histórico"/"Resumo Histórico" não encontradas. Rode setupEstrutura() primeiro.');
  }

  const ultimaLinha = abaHistorico.getLastRow();
  if (ultimaLinha < 2) return { linhasExpurgadas: 0, chavesAtualizadas: 0 };

  const cabecalho = abaHistorico.getRange(1, 1, 1, abaHistorico.getLastColumn()).getValues()[0];
  const colData = obterIndiceColuna_(cabecalho, 'Data', NOME_ABA_HISTORICO);
  const colLoja = obterIndiceColuna_(cabecalho, 'Loja', NOME_ABA_HISTORICO);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_HISTORICO);
  const colDescricao = obterIndiceColuna_(cabecalho, 'Descrição', NOME_ABA_HISTORICO);

  const dataCorte = new Date();
  dataCorte.setFullYear(dataCorte.getFullYear() - ANOS_RETENCAO_HISTORICO);

  const todasAsLinhas = abaHistorico.getRange(2, 1, ultimaLinha - 1, abaHistorico.getLastColumn()).getValues();
  const linhasParaManter = [];
  const linhasParaExpurgar = [];

  todasAsLinhas.forEach(function (linha) {
    const data = linha[colData];
    if (data instanceof Date && data < dataCorte) {
      linhasParaExpurgar.push(linha);
    } else {
      linhasParaManter.push(linha);
    }
  });

  if (linhasParaExpurgar.length === 0) return { linhasExpurgadas: 0, chavesAtualizadas: 0 };

  const indiceBaseDados = construirIndiceBaseDados_();
  const somaPorChave = new Map(); // "produto||anoMes||loja" -> {produto, anoMes, loja, quantidade}

  linhasParaExpurgar.forEach(function (linha) {
    const data = linha[colData];
    const loja = linha[colLoja];
    const quantidade = Number(linha[colQuantidade]) || 0;
    const anoMes = formatarAnoMes_(data);

    const chaveTitulo = normalizarTitulo_(linha[colDescricao]);
    const produtos = indiceBaseDados.has(chaveTitulo) ? indiceBaseDados.get(chaveTitulo) : [];

    produtos.forEach(function (produto) {
      const chave = produto + '||' + anoMes + '||' + loja;
      if (!somaPorChave.has(chave)) {
        somaPorChave.set(chave, { produto: produto, anoMes: anoMes, loja: loja, quantidade: 0 });
      }
      somaPorChave.get(chave).quantidade += quantidade;
    });
  });

  const chavesAtualizadas = mesclarNoResumoHistorico_(abaResumo, somaPorChave);
  reescreverCorpoDaAba_(abaHistorico, cabecalho.length, linhasParaManter);

  return { linhasExpurgadas: linhasParaExpurgar.length, chavesAtualizadas: chavesAtualizadas };
}

// Soma (não sobrescreve) as novas quantidades expurgadas às que já
// existiam em "Resumo Histórico" — importante porque esta função roda
// todo mês, e um mesmo Produto+Ano-Mês+Loja pode já ter uma linha de uma
// expurgação anterior (ex.: dois lojas diferentes do mesmo produto no
// mesmo mês, expurgadas em execuções separadas).
function mesclarNoResumoHistorico_(abaResumo, somaNovaPorChave) {
  const cabecalho = abaResumo.getRange(1, 1, 1, abaResumo.getLastColumn()).getValues()[0];
  const colProduto = obterIndiceColuna_(cabecalho, 'Produto (canônico)', NOME_ABA_RESUMO_HISTORICO);
  const colAnoMes = obterIndiceColuna_(cabecalho, 'Ano-Mês', NOME_ABA_RESUMO_HISTORICO);
  const colLoja = obterIndiceColuna_(cabecalho, 'Loja', NOME_ABA_RESUMO_HISTORICO);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade Total', NOME_ABA_RESUMO_HISTORICO);

  const ultimaLinha = abaResumo.getLastRow();
  const somaAcumulada = new Map();

  if (ultimaLinha >= 2) {
    abaResumo
      .getRange(2, 1, ultimaLinha - 1, abaResumo.getLastColumn())
      .getValues()
      .forEach(function (linha) {
        const chave = linha[colProduto] + '||' + linha[colAnoMes] + '||' + linha[colLoja];
        somaAcumulada.set(chave, {
          produto: linha[colProduto],
          anoMes: linha[colAnoMes],
          loja: linha[colLoja],
          quantidade: Number(linha[colQuantidade]) || 0,
        });
      });
  }

  somaNovaPorChave.forEach(function (valor, chave) {
    if (!somaAcumulada.has(chave)) {
      somaAcumulada.set(chave, { produto: valor.produto, anoMes: valor.anoMes, loja: valor.loja, quantidade: 0 });
    }
    somaAcumulada.get(chave).quantidade += valor.quantidade;
  });

  const linhas = Array.from(somaAcumulada.values())
    .sort(function (a, b) {
      return (
        a.produto.localeCompare(b.produto) || a.anoMes.localeCompare(b.anoMes) || String(a.loja).localeCompare(String(b.loja))
      );
    })
    .map(function (item) {
      return [item.produto, item.anoMes, item.loja, item.quantidade];
    });

  reescreverCorpoDaAba_(abaResumo, abaResumo.getLastColumn(), linhas);

  return somaNovaPorChave.size;
}

// Limpa tudo abaixo do cabeçalho e regrava só as linhas passadas — usado
// pelas duas operações acima pra "remover" linhas sem depender de
// deleteRow linha a linha (mais lento e sujeito a erro de índice).
function reescreverCorpoDaAba_(aba, quantidadeColunas, linhas) {
  const ultimaLinhaAtual = aba.getLastRow();
  if (ultimaLinhaAtual >= 2) {
    aba.getRange(2, 1, ultimaLinhaAtual - 1, quantidadeColunas).clearContent();
  }
  if (linhas.length > 0) {
    aba.getRange(2, 1, linhas.length, quantidadeColunas).setValues(linhas);
  }
}

function formatarAnoMes_(data) {
  return Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM');
}
