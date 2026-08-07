/**
 * Leitura combinada de vendas: "Pedidos" (mês corrente) + "Histórico"
 * (até 2 anos, ver Arquivamento.gs) — fonte única usada por todo módulo
 * de análise (Dashboard, Estoque Inteligente, Previsão, Rankings, Kits,
 * Marketplace, Alertas — Categorias e Calendário usam essas por tabela,
 * então também se beneficiam sem precisar de nenhuma mudança própria).
 *
 * Antes de existir "Histórico", cada módulo lia "Pedidos" direto. Agora
 * que "Pedidos" só guarda o mês corrente (arquivamento mensal move o
 * resto pra "Histórico"), qualquer leitura que olhasse só pra "Pedidos"
 * ficaria incompleta pra janelas maiores que 1 mês (ex.: "últimos 12
 * meses", "todo o histórico" em Rankings). Esta função existe pra
 * centralizar o conserto num lugar só.
 *
 * "Resumo Histórico" (dados com mais de 2 anos) NÃO entra aqui — não tem
 * granularidade de pedido/kit/loja por linha, só Produto+Ano-Mês+Loja já
 * agregado, então não serve pras análises que precisam de linha de
 * pedido individual.
 *
 * `resolverIntervaloPeriodo_`/`lerPedidosNoPeriodo_` centralizam o filtro
 * de período do Web App (dropdown "período"): aceita tanto dias corridos
 * (número, ex.: 30 = últimos 30 dias a partir de hoje) quanto período de
 * calendário fechado ('mesAtual' = dia 1 do mês corrente até hoje;
 * 'mesAnterior' = dia 1 até o último dia do mês anterior). Usada por
 * Dashboard, Marketplace, Categorias e Etiquetas — as únicas seções que
 * respeitam esse filtro (Rankings, Estoque Inteligente, Previsão,
 * Calendário e Alertas usam janelas fixas próprias, por design — ver
 * comentários em cada arquivo).
 */

function lerLinhasBrutasDeAba_(aba) {
  if (!aba) return [];
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return [];

  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const colData = obterIndiceColuna_(cabecalho, 'Data', aba.getName());
  const colPedido = obterIndiceColuna_(cabecalho, 'Pedido', aba.getName());
  const colLoja = obterIndiceColuna_(cabecalho, 'Loja', aba.getName());
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', aba.getName());
  const colDescricao = obterIndiceColuna_(cabecalho, 'Descrição', aba.getName());

  return aba
    .getRange(2, 1, ultimaLinha - 1, aba.getLastColumn())
    .getValues()
    .filter(function (linha) {
      return linha[colData] instanceof Date;
    })
    .map(function (linha) {
      return {
        data: linha[colData],
        pedido: linha[colPedido],
        loja: linha[colLoja],
        quantidade: Number(linha[colQuantidade]) || 0,
        descricao: linha[colDescricao],
      };
    });
}

// Todo módulo de análise deveria chamar esta função em vez de ler
// "Pedidos" (ou "Histórico") diretamente.
function lerPedidosCombinados_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaPedidos = planilha.getSheetByName(NOME_ABA_PEDIDOS);
  const abaHistorico = planilha.getSheetByName(NOME_ABA_HISTORICO);
  return lerLinhasBrutasDeAba_(abaPedidos).concat(lerLinhasBrutasDeAba_(abaHistorico));
}

// periodo: número de dias corridos (ex.: 30) ou 'mesAtual' / 'mesAnterior'.
function resolverIntervaloPeriodo_(periodo) {
  const hoje = new Date();
  const inicioHoje = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  const amanha = new Date(inicioHoje.getFullYear(), inicioHoje.getMonth(), inicioHoje.getDate() + 1);

  if (periodo === 'mesAtual') {
    return { inicio: new Date(hoje.getFullYear(), hoje.getMonth(), 1), fimExclusivo: amanha };
  }

  if (periodo === 'mesAnterior') {
    return {
      inicio: new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1),
      fimExclusivo: new Date(hoje.getFullYear(), hoje.getMonth(), 1),
    };
  }

  const dias = Number(periodo) || DIAS_PERIODO_KPI;
  const inicio = new Date(inicioHoje.getFullYear(), inicioHoje.getMonth(), inicioHoje.getDate() - dias);
  return { inicio: inicio, fimExclusivo: amanha };
}

function lerPedidosNoPeriodo_(periodo) {
  const intervalo = resolverIntervaloPeriodo_(periodo);
  return lerPedidosCombinados_().filter(function (linha) {
    return linha.data >= intervalo.inicio && linha.data < intervalo.fimExclusivo;
  });
}

function rotuloPeriodo_(periodo) {
  if (periodo === 'mesAtual') return 'mês atual';
  if (periodo === 'mesAnterior') return 'mês anterior';
  return 'últimos ' + (Number(periodo) || DIAS_PERIODO_KPI) + ' dias';
}
