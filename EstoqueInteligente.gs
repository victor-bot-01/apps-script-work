/**
 * Utilitários de leitura de vendas por produto, por janela de dias
 * corridos a partir de hoje. Historicamente viviam junto com o cálculo de
 * níveis de estoque (daí o nome do arquivo), mas esse cálculo foi
 * substituído pela aba "Produção" (ver Producao.gs) — essas três funções
 * continuam aqui porque Alertas.gs (crescimento/queda de vendas, kits
 * crescendo) e Calendario.gs (classificação de produtos) ainda dependem
 * delas, sem relação nenhuma com o cálculo que foi trocado.
 *
 * Lê de "Pedidos" + "Histórico" combinados (ver lerPedidosCombinados_ em
 * LeituraVendas.gs), não só "Pedidos".
 */

// Quantidade atribuída inteira a cada produto do kit, igual à lógica já
// usada no Dashboard — um kit de 2 produtos vendido 1x conta 1 unidade
// pra cada produto, não 0,5.
function obterVendasPorProduto_(diasHistorico) {
  const eventosPorProduto = new Map();

  const dataLimite = new Date();
  dataLimite.setDate(dataLimite.getDate() - diasHistorico);

  const indiceBaseDados = construirIndiceBaseDados_();

  lerPedidosCombinados_()
    .filter(function (linha) {
      return linha.data >= dataLimite;
    })
    .forEach(function (linha) {
      const chave = normalizarTitulo_(linha.descricao);
      const produtos = indiceBaseDados.has(chave) ? indiceBaseDados.get(chave) : [];
      if (produtos.length === 0) return;

      produtos.forEach(function (produto) {
        if (!eventosPorProduto.has(produto)) eventosPorProduto.set(produto, []);
        eventosPorProduto.get(produto).push({ data: linha.data, quantidade: linha.quantidade });
      });
    });

  return eventosPorProduto;
}

function somarQuantidadeDesde_(eventos, dias) {
  return somarQuantidadeEntreDias_(eventos, 0, dias);
}

// Reaproveitado por Alertas.gs e Calendario.gs pra comparar uma janela
// (ex.: 30-60 dias atrás) sem duplicar a lógica de filtro por data.
function somarQuantidadeEntreDias_(eventos, diasAtrasInicio, diasAtrasFim) {
  const hoje = new Date();
  return eventos.reduce(function (soma, evento) {
    const diasAtras = Math.floor((hoje - evento.data) / 86400000);
    return diasAtras >= diasAtrasInicio && diasAtras < diasAtrasFim ? soma + evento.quantidade : soma;
  }, 0);
}
