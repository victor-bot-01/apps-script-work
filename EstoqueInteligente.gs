/**
 * Estoque inteligente: níveis mínimo/ideal/segurança/máximo por produto,
 * dias restantes até acabar, risco de ruptura, excesso, parado, dias
 * médios para vender 1 unidade e giro semanal/mensal.
 *
 * Metodologia (documentada porque envolve decisões de negócio ajustáveis):
 * - Demanda semanal de cada produto = série de SEMANAS_HISTORICO_ESTOQUE
 *   semanas móveis contadas a partir de hoje (não semana de calendário),
 *   somando pedidos já casados com produto canônico via Base_Dados.
 * - Nível de segurança = Z_SERVICO_70 * desvio-padrão da demanda durante o
 *   lead time assumido (SEMANAS_LEAD_TIME). Z_SERVICO_70 = 0,5244 é o
 *   percentil 70 da normal padrão, isto é, "nível de segurança" cobre 70%
 *   de probabilidade de não faltar produto durante o lead time — a base
 *   pedida no prompt.
 * - Nível mínimo (ponto de reposição) = demanda média durante o lead time
 *   + nível de segurança.
 * - Nível ideal = demanda de SEMANAS_COBERTURA_IDEAL semanas + segurança.
 * - Nível máximo = demanda de 4 semanas (~1 mês) + segurança; acima disso
 *   o produto é considerado em excesso.
 * - Giro = vendas do período ÷ estoque atual (proxy de giro de estoque;
 *   não há série histórica de nível de estoque pra calcular giro contra
 *   estoque médio real).
 *
 * Não há lead time de produção real informado pelo cliente — ajuste
 * SEMANAS_LEAD_TIME se 1 semana não refletir a realidade da produção.
 *
 * `obterVendasPorProduto_` lê de "Pedidos" + "Histórico" combinados (ver
 * lerPedidosCombinados_ em LeituraVendas.gs), não só "Pedidos" — a janela
 * de SEMANAS_HISTORICO_ESTOQUE (12 semanas = ~84 dias) costuma cruzar a
 * virada de mês, e depois do arquivamento mensal "Pedidos" sozinha só tem
 * o mês corrente.
 */

const NOME_ABA_ESTOQUE_INTELIGENTE = 'Estoque Inteligente';
const SEMANAS_HISTORICO_ESTOQUE = 12;
const Z_SERVICO_70 = 0.5244;
const SEMANAS_LEAD_TIME = 1;
const SEMANAS_COBERTURA_IDEAL = 2;
const DIAS_PARADO = 30;

function atualizarEstoqueInteligente() {
  const linhas = calcularEstoqueInteligente();
  escreverAbaEstoqueInteligente_(linhas);
  return linhas;
}

function calcularEstoqueInteligente() {
  const eventosPorProduto = obterVendasPorProduto_(SEMANAS_HISTORICO_ESTOQUE * 7);
  const estoque = lerEstoque_();

  const produtos = new Set(estoque.keys());
  eventosPorProduto.forEach(function (_eventos, produto) {
    produtos.add(produto);
  });

  const linhas = [];
  produtos.forEach(function (produto) {
    const eventos = eventosPorProduto.get(produto) || [];
    const estoqueAtual = estoque.has(produto) ? estoque.get(produto) : null;

    const baldesSemanais = distribuirEmSemanas_(eventos, SEMANAS_HISTORICO_ESTOQUE);
    const estatisticas = calcularMediaEDesvio_(baldesSemanais);
    const niveis = calcularNiveisDeEstoque_(estatisticas);

    const mediaDiaria = estatisticas.mediaSemanal / 7;
    const diasRestantes = estoqueAtual !== null && mediaDiaria > 0 ? estoqueAtual / mediaDiaria : null;
    const diasParaVenderUmaUnidade = mediaDiaria > 0 ? 1 / mediaDiaria : null;
    const vendidoUltimosDiasParado = somarQuantidadeDesde_(eventos, DIAS_PARADO);

    const giroSemanal = estoqueAtual ? estatisticas.mediaSemanal / estoqueAtual : null;
    const giroMensal = estoqueAtual ? (estatisticas.mediaSemanal * 4) / estoqueAtual : null;

    const emRiscoDeRuptura = estoqueAtual !== null && estoqueAtual <= niveis.nivelMinimo;
    const emExcesso = estoqueAtual !== null && estoqueAtual > niveis.nivelMaximo;
    const parado = estoqueAtual !== null && estoqueAtual > 0 && vendidoUltimosDiasParado === 0;

    linhas.push({
      produto: produto,
      estoqueAtual: estoqueAtual,
      nivelMinimo: niveis.nivelMinimo,
      nivelIdeal: niveis.nivelIdeal,
      nivelSeguranca: niveis.nivelSeguranca,
      nivelMaximo: niveis.nivelMaximo,
      diasRestantes: diasRestantes,
      diasParaVenderUmaUnidade: diasParaVenderUmaUnidade,
      giroSemanal: giroSemanal,
      giroMensal: giroMensal,
      status: calcularStatusEstoque_(emRiscoDeRuptura, emExcesso, parado, estoqueAtual),
    });
  });

  return linhas.sort(function (a, b) {
    return a.produto.localeCompare(b.produto);
  });
}

function calcularStatusEstoque_(emRiscoDeRuptura, emExcesso, parado, estoqueAtual) {
  if (estoqueAtual === null) return 'Sem registro em Estoque';

  const rotulos = [];
  if (emRiscoDeRuptura) rotulos.push('Risco de ruptura');
  if (emExcesso) rotulos.push('Excesso');
  if (parado) rotulos.push('Parado');
  return rotulos.length > 0 ? rotulos.join(' / ') : 'Normal';
}

function obterProdutosEmRiscoDeRuptura_() {
  return calcularEstoqueInteligente()
    .filter(function (linha) {
      return linha.status.indexOf('Risco de ruptura') !== -1;
    })
    .map(function (linha) {
      return linha.produto;
    });
}

// Quantidade atribuída inteira a cada produto do kit, igual à lógica já
// usada no Dashboard — um kit de 2 produtos vendido 1x conta 1 unidade
// pra cada produto, não 0,5.
//
// Lê de "Pedidos" + "Histórico" combinados (ver lerPedidosCombinados_ em
// LeituraVendas.gs) — não só "Pedidos".
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

function distribuirEmSemanas_(eventos, semanas) {
  const hoje = new Date();
  const totalDias = semanas * 7;
  const baldes = new Array(semanas).fill(0);

  eventos.forEach(function (evento) {
    const diasAtras = Math.floor((hoje - evento.data) / 86400000);
    if (diasAtras < 0 || diasAtras >= totalDias) return;
    baldes[Math.floor(diasAtras / 7)] += evento.quantidade;
  });

  return baldes;
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

function calcularMediaEDesvio_(valores) {
  const n = valores.length;
  const mediaSemanal =
    valores.reduce(function (soma, valor) {
      return soma + valor;
    }, 0) / n;
  const variancia =
    valores.reduce(function (soma, valor) {
      return soma + Math.pow(valor - mediaSemanal, 2);
    }, 0) / n;
  return { mediaSemanal: mediaSemanal, desvioSemanal: Math.sqrt(variancia) };
}

function calcularNiveisDeEstoque_(estatisticas) {
  const desvioLeadTime = estatisticas.desvioSemanal * Math.sqrt(SEMANAS_LEAD_TIME);
  const demandaLeadTime = estatisticas.mediaSemanal * SEMANAS_LEAD_TIME;

  const nivelSeguranca = Math.round(Z_SERVICO_70 * desvioLeadTime);
  const nivelMinimo = Math.round(demandaLeadTime) + nivelSeguranca;
  const nivelIdeal = Math.round(estatisticas.mediaSemanal * SEMANAS_COBERTURA_IDEAL) + nivelSeguranca;
  const nivelMaximo = Math.round(estatisticas.mediaSemanal * 4) + nivelSeguranca;

  return {
    nivelSeguranca: nivelSeguranca,
    nivelMinimo: nivelMinimo,
    nivelIdeal: nivelIdeal,
    nivelMaximo: nivelMaximo,
  };
}

function escreverAbaEstoqueInteligente_(linhas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_ESTOQUE_INTELIGENTE);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_ESTOQUE_INTELIGENTE);

  aba.clear();
  aba.clearConditionalFormatRules();

  const cabecalho = [
    'Produto (canônico)',
    'Estoque Atual',
    'Nível Mínimo',
    'Nível Ideal',
    'Nível Segurança',
    'Nível Máximo',
    'Dias Restantes',
    'Dias p/ Vender 1 Un.',
    'Giro Semanal',
    'Giro Mensal',
    'Status',
  ];
  aba
    .getRange(1, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  aba.setFrozenRows(1);

  if (linhas.length === 0) {
    aba.autoResizeColumns(1, cabecalho.length);
    return;
  }

  const corpo = linhas.map(function (linha) {
    return [
      linha.produto,
      linha.estoqueAtual,
      linha.nivelMinimo,
      linha.nivelIdeal,
      linha.nivelSeguranca,
      linha.nivelMaximo,
      linha.diasRestantes !== null ? Math.round(linha.diasRestantes) : '—',
      linha.diasParaVenderUmaUnidade !== null ? Math.round(linha.diasParaVenderUmaUnidade * 10) / 10 : '—',
      linha.giroSemanal !== null ? Math.round(linha.giroSemanal * 100) / 100 : '—',
      linha.giroMensal !== null ? Math.round(linha.giroMensal * 100) / 100 : '—',
      linha.status,
    ];
  });

  aba.getRange(2, 1, corpo.length, cabecalho.length).setValues(corpo);
  aba.autoResizeColumns(1, cabecalho.length);

  aplicarFormatacaoCondicionalStatus_(aba, corpo.length);
}

function aplicarFormatacaoCondicionalStatus_(aba, totalLinhas) {
  const faixaStatus = aba.getRange(2, 11, totalLinhas, 1);
  const regras = [
    { texto: 'Risco de ruptura', cor: '#fecaca' },
    { texto: 'Excesso', cor: '#fed7aa' },
    { texto: 'Parado', cor: '#e5e7eb' },
    { texto: 'Sem registro em Estoque', cor: '#fde68a' },
  ];

  aba.setConditionalFormatRules(
    regras.map(function (regra) {
      return SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(regra.texto)
        .setBackground(regra.cor)
        .setRanges([faixaStatus])
        .build();
    })
  );
}
