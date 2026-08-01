/**
 * Previsão de demanda por produto: próxima semana, 15 dias, próximo mês,
 * 3 e 6 meses, com intervalo de confiança, média móvel, tendência linear,
 * crescimento percentual e sugestão de quantidade a produzir.
 *
 * Reaproveita a mesma janela e os mesmos buckets semanais do Estoque
 * Inteligente (EstoqueInteligente.gs: obterVendasPorProduto_,
 * distribuirEmSemanas_, calcularMediaEDesvio_, SEMANAS_HISTORICO_ESTOQUE)
 * pra manter as duas análises consistentes entre si.
 *
 * Metodologia:
 * - Regressão linear simples sobre as SEMANAS_HISTORICO_ESTOQUE semanas
 *   (x = índice cronológico da semana, y = quantidade vendida) dá a
 *   "tendência linear" (inclinação) e a base pra extrapolar o futuro.
 * - A previsão de cada horizonte soma uma taxa diária contínua derivada
 *   da reta (em vez de arredondar pra semanas inteiras), então 15 dias ou
 *   91 dias saem exatos, não aproximados por semana.
 * - "Crescimento %" = inclinação da reta expressa como % da média
 *   semanal (positivo = tendência de alta, negativo = queda). É uma
 *   leitura da força da tendência, não o mesmo crescimento histórico
 *   real vs. período anterior que já existe em Rankings.gs.
 * - Intervalo de confiança (95%, Z=1,96 — ajustável) é calculado só para
 *   o horizonte de 30 dias, propagando o desvio-padrão semanal.
 * - Sugestão de produção (30 dias) = previsão do próximo mês + nível de
 *   segurança (do Estoque Inteligente) − estoque atual, nunca negativa.
 */

const NOME_ABA_PREVISAO = 'Previsão';
const Z_CONFIANCA_PREVISAO = 1.96;
const HORIZONTES_PREVISAO = [
  { rotulo: 'Próx. Semana (7d)', dias: 7 },
  { rotulo: '15 Dias', dias: 15 },
  { rotulo: 'Próx. Mês (30d)', dias: 30 },
  { rotulo: '3 Meses (91d)', dias: 91 },
  { rotulo: '6 Meses (182d)', dias: 182 },
];

function atualizarPrevisao() {
  const linhas = calcularPrevisaoDemanda();
  escreverAbaPrevisao_(linhas);
  return linhas;
}

function calcularPrevisaoDemanda() {
  const eventosPorProduto = obterVendasPorProduto_(SEMANAS_HISTORICO_ESTOQUE * 7);

  return calcularEstoqueInteligente()
    .map(function (infoEstoque) {
      const eventos = eventosPorProduto.get(infoEstoque.produto) || [];
      const baldesRecentesPrimeiro = distribuirEmSemanas_(eventos, SEMANAS_HISTORICO_ESTOQUE);
      const baldesCronologicos = baldesRecentesPrimeiro.slice().reverse();

      const estatisticas = calcularMediaEDesvio_(baldesRecentesPrimeiro);
      const regressao = calcularRegressaoLinear_(baldesCronologicos);
      const diariosPrevistos = calcularPrevisoesDiarias_(
        regressao.intercepto,
        regressao.inclinacao,
        baldesCronologicos.length,
        182
      );

      const previsoes = {};
      HORIZONTES_PREVISAO.forEach(function (horizonte) {
        previsoes[horizonte.rotulo] = somarPrefixo_(diariosPrevistos, horizonte.dias);
      });

      const previsaoProximoMes = previsoes['Próx. Mês (30d)'];
      const margemConfianca = Z_CONFIANCA_PREVISAO * estatisticas.desvioSemanal * Math.sqrt(30 / 7);
      const intervaloConfianca = [
        Math.max(0, Math.round(previsaoProximoMes - margemConfianca)),
        Math.round(previsaoProximoMes + margemConfianca),
      ];

      const crescimentoPercentual =
        estatisticas.mediaSemanal > 0 ? (regressao.inclinacao / estatisticas.mediaSemanal) * 100 : null;

      const estoqueAtual = infoEstoque.estoqueAtual;
      const sugestaoProducao =
        estoqueAtual !== null
          ? Math.max(0, Math.round(previsaoProximoMes + infoEstoque.nivelSeguranca - estoqueAtual))
          : Math.round(previsaoProximoMes);

      return {
        produto: infoEstoque.produto,
        mediaMovelSemanal: Math.round(estatisticas.mediaSemanal * 10) / 10,
        tendenciaSemanal: Math.round(regressao.inclinacao * 100) / 100,
        crescimentoPercentual: crescimentoPercentual,
        previsoes: previsoes,
        intervaloConfianca: intervaloConfianca,
        sugestaoProducao: sugestaoProducao,
      };
    })
    .sort(function (a, b) {
      return a.produto.localeCompare(b.produto);
    });
}

function calcularRegressaoLinear_(valoresCronologicos) {
  const n = valoresCronologicos.length;
  if (n === 0) return { inclinacao: 0, intercepto: 0 };

  let somaX = 0;
  let somaY = 0;
  let somaXY = 0;
  let somaXX = 0;
  valoresCronologicos.forEach(function (y, x) {
    somaX += x;
    somaY += y;
    somaXY += x * y;
    somaXX += x * x;
  });

  const denominador = n * somaXX - somaX * somaX;
  const inclinacao = denominador !== 0 ? (n * somaXY - somaX * somaY) / denominador : 0;
  const intercepto = (somaY - inclinacao * somaX) / n;
  return { inclinacao: inclinacao, intercepto: intercepto };
}

function calcularPrevisoesDiarias_(intercepto, inclinacao, semanasHistoricas, diasMax) {
  const diarios = [];
  for (let dia = 1; dia <= diasMax; dia++) {
    const xSemanaContinuo = semanasHistoricas + (dia - 0.5) / 7;
    const taxaSemanal = intercepto + inclinacao * xSemanaContinuo;
    diarios.push(Math.max(taxaSemanal, 0) / 7);
  }
  return diarios;
}

function somarPrefixo_(diarios, quantidadeDias) {
  let soma = 0;
  for (let i = 0; i < quantidadeDias && i < diarios.length; i++) soma += diarios[i];
  return Math.round(soma);
}

function escreverAbaPrevisao_(linhas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_PREVISAO);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_PREVISAO);
  aba.clear();

  const cabecalho = ['Produto', 'Média Móvel (un/sem)', 'Tendência (un/sem)', 'Crescimento %']
    .concat(
      HORIZONTES_PREVISAO.map(function (horizonte) {
        return horizonte.rotulo;
      })
    )
    .concat(['Intervalo de Confiança (30d)', 'Sugestão de Produção (30d)']);

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
      linha.mediaMovelSemanal,
      linha.tendenciaSemanal,
      linha.crescimentoPercentual !== null ? Math.round(linha.crescimentoPercentual * 10) / 10 + '%' : '—',
    ]
      .concat(
        HORIZONTES_PREVISAO.map(function (horizonte) {
          return linha.previsoes[horizonte.rotulo];
        })
      )
      .concat([
        '[' + linha.intervaloConfianca[0] + ' – ' + linha.intervaloConfianca[1] + ']',
        linha.sugestaoProducao,
      ]);
  });

  aba.getRange(2, 1, corpo.length, cabecalho.length).setValues(corpo);
  aba.autoResizeColumns(1, cabecalho.length);
}
