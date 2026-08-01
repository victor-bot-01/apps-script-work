/**
 * Alertas inteligentes automáticos: combina os módulos já existentes
 * (Estoque Inteligente, vendas por produto, kits) pra gerar uma lista de
 * alertas em linguagem natural, em vez de recalcular tudo de novo.
 *
 * Limiares (% de crescimento/queda, dias de estoque) são constantes no
 * topo, ajustáveis — o prompt cita os alertas como exemplo, sem definir
 * os números exatos que disparam cada um.
 */

const NOME_ABA_ALERTAS = 'Alertas';
const LIMIAR_CRESCIMENTO_ACIMA_DA_MEDIA = 30;
const LIMIAR_QUEDA_ALERTA = -30;
const LIMIAR_DIAS_ESTOQUE_ACABANDO = 7;
const LIMIAR_CRESCIMENTO_KIT_RAPIDO = 50;

// Agregador puro usado pelo Web App (WebAppApi.gs).
function calcularAlertas() {
  const alertas = []
    .concat(gerarAlertasDeEstoque_())
    .concat(gerarAlertasDeCrescimentoEQueda_())
    .concat(gerarAlertasDeKitsCrescendo_());

  try {
    alertas.push.apply(alertas, gerarAlertasDeEtiqueta_());
  } catch (erro) {
    alertas.push({
      tipo: 'Aviso do sistema',
      mensagem: 'Não foi possível checar etiquetas: ' + (erro && erro.message ? erro.message : erro),
    });
  }

  return alertas;
}

// Utilitário manual (menu "Essência do Brasil"): grava um retrato dos
// alertas numa aba. Não é chamado automaticamente pela importação diária —
// a planilha "Análise e Controle" é só banco de dados (ver prompt.md); a
// visão viva é o dashboard (Web App).
function atualizarAlertas() {
  const alertas = calcularAlertas();
  escreverAbaAlertas_(alertas);
  return alertas;
}

function gerarAlertasDeEstoque_() {
  return calcularEstoqueInteligente()
    .filter(function (linha) {
      return linha.diasRestantes !== null && linha.diasRestantes <= LIMIAR_DIAS_ESTOQUE_ACABANDO;
    })
    .map(function (linha) {
      return {
        tipo: 'Estoque acabando',
        mensagem: 'Estoque de "' + linha.produto + '" acaba em ' + Math.round(linha.diasRestantes) + ' dias.',
      };
    });
}

function gerarAlertasDeCrescimentoEQueda_() {
  const eventosPorProduto = obterVendasPorProduto_(90);
  const alertas = [];

  eventosPorProduto.forEach(function (eventos, produto) {
    const vendidoUltimos30 = somarQuantidadeDesde_(eventos, 30);
    const vendidoAnterior30 = somarQuantidadeEntreDias_(eventos, 30, 60);
    if (vendidoAnterior30 === 0) return;

    const variacao = ((vendidoUltimos30 - vendidoAnterior30) / vendidoAnterior30) * 100;
    if (variacao >= LIMIAR_CRESCIMENTO_ACIMA_DA_MEDIA) {
      alertas.push({
        tipo: 'Acima da média',
        mensagem: 'Produto "' + produto + '" vendeu ' + Math.round(variacao) + '% a mais que no período anterior.',
      });
    } else if (variacao <= LIMIAR_QUEDA_ALERTA) {
      alertas.push({
        tipo: 'Queda de vendas',
        mensagem:
          'Produto "' + produto + '" caiu ' + Math.round(Math.abs(variacao)) + '% em relação ao período anterior.',
      });
    }
  });

  return alertas;
}

function gerarAlertasDeKitsCrescendo_() {
  const vendasDeKits = obterVendasDeKitsComData_();
  const porKit = new Map();

  vendasDeKits.forEach(function (venda) {
    if (!porKit.has(venda.chaveTitulo)) porKit.set(venda.chaveTitulo, { rotulo: venda.rotulo, eventos: [] });
    porKit.get(venda.chaveTitulo).eventos.push({ data: venda.data, quantidade: venda.quantidade });
  });

  const alertas = [];
  porKit.forEach(function (info) {
    const vendidoUltimos30 = somarQuantidadeDesde_(info.eventos, 30);
    const vendidoAnterior30 = somarQuantidadeEntreDias_(info.eventos, 30, 60);
    if (vendidoAnterior30 === 0) return;

    const variacao = ((vendidoUltimos30 - vendidoAnterior30) / vendidoAnterior30) * 100;
    if (variacao >= LIMIAR_CRESCIMENTO_KIT_RAPIDO) {
      alertas.push({
        tipo: 'Kit crescendo rapidamente',
        mensagem: 'Kit "' + info.rotulo + '" cresceu ' + Math.round(variacao) + '% em relação ao período anterior.',
      });
    }
  });

  return alertas;
}

function obterVendasDeKitsComData_() {
  const abaPedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_PEDIDOS);
  if (!abaPedidos) return [];

  const ultimaLinha = abaPedidos.getLastRow();
  if (ultimaLinha < 2) return [];

  const cabecalho = abaPedidos.getRange(1, 1, 1, abaPedidos.getLastColumn()).getValues()[0];
  const colData = obterIndiceColuna_(cabecalho, 'Data', NOME_ABA_PEDIDOS);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_PEDIDOS);
  const colDescricao = obterIndiceColuna_(cabecalho, 'Descrição', NOME_ABA_PEDIDOS);

  const indiceBaseDados = construirIndiceBaseDados_();
  const dataLimite = new Date();
  dataLimite.setDate(dataLimite.getDate() - 60);

  const resultado = [];
  abaPedidos
    .getRange(2, 1, ultimaLinha - 1, abaPedidos.getLastColumn())
    .getValues()
    .forEach(function (linha) {
      const data = linha[colData];
      if (!(data instanceof Date) || data < dataLimite) return;

      const descricao = linha[colDescricao];
      const chave = normalizarTitulo_(descricao);
      const produtos = indiceBaseDados.has(chave) ? indiceBaseDados.get(chave) : [];
      if (produtos.length <= 1) return;

      resultado.push({
        data: data,
        chaveTitulo: chave,
        rotulo: produtos.join(' + '),
        quantidade: Number(linha[colQuantidade]) || 0,
      });
    });

  return resultado;
}

function gerarAlertasDeEtiqueta_() {
  return obterProdutosSemEtiqueta_().map(function (produto) {
    return { tipo: 'Sem etiqueta', mensagem: 'Produto "' + produto + '" está sem etiqueta disponível.' };
  });
}

function escreverAbaAlertas_(alertas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_ALERTAS);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_ALERTAS);
  aba.clear();

  const cabecalho = ['Gerado em', 'Tipo', 'Alerta'];
  aba
    .getRange(1, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  aba.setFrozenRows(1);

  if (alertas.length === 0) {
    aba.autoResizeColumns(1, cabecalho.length);
    return;
  }

  const agora = new Date();
  const linhas = alertas.map(function (alerta) {
    return [agora, alerta.tipo, alerta.mensagem];
  });
  aba.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
  aba.autoResizeColumns(1, cabecalho.length);
}
