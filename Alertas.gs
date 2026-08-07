/**
 * Alertas inteligentes automáticos: combina os módulos já existentes
 * (Estoque Inteligente, vendas por produto, kits) pra gerar uma lista de
 * alertas em linguagem natural, em vez de recalcular tudo de novo.
 *
 * Limiares (% de crescimento/queda, dias de estoque) são constantes no
 * topo, ajustáveis — o prompt cita os alertas como exemplo, sem definir
 * os números exatos que disparam cada um.
 *
 * `obterVendasDeKitsComData_` lê "Pedidos" + "Histórico" combinados (ver
 * lerPedidosCombinados_ em LeituraVendas.gs) — a janela de 60 dias pode
 * cruzar a virada de mês.
 *
 * O alerta de "sem etiqueta" que existia aqui foi removido — virou a aba
 * própria "Etiquetas" no Web App (ver Etiquetas.gs: calcularEtiquetas),
 * que é mais completa (com/sem etiqueta + sem correspondência, filtrado
 * pelo período selecionado, não o catálogo inteiro).
 */

const NOME_ABA_ALERTAS = 'Alertas';
const LIMIAR_CRESCIMENTO_ACIMA_DA_MEDIA = 30;
const LIMIAR_QUEDA_ALERTA = -30;
const LIMIAR_DIAS_ESTOQUE_ACABANDO = 7;
const LIMIAR_CRESCIMENTO_KIT_RAPIDO = 50;

// Agregador puro usado pelo Web App (WebAppApi.gs).
function calcularAlertas() {
  return []
    .concat(gerarAlertasDeEstoque_())
    .concat(gerarAlertasDeCrescimentoEQueda_())
    .concat(gerarAlertasDeKitsCrescendo_());
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

// Lê de "Pedidos" + "Histórico" combinados (ver lerPedidosCombinados_ em
// LeituraVendas.gs).
function obterVendasDeKitsComData_() {
  const indiceBaseDados = construirIndiceBaseDados_();
  const dataLimite = new Date();
  dataLimite.setDate(dataLimite.getDate() - 60);

  const resultado = [];
  lerPedidosCombinados_().forEach(function (linha) {
    if (linha.data < dataLimite) return;

    const chave = normalizarTitulo_(linha.descricao);
    const produtos = indiceBaseDados.has(chave) ? indiceBaseDados.get(chave) : [];
    if (produtos.length <= 1) return;

    resultado.push({ data: linha.data, chaveTitulo: chave, rotulo: produtos.join(' + '), quantidade: linha.quantidade });
  });

  return resultado;
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
