/**
 * Pipeline de importação diária de pedidos: lê "Mês Atual" (planilha externa
 * "Teste Fase 2 Vendas Essência do Brasil"), copia os pedidos novos para a
 * aba "Pedidos" desta planilha, e registra cada rodada na aba
 * "Log de Execução".
 *
 * Regras aplicadas aqui (ver prompt.md):
 * - Antes de importar, roda o arquivamento mensal (ver Arquivamento.gs):
 *   move pedidos de mês fechado de "Pedidos" para "Histórico", e expurga
 *   de "Histórico" o que passou de 2 anos (resumindo em "Resumo
 *   Histórico"). Isso deixa "Pedidos" sempre enxuta (só mês corrente).
 * - Data de cada pedido novo = dia anterior ao dia da execução.
 * - Deduplicação por (Número do pedido multiloja + Descrição), verificando
 *   tanto "Pedidos" quanto "Histórico" — não só "Pedidos", porque depois
 *   que o arquivamento existe um pedido do fim do mês passado pode já ter
 *   sido movido pra "Histórico" antes da próxima execução rodar; sem
 *   checar lá também, ele pareceria "novo" e seria reimportado duplicado.
 *   Também não é por "último ID": números de pedido não são um contador
 *   único entre as 18 lojas/marketplaces.
 * - Sem coluna/dado de cliente na aba Pedidos.
 * - Sem análise de preço/faturamento — as colunas de valor de "Mês Atual"
 *   nem são lidas.
 * - Se uma tentativa do dia já tiver concluído com sucesso, as tentativas
 *   de segurança seguintes abortam sem reprocessar.
 * - Se todas as tentativas de um dia falharem, o lote acumulado é
 *   processado no dia seguinte com aviso explícito no Log de Execução.
 * - Backup da aba Pedidos antes de cada gravação; alerta por e-mail em
 *   caso de falha.
 */

const ID_PLANILHA_VENDAS = '1KCTrbN6Jx9sg1H-ho6eIZ4we1rm5HL0mjd6OtlpWMMA';
const NOME_ABA_MES_ATUAL = 'Mês Atual';
const NOME_ABA_LOG = 'Log de Execução';
const EMAIL_ALERTA_FALHA = 'victor@gigaimports.com';
const PROP_ULTIMA_DATA_IMPORTADA = 'ultimaDataAlvoImportada';
const MAX_BACKUPS_PEDIDOS = 30;

/**
 * Ponto de entrada usado pelos 5 gatilhos diários (02:00, 02:30, 03:00,
 * 03:30, 04:00). Aborta silenciosamente (só registrando no Log) se a
 * importação do dia-alvo já tiver sido concluída com sucesso por uma
 * tentativa anterior.
 */
function executarImportacaoAgendada() {
  const dataAlvo = calcularDataAlvo_();
  const jaConcluida =
    PropertiesService.getScriptProperties().getProperty(PROP_ULTIMA_DATA_IMPORTADA) ===
    formatarDataISO_(dataAlvo);

  if (jaConcluida) {
    registrarLog_({
      inicio: new Date(),
      fim: new Date(),
      pedidosImportados: 0,
      erros: '',
      avisos: 'Importação de ' + formatarDataISO_(dataAlvo) + ' já concluída — tentativa abortada.',
    });
    return;
  }

  executarImportacaoComTratamento_(dataAlvo);
}

/**
 * Ponto de entrada do botão de atualização manual (menu "Essência do
 * Brasil"). Roda mesmo que a importação do dia já tenha sido concluída,
 * pois é um pedido explícito do usuário.
 */
function executarImportacaoManual() {
  const dataAlvo = calcularDataAlvo_();
  executarImportacaoComTratamento_(dataAlvo);
  SpreadsheetApp.getActiveSpreadsheet().toast('Importação concluída. Veja o Log de Execução.', 'Essência do Brasil');
}

function executarImportacaoComTratamento_(dataAlvo) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    registrarLog_({
      inicio: new Date(),
      fim: new Date(),
      pedidosImportados: 0,
      erros: '',
      avisos: 'Outra importação já em andamento — tentativa abortada.',
    });
    return;
  }

  const inicio = new Date();
  try {
    // Arquivamento roda antes da importação — assim "Pedidos" já está
    // enxuta (só mês corrente) quando a checagem de duplicidade e a
    // gravação dos pedidos novos acontecem logo abaixo.
    const avisosArquivamento = executarArquivamento_();

    const resultado = importarPedidosNovos_(dataAlvo);
    executarModulosPosImportacao_(resultado);

    registrarLog_({
      inicio: inicio,
      fim: new Date(),
      pedidosImportados: resultado.totalImportado,
      erros: '',
      avisos: avisosArquivamento.concat(resultado.avisos).join(' | '),
    });
    PropertiesService.getScriptProperties().setProperty(
      PROP_ULTIMA_DATA_IMPORTADA,
      formatarDataISO_(dataAlvo)
    );
  } catch (erro) {
    registrarLog_({
      inicio: inicio,
      fim: new Date(),
      pedidosImportados: 0,
      erros: String(erro && erro.message ? erro.message : erro),
      avisos: '',
    });
    enviarAlertaFalha_(erro);
  } finally {
    lock.releaseLock();
  }
}

// A planilha "Análise e Controle" é só banco de dados (ver prompt.md) — as
// análises (Estoque Inteligente, Previsão, Rankings, Marketplace,
// Categorias, Calendário, Inventário Rápido, Alertas, Dashboard) não
// gravam mais abas aqui a cada importação. Elas são recalculadas sob
// demanda pelo dashboard (Web App, ver WebApp.gs/WebAppApi.gs) quando o
// usuário abre a página, sempre a partir dos dados mais recentes de
// Pedidos/Histórico/Estoque/Ficha Técnica. As funções `atualizar*` que
// gravam abas continuam disponíveis nos respectivos arquivos como
// utilitário manual (menu "Essência do Brasil"), pra quem quiser um
// retrato pontual dentro do próprio Sheets.
function executarModulosPosImportacao_(resultado) {
  // Sem módulos pós-importação por enquanto — mantido como ponto de
  // extensão caso surja alguma análise que precise mesmo ser persistida
  // em planilha (ex.: cache pesado demais pra recalcular por requisição).
}

function importarPedidosNovos_(dataAlvo) {
  const planilhaAtual = SpreadsheetApp.getActiveSpreadsheet();
  const abaPedidos = planilhaAtual.getSheetByName(NOME_ABA_PEDIDOS);
  if (!abaPedidos) {
    throw new Error('Aba "' + NOME_ABA_PEDIDOS + '" não encontrada. Rode setupEstrutura() primeiro.');
  }

  const planilhaVendas = SpreadsheetApp.openById(ID_PLANILHA_VENDAS);
  const abaMesAtual = planilhaVendas.getSheetByName(NOME_ABA_MES_ATUAL);
  if (!abaMesAtual) {
    throw new Error(
      'Aba "' + NOME_ABA_MES_ATUAL + '" não encontrada em "Teste Fase 2 Vendas Essência do Brasil".'
    );
  }

  const dadosOrigem = abaMesAtual.getDataRange().getValues();
  const cabecalhoOrigem = dadosOrigem[0];
  const colPedido = obterIndiceColuna_(cabecalhoOrigem, 'Número do pedido multiloja - Venda', NOME_ABA_MES_ATUAL);
  const colLoja = obterIndiceColuna_(cabecalhoOrigem, 'Loja', NOME_ABA_MES_ATUAL);
  const colQuantidade = obterIndiceColuna_(cabecalhoOrigem, 'Quantidade', NOME_ABA_MES_ATUAL);
  const colDescricao = obterIndiceColuna_(cabecalhoOrigem, 'Descrição', NOME_ABA_MES_ATUAL);

  const chavesExistentes = obterChavesExistentesEmPedidosEHistorico_();
  const chavesNesteLote = new Set();
  const linhasParaGravar = [];

  dadosOrigem.slice(1).forEach(function (linha) {
    const pedido = linha[colPedido];
    const descricao = linha[colDescricao];
    if (!pedido || !descricao) return;

    const chave = normalizarChavePedidoProduto_(pedido, descricao);
    if (chavesExistentes.has(chave) || chavesNesteLote.has(chave)) return;

    chavesNesteLote.add(chave);
    // A Descrição de "Mês Atual" vem com "(Qtd: N)" colado no fim (ver
    // removerSufixoQuantidade_ em BaseDadosMatching.gs) — removido aqui
    // pra não duplicar a informação que já está na coluna Quantidade.
    const descricaoLimpa = removerSufixoQuantidade_(descricao).trim();
    linhasParaGravar.push([dataAlvo, pedido, linha[colLoja], linha[colQuantidade], descricaoLimpa]);
  });

  const avisos = [];

  if (linhasParaGravar.length > 0) {
    criarBackupPedidos_(abaPedidos);

    abaPedidos
      .getRange(abaPedidos.getLastRow() + 1, 1, linhasParaGravar.length, linhasParaGravar[0].length)
      .setValues(linhasParaGravar);

    const titulosNovos = linhasParaGravar.map(function (linha) {
      return linha[4];
    });
    const semCorrespondencia = Array.from(
      new Set(
        corresponderTitulosPedidos(titulosNovos)
          .filter(function (resultado) {
            return !resultado.encontrado;
          })
          .map(function (resultado) {
            return resultado.titulo;
          })
      )
    );
    if (semCorrespondencia.length > 0) {
      avisos.push('Títulos sem correspondência no Base_Dados: ' + semCorrespondencia.join('; '));
    }
  }

  const avisoGap = verificarGapDeDias_(dataAlvo);
  if (avisoGap) avisos.push(avisoGap);

  return { totalImportado: linhasParaGravar.length, avisos: avisos };
}

// Verifica duplicidade contra "Pedidos" + "Histórico" juntos (ver
// lerPedidosCombinados_ em LeituraVendas.gs) — necessário porque um
// pedido pode ter sido arquivado de "Pedidos" para "Histórico" pelo
// arquivamento mensal desde a última importação.
function obterChavesExistentesEmPedidosEHistorico_() {
  const chaves = new Set();
  lerPedidosCombinados_().forEach(function (linha) {
    chaves.add(normalizarChavePedidoProduto_(linha.pedido, linha.descricao));
  });
  return chaves;
}

function normalizarChavePedidoProduto_(pedido, descricao) {
  return String(pedido).trim().toLowerCase() + '||' + normalizarTitulo_(descricao);
}

function obterIndiceColuna_(cabecalho, nomeColuna, nomeAba) {
  const indice = cabecalho.indexOf(nomeColuna);
  if (indice === -1) {
    throw new Error('Coluna "' + nomeColuna + '" não encontrada na aba "' + nomeAba + '".');
  }
  return indice;
}

function calcularDataAlvo_() {
  const agora = new Date();
  return new Date(agora.getFullYear(), agora.getMonth(), agora.getDate() - 1);
}

function formatarDataISO_(data) {
  return Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function verificarGapDeDias_(dataAlvo) {
  const ultima = PropertiesService.getScriptProperties().getProperty(PROP_ULTIMA_DATA_IMPORTADA);
  if (!ultima) return null;

  const dataUltima = new Date(ultima + 'T00:00:00');
  const diffDias = Math.round((dataAlvo.getTime() - dataUltima.getTime()) / 86400000);

  if (diffDias > 1) {
    return (
      'Atenção: a última importação concluída com sucesso foi em ' +
      ultima +
      '. Este lote pode conter pedidos de mais de 1 dia — ajuste manual de data pode ser necessário.'
    );
  }
  return null;
}

function criarBackupPedidos_(abaPedidos) {
  const planilha = abaPedidos.getParent();
  const abaAtivaOriginal = planilha.getActiveSheet();
  const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

  const copia = abaPedidos.copyTo(planilha);
  copia.setName('Backup_Pedidos_' + carimbo);
  copia.hideSheet();
  planilha.setActiveSheet(abaAtivaOriginal);

  limparBackupsAntigos_(planilha);
}

function limparBackupsAntigos_(planilha) {
  const backups = planilha
    .getSheets()
    .filter(function (aba) {
      return aba.getName().indexOf('Backup_Pedidos_') === 0;
    })
    .sort(function (a, b) {
      return a.getName().localeCompare(b.getName());
    });

  if (backups.length <= MAX_BACKUPS_PEDIDOS) return;

  backups.slice(0, backups.length - MAX_BACKUPS_PEDIDOS).forEach(function (aba) {
    planilha.deleteSheet(aba);
  });
}

function registrarLog_(entrada) {
  const abaLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_LOG);
  if (!abaLog) return;

  const duracaoSegundos = Math.round((entrada.fim.getTime() - entrada.inicio.getTime()) / 1000);
  abaLog.appendRow([
    entrada.inicio,
    entrada.fim,
    duracaoSegundos,
    entrada.pedidosImportados,
    entrada.erros,
    entrada.avisos,
  ]);
}

function enviarAlertaFalha_(erro) {
  const mensagem =
    'Falha na importação de pedidos em ' +
    new Date().toString() +
    '.\n\nErro: ' +
    (erro && erro.message ? erro.message : erro);
  MailApp.sendEmail(EMAIL_ALERTA_FALHA, 'Essência do Brasil — Falha na importação de pedidos', mensagem);
}
