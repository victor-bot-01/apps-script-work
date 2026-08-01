/**
 * Gatilhos de automação: execução principal às 02:00 + 4 tentativas de
 * segurança (02:30, 03:00, 03:30, 04:00), todas chamando
 * executarImportacaoAgendada (que aborta sozinha se o dia já foi
 * concluído). Também um gatilho diário às 06:00 pra relatórios em PDF
 * (que só efetivamente gera algo às segundas-feiras e no dia 1 do mês —
 * ver RelatoriosPDF.gs). E o menu com todos os botões manuais.
 *
 * Rode `configurarGatilhosImportacao` e `configurarGatilhosRelatorios`
 * uma vez cada pra instalar os gatilhos (ambos idempotentes: removem os
 * antigos antes de recriar, então dá pra rodar de novo sem duplicar).
 *
 * Aviso: gatilhos de tempo do Apps Script disparam dentro de uma janela
 * de alguns minutos ao redor do horário pedido, não no segundo exato.
 */

function configurarGatilhosImportacao() {
  removerGatilhosPorFuncao_('executarImportacaoAgendada');

  const horarios = [
    { hora: 2, minuto: 0 },
    { hora: 2, minuto: 30 },
    { hora: 3, minuto: 0 },
    { hora: 3, minuto: 30 },
    { hora: 4, minuto: 0 },
  ];

  horarios.forEach(function (horario) {
    ScriptApp.newTrigger('executarImportacaoAgendada')
      .timeBased()
      .atHour(horario.hora)
      .nearMinute(horario.minuto)
      .everyDays(1)
      .create();
  });
}

function configurarGatilhosRelatorios() {
  removerGatilhosPorFuncao_('executarRelatoriosAgendados');

  ScriptApp.newTrigger('executarRelatoriosAgendados').timeBased().atHour(6).nearMinute(0).everyDays(1).create();
}

function removerGatilhosPorFuncao_(nomeFuncao) {
  ScriptApp.getProjectTriggers().forEach(function (gatilho) {
    if (gatilho.getHandlerFunction() === nomeFuncao) {
      ScriptApp.deleteTrigger(gatilho);
    }
  });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Essência do Brasil')
    .addItem('Atualizar Pedidos Agora', 'executarImportacaoManual')
    .addSeparator()
    .addItem('Atualizar Estoque Inteligente', 'atualizarEstoqueInteligente')
    .addItem('Atualizar Previsão', 'atualizarPrevisao')
    .addItem('Atualizar Rankings', 'atualizarRankings')
    .addItem('Atualizar Marketplace', 'atualizarMarketplace')
    .addItem('Atualizar Categorias', 'atualizarCategorias')
    .addItem('Atualizar Calendário', 'atualizarCalendario')
    .addItem('Atualizar Inventário Rápido', 'atualizarInventarioRapido')
    .addItem('Atualizar Alertas', 'atualizarAlertas')
    .addItem('Atualizar Dashboard', 'atualizarDashboard')
    .addSeparator()
    .addItem('Gerar Relatórios Semanais Agora', 'gerarTodosRelatoriosSemanais')
    .addItem('Gerar Relatório Mensal Agora', 'gerarRelatorioResumoMensal')
    .addSeparator()
    .addItem('Configurar Coluna "Etapa de Produção"', 'adicionarColunaEtapaDeProducao')
    .addToUi();
}
