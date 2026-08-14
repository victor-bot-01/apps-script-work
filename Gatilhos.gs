/**
 * Gatilhos de automação: execução principal às 02:00 + 4 tentativas de
 * segurança (02:30, 03:00, 03:30, 04:00), todas chamando
 * executarImportacaoAgendada (que aborta sozinha se o dia já foi
 * concluído). Também um gatilho diário às 06:00 pra relatórios em PDF
 * (que só efetivamente gera algo às segundas-feiras e no dia 1 do mês —
 * ver RelatoriosPDF.gs), e um gatilho diário às 15h pra verificação de
 * eventos (ver Eventos.gs: executarVerificacaoEventos). E o menu com
 * todos os botões manuais.
 *
 * Rode `configurarGatilhosImportacao`, `configurarGatilhosRelatorios` e
 * `configurarGatilhosEventos` uma vez cada pra instalar os gatilhos
 * (todos idempotentes: removem os antigos antes de recriar, então dá pra
 * rodar de novo sem duplicar).
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

function configurarGatilhosEventos() {
  removerGatilhosPorFuncao_('executarVerificacaoEventos');

  ScriptApp.newTrigger('executarVerificacaoEventos').timeBased().atHour(15).nearMinute(0).everyDays(1).create();
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
    .addItem('Abrir Dashboard', 'abrirDashboard_')
    .addItem('Atualizar Pedidos Agora', 'executarImportacaoManual')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Retrato manual em planilha (opcional)')
        .addItem('Atualizar Estoque Inteligente', 'atualizarEstoqueInteligente')
        .addItem('Atualizar Previsão', 'atualizarPrevisao')
        .addItem('Atualizar Rankings', 'atualizarRankings')
        .addItem('Atualizar Marketplace', 'atualizarMarketplace')
        .addItem('Atualizar Categorias', 'atualizarCategorias')
        .addItem('Atualizar Calendário', 'atualizarCalendario')
        .addItem('Atualizar Inventário Rápido', 'atualizarInventarioRapido')
        .addItem('Atualizar Alertas', 'atualizarAlertas')
        .addItem('Atualizar Dashboard', 'atualizarDashboard')
    )
    .addSeparator()
    .addItem('Gerar Relatórios Semanais Agora', 'gerarTodosRelatoriosSemanais')
    .addItem('Gerar Relatório Mensal Agora', 'gerarRelatorioResumoMensal')
    .addItem('Verificar Eventos Agora', 'executarVerificacaoEventos')
    .addSeparator()
    .addItem('Configurar Coluna "Etapa de Produção"', 'adicionarColunaEtapaDeProducao')
    .addItem('Remover Backups Antigos de Pedidos', 'removerBackupsPedidosAntigos')
    .addToUi();
}

// Mostra a URL do dashboard (Web App). Só existe depois do primeiro
// "Implantar > Nova implantação" (ver instruções no topo de WebApp.gs);
// antes disso, ScriptApp.getService().getUrl() vem vazio.
function abrirDashboard_() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();

  if (!url) {
    ui.alert(
      'Dashboard ainda não implantado',
      'Vá em Implantar > Nova implantação > App da Web (execute como "Eu", acesso "Somente eu") e rode esta opção de novo.',
      ui.ButtonSet.OK
    );
    return;
  }

  const html = HtmlService.createHtmlOutput(
    '<a href="' + url + '" target="_blank" style="font-family: sans-serif; font-size: 14px;">Abrir dashboard em nova aba</a>'
  )
    .setWidth(320)
    .setHeight(60);
  ui.showModalDialog(html, 'Dashboard Essência do Brasil');
}
