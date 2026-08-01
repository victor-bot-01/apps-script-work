/**
 * Web App do dashboard: página HTML própria (Extensões > Apps Script da
 * planilha "Análise e Controle"), aberta no navegador — não é dashboard
 * dentro do Sheets. Acesso restrito a EMAIL_ALERTA_FALHA (victor@gigaimports.com).
 *
 * Deploy (uma vez, e de novo a cada mudança de código):
 * 1. No editor do Apps Script: Implantar > Nova implantação.
 * 2. Tipo: App da Web.
 * 3. Executar como: "Eu" (o dono do script).
 * 4. Quem tem acesso: "Somente eu" — essa é a barreira real de acesso;
 *    a checagem de e-mail abaixo é só uma segunda camada.
 * 5. Copie a URL gerada e abra no navegador.
 *
 * "Quem tem acesso: Somente eu" faz o Apps Script recusar a requisição
 * antes mesmo de rodar doGet() para qualquer conta que não seja a do
 * dono — por isso a checagem de e-mail aqui é defesa em profundidade,
 * útil principalmente se um dia o modo de acesso for trocado por engano.
 */

function doGet() {
  const emailAtivo = Session.getActiveUser().getEmail();
  if (emailAtivo && emailAtivo !== EMAIL_ALERTA_FALHA) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family: sans-serif; padding: 40px;">Acesso restrito.</p>'
    );
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Essência do Brasil — Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function incluirArquivo_(nomeArquivo) {
  return HtmlService.createHtmlOutputFromFile(nomeArquivo).getContent();
}
