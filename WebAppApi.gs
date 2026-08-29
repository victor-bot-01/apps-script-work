/**
 * Camada de API chamada pelo cliente (google.script.run) a partir do
 * Index.html/JavaScript.html. Cada função aqui só adapta o resultado das
 * funções de cálculo que já existem nos outros módulos (Dashboard.gs,
 * Producao.gs, RankingsEKits.gs, Marketplace.gs, Categorias.gs,
 * Calendario.gs, Alertas.gs, InventarioRapido.gs, Etiquetas.gs,
 * Eventos.gs, Detalhamento.gs) — nenhuma lógica de negócio nova mora aqui.
 *
 * Tudo retornado precisa ser serializável em JSON (google.script.run não
 * consegue devolver Map/Set/Date para o cliente) — por isso os
 * agregadores por seção (calcularRankings, calcularMarketplace,
 * calcularCalendario, calcularCategorias, calcularAlertas, calcularEtiquetas)
 * moram no arquivo de domínio de cada um, perto das funções de cálculo que
 * já devolvem arrays/objetos simples.
 *
 * `periodo` (KPIs, Marketplace, Categorias, Etiquetas): dias corridos
 * (número/string, ex.: "30") ou 'mesAtual' / 'mesAnterior' — ver
 * resolverIntervaloPeriodo_ em LeituraVendas.gs.
 */

function apiObterKPIs(periodo) {
  return calcularKPIs_(periodo);
}

function apiObterProducao() {
  return calcularProducao();
}

function apiDefinirEstoqueProducao(produto, tipo, quantidade, codigo) {
  definirEstoqueProducao_(produto, tipo, quantidade, codigo);
  return calcularProducao();
}

// `itens`: array de { produto, tipo, codigo, sugestaoProducao }, exatamente
// como exibido na tela no momento do clique (ver gerarEtiquetasProducao_ em
// JavaScript.html). Não retorna nada — sucesso/erro tratado pelo
// resolve/reject da própria chamada.
function apiGerarEtiquetasProducaoPdf(itens) {
  gerarEnviarEtiquetasProducaoPdf_(itens);
}

// `destinatarios`: array de e-mails (digitado + marcados da lista de
// salvos — mesma lista da aba Etiquetas). `itens`: array de
// { produto, tipo, vendidoMesPassado, vendidoMesAtual, sugestaoProducao },
// exatamente como exibido na tela no momento do clique.
function apiEnviarSolicitacaoProducao(destinatarios, itens) {
  gerarEnviarSolicitacaoProducao_(destinatarios, itens);
}

function apiObterRankings() {
  return calcularRankings();
}

function apiObterMarketplace(periodo) {
  return calcularMarketplace(periodo);
}

function apiObterCategorias(periodo) {
  return calcularCategorias(periodo);
}

function apiObterCalendario() {
  return calcularCalendario();
}

function apiObterAlertas() {
  return calcularAlertas();
}

function apiObterInventarioRapido() {
  return {
    itens: calcularInventarioRapido_(),
    checklistProducao: calcularChecklistProducao_(),
  };
}

function apiObterEtiquetas(periodo) {
  return calcularEtiquetas(periodo);
}

// `destinatarios`: array de e-mails (digitado + marcados da lista de
// salvos, já combinados pelo cliente). `itens`: array de
// { produto, total, individual, kit }, exatamente como exibido na tela no
// momento do envio (ver enviarRelatorioEtiquetas_ em JavaScript.html). Não
// retorna nada — o cliente trata sucesso/erro via resolve/reject da
// própria chamada.
function apiEnviarRelatorioEtiquetas(destinatarios, rotuloPeriodo, itens) {
  enviarRelatorioEtiquetas_(destinatarios, rotuloPeriodo, itens);
}

// `dataEnvio`: string 'yyyy-MM-dd' (input type="date" do cliente). Envia
// na hora uma prévia pro e-mail fixo do Victor e devolve a lista
// atualizada de programações ativas (ver programarRelatorioEtiquetas_ em
// Etiquetas.gs).
function apiProgramarRelatorioEtiquetas(destinatarios, rotuloPeriodo, itens, dataEnvio) {
  return programarRelatorioEtiquetas_(destinatarios, rotuloPeriodo, itens, dataEnvio);
}

function apiObterProgramacoesEtiquetas() {
  return obterProgramacoesAtivas_();
}

function apiConfirmarProgramacaoEtiquetas(id) {
  return confirmarProgramacaoEtiquetas_(id);
}

function apiCancelarProgramacaoEtiquetas(id) {
  return cancelarProgramacaoEtiquetas_(id);
}

function apiObterEmailsSalvos() {
  return obterEmailsSalvos_();
}

function apiSalvarEmail(email) {
  return salvarEmail_(email);
}

function apiRemoverEmailSalvo(email) {
  return removerEmailSalvo_(email);
}

function apiBuscarDetalhamento(termo) {
  return buscarDetalhamento_(termo);
}

// `mesesGrafico`: 6, 12 ou 24 — janela do gráfico de vendas por mês.
function apiObterDetalhamento(chave, mesesGrafico) {
  return calcularDetalhamentoUnidade_(chave, mesesGrafico);
}

// Modo "por Item" (soma produto individual + todos os kits que o contêm).
function apiBuscarDetalhamentoItem(termo) {
  return buscarDetalhamentoPorItem_(termo);
}

function apiObterDetalhamentoItem(nomeProduto, mesesGrafico) {
  return calcularDetalhamentoItem_(nomeProduto, mesesGrafico);
}

function apiObterEventos() {
  return obterEventos_();
}

function apiCriarEvento(dados) {
  return criarEvento_(dados);
}

function apiRemoverEvento(id) {
  return removerEvento_(id);
}
