/**
 * Camada de API chamada pelo cliente (google.script.run) a partir do
 * Index.html/JavaScript.html. Cada função aqui só adapta o resultado das
 * funções de cálculo que já existem nos outros módulos (Dashboard.gs,
 * EstoqueInteligente.gs, PrevisaoDemanda.gs, RankingsEKits.gs,
 * Marketplace.gs, Categorias.gs, Calendario.gs, Alertas.gs,
 * InventarioRapido.gs, Etiquetas.gs) — nenhuma lógica de negócio nova
 * mora aqui.
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

function apiObterEstoqueInteligente() {
  return calcularEstoqueInteligente();
}

function apiObterPrevisao() {
  return calcularPrevisaoDemanda();
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

function apiObterEmailsSalvos() {
  return obterEmailsSalvos_();
}

function apiSalvarEmail(email) {
  return salvarEmail_(email);
}

function apiRemoverEmailSalvo(email) {
  return removerEmailSalvo_(email);
}
