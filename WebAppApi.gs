/**
 * Camada de API chamada pelo cliente (google.script.run) a partir do
 * Index.html/JavaScript.html. Cada função aqui só adapta o resultado das
 * funções de cálculo que já existem nos outros módulos (Dashboard.gs,
 * EstoqueInteligente.gs, PrevisaoDemanda.gs, RankingsEKits.gs,
 * Marketplace.gs, Categorias.gs, Calendario.gs, Alertas.gs,
 * InventarioRapido.gs) — nenhuma lógica de negócio nova mora aqui.
 *
 * Tudo retornado precisa ser serializável em JSON (google.script.run não
 * consegue devolver Map/Set/Date para o cliente) — por isso os
 * agregadores por seção (calcularRankings, calcularMarketplace,
 * calcularCalendario, calcularCategorias, calcularAlertas) moram no
 * arquivo de domínio de cada um, perto das funções de cálculo que já
 * devolvem arrays/objetos simples.
 */

function apiObterKPIs(diasPeriodo) {
  return calcularKPIs_(diasPeriodo);
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

function apiObterMarketplace() {
  return calcularMarketplace();
}

function apiObterCategorias() {
  return calcularCategorias();
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
