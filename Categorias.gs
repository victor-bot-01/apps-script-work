/**
 * Categorias (via Ficha Técnica): ranking por categoria, perfumes por
 * gênero/coleção/família olfativa/volumetria, óleo essencial x vegetal,
 * tamanho e "misturas" mais vendidas entre óleos, e participação de cada
 * família olfativa dentro da categoria essência.
 *
 * Tudo aqui depende da Ficha Técnica estar preenchida pelo cliente — se
 * estiver vazia, as seções saem vazias (não é erro).
 *
 * "Misturas" = kits (Base_Dados com ";") cujos produtos são todos óleo
 * essencial/vegetal — não existe um campo explícito de "mistura" nos
 * dados, essa foi a interpretação mais direta do pedido.
 */

const NOME_ABA_CATEGORIAS = 'Categorias';

function lerFichaTecnica_() {
  const mapa = new Map();
  const abaFicha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_FICHA_TECNICA);
  if (!abaFicha) return mapa;

  const ultimaLinha = abaFicha.getLastRow();
  if (ultimaLinha < 2) return mapa;

  const cabecalho = abaFicha.getRange(1, 1, 1, abaFicha.getLastColumn()).getValues()[0];
  const colProduto = obterIndiceColuna_(cabecalho, 'Produto (canônico)', NOME_ABA_FICHA_TECNICA);
  const colGenero = obterIndiceColuna_(cabecalho, 'Gênero', NOME_ABA_FICHA_TECNICA);
  const colColecao = obterIndiceColuna_(cabecalho, 'Coleção', NOME_ABA_FICHA_TECNICA);
  const colFamilia = obterIndiceColuna_(cabecalho, 'Família Olfativa', NOME_ABA_FICHA_TECNICA);
  const colCategoria = obterIndiceColuna_(cabecalho, 'Categoria', NOME_ABA_FICHA_TECNICA);
  const colVolumetria = obterIndiceColuna_(cabecalho, 'Volumetria', NOME_ABA_FICHA_TECNICA);

  abaFicha
    .getRange(2, 1, ultimaLinha - 1, abaFicha.getLastColumn())
    .getValues()
    .forEach(function (linha) {
      const produto = linha[colProduto];
      if (!produto) return;
      mapa.set(produto, {
        genero: linha[colGenero],
        colecao: linha[colColecao],
        familiaOlfativa: linha[colFamilia],
        categoria: linha[colCategoria],
        volumetria: linha[colVolumetria],
      });
    });

  return mapa;
}

// Agregador puro usado pelo Web App (WebAppApi.gs).
function calcularCategorias() {
  const vendas = obterVendasResolvidas_();
  const ficha = lerFichaTecnica_();

  return {
    rankingPorCategoria: calcularRankingPorCampo_(vendas, ficha, 'categoria'),
    perfumesPorGenero: calcularRankingPorCampoDeCategoria_(vendas, ficha, 'perfume', 'genero'),
    perfumesPorColecao: calcularRankingPorCampoDeCategoria_(vendas, ficha, 'perfume', 'colecao'),
    perfumesPorFamiliaOlfativa: calcularRankingPorCampoDeCategoria_(vendas, ficha, 'perfume', 'familiaOlfativa'),
    perfumesPorVolumetria: calcularRankingPorCampoDeCategoria_(vendas, ficha, 'perfume', 'volumetria'),
    oleoEssencialXVegetal: calcularOleoEssencialXVegetal_(vendas, ficha),
    volumetriaDeOleos: calcularVolumetriaDeOleos_(vendas, ficha),
    misturasDeOleos: calcularMisturasDeOleos_(vendas, ficha),
    participacaoFamiliaEssencias: calcularParticipacaoFamiliaEmEssencias_(vendas, ficha),
  };
}

// Utilitário manual (menu "Essência do Brasil"): grava um retrato das
// categorias numa aba. Não é chamado automaticamente pela importação
// diária — a planilha "Análise e Controle" é só banco de dados (ver
// prompt.md); a visão viva é o dashboard (Web App).
function atualizarCategorias() {
  const dados = calcularCategorias();

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_CATEGORIAS);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_CATEGORIAS);
  aba.clear();

  let linha = 1;
  linha = escreverSecaoTabela_(aba, linha, 'Ranking por Categoria', ['Categoria', 'Quantidade'], dados.rankingPorCategoria);
  linha = escreverSecaoTabela_(aba, linha, 'Perfumes por Gênero', ['Gênero', 'Quantidade'], dados.perfumesPorGenero);
  linha = escreverSecaoTabela_(aba, linha, 'Perfumes por Coleção', ['Coleção', 'Quantidade'], dados.perfumesPorColecao);
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Perfumes por Família Olfativa',
    ['Família Olfativa', 'Quantidade'],
    dados.perfumesPorFamiliaOlfativa
  );
  linha = escreverSecaoTabela_(aba, linha, 'Perfumes por Volumetria', ['Volumetria', 'Quantidade'], dados.perfumesPorVolumetria);
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Óleo Essencial x Óleo Vegetal',
    ['Tipo de Óleo', 'Quantidade'],
    dados.oleoEssencialXVegetal
  );
  linha = escreverSecaoTabela_(
    aba,
    linha,
    'Tamanho Mais Vendido entre Óleos',
    ['Volumetria', 'Quantidade'],
    dados.volumetriaDeOleos
  );
  linha = escreverSecaoTabela_(aba, linha, 'Misturas de Óleos Mais Vendidas', ['Mistura', 'Quantidade'], dados.misturasDeOleos);
  escreverSecaoTabela_(
    aba,
    linha,
    'Participação por Família Olfativa (dentro de Essências)',
    ['Família Olfativa', 'Quantidade', '% do Total de Essências'],
    dados.participacaoFamiliaEssencias
  );

  aba.autoResizeColumns(1, 3);
}

function calcularRankingPorCampo_(vendas, ficha, campo) {
  const totais = new Map();
  vendas.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      const info = ficha.get(produto);
      const valor = info ? info[campo] : null;
      if (!valor) return;
      totais.set(valor, (totais.get(valor) || 0) + venda.quantidade);
    });
  });
  return Array.from(totais.entries()).sort(function (a, b) {
    return b[1] - a[1];
  });
}

function calcularRankingPorCampoDeCategoria_(vendas, ficha, categoriaFiltro, campo) {
  const totais = new Map();
  vendas.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      const info = ficha.get(produto);
      if (!info || normalizarTitulo_(info.categoria || '') !== normalizarTitulo_(categoriaFiltro)) return;
      const valor = info[campo];
      if (!valor) return;
      totais.set(valor, (totais.get(valor) || 0) + venda.quantidade);
    });
  });
  return Array.from(totais.entries()).sort(function (a, b) {
    return b[1] - a[1];
  });
}

function calcularOleoEssencialXVegetal_(vendas, ficha) {
  const totais = new Map([
    ['Óleo Essencial', 0],
    ['Óleo Vegetal', 0],
  ]);
  vendas.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      const info = ficha.get(produto);
      if (!info) return;
      const categoria = normalizarTitulo_(info.categoria || '');
      if (categoria === 'óleo essencial') totais.set('Óleo Essencial', totais.get('Óleo Essencial') + venda.quantidade);
      else if (categoria === 'óleo vegetal') totais.set('Óleo Vegetal', totais.get('Óleo Vegetal') + venda.quantidade);
    });
  });
  return Array.from(totais.entries());
}

function calcularVolumetriaDeOleos_(vendas, ficha) {
  const totais = new Map();
  vendas.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      const info = ficha.get(produto);
      if (!info) return;
      const categoria = normalizarTitulo_(info.categoria || '');
      if (categoria !== 'óleo essencial' && categoria !== 'óleo vegetal') return;
      if (!info.volumetria) return;
      totais.set(info.volumetria, (totais.get(info.volumetria) || 0) + venda.quantidade);
    });
  });
  return Array.from(totais.entries()).sort(function (a, b) {
    return b[1] - a[1];
  });
}

function calcularMisturasDeOleos_(vendas, ficha) {
  const totais = new Map();
  vendas
    .filter(function (venda) {
      return venda.isKit;
    })
    .forEach(function (venda) {
      const todosSaoOleos = venda.produtos.every(function (produto) {
        const info = ficha.get(produto);
        if (!info) return false;
        const categoria = normalizarTitulo_(info.categoria || '');
        return categoria === 'óleo essencial' || categoria === 'óleo vegetal';
      });
      if (!todosSaoOleos) return;

      const rotulo = venda.produtos.join(' + ');
      totais.set(rotulo, (totais.get(rotulo) || 0) + venda.quantidade);
    });
  return Array.from(totais.entries()).sort(function (a, b) {
    return b[1] - a[1];
  });
}

function calcularParticipacaoFamiliaEmEssencias_(vendas, ficha) {
  const totais = new Map();
  let totalEssencias = 0;

  vendas.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      const info = ficha.get(produto);
      if (!info || normalizarTitulo_(info.categoria || '') !== 'essência') return;
      const familia = info.familiaOlfativa || '(não classificado)';
      totais.set(familia, (totais.get(familia) || 0) + venda.quantidade);
      totalEssencias += venda.quantidade;
    });
  });

  return Array.from(totais.entries())
    .map(function (par) {
      const percentual = totalEssencias > 0 ? Math.round((par[1] / totalEssencias) * 1000) / 10 : 0;
      return [par[0], par[1], percentual + '%'];
    })
    .sort(function (a, b) {
      return b[1] - a[1];
    });
}
