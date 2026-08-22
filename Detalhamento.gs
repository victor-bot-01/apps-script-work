/**
 * Aba "Detalhamento": busca livre por um produto individual ou por uma
 * combinação de kit específica (ex.: "Patchouli 100ml" ou "Patchouli
 * 100ml + Dark Tabac 30ml") em TODO o histórico de vendas já importado —
 * não só o Top 15 do mês passado (esse é o recorte da aba Produção) — e
 * mostra tudo que o sistema sabe sobre aquela unidade específica.
 *
 * Fluxo: `buscarDetalhamento_` (botão "Buscar") devolve os resultados que
 * batem com o termo digitado, ordenados do mais vendido pro menos vendido
 * (produtos individuais e kits misturados na mesma lista — cada um é sua
 * própria "unidade", mesmo conceito de Producao.gs). O usuário clica num
 * resultado e `calcularDetalhamentoUnidade_` monta o painel completo
 * daquela unidade, cruzando com todos os outros módulos:
 * - Vendas: total histórico, mês passado, mês atual, gráfico mensal (janela
 *   escolhida pelo usuário — 6/12/24 meses).
 * - Kits relacionados (só pra produto Individual): quais kits contêm esse
 *   produto, e quanto cada kit vendeu — cruza via `produtosComponentes` que
 *   `agruparVendasPorUnidade_` (Producao.gs) guarda por unidade.
 * - Componentes do kit (só pra unidade Kit): os produtos que formam esse
 *   kit, com o total que cada um vende SOZINHO (fora do kit).
 * - Marketplace: quantidade total histórica por loja + todos os títulos de
 *   anúncio distintos vistos ali (mesmo padrão de
 *   calcularProducaoPorMarketplace_ em RankingsEKits.gs, mas sem
 *   restringir ao Top 15 nem à janela mês passado/atual).
 * - Categoria: gênero/coleção/família olfativa/categoria/volumetria da
 *   Ficha Técnica, por componente (1 linha se Individual, N se Kit).
 * - Produção: se está no Top 15 atual (calcularProducao), mostra
 *   estoque/código/meta/sugestão de lá; se não está, ainda mostra
 *   estoque/código direto da aba Estoque (existem independente do Top 15).
 * - Etiquetas: disponibilidade de etiqueta física por produto componente.
 * - Alertas: alertas ativos agora que mencionam essa unidade.
 */

const TAMANHO_MAX_RESULTADOS_DETALHAMENTO = 30;
const MESES_CURTOS_DETALHAMENTO = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

// Usada pelo botão "Buscar" — varre TODO o histórico (não só o Top 15),
// produtos individuais e kits juntos, ordenado por total vendido desc.
function buscarDetalhamento_(termo) {
  const termoNormalizado = normalizarTitulo_(String(termo || ''));
  if (!termoNormalizado) return [];

  const vendas = obterVendasResolvidas_();
  const unidades = agruparVendasPorUnidade_(vendas);

  const resultados = [];
  unidades.forEach(function (unidade, chave) {
    if (normalizarTitulo_(unidade.rotulo).indexOf(termoNormalizado) === -1) return;
    resultados.push({
      chave: chave,
      rotulo: unidade.rotulo,
      tipo: unidade.tipo,
      totalVendido: unidade.eventos.reduce(function (soma, evento) {
        return soma + evento.quantidade;
      }, 0),
    });
  });

  return resultados
    .sort(function (a, b) {
      return b.totalVendido - a.totalVendido;
    })
    .slice(0, TAMANHO_MAX_RESULTADOS_DETALHAMENTO);
}

// Painel completo de uma unidade específica (produto individual ou kit),
// identificada pela `chave` que já vem do resultado da busca.
// `mesesGrafico`: 6, 12 ou 24 — janela do gráfico de vendas por mês.
function calcularDetalhamentoUnidade_(chave, mesesGrafico) {
  const vendas = obterVendasResolvidas_();
  const unidades = agruparVendasPorUnidade_(vendas);
  const unidade = unidades.get(chave);
  if (!unidade) {
    throw new Error('Produto/kit não encontrado — busque de novo, os dados podem ter mudado.');
  }

  const totalHistorico = unidade.eventos.reduce(function (soma, evento) {
    return soma + evento.quantidade;
  }, 0);

  const mesPassadoInfo = calcularMesesFechados_(1)[0];
  const vendidoMesPassado = somarPorMes_(unidade.eventos, mesPassadoInfo.ano, mesPassadoInfo.mes);

  const hoje = new Date();
  const inicioMesAtual = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const vendidoMesAtual = unidade.eventos.reduce(function (soma, evento) {
    return evento.data >= inicioMesAtual ? soma + evento.quantidade : soma;
  }, 0);

  const kitsRelacionados =
    unidade.tipo === 'Individual'
      ? Array.from(unidades.values())
          .filter(function (u) {
            return u.tipo === 'Kit' && u.produtosComponentes.indexOf(unidade.rotulo) !== -1;
          })
          .map(function (u) {
            return {
              rotulo: u.rotulo,
              totalVendido: u.eventos.reduce(function (soma, evento) {
                return soma + evento.quantidade;
              }, 0),
            };
          })
          .sort(function (a, b) {
            return b.totalVendido - a.totalVendido;
          })
      : [];

  const componentesDoKit =
    unidade.tipo === 'Kit'
      ? unidade.produtosComponentes.map(function (nome) {
          const unidadeIndividual = unidades.get(chaveEstoqueProducao_('Individual', nome));
          return {
            produto: nome,
            vendidoIndividualHistorico: unidadeIndividual
              ? unidadeIndividual.eventos.reduce(function (soma, evento) {
                  return soma + evento.quantidade;
                }, 0)
              : 0,
          };
        })
      : [];

  return {
    rotulo: unidade.rotulo,
    tipo: unidade.tipo,
    totalHistorico: totalHistorico,
    vendidoMesPassado: vendidoMesPassado,
    vendidoMesAtual: vendidoMesAtual,
    graficoMensal: construirGraficoMensalDetalhamento_(unidade.eventos, mesesGrafico || 12),
    kitsRelacionados: kitsRelacionados,
    componentesDoKit: componentesDoKit,
    marketplace: calcularMarketplacePorUnidadeDetalhamento_(chave),
    categoria: calcularCategoriaDetalhamento_(unidade.produtosComponentes),
    producao: calcularProducaoInfoDetalhamento_(chave),
    etiquetas: calcularEtiquetaDetalhamento_(unidade.produtosComponentes),
    alertas: calcularAlertasDetalhamento_(unidade.rotulo),
  };
}

// Cronológico (mais antigo primeiro) — calcularMesesFechados_ devolve do
// mês passado pro mais antigo, por isso o .reverse() no final.
function construirGraficoMensalDetalhamento_(eventos, quantidadeMeses) {
  return calcularMesesFechados_(quantidadeMeses)
    .map(function (info) {
      return {
        rotulo: MESES_CURTOS_DETALHAMENTO[info.mes] + '/' + String(info.ano).slice(2),
        quantidade: somarPorMes_(eventos, info.ano, info.mes),
      };
    })
    .reverse();
}

// Total histórico por loja (não restrito ao mês passado/atual, ao
// contrário de calcularProducaoPorMarketplace_ em RankingsEKits.gs) +
// todos os títulos de anúncio distintos vistos naquele marketplace.
function calcularMarketplacePorUnidadeDetalhamento_(chave) {
  const vendas = obterVendasComLoja_();
  const porLoja = new Map();

  vendas.forEach(function (venda) {
    if (venda.produtos.length === 0) return;
    const rotulo = venda.isKit ? venda.produtos.join(' + ') : venda.produtos[0];
    const tipo = venda.isKit ? 'Kit' : 'Individual';
    if (chaveEstoqueProducao_(tipo, rotulo) !== chave) return;

    if (!porLoja.has(venda.loja)) porLoja.set(venda.loja, { total: 0, titulos: new Set() });
    const registro = porLoja.get(venda.loja);
    registro.total += venda.quantidade;
    if (venda.titulo) registro.titulos.add(venda.titulo);
  });

  return Array.from(porLoja.entries())
    .map(function (par) {
      return { loja: par[0], total: par[1].total, titulos: Array.from(par[1].titulos).sort() };
    })
    .sort(function (a, b) {
      return b.total - a.total;
    });
}

// Ficha Técnica por produto componente (1 linha se Individual, N se Kit).
function calcularCategoriaDetalhamento_(nomesProduto) {
  const ficha = lerFichaTecnica_();
  return nomesProduto.map(function (nome) {
    const info = ficha.get(nome);
    return {
      produto: nome,
      genero: info ? info.genero : '',
      colecao: info ? info.colecao : '',
      familiaOlfativa: info ? info.familiaOlfativa : '',
      categoria: info ? info.categoria : '',
      volumetria: info ? info.volumetria : '',
    };
  });
}

// Se a unidade está no Top 15 atual, usa os números de lá (meta/sugestão
// só existem pra quem está no Top 15); senão, ainda mostra estoque/código,
// que existem na aba Estoque independente de estar no Top 15 ou não.
function calcularProducaoInfoDetalhamento_(chave) {
  const linhaTop15 = calcularProducao().filter(function (linha) {
    return linha.chave === chave;
  })[0];

  if (linhaTop15) {
    return {
      noTop15: true,
      estoqueAtual: linhaTop15.estoqueAtual,
      codigo: linhaTop15.codigo,
      meta: linhaTop15.meta,
      sugestaoProducao: linhaTop15.sugestaoProducao,
    };
  }

  const registroEstoque = lerEstoqueComTipo_().get(chave);
  return {
    noTop15: false,
    estoqueAtual: registroEstoque ? registroEstoque.quantidade : 0,
    codigo: registroEstoque ? registroEstoque.codigo : '',
    meta: null,
    sugestaoProducao: null,
  };
}

function calcularEtiquetaDetalhamento_(nomesProduto) {
  const comEtiqueta = obterProdutosComEtiquetaDisponivel_();
  return nomesProduto.map(function (nome) {
    return { produto: nome, temEtiqueta: comEtiqueta.has(normalizarTitulo_(nome)) };
  });
}

function calcularAlertasDetalhamento_(rotulo) {
  return calcularAlertas().filter(function (alerta) {
    return alerta.mensagem.indexOf('"' + rotulo + '"') !== -1;
  });
}
