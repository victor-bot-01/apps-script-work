/**
 * Marketplace: quantidade vendida e produto/kit/categoria mais vendidos,
 * por loja/marketplace.
 *
 * `obterVendasComLoja_` lê "Pedidos" + "Histórico" combinados (ver
 * lerPedidosCombinados_ em LeituraVendas.gs) — sem limite de data, então
 * precisa das duas fontes pra não ficar restrita ao mês corrente depois
 * do arquivamento mensal.
 */

const NOME_ABA_MARKETPLACE = 'Marketplace';

// Agregador puro usado pelo Web App (WebAppApi.gs).
function calcularMarketplace() {
  const vendas = obterVendasComLoja_();
  const ficha = lerFichaTecnica_();

  const porLoja = new Map();
  vendas.forEach(function (venda) {
    if (!porLoja.has(venda.loja)) porLoja.set(venda.loja, []);
    porLoja.get(venda.loja).push(venda);
  });

  const categoriasPresentes = obterCategoriasDistintas_(ficha);
  const linhas = Array.from(porLoja.entries())
    .map(function (par) {
      return montarLinhaMarketplace_(par[0], par[1], ficha, categoriasPresentes);
    })
    .sort(function (a, b) {
      return b[1] - a[1];
    });

  return { categorias: categoriasPresentes, linhas: linhas };
}

// Utilitário manual (menu "Essência do Brasil"): grava um retrato do
// marketplace numa aba. Não é chamado automaticamente pela importação
// diária — a planilha "Análise e Controle" é só banco de dados (ver
// prompt.md); a visão viva é o dashboard (Web App).
function atualizarMarketplace() {
  const dados = calcularMarketplace();

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_MARKETPLACE);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_MARKETPLACE);
  aba.clear();

  const cabecalho = ['Loja', 'Quantidade Total', 'Produto Mais Vendido', 'Kit Mais Vendido'].concat(
    dados.categorias.map(function (categoria) {
      return 'Mais Vendido (' + categoria + ')';
    })
  );

  aba
    .getRange(1, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  aba.setFrozenRows(1);

  if (dados.linhas.length > 0) {
    aba.getRange(2, 1, dados.linhas.length, cabecalho.length).setValues(dados.linhas);
  }
  aba.autoResizeColumns(1, cabecalho.length);
}

function obterVendasComLoja_() {
  const indiceBaseDados = construirIndiceBaseDados_();

  return lerPedidosCombinados_()
    .map(function (linha) {
      const chave = normalizarTitulo_(linha.descricao);
      const produtos = indiceBaseDados.has(chave) ? indiceBaseDados.get(chave) : [];
      return {
        loja: linha.loja,
        quantidade: linha.quantidade,
        produtos: produtos,
        isKit: produtos.length > 1,
      };
    })
    .filter(function (venda) {
      return venda.produtos.length > 0;
    });
}

function obterCategoriasDistintas_(ficha) {
  const categorias = new Set();
  ficha.forEach(function (info) {
    if (info.categoria) categorias.add(info.categoria);
  });
  return Array.from(categorias).sort();
}

function montarLinhaMarketplace_(loja, vendasDaLoja, ficha, categoriasPresentes) {
  const quantidadeTotal = vendasDaLoja.reduce(function (soma, venda) {
    return soma + venda.quantidade;
  }, 0);

  const quantidadePorProduto = new Map();
  const quantidadePorKit = new Map();
  const quantidadePorCategoriaEProduto = new Map();

  vendasDaLoja.forEach(function (venda) {
    venda.produtos.forEach(function (produto) {
      quantidadePorProduto.set(produto, (quantidadePorProduto.get(produto) || 0) + venda.quantidade);

      const info = ficha.get(produto);
      const categoria = info ? info.categoria : null;
      if (categoria) {
        if (!quantidadePorCategoriaEProduto.has(categoria)) quantidadePorCategoriaEProduto.set(categoria, new Map());
        const mapaCategoria = quantidadePorCategoriaEProduto.get(categoria);
        mapaCategoria.set(produto, (mapaCategoria.get(produto) || 0) + venda.quantidade);
      }
    });

    if (venda.isKit) {
      const rotulo = venda.produtos.join(' + ');
      quantidadePorKit.set(rotulo, (quantidadePorKit.get(rotulo) || 0) + venda.quantidade);
    }
  });

  const linha = [
    loja,
    quantidadeTotal,
    obterChaveComMaiorValor_(quantidadePorProduto) || '—',
    obterChaveComMaiorValor_(quantidadePorKit) || '—',
  ];

  categoriasPresentes.forEach(function (categoria) {
    const mapaCategoria = quantidadePorCategoriaEProduto.get(categoria);
    linha.push(mapaCategoria ? obterChaveComMaiorValor_(mapaCategoria) || '—' : '—');
  });

  return linha;
}
