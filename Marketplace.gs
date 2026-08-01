/**
 * Marketplace: quantidade vendida e produto/kit/categoria mais vendidos,
 * por loja/marketplace.
 */

const NOME_ABA_MARKETPLACE = 'Marketplace';

function atualizarMarketplace() {
  const vendas = obterVendasComLoja_();
  const ficha = lerFichaTecnica_();

  const porLoja = new Map();
  vendas.forEach(function (venda) {
    if (!porLoja.has(venda.loja)) porLoja.set(venda.loja, []);
    porLoja.get(venda.loja).push(venda);
  });

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_MARKETPLACE);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_MARKETPLACE);
  aba.clear();

  const categoriasPresentes = obterCategoriasDistintas_(ficha);
  const cabecalho = ['Loja', 'Quantidade Total', 'Produto Mais Vendido', 'Kit Mais Vendido'].concat(
    categoriasPresentes.map(function (categoria) {
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

  const linhas = Array.from(porLoja.entries())
    .map(function (par) {
      return montarLinhaMarketplace_(par[0], par[1], ficha, categoriasPresentes);
    })
    .sort(function (a, b) {
      return b[1] - a[1];
    });

  if (linhas.length > 0) {
    aba.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
  }
  aba.autoResizeColumns(1, cabecalho.length);
}

function obterVendasComLoja_() {
  const abaPedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_PEDIDOS);
  if (!abaPedidos) throw new Error('Aba "' + NOME_ABA_PEDIDOS + '" não encontrada.');

  const ultimaLinha = abaPedidos.getLastRow();
  if (ultimaLinha < 2) return [];

  const cabecalho = abaPedidos.getRange(1, 1, 1, abaPedidos.getLastColumn()).getValues()[0];
  const colLoja = obterIndiceColuna_(cabecalho, 'Loja', NOME_ABA_PEDIDOS);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_PEDIDOS);
  const colDescricao = obterIndiceColuna_(cabecalho, 'Descrição', NOME_ABA_PEDIDOS);

  const indiceBaseDados = construirIndiceBaseDados_();

  return abaPedidos
    .getRange(2, 1, ultimaLinha - 1, abaPedidos.getLastColumn())
    .getValues()
    .map(function (linha) {
      const descricao = linha[colDescricao];
      const chave = normalizarTitulo_(descricao);
      const produtos = indiceBaseDados.has(chave) ? indiceBaseDados.get(chave) : [];
      return {
        loja: linha[colLoja],
        quantidade: Number(linha[colQuantidade]) || 0,
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
