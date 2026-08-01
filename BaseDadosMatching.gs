/**
 * Matching entre título de anúncio (Base_Dados, coluna A) e produto(s)
 * canônico(s) (Base_Dados, coluna B), usado para resolver o(s) produto(s)
 * de cada pedido a partir da Descrição importada de "Mês Atual".
 *
 * Formato do Base_Dados (6.144 linhas, sem cabeçalho):
 *   Coluna A = título completo do anúncio.
 *   Coluna B = produto(s) canônico(s), separados por ";" quando é kit.
 *              Pode vir vazia — isso identifica itens fora de escopo
 *              (ex.: joias Ponte Vecchio), que aparecem no Base_Dados mas
 *              não devem entrar nas análises da Essência do Brasil.
 *
 * Três resultados possíveis para um título de pedido:
 *   1. encontrado=true,  produtos=[...]  -> produto(s) Essência do Brasil.
 *   2. encontrado=true,  produtos=[]     -> existe no Base_Dados mas sem
 *                                           produto canônico (fora de
 *                                           escopo) -> ignorar silenciosamente.
 *   3. encontrado=false, produtos=[]     -> título não existe no Base_Dados
 *                                           -> sinalizar (sem correspondência).
 */

const NOME_ABA_BASE_DADOS = 'Base_Dados';
const NOME_ABA_PEDIDOS = 'Pedidos';

/**
 * Lê o Base_Dados inteiro uma única vez e monta um índice título -> produtos.
 * Matching é case-insensitive e tolera espaços extras/inconsistentes no
 * título (comuns em títulos de marketplace).
 *
 * Em caso de título duplicado no Base_Dados, a primeira ocorrência prevalece.
 */
function construirIndiceBaseDados_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(NOME_ABA_BASE_DADOS);
  if (!aba) {
    throw new Error('Aba "' + NOME_ABA_BASE_DADOS + '" não encontrada.');
  }

  const dados = aba.getDataRange().getValues();
  const indice = new Map();

  dados.forEach(function (linha) {
    const titulo = linha[0];
    const produtosBrutos = linha[1];
    if (!titulo) return;

    const chave = normalizarTitulo_(titulo);
    if (indice.has(chave)) return;

    const produtos = produtosBrutos
      ? String(produtosBrutos)
          .split(';')
          .map(function (produto) {
            return produto.trim();
          })
          .filter(function (produto) {
            return produto !== '';
          })
      : [];

    indice.set(chave, produtos);
  });

  return indice;
}

// A integração Bling existente (importarPedidosUltimos3Dias, na planilha
// "Teste Fase 2 Vendas Essência do Brasil") grava a Descrição como
// `${descricao} (Qtd: ${quantidade})` — o sufixo "(Qtd: N)" não existe no
// título do Base_Dados e quebraria o matching de toda linha se não fosse
// removido antes de comparar.
function removerSufixoQuantidade_(texto) {
  return String(texto).replace(/\s*\(qtd:\s*[\d.,]+\)\s*$/i, '');
}

function normalizarTitulo_(titulo) {
  return removerSufixoQuantidade_(titulo).replace(/\s+/g, ' ').trim().toLowerCase();
}

function corresponderTitulo_(titulo, indice) {
  const chave = normalizarTitulo_(titulo);
  if (!indice.has(chave)) {
    return { titulo: titulo, encontrado: false, foraDeEscopo: false, produtos: [] };
  }

  const produtos = indice.get(chave);
  return {
    titulo: titulo,
    encontrado: true,
    foraDeEscopo: produtos.length === 0,
    produtos: produtos,
  };
}

/**
 * Resolve uma lista de títulos de pedido de uma vez (constrói o índice uma
 * única vez). Use esta função em lote em vez de chamar o matching título a
 * título dentro de um loop.
 */
function corresponderTitulosPedidos(titulos) {
  const indice = construirIndiceBaseDados_();
  return titulos.map(function (titulo) {
    return corresponderTitulo_(titulo, indice);
  });
}

/**
 * Varre a coluna "Descrição" da aba Pedidos e devolve os títulos distintos
 * sem correspondência no Base_Dados (para registrar como aviso no Log de
 * Execução). Títulos fora de escopo (ex.: joias) não entram nessa lista.
 */
function listarTitulosSemCorrespondencia() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaPedidos = planilha.getSheetByName(NOME_ABA_PEDIDOS);
  if (!abaPedidos) {
    throw new Error('Aba "' + NOME_ABA_PEDIDOS + '" não encontrada.');
  }

  const cabecalho = abaPedidos.getRange(1, 1, 1, abaPedidos.getLastColumn()).getValues()[0];
  const colDescricao = cabecalho.indexOf('Descrição');
  if (colDescricao === -1) {
    throw new Error('Coluna "Descrição" não encontrada na aba Pedidos.');
  }

  const ultimaLinha = abaPedidos.getLastRow();
  if (ultimaLinha < 2) return [];

  const titulos = abaPedidos
    .getRange(2, colDescricao + 1, ultimaLinha - 1, 1)
    .getValues()
    .map(function (linha) {
      return linha[0];
    })
    .filter(function (titulo) {
      return titulo !== '';
    });

  const resultados = corresponderTitulosPedidos(titulos);
  const titulosSemCorrespondencia = resultados
    .filter(function (resultado) {
      return !resultado.encontrado;
    })
    .map(function (resultado) {
      return resultado.titulo;
    });

  return Array.from(new Set(titulosSemCorrespondencia));
}
