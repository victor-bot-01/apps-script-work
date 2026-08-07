/**
 * Cruzamento com a planilha externa "Etiquetas" (aba "Inventário") — só
 * leitura, nunca escreve nela.
 *
 * Essa aba NÃO é uma tabela normal (Produto | Tem Etiqueta) — é uma grade
 * larga: colunas A e B são "Lista de Pedidos Pendentes" e "Localização"
 * (não usadas aqui), e a partir daí o cabeçalho se repete em pares (ex.:
 * "Pasta P1", "Qt.", "Pasta P1", "Qt.", "Pasta P2", "Qt.", ..., "Caixa
 * Profumos 65ml", "Qt."...), um par de colunas por slot físico de
 * pasta/caixa. Dentro de cada par, a CÉLULA (não o cabeçalho) traz o nome
 * do produto atribuído àquele slot, e a coluna "Qt." ao lado é 1 (tem
 * etiqueta disponível) ou 0 (não tem), linha a linha. Confirmado lendo a
 * planilha real — não é mais um chute.
 *
 * Os pares são identificados dinamicamente pelo cabeçalho (qualquer
 * coluna que não seja "Qt." mas cuja próxima coluna seja "Qt.") em vez de
 * índices fixos, porque o número de pastas/caixas muda conforme o
 * cliente reorganiza o estoque físico.
 *
 * Limitação: o nome do produto aqui é digitado à mão nessa planilha
 * separada e pode não bater 100% com o nome canônico do Base_Dados/Ficha
 * Técnica — o matching é case-insensitive (mesma normalização usada em
 * todo o resto do projeto) mas não tolera grafias muito diferentes.
 *
 * `calcularEtiquetas` (usada pela seção "Etiquetas" do Web App) cruza o
 * que VENDEU no período filtrado (não o catálogo inteiro) com essa
 * disponibilidade de etiqueta — substituiu o antigo alerta "Sem etiqueta"
 * de Alertas.gs, que rodava sobre o catálogo inteiro sem filtro de
 * período; essa informação agora vive só aqui, numa aba própria e mais
 * completa (com/sem etiqueta + sem correspondência no Base_Dados).
 */

const ID_PLANILHA_ETIQUETAS = '11yovl7B1uy6YxDjAAvoBfmeWv4t9hcNFibe5_UPA4D4';
const NOME_ABA_INVENTARIO_ETIQUETAS = 'Inventário';
const CABECALHO_COLUNA_QUANTIDADE = 'Qt.';

// Títulos vendidos sem correspondência no Base_Dados que contenham
// qualquer uma dessas palavras são joia (Ponte Vecchio) fora de escopo —
// não aparecem nem na lista "sem etiqueta" nem na lista "sem
// correspondência" da aba Etiquetas. Diferente dos itens que JÁ têm linha
// no Base_Dados com produto canônico vazio (esses já são ignorados pelo
// resto do sistema por padrão); esse filtro cobre os que nem chegaram a
// ser cadastrados no Base_Dados.
const PALAVRAS_FORA_DE_ESCOPO_ETIQUETAS = ['pingente', 'colar', 'pulseira', 'brinco', 'gema', 'anel'];

function ehJoiaForaDeEscopo_(titulo) {
  const tituloNormalizado = normalizarTitulo_(titulo);
  return PALAVRAS_FORA_DE_ESCOPO_ETIQUETAS.some(function (palavra) {
    return tituloNormalizado.indexOf(palavra) !== -1;
  });
}

function obterProdutosComEtiquetaDisponivel_() {
  const planilhaEtiquetas = SpreadsheetApp.openById(ID_PLANILHA_ETIQUETAS);
  const abaInventario = planilhaEtiquetas.getSheetByName(NOME_ABA_INVENTARIO_ETIQUETAS);
  if (!abaInventario) {
    throw new Error('Aba "' + NOME_ABA_INVENTARIO_ETIQUETAS + '" não encontrada na planilha "Etiquetas".');
  }

  const comEtiqueta = new Set();
  const ultimaLinha = abaInventario.getLastRow();
  const ultimaColuna = abaInventario.getLastColumn();
  if (ultimaLinha < 2) return comEtiqueta;

  const cabecalho = abaInventario.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  const paresDeColunas = identificarParesProdutoQuantidade_(cabecalho);

  abaInventario
    .getRange(2, 1, ultimaLinha - 1, ultimaColuna)
    .getValues()
    .forEach(function (linha) {
      paresDeColunas.forEach(function (par) {
        const produto = linha[par.colunaProduto];
        const quantidade = Number(linha[par.colunaQuantidade]) || 0;
        if (produto && quantidade > 0) comEtiqueta.add(normalizarTitulo_(produto));
      });
    });

  return comEtiqueta;
}

function identificarParesProdutoQuantidade_(cabecalho) {
  const pares = [];
  for (let coluna = 0; coluna < cabecalho.length - 1; coluna++) {
    const atual = String(cabecalho[coluna] || '').trim();
    const proxima = String(cabecalho[coluna + 1] || '').trim();
    if (atual !== CABECALHO_COLUNA_QUANTIDADE && proxima === CABECALHO_COLUNA_QUANTIDADE) {
      pares.push({ colunaProduto: coluna, colunaQuantidade: coluna + 1 });
    }
  }
  return pares;
}

// Agregador puro usado pelo Web App (WebAppApi.gs). `periodo` opcional —
// ver lerPedidosNoPeriodo_ em LeituraVendas.gs; omitido, lê todo o
// histórico.
function calcularEtiquetas(periodo) {
  const indiceBaseDados = construirIndiceBaseDados_();
  const comEtiqueta = obterProdutosComEtiquetaDisponivel_();
  const pedidos = periodo ? lerPedidosNoPeriodo_(periodo) : lerPedidosCombinados_();

  const produtosVendidos = new Set();
  const titulosSemCorrespondencia = new Set();

  pedidos.forEach(function (linha) {
    const chave = normalizarTitulo_(linha.descricao);

    if (!indiceBaseDados.has(chave)) {
      if (!ehJoiaForaDeEscopo_(linha.descricao)) {
        titulosSemCorrespondencia.add(linha.descricao);
      }
      return;
    }

    const produtos = indiceBaseDados.get(chave);
    produtos.forEach(function (produto) {
      produtosVendidos.add(produto);
    });
  });

  const comEtiquetaLista = [];
  const semEtiquetaLista = [];
  produtosVendidos.forEach(function (produto) {
    if (comEtiqueta.has(normalizarTitulo_(produto))) comEtiquetaLista.push(produto);
    else semEtiquetaLista.push(produto);
  });

  return {
    comEtiqueta: comEtiquetaLista.sort(),
    semEtiqueta: semEtiquetaLista.sort(),
    semCorrespondencia: Array.from(titulosSemCorrespondencia).sort(),
  };
}
