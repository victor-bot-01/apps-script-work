/**
 * Cria/atualiza a estrutura de abas da planilha "Análise e Controle".
 * Idempotente: pode rodar quantas vezes quiser sem duplicar abas nem
 * apagar dados já existentes (só ajusta o cabeçalho se ele estiver
 * ausente ou divergente).
 *
 * Como usar:
 * 1. Abra a planilha "Análise e Controle" (conta victor@gigaimports.com).
 * 2. Extensões > Apps Script.
 * 3. Cole este arquivo como um novo script (ex.: SetupEstrutura.gs).
 * 4. Rode a função `setupEstrutura` uma vez (Executar > setupEstrutura).
 *    Na primeira execução o Google vai pedir autorização de permissões.
 *
 * "Histórico" e "Resumo Histórico" (ver Arquivamento.gs) foram
 * adicionadas depois do restante do projeto — guardam, respectivamente,
 * os pedidos de meses já fechados (até 2 anos) e um resumo compacto
 * (Produto + Ano-Mês + Loja) do que passou dos 2 anos, antes de descartar
 * o pedido detalhado.
 */

const DEFINICAO_ABAS = [
  {
    nome: 'Pedidos',
    // Sem coluna Cliente: a regra de negócio é "sem dados nem análises de
    // cliente", então esse dado nem entra na planilha, mesmo em bruto.
    cabecalho: ['Data', 'Pedido', 'Loja', 'Quantidade', 'Descrição'],
  },
  {
    nome: 'Histórico',
    // Mesmo formato de "Pedidos" — é só o destino dos meses já fechados
    // (ver arquivarPedidosDoMesFechado_ em Arquivamento.gs).
    cabecalho: ['Data', 'Pedido', 'Loja', 'Quantidade', 'Descrição'],
  },
  {
    nome: 'Resumo Histórico',
    // Memória de longuíssimo prazo (>2 anos), sem granularidade de
    // pedido/kit — só Produto + Ano-Mês + Loja com a soma vendida.
    cabecalho: ['Produto (canônico)', 'Ano-Mês', 'Loja', 'Quantidade Total'],
  },
  {
    nome: 'Estoque',
    cabecalho: ['Produto (canônico)', 'Quantidade', 'Última Atualização'],
  },
  {
    nome: 'Ficha Técnica',
    cabecalho: [
      'Produto (canônico)',
      'Gênero',
      'Coleção',
      'Família Olfativa',
      'Categoria',
      'Volumetria',
    ],
  },
  {
    nome: 'Log de Execução',
    cabecalho: [
      'Início',
      'Fim',
      'Duração (s)',
      'Pedidos Importados',
      'Erros',
      'Avisos',
    ],
  },
];

function setupEstrutura() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  DEFINICAO_ABAS.forEach(function (definicao) {
    criarOuAtualizarAba(planilha, definicao.nome, definicao.cabecalho);
  });
}

function criarOuAtualizarAba(planilha, nomeAba, cabecalho) {
  let aba = planilha.getSheetByName(nomeAba);
  if (!aba) {
    aba = planilha.insertSheet(nomeAba);
  }

  const faixaCabecalho = aba.getRange(1, 1, 1, cabecalho.length);
  const cabecalhoAtual = faixaCabecalho.getValues()[0];
  const cabecalhoDivergente = cabecalho.some(function (valor, indice) {
    return cabecalhoAtual[indice] !== valor;
  });

  if (cabecalhoDivergente) {
    faixaCabecalho.setValues([cabecalho]);
  }

  faixaCabecalho.setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, cabecalho.length);
}
