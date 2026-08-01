/**
 * Inventário rápido: separa produtos por etapa de produção cruzada com a
 * presença de etiqueta (planilha externa "Etiquetas"), e gera um
 * checklist de produção (☐ Produto) na quantidade sugerida pela Previsão.
 *
 * O prompt pede separar produtos "embalados/rotulados/sem rótulo/
 * prontos/em produção/aguardando embalagem-etiqueta", mas "rótulo" já vem
 * pronto da planilha "Etiquetas" (só leitura) — a ETAPA DE PRODUÇÃO em si
 * não existia em nenhum lugar dos dados. Por isso adicionei uma coluna
 * manual nova "Etapa de Produção" na aba Estoque (rode
 * `adicionarColunaEtapaDeProducao` uma vez), com 3 estágios: Em Produção
 * / Aguardando Embalagem-Etiqueta / Pronto. "Embalado" virou sinônimo de
 * "Pronto" pra não multiplicar estados demais pro cliente manter.
 */

const COLUNA_ETAPA_PRODUCAO = 'Etapa de Produção';
const ETAPAS_PRODUCAO_VALIDAS = ['Em Produção', 'Aguardando Embalagem-Etiqueta', 'Pronto'];
const NOME_ABA_INVENTARIO_RAPIDO = 'Inventário Rápido';
const NOME_ABA_CHECKLIST_PRODUCAO = 'Checklist de Produção';

function adicionarColunaEtapaDeProducao() {
  const abaEstoque = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_ESTOQUE);
  if (!abaEstoque) {
    throw new Error('Aba "' + NOME_ABA_ESTOQUE + '" não encontrada. Rode setupEstrutura() primeiro.');
  }

  const cabecalho = abaEstoque.getRange(1, 1, 1, abaEstoque.getLastColumn()).getValues()[0];
  if (cabecalho.indexOf(COLUNA_ETAPA_PRODUCAO) !== -1) return;

  const proximaColuna = abaEstoque.getLastColumn() + 1;
  abaEstoque
    .getRange(1, proximaColuna)
    .setValue(COLUNA_ETAPA_PRODUCAO)
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');

  const quantidadeLinhas = Math.max(abaEstoque.getLastRow() - 1, 1);
  const regraValidacao = SpreadsheetApp.newDataValidation()
    .requireValueInList(ETAPAS_PRODUCAO_VALIDAS, true)
    .setAllowInvalid(false)
    .build();
  abaEstoque.getRange(2, proximaColuna, quantidadeLinhas, 1).setDataValidation(regraValidacao);
}

function atualizarInventarioRapido() {
  const linhas = calcularInventarioRapido_();
  escreverAbaInventarioRapido_(linhas);
  gerarChecklistDeProducao_();
  return linhas;
}

function calcularInventarioRapido_() {
  const abaEstoque = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_ESTOQUE);
  if (!abaEstoque) throw new Error('Aba "' + NOME_ABA_ESTOQUE + '" não encontrada.');

  const ultimaLinha = abaEstoque.getLastRow();
  if (ultimaLinha < 2) return [];

  const cabecalho = abaEstoque.getRange(1, 1, 1, abaEstoque.getLastColumn()).getValues()[0];
  const colProduto = obterIndiceColuna_(cabecalho, 'Produto (canônico)', NOME_ABA_ESTOQUE);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_ESTOQUE);
  const indiceEtapa = cabecalho.indexOf(COLUNA_ETAPA_PRODUCAO);

  let comEtiqueta = new Set();
  let etiquetaDisponivel = true;
  try {
    comEtiqueta = obterProdutosComEtiquetaDisponivel_();
  } catch (erro) {
    etiquetaDisponivel = false;
  }

  return abaEstoque
    .getRange(2, 1, ultimaLinha - 1, abaEstoque.getLastColumn())
    .getValues()
    .filter(function (linha) {
      return linha[colProduto];
    })
    .map(function (linha) {
      const produto = linha[colProduto];
      const etapa = indiceEtapa !== -1 ? linha[indiceEtapa] : '';
      const rotulado = etiquetaDisponivel ? comEtiqueta.has(normalizarTitulo_(produto)) : null;
      return {
        produto: produto,
        quantidade: Number(linha[colQuantidade]) || 0,
        etapa: etapa || '(não preenchido)',
        rotulado: rotulado,
      };
    });
}

function escreverAbaInventarioRapido_(linhas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_INVENTARIO_RAPIDO);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_INVENTARIO_RAPIDO);
  aba.clear();

  const cabecalho = ['Produto', 'Quantidade', 'Etapa de Produção', 'Etiqueta'];
  aba
    .getRange(1, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  aba.setFrozenRows(1);

  if (linhas.length === 0) {
    aba.autoResizeColumns(1, cabecalho.length);
    return;
  }

  const corpo = linhas.map(function (linha) {
    return [
      linha.produto,
      linha.quantidade,
      linha.etapa,
      linha.rotulado === null ? '—' : linha.rotulado ? 'Rotulado' : 'Sem Rótulo',
    ];
  });
  aba.getRange(2, 1, corpo.length, cabecalho.length).setValues(corpo);
  aba.autoResizeColumns(1, cabecalho.length);
}

function gerarChecklistDeProducao_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_CHECKLIST_PRODUCAO);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_CHECKLIST_PRODUCAO);
  aba.clear();

  aba.getRange(1, 1).setValue('Checklist de Produção Sugerida').setFontWeight('bold').setFontSize(14);

  const sugestoes = calcularPrevisaoDemanda().filter(function (linha) {
    return linha.sugestaoProducao > 0;
  });

  let linhaAtual = 3;
  sugestoes.forEach(function (sugestao) {
    aba
      .getRange(linhaAtual, 1)
      .setValue(sugestao.produto + ' (produzir ' + sugestao.sugestaoProducao + ')')
      .setFontWeight('bold');
    linhaAtual += 1;

    for (let unidade = 1; unidade <= sugestao.sugestaoProducao; unidade++) {
      aba.getRange(linhaAtual, 1).setValue('☐ ' + sugestao.produto);
      linhaAtual += 1;
    }
    linhaAtual += 1;
  });

  aba.autoResizeColumns(1, 1);
}
