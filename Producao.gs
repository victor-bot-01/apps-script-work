/**
 * Sugestão de produção mensal, por unidade vendida — produto individual
 * (anúncio próprio, sem estar em nenhum kit) OU combinação específica de
 * kit (ex.: "Perfume A 100ml + Perfume B 30ml"), cada um tratado como sua
 * própria linha, ranqueados juntos. Substitui Estoque Inteligente.gs
 * (níveis semanais) + PrevisaoDemanda.gs (regressão linear) — os dois
 * foram removidos; as funções utilitárias que outras partes do sistema
 * ainda usam (obterVendasPorProduto_, somarQuantidadeDesde_,
 * somarQuantidadeEntreDias_) continuam em EstoqueInteligente.gs, que
 * ficou só com elas.
 *
 * Metodologia (decisão consciente, ver conversa):
 * - Base do mês = quanto vendeu no mês FECHADO anterior (não uma janela
 *   móvel de dias como antes).
 * - Margem de segurança (confiança de 70%) = Z_SERVICO_70_PRODUCAO ×
 *   desvio-padrão dos totais mensais dos últimos MESES_HISTORICO_PRODUCAO
 *   meses fechados — histórico mais longo só pra medir a variação, a base
 *   em si continua sendo só o mês passado.
 * - Meta do mês = base + margem.
 * - Sugestão de produção = Meta − (já vendido esse mês, do dia 1 até
 *   hoje) − (estoque atual já pronto/embalado), nunca negativa.
 * - Só entram na lista as TAMANHO_TOP_PRODUCAO (15) unidades que mais
 *   venderam no mês passado — o resto ("pouca venda") fica de fora.
 *
 * Estoque: aba "Estoque" ganhou colunas "Tipo" (Individual/Kit) e "Código"
 * (código da caixa, digitado manualmente pelo usuário uma vez por
 * produto/kit e reaproveitado depois) — a mesma aba serve pros dois, a
 * coluna "Quantidade" significa "pronto e embalado pra enviar" pros dois
 * tipos. Editado direto pela aba Produção do Web App (ver
 * apiDefinirEstoqueProducao em WebAppApi.gs), sem precisar abrir a
 * planilha. Produtos individuais continuam contando pro KPI "Estoque
 * total" do Dashboard (ver lerEstoque_ em Dashboard.gs, que ignora linhas
 * Tipo=Kit pra não misturar as duas coisas); kits não têm esse KPI.
 *
 * Etiquetas em PDF (gerarEnviarEtiquetasProducaoPdf_): a partir dos itens
 * do Top 15 marcados pelo usuário no Web App, gera um PDF (A4 paisagem,
 * grade de etiquetas de 5,7cm x 1,5cm com borda, 5 colunas por página) e
 * envia por e-mail pra vitor@... (EMAIL_DESTINATARIO_ETIQUETAS_PRODUCAO).
 * Cada etiqueta tem o tamanho de fonte calculado conforme o comprimento
 * do texto (ver estiloTextoEtiquetaProducao_) — texto curto usa fonte
 * grande numa linha só, texto longo usa fonte menor e quebra em até 3
 * linhas (sem abreviar palavra nenhuma, só quebra normal de texto).
 * Pra cada item selecionado, a quantidade de etiquetas = sugestaoProducao
 * daquele item: N etiquetas com o código da caixa + N etiquetas por
 * produto dentro do rótulo (1 produto se for Individual, 2+ se for Kit,
 * pegos separando o rótulo por " + "). Todas as etiquetas de caixa vêm
 * primeiro (de todos os itens), depois todas as de produto — cada bloco
 * de N repetições do mesmo texto fica contíguo, sem intercalar. Item sem
 * código cadastrado bloqueia a geração com um erro explicando qual(is)
 * item(ns) precisam de código antes. PDF paginado em blocos fixos de
 * COLUNAS_ETIQUETA_PRODUCAO x LINHAS_ETIQUETA_PRODUCAO_POR_PAGINA (5x13),
 * com quebra de página forçada — nunca corta linha na borda da folha.
 *
 * Solicitar Produção (gerarEnviarSolicitacaoProducao_): segundo botão da
 * aba Produção, usa os MESMOS checkboxes da geração de etiquetas. Envia
 * e-mail (destinatário digitado + e-mails salvos, mesma lista compartilhada
 * com a aba Etiquetas) com a sugestão de produção (valor cheio, igual ao
 * mostrado na coluna "Sugestão de Produção") de cada item marcado — itens
 * com sugestão 0 são ignorados sem bloquear o envio dos demais.
 */

const NOME_ABA_PRODUCAO = 'Produção';
const Z_SERVICO_70_PRODUCAO = 0.5244;
const MESES_HISTORICO_PRODUCAO = 6;
const TAMANHO_TOP_PRODUCAO = 15;
const EMAIL_DESTINATARIO_ETIQUETAS_PRODUCAO = 'victor@gigaimports.com';
// 5 colunas de 5,7cm = 28,5cm + margens de 0,6cm dos dois lados = 29,7cm,
// exatamente a largura de uma A4 paisagem.
const LARGURA_ETIQUETA_PRODUCAO_CM = 5.7;
const ALTURA_ETIQUETA_PRODUCAO_CM = 1.5;
const COLUNAS_ETIQUETA_PRODUCAO = 5;
// 13 linhas x 1,5cm = 19,5cm, cabe dentro dos 19,8cm úteis (21cm - 2x0,6cm
// de margem) — cada página sempre fecha numa linha cheia, nunca corta.
const LINHAS_ETIQUETA_PRODUCAO_POR_PAGINA = 13;

function atualizarProducao() {
  const linhas = calcularProducao();
  escreverAbaProducao_(linhas);
  return linhas;
}

function calcularProducao() {
  const vendas = obterVendasResolvidas_();
  const unidades = agruparVendasPorUnidade_(vendas);
  const mesesFechados = calcularMesesFechados_(MESES_HISTORICO_PRODUCAO);

  const hoje = new Date();
  const inicioMesAtual = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const estoque = lerEstoqueComTipo_();

  const linhas = [];
  unidades.forEach(function (unidade) {
    const totaisMensais = mesesFechados.map(function (mesInfo) {
      return somarPorMes_(unidade.eventos, mesInfo.ano, mesInfo.mes);
    });
    const vendidoMesPassado = totaisMensais[0];
    if (vendidoMesPassado <= 0) return; // sem venda no mês passado, fora da lista

    const estatisticas = calcularMediaEDesvioMensal_(totaisMensais);
    const meta = vendidoMesPassado + Z_SERVICO_70_PRODUCAO * estatisticas.desvio;

    const vendidoMesAtual = unidade.eventos.reduce(function (soma, evento) {
      return evento.data >= inicioMesAtual ? soma + evento.quantidade : soma;
    }, 0);

    const chaveEstoque = chaveEstoqueProducao_(unidade.tipo, unidade.rotulo);
    const estoqueAtual = estoque.has(chaveEstoque) ? estoque.get(chaveEstoque).quantidade : 0;

    const sugestaoProducao = Math.max(0, Math.round(meta - vendidoMesAtual - estoqueAtual));
    const codigo = estoque.has(chaveEstoque) ? estoque.get(chaveEstoque).codigo : '';

    linhas.push({
      produto: unidade.rotulo,
      tipo: unidade.tipo,
      chave: chaveEstoque,
      vendidoMesPassado: vendidoMesPassado,
      vendidoMesAtual: vendidoMesAtual,
      meta: Math.round(meta),
      estoqueAtual: estoqueAtual,
      codigo: codigo,
      sugestaoProducao: sugestaoProducao,
    });
  });

  return linhas
    .sort(function (a, b) {
      return b.vendidoMesPassado - a.vendidoMesPassado;
    })
    .slice(0, TAMANHO_TOP_PRODUCAO);
}

// Agrupa vendas por unidade de análise: produto vendido sozinho (isKit
// false) vira uma linha com o nome do próprio produto; produto vendido
// dentro de kit vira uma linha com o rótulo do kit inteiro (ex.: "A + B")
// — o MESMO produto pode aparecer em duas linhas diferentes (uma como
// individual, outra dentro de um ou mais kits), de propósito, cada
// combinação é sua própria unidade produzível. `produtosComponentes` guarda
// a lista de produtos canônicos que compõem a unidade (1 item se
// Individual, 2+ se Kit) — usada por Detalhamento.gs pra cruzar "quais
// kits incluem este produto" e "quais produtos compõem este kit".
function agruparVendasPorUnidade_(vendas) {
  const mapa = new Map();
  vendas.forEach(function (venda) {
    const rotulo = venda.isKit ? venda.produtos.join(' + ') : venda.produtos[0];
    const tipo = venda.isKit ? 'Kit' : 'Individual';
    const chave = chaveEstoqueProducao_(tipo, rotulo);
    if (!mapa.has(chave)) mapa.set(chave, { rotulo: rotulo, tipo: tipo, produtosComponentes: venda.produtos, eventos: [] });
    mapa.get(chave).eventos.push({ data: venda.data, quantidade: venda.quantidade });
  });
  return mapa;
}

function chaveEstoqueProducao_(tipo, rotulo) {
  return tipo + '||' + normalizarTitulo_(rotulo);
}

// mesesFechados[0] = mês passado (o mais recente já fechado), até
// mesesFechados[quantidade-1] = o mais antigo da janela.
function calcularMesesFechados_(quantidade) {
  const hoje = new Date();
  const meses = [];
  for (let i = 1; i <= quantidade; i++) {
    const referencia = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
    meses.push({ ano: referencia.getFullYear(), mes: referencia.getMonth() });
  }
  return meses;
}

function somarPorMes_(eventos, ano, mes) {
  return eventos.reduce(function (soma, evento) {
    return evento.data.getFullYear() === ano && evento.data.getMonth() === mes ? soma + evento.quantidade : soma;
  }, 0);
}

function calcularMediaEDesvioMensal_(totaisMensais) {
  const n = totaisMensais.length;
  const media =
    totaisMensais.reduce(function (soma, valor) {
      return soma + valor;
    }, 0) / n;
  const variancia =
    totaisMensais.reduce(function (soma, valor) {
      return soma + Math.pow(valor - media, 2);
    }, 0) / n;
  return { media: media, desvio: Math.sqrt(variancia) };
}

// Garante que a aba Estoque exista e tenha as colunas "Tipo" e "Código" —
// migração idempotente: coluna que já existe não é tocada; coluna que
// falta é criada (Tipo preenchendo linhas existentes como "Individual",
// Código ficando em branco pra o usuário cadastrar aos poucos).
function obterOuMigrarAbaEstoque_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_ESTOQUE);
  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA_ESTOQUE);
    aba
      .getRange(1, 1, 1, 4)
      .setValues([['Produto (canônico)', 'Quantidade', 'Tipo', 'Código']])
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#ffffff');
    aba.setFrozenRows(1);
    return aba;
  }

  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

  if (cabecalho.indexOf('Tipo') === -1) {
    const proximaColuna = aba.getLastColumn() + 1;
    aba
      .getRange(1, proximaColuna)
      .setValue('Tipo')
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#ffffff');

    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha >= 2) {
      const valoresPadrao = [];
      for (let i = 0; i < ultimaLinha - 1; i++) valoresPadrao.push(['Individual']);
      aba.getRange(2, proximaColuna, ultimaLinha - 1, 1).setValues(valoresPadrao);
    }
  }

  const cabecalhoAtualizado = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  if (cabecalhoAtualizado.indexOf('Código') === -1) {
    const proximaColuna = aba.getLastColumn() + 1;
    aba
      .getRange(1, proximaColuna)
      .setValue('Código')
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#ffffff');
  }

  return aba;
}

// Mapa chave (tipo+produto normalizado) -> { linha, quantidade, codigo }.
function lerEstoqueComTipo_() {
  const aba = obterOuMigrarAbaEstoque_();
  const mapa = new Map();

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return mapa;

  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const colProduto = obterIndiceColuna_(cabecalho, 'Produto (canônico)', NOME_ABA_ESTOQUE);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_ESTOQUE);
  const colTipo = cabecalho.indexOf('Tipo');
  const colCodigo = cabecalho.indexOf('Código');

  aba
    .getRange(2, 1, ultimaLinha - 1, aba.getLastColumn())
    .getValues()
    .forEach(function (linha, indice) {
      const produto = linha[colProduto];
      if (!produto) return;
      const tipo = colTipo !== -1 && linha[colTipo] ? linha[colTipo] : 'Individual';
      const codigo = colCodigo !== -1 && linha[colCodigo] ? String(linha[colCodigo]).trim() : '';
      mapa.set(chaveEstoqueProducao_(tipo, produto), { linha: indice + 2, quantidade: Number(linha[colQuantidade]) || 0, codigo: codigo });
    });

  return mapa;
}

// Atualiza (ou cria, se ainda não existir linha pra essa unidade) a
// quantidade em estoque e o código da caixa de um produto/kit. Chamada
// pelo Web App quando o usuário edita "Estoque Atual"/"Código da Caixa" e
// clica em salvar na aba Produção — os dois campos são salvos juntos num
// só clique.
function definirEstoqueProducao_(produto, tipo, quantidade, codigo) {
  const tituloProduto = String(produto || '').trim();
  const tipoNormalizado = tipo === 'Kit' ? 'Kit' : 'Individual';
  const qtd = Math.max(0, Math.round(Number(quantidade) || 0));
  const codigoNormalizado = String(codigo || '').trim();

  if (!tituloProduto) throw new Error('Produto/kit não informado.');

  const aba = obterOuMigrarAbaEstoque_();
  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const colProduto = obterIndiceColuna_(cabecalho, 'Produto (canônico)', NOME_ABA_ESTOQUE);
  const colQuantidade = obterIndiceColuna_(cabecalho, 'Quantidade', NOME_ABA_ESTOQUE);
  const colTipo = cabecalho.indexOf('Tipo');
  const colCodigo = cabecalho.indexOf('Código');

  const ultimaLinha = aba.getLastRow();
  let linhaEncontrada = -1;
  if (ultimaLinha >= 2) {
    const dados = aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn()).getValues();
    for (let i = 0; i < dados.length; i++) {
      const tipoLinha = colTipo !== -1 && dados[i][colTipo] ? dados[i][colTipo] : 'Individual';
      if (tipoLinha === tipoNormalizado && normalizarTitulo_(dados[i][colProduto]) === normalizarTitulo_(tituloProduto)) {
        linhaEncontrada = i + 2;
        break;
      }
    }
  }

  if (linhaEncontrada !== -1) {
    aba.getRange(linhaEncontrada, colQuantidade + 1).setValue(qtd);
    if (colCodigo !== -1) aba.getRange(linhaEncontrada, colCodigo + 1).setValue(codigoNormalizado);
  } else {
    const novaLinha = new Array(aba.getLastColumn()).fill('');
    novaLinha[colProduto] = tituloProduto;
    novaLinha[colQuantidade] = qtd;
    if (colTipo !== -1) novaLinha[colTipo] = tipoNormalizado;
    if (colCodigo !== -1) novaLinha[colCodigo] = codigoNormalizado;
    aba.appendRow(novaLinha);
  }
}

// "Em risco": estoque atual não cobre o resto da meta do mês em
// andamento — mesma condição de sugestaoProducao > 0. Usada pelo KPI
// "Produtos em risco de ruptura" (Dashboard.gs).
function obterProdutosEmRiscoDeRuptura_() {
  return calcularProducao()
    .filter(function (linha) {
      return linha.sugestaoProducao > 0;
    })
    .map(function (linha) {
      return linha.produto;
    });
}

function escreverAbaProducao_(linhas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_PRODUCAO);
  if (!aba) aba = planilha.insertSheet(NOME_ABA_PRODUCAO);
  aba.clear();

  const cabecalho = [
    'Produto/Kit',
    'Tipo',
    'Vendido Mês Passado',
    'Vendido Este Mês',
    'Meta do Mês',
    'Estoque Atual',
    'Código da Caixa',
    'Sugestão de Produção',
  ];
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
      linha.tipo,
      linha.vendidoMesPassado,
      linha.vendidoMesAtual,
      linha.meta,
      linha.estoqueAtual,
      linha.codigo,
      linha.sugestaoProducao,
    ];
  });
  aba.getRange(2, 1, corpo.length, cabecalho.length).setValues(corpo);
  aba.autoResizeColumns(1, cabecalho.length);
}

// Recebe os itens do Top 15 marcados pelo usuário no Web App (cada um já
// com produto/tipo/código/sugestaoProducao exatamente como exibido na
// tela — mesmo padrão de enviarRelatorioEtiquetas_ em Etiquetas.gs, o
// servidor não recalcula nada), mais `incluirCaixa`/`incluirIndividual`
// (booleans, um por item — escolhidos no modal de seleção antes do envio;
// ausente/undefined é tratado como true, pra não quebrar chamadas antigas).
// Item com os dois false simplesmente não contribui etiqueta nenhuma.
// Gera um único PDF com todas as etiquetas dos itens/tipos selecionados e
// envia por e-mail.
function gerarEnviarEtiquetasProducaoPdf_(itensRecebidos) {
  const itens = Array.isArray(itensRecebidos) ? itensRecebidos : [];
  if (itens.length === 0) throw new Error('Nenhum produto/kit selecionado.');

  // Só exige código de itens que realmente vão gerar etiqueta de caixa.
  const semCodigo = itens.filter(function (item) {
    return item.incluirCaixa !== false && !String((item && item.codigo) || '').trim();
  });
  if (semCodigo.length > 0) {
    throw new Error(
      'Cadastre o código da caixa antes de gerar as etiquetas: ' +
        semCodigo
          .map(function (item) {
            return item.produto;
          })
          .join(', ')
    );
  }

  const etiquetasCaixa = [];
  const etiquetasProduto = [];

  itens.forEach(function (item) {
    const quantidade = Math.max(0, Math.round(Number(item.sugestaoProducao) || 0));
    if (quantidade === 0) return;

    if (item.incluirCaixa !== false) {
      const codigo = String(item.codigo).trim();
      for (let i = 0; i < quantidade; i++) etiquetasCaixa.push(codigo);
    }

    if (item.incluirIndividual !== false) {
      const nomesProduto = item.tipo === 'Kit' ? String(item.produto).split(' + ') : [String(item.produto)];
      nomesProduto.forEach(function (nome) {
        const nomeAparado = nome.trim();
        if (!nomeAparado) return;
        for (let i = 0; i < quantidade; i++) etiquetasProduto.push(nomeAparado);
      });
    }
  });

  if (etiquetasCaixa.length === 0 && etiquetasProduto.length === 0) {
    throw new Error(
      'Nenhuma etiqueta a gerar — confira se algum item tem sugestão de produção 0 ou se "Caixa"/"Individual" ficaram desmarcados em todos.'
    );
  }

  const blobPdf = construirPdfEtiquetasProducao_(etiquetasCaixa.concat(etiquetasProduto));

  MailApp.sendEmail({
    to: EMAIL_DESTINATARIO_ETIQUETAS_PRODUCAO,
    subject: '🏷️ Etiquetas de Produção — Essência do Brasil',
    body:
      'Segue em anexo o PDF com as etiquetas para colar nas caixas e nos produtos (' +
      etiquetasCaixa.length +
      ' etiqueta(s) de caixa, ' +
      etiquetasProduto.length +
      ' etiqueta(s) de produto).',
    attachments: [blobPdf],
  });
}

// Tamanho de fonte por etiqueta, baseado no comprimento do texto — texto
// curto (ex.: código de caixa) usa fonte grande, preenchendo bem a
// etiqueta; texto longo (ex.: nome de produto grande) usa fonte menor e
// quebra em várias linhas (até 3, a altura da etiqueta permite), sem
// abreviar palavras. `overflow: hidden` na etiqueta fica só como rede de
// segurança pro caso raríssimo de um texto não caber nem assim.
function estiloTextoEtiquetaProducao_(texto) {
  const comprimento = String(texto).length;
  if (comprimento <= 12) return { fontSize: 15, lineHeight: 16 };
  if (comprimento <= 24) return { fontSize: 12, lineHeight: 13 };
  if (comprimento <= 40) return { fontSize: 10, lineHeight: 11 };
  return { fontSize: 8, lineHeight: 9 };
}

// Grade de etiquetas de texto simples (sem QR code/imagem) em A4 paisagem
// — @page/CSS em cm garante o tamanho físico exato de cada etiqueta
// (LARGURA/ALTURA_ETIQUETA_PRODUCAO_CM), a borda ajuda no corte. `textos`
// já vem na ordem final (blocos contíguos de caixa, depois de produto).
// Dividido em páginas de exatamente COLUNAS x LINHAS_POR_PAGINA etiquetas
// (5x13=65), com quebra de página forçada entre blocos — assim a grade
// nunca "flui" além do que cabe e corta uma linha na borda da folha.
function construirPdfEtiquetasProducao_(textos) {
  const etiquetasPorPagina = COLUNAS_ETIQUETA_PRODUCAO * LINHAS_ETIQUETA_PRODUCAO_POR_PAGINA;
  const paginas = [];
  for (let inicio = 0; inicio < textos.length; inicio += etiquetasPorPagina) {
    paginas.push(textos.slice(inicio, inicio + etiquetasPorPagina));
  }

  const htmlPaginas = paginas
    .map(function (pagina, indice) {
      const celulas = pagina
        .map(function (texto) {
          const estilo = estiloTextoEtiquetaProducao_(texto);
          return (
            '<div class="etiqueta" style="font-size:' +
            estilo.fontSize +
            'pt;line-height:' +
            estilo.lineHeight +
            'pt;">' +
            escaparHtml_(texto) +
            '</div>'
          );
        })
        .join('');
      const estiloQuebra = indice < paginas.length - 1 ? ' style="page-break-after: always;"' : '';
      return '<div class="pagina"' + estiloQuebra + '><div class="grade">' + celulas + '</div></div>';
    })
    .join('');

  const html =
    '<html><head><meta charset="utf-8"><style>' +
    '@page { size: A4 landscape; margin: 0.6cm; }' +
    'body { margin: 0; font-family: Arial, Helvetica, sans-serif; }' +
    '.grade { display: flex; flex-wrap: wrap; }' +
    '.etiqueta {' +
    'width: ' +
    LARGURA_ETIQUETA_PRODUCAO_CM +
    'cm; height: ' +
    ALTURA_ETIQUETA_PRODUCAO_CM +
    'cm; box-sizing: border-box; border: 1px solid #e1e1e1; ' +
    'display: flex; align-items: center; justify-content: center; text-align: center; ' +
    'font-weight: bold; overflow: hidden; padding: 0 3px; white-space: normal; word-wrap: break-word;' +
    '}' +
    '</style></head><body>' +
    htmlPaginas +
    '</body></html>';

  const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return HtmlService.createHtmlOutput(html).getAs('application/pdf').setName('Etiquetas_Producao_' + carimbo + '.pdf');
}

// Recebe destinatários + itens do Top 15 marcados (mesmos checkboxes que
// alimentam a geração de etiquetas — ver gerarEnviarEtiquetasProducaoPdf_
// acima) e envia um e-mail com a sugestão de produção (valor cheio) de
// cada um. Itens com sugestaoProducao 0 (nada a produzir) são simplesmente
// ignorados, sem bloquear o envio dos demais —
// só dá erro se sobrar zero item depois desse filtro. Usa os mesmos
// e-mails salvos da aba Etiquetas (obterEmailsSalvos_/salvarEmail_ em
// Etiquetas.gs) — lista compartilhada entre as duas telas de propósito.
function gerarEnviarSolicitacaoProducao_(destinatarios, itensRecebidos) {
  const listaDestinatarios = (Array.isArray(destinatarios) ? destinatarios : [destinatarios])
    .map(function (email) {
      return String(email || '').trim();
    })
    .filter(function (email) {
      return email !== '';
    });

  if (listaDestinatarios.length === 0) {
    throw new Error('Nenhum destinatário informado.');
  }
  listaDestinatarios.forEach(function (email) {
    if (!validarEmail_(email)) {
      throw new Error('E-mail do destinatário inválido: "' + email + '".');
    }
  });

  const itens = (Array.isArray(itensRecebidos) ? itensRecebidos : [])
    .map(function (item) {
      return {
        produto: item.produto,
        tipo: item.tipo,
        vendidoMesPassado: Number(item.vendidoMesPassado) || 0,
        vendidoMesAtual: Number(item.vendidoMesAtual) || 0,
        quantidadeSolicitada: Math.round(Number(item.sugestaoProducao) || 0),
      };
    })
    .filter(function (item) {
      return item.quantidadeSolicitada > 0;
    });

  if (itens.length === 0) {
    throw new Error('Nenhum item selecionado tem sugestão de produção — nada a solicitar.');
  }

  MailApp.sendEmail({
    to: listaDestinatarios.join(','),
    subject: '📋 Sugestão de Produção — Essência do Brasil',
    htmlBody: construirEmailSolicitacaoProducao_(itens),
  });
}

// HTML de e-mail com tabelas + estilo inline, mesmo estilo visual do
// relatório de Etiquetas (ver construirEmailEtiquetas_ em Etiquetas.gs),
// com cor de destaque azul pra diferenciar visualmente do relatório
// laranja de etiquetas pendentes.
function construirEmailSolicitacaoProducao_(itens) {
  const totalSolicitado = itens.reduce(function (soma, item) {
    return soma + item.quantidadeSolicitada;
  }, 0);

  const linhas = itens
    .map(function (item, indice) {
      const corFundo = indice % 2 === 0 ? '#17233d' : '#111a2e';
      return (
        '<tr>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#e6ecff;font-size:13px;">🏭 ' +
        escaparHtml_(item.produto) +
        ' <span style="color:#93a2c7;font-size:11px;">(' +
        escaparHtml_(item.tipo) +
        ')</span></td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#93a2c7;font-size:13px;text-align:center;">' +
        fmtNumero_(item.vendidoMesPassado) +
        '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#93a2c7;font-size:13px;text-align:center;">' +
        fmtNumero_(item.vendidoMesAtual) +
        '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#86efac;font-size:13px;text-align:center;font-weight:bold;">' +
        fmtNumero_(item.quantidadeSolicitada) +
        '</td>' +
        '</tr>'
      );
    })
    .join('');

  return (
    '<div style="background:#0b1220;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#111a2e;border-radius:14px;overflow:hidden;border:1px solid #263353;">' +
    '<tr><td style="background:#2563eb;padding:32px 32px 28px 32px;">' +
    '<div style="font-size:30px;line-height:1.2;">📋</div>' +
    '<div style="font-size:22px;font-weight:bold;color:#ffffff;margin-top:10px;">Sugestão de Produção</div>' +
    '<div style="font-size:14px;color:#dbeafe;margin-top:4px;">Essência do Brasil &middot; Painel de Vendas &amp; Estoque</div>' +
    '</td></tr>' +
    '<tr><td style="padding:28px 32px 8px 32px;">' +
    '<div style="color:#e6ecff;font-size:14px;line-height:1.6;">' +
    'Com base nas vendas recentes, aqui vai uma sugestão de produção. Fica a sugestão, sem compromisso.' +
    '</div>' +
    '</td></tr>' +
    '<tr><td style="padding:16px 32px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td width="50%" style="background:#1f2937;border-radius:10px;padding:16px;text-align:center;">' +
    '<div style="color:#93a2c7;font-size:11px;text-transform:uppercase;letter-spacing:.04em;">Produtos/kits</div>' +
    '<div style="color:#93c5fd;font-size:26px;font-weight:bold;margin-top:6px;">' +
    itens.length +
    '</div>' +
    '</td>' +
    '<td width="16">&nbsp;</td>' +
    '<td width="50%" style="background:#1f2937;border-radius:10px;padding:16px;text-align:center;">' +
    '<div style="color:#93a2c7;font-size:11px;text-transform:uppercase;letter-spacing:.04em;">Total sugerido</div>' +
    '<div style="color:#86efac;font-size:26px;font-weight:bold;margin-top:6px;">' +
    fmtNumero_(totalSolicitado) +
    '</div>' +
    '</td>' +
    '</tr></table>' +
    '</td></tr>' +
    '<tr><td style="padding:8px 32px 24px 32px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-radius:10px;overflow:hidden;border:1px solid #263353;">' +
    '<tr>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;text-align:left;">Produto/Kit</th>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;">Vendido Mês Passado</th>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;">Vendido Este Mês</th>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;">Minha Sugestão Para Produção</th>' +
    '</tr>' +
    linhas +
    '</table>' +
    '</td></tr>' +
    '<tr><td style="padding:0 32px 28px 32px;">' +
    '<div style="color:#93a2c7;font-size:12.5px;line-height:1.6;font-style:italic;">' +
    '🥖 Assim como ninguém liga o forno pra assar um pãozinho só, produzir de uma vez só rende mais tempo livre depois — e menos idas à cozinha (ou à produção) no meio da semana.' +
    '</div>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</div>'
  );
}
