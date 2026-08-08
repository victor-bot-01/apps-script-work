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
 * completa (com/sem etiqueta, quantidade total/individual/kit por
 * produto, + sem correspondência no Base_Dados).
 *
 * `enviarRelatorioEtiquetas_` envia por e-mail (MailApp) um subconjunto
 * escolhido pelo usuário da lista "sem etiqueta" — o servidor não
 * recalcula nada, só recebe os itens (produto + quantidades) que o
 * cliente já tinha na tela e monta o e-mail. Isso mantém o número exibido
 * na tela idêntico ao número enviado, sem risco de os dois divergirem por
 * causa de uma mudança de filtro entre o carregamento e o envio. Aceita
 * mais de um destinatário de uma vez (array).
 *
 * `obterEmailsSalvos_`/`salvarEmail_`/`removerEmailSalvo_` guardam a
 * lista de e-mails salvos em PropertiesService (ScriptProperties) — é
 * global, qualquer um que abrir o painel vê e usa a mesma lista, não
 * depende de navegador/máquina.
 */

const ID_PLANILHA_ETIQUETAS = '11yovl7B1uy6YxDjAAvoBfmeWv4t9hcNFibe5_UPA4D4';
const NOME_ABA_INVENTARIO_ETIQUETAS = 'Inventário';
const CABECALHO_COLUNA_QUANTIDADE = 'Qt.';
const CHAVE_PROPRIEDADE_EMAILS_SALVOS = 'EMAILS_SALVOS_ETIQUETAS';

// Títulos vendidos que contenham qualquer uma dessas palavras são joia
// (Ponte Vecchio) fora de escopo — não aparecem em nenhuma das três
// listas da aba Etiquetas (com etiqueta, sem etiqueta, sem
// correspondência), mesmo quando têm produto canônico cadastrado no
// Base_Dados.
const PALAVRAS_FORA_DE_ESCOPO_ETIQUETAS = [
  'pingente',
  'colar',
  'pulseira',
  'brinco',
  'gema',
  'anel',
  'larimar',
  'corrente',
  'carteira',
  'nécessaire',
  'bolsa',
];

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
// histórico. "comEtiqueta"/"semEtiqueta" trazem, por produto, a
// quantidade total vendida no período e a divisão kit/individual (mesma
// lógica de agruparKitIndividualPorProduto_ em RankingsEKits.gs).
// "semCorrespondencia" não tem quantidade — títulos que não bateram com
// nenhum produto canônico no Base_Dados não têm como ser classificados em
// kit/individual.
function calcularEtiquetas(periodo) {
  const indiceBaseDados = construirIndiceBaseDados_();
  const comEtiqueta = obterProdutosComEtiquetaDisponivel_();
  const pedidos = periodo ? lerPedidosNoPeriodo_(periodo) : lerPedidosCombinados_();

  const titulosSemCorrespondencia = new Set();
  const vendas = [];

  pedidos.forEach(function (linha) {
    const chave = normalizarTitulo_(linha.descricao);

    if (!indiceBaseDados.has(chave)) {
      if (!ehJoiaForaDeEscopo_(linha.descricao)) {
        titulosSemCorrespondencia.add(linha.descricao);
      }
      return;
    }

    const produtos = indiceBaseDados.get(chave);
    if (produtos.length === 0) return; // fora de escopo já cadastrado (ex.: joia Ponte Vecchio)

    // isKit reflete como o pedido foi vendido originalmente (kit com N
    // produtos), mesmo que um dos produtos do kit seja joia e acabe
    // filtrado abaixo — só não atribuímos quantidade a ele.
    const isKitOriginal = produtos.length > 1;
    const produtosValidos = produtos.filter(function (produto) {
      return !ehJoiaForaDeEscopo_(produto);
    });
    if (produtosValidos.length === 0) return;

    vendas.push({ quantidade: linha.quantidade, produtos: produtosValidos, isKit: isKitOriginal });
  });

  const dadosPorProduto = agruparKitIndividualPorProduto_(vendas);

  const comEtiquetaLista = [];
  const semEtiquetaLista = [];
  dadosPorProduto.forEach(function (registro, produto) {
    const item = {
      produto: produto,
      total: registro.kit + registro.individual,
      individual: registro.individual,
      kit: registro.kit,
    };
    if (comEtiqueta.has(normalizarTitulo_(produto))) comEtiquetaLista.push(item);
    else semEtiquetaLista.push(item);
  });

  comEtiquetaLista.sort(function (a, b) {
    return a.produto.localeCompare(b.produto);
  });
  semEtiquetaLista.sort(function (a, b) {
    return a.produto.localeCompare(b.produto);
  });

  return {
    comEtiqueta: comEtiquetaLista,
    semEtiqueta: semEtiquetaLista,
    semCorrespondencia: Array.from(titulosSemCorrespondencia).sort(),
  };
}

// `destinatarios`: string única ou array de e-mails (digitado + marcados
// da lista de salvos, já combinados pelo cliente). `itens`: array de
// { produto, total, individual, kit } — exatamente o que o cliente já
// tinha renderizado na tela (ver JavaScript.html: enviarRelatorioEtiquetas_).
// O servidor não recalcula nada, só valida e monta o e-mail.
function enviarRelatorioEtiquetas_(destinatarios, rotuloPeriodo, itens) {
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
  if (!itens || itens.length === 0) {
    throw new Error('Nenhum produto selecionado para o relatório.');
  }

  const assunto = '🏷️ Relatório de Etiquetas Pendentes — Essência do Brasil (' + rotuloPeriodo + ')';
  const corpoHtml = construirEmailEtiquetas_(itens, rotuloPeriodo);

  MailApp.sendEmail({
    to: listaDestinatarios.join(','),
    subject: assunto,
    htmlBody: corpoHtml,
  });
}

function validarEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function obterEmailsSalvos_() {
  const bruto = PropertiesService.getScriptProperties().getProperty(CHAVE_PROPRIEDADE_EMAILS_SALVOS);
  if (!bruto) return [];
  try {
    const lista = JSON.parse(bruto);
    return Array.isArray(lista) ? lista : [];
  } catch (erro) {
    return [];
  }
}

function gravarEmailsSalvos_(lista) {
  PropertiesService.getScriptProperties().setProperty(CHAVE_PROPRIEDADE_EMAILS_SALVOS, JSON.stringify(lista));
}

// Retorna a lista atualizada (o cliente re-renderiza direto com o retorno,
// sem precisar de uma segunda chamada pra buscar de novo).
function salvarEmail_(email) {
  const emailNormalizado = String(email || '').trim().toLowerCase();
  if (!validarEmail_(emailNormalizado)) {
    throw new Error('E-mail inválido: "' + email + '".');
  }

  const lista = obterEmailsSalvos_();
  if (lista.indexOf(emailNormalizado) === -1) {
    lista.push(emailNormalizado);
    lista.sort();
    gravarEmailsSalvos_(lista);
  }
  return lista;
}

function removerEmailSalvo_(email) {
  const emailNormalizado = String(email || '').trim().toLowerCase();
  const lista = obterEmailsSalvos_().filter(function (existente) {
    return existente !== emailNormalizado;
  });
  gravarEmailsSalvos_(lista);
  return lista;
}

// HTML de e-mail com tabelas + estilo inline (não CSS externo, não
// gradiente) de propósito — é o que funciona de forma consistente no
// maior número de clientes de e-mail (Gmail, Outlook, etc.), que têm
// suporte bem mais limitado a CSS moderno do que um navegador.
function construirEmailEtiquetas_(itens, rotuloPeriodo) {
  const totalUnidades = itens.reduce(function (soma, item) {
    return soma + (Number(item.total) || 0);
  }, 0);

  const linhas = itens
    .map(function (item, indice) {
      const corFundo = indice % 2 === 0 ? '#17233d' : '#111a2e';
      return (
        '<tr>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#e6ecff;font-size:13px;">🚫 ' +
        escaparHtml_(item.produto) +
        '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#e6ecff;font-size:13px;text-align:center;font-weight:bold;">' +
        fmtNumero_(item.total) +
        '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#93a2c7;font-size:13px;text-align:center;">' +
        fmtNumero_(item.individual) +
        '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #263353;background:' +
        corFundo +
        ';color:#93a2c7;font-size:13px;text-align:center;">' +
        fmtNumero_(item.kit) +
        '</td>' +
        '</tr>'
      );
    })
    .join('');

  return (
    '<div style="background:#0b1220;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#111a2e;border-radius:14px;overflow:hidden;border:1px solid #263353;">' +
    // Cabeçalho
    '<tr><td style="background:#c2410c;padding:32px 32px 28px 32px;">' +
    '<div style="font-size:30px;line-height:1.2;">🏷️</div>' +
    '<div style="font-size:22px;font-weight:bold;color:#ffffff;margin-top:10px;">Relatório de Etiquetas Pendentes</div>' +
    '<div style="font-size:14px;color:#ffe4d1;margin-top:4px;">Essência do Brasil &middot; Painel de Vendas &amp; Estoque</div>' +
    '</td></tr>' +
    // Resumo
    '<tr><td style="padding:28px 32px 8px 32px;">' +
    '<div style="color:#e6ecff;font-size:14px;line-height:1.6;">' +
    'Os produtos abaixo venderam no período <strong>' +
    escaparHtml_(rotuloPeriodo) +
    '</strong> e não têm etiqueta disponível no estoque.' +
    '</div>' +
    '</td></tr>' +
    // Cards de resumo
    '<tr><td style="padding:16px 32px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td width="50%" style="background:#1f2937;border-radius:10px;padding:16px;text-align:center;">' +
    '<div style="color:#93a2c7;font-size:11px;text-transform:uppercase;letter-spacing:.04em;">Produtos sem etiqueta</div>' +
    '<div style="color:#fca5a5;font-size:26px;font-weight:bold;margin-top:6px;">' +
    itens.length +
    '</div>' +
    '</td>' +
    '<td width="16">&nbsp;</td>' +
    '<td width="50%" style="background:#1f2937;border-radius:10px;padding:16px;text-align:center;">' +
    '<div style="color:#93a2c7;font-size:11px;text-transform:uppercase;letter-spacing:.04em;">Unidades vendidas</div>' +
    '<div style="color:#fdba74;font-size:26px;font-weight:bold;margin-top:6px;">' +
    fmtNumero_(totalUnidades) +
    '</div>' +
    '</td>' +
    '</tr></table>' +
    '</td></tr>' +
    // Tabela
    '<tr><td style="padding:8px 32px 24px 32px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-radius:10px;overflow:hidden;border:1px solid #263353;">' +
    '<tr>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;text-align:left;">Produto</th>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;">Total</th>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;">Individual</th>' +
    '<th style="padding:10px 16px;background:#1f2937;color:#93a2c7;font-size:11px;text-transform:uppercase;">Kit</th>' +
    '</tr>' +
    linhas +
    '</table>' +
    '</td></tr>' +
    // Aviso
    '<tr><td style="padding:0 32px 28px 32px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#4a3c12;border-radius:10px;border:1px solid #eab308;">' +
    '<tr><td style="padding:16px 18px;color:#fde68a;font-size:12.5px;line-height:1.6;">' +
    '⚠️ Atenção: este relatório é baseado no estoque de etiquetas registrado no momento da consulta. Verifique se o item não está em produção antes de solicitar a impressão, para evitar duplicidade.' +
    '</td></tr>' +
    '</table>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</div>'
  );
}

function fmtNumero_(valor) {
  return Number(valor || 0).toLocaleString('pt-BR');
}

function escaparHtml_(texto) {
  return String(texto)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
