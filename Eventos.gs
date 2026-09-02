/**
 * Aba "Eventos": lembretes de marketing/operacional com envio automático
 * de e-mail. Guardados na aba "Eventos" da própria planilha (criada
 * automaticamente na primeira vez que qualquer função daqui roda).
 *
 * Dois tipos:
 * - "Personalizado": data anual fixa (dia + mês, sem ano — recorre todo
 *   ano no mesmo dia/mês). Dispara DOIS e-mails: um "faltam 7 dias" e um
 *   "hoje é o dia". Pensado pra datas de marketing (Dia das Mães, Black
 *   Friday etc.).
 * - "Simples": recorrência MENSAL (só dia do mês, sem mês fixo — repete
 *   todo mês nesse dia). Dispara UM e-mail com o texto livre cadastrado.
 *   Pensado pra lembretes operacionais recorrentes (ex.: "fazer relatório
 *   de etiquetas dia 20").
 *
 * Datas móveis (Dia das Mães/Pais, Black Friday, Semana do Brasil, Dia do
 * Frete Grátis, Carnaval, Páscoa, Enem) NÃO têm cálculo automático (nem de
 * "2º domingo do mês" nem da Páscoa/INEP) — ficam com o dia fixo do ano em
 * que foram cadastradas (decisão consciente, ver conversa). Por isso
 * existe um evento Personalizado próprio (1º de dezembro) lembrando de
 * reajustar essas datas todo ano.
 *
 * "A definir" (dia/mês 0): eventos sem data confirmada ainda (Copa do
 * Mundo, Olimpíadas, Rock in Rio, G20 etc.) — dia/mês 0 é um sentinela
 * seguro: `processarEventoPersonalizado_` sai antes de calcular qualquer
 * data (sem esse corte, `new Date(ano, mes-1, 0)` rolaria pro último dia
 * do mês anterior, uma data REAL que eventualmente bateria com "hoje" e
 * dispararia e-mail por engano). Também nunca aparecem no calendário
 * mensal (dia/mês fora do intervalo 1-31/1-12 nunca bate com nenhuma
 * célula), só na tabela "Todos os eventos cadastrados" do Web App — que
 * tem um botão "Copiar Eventos Sem Data" (copia todos de uma vez, não um
 * por um) pra colar no chat e confirmar a data real quando descobrir.
 *
 * `executarVerificacaoEventos` é o handler do gatilho diário (ver
 * Gatilhos.gs: configurarGatilhosEventos, 15h) — idempotente por dia/mês
 * via as colunas de controle "Último Aviso Prévio" / "Último Aviso No
 * Dia" / "Último Envio Mensal", pro caso do gatilho disparar mais de uma
 * vez no mesmo dia (Apps Script não garante disparo único exato).
 */

const NOME_ABA_EVENTOS = 'Eventos';
const EMAIL_DESTINATARIO_EVENTOS = 'victor@gigaimports.com';
const CABECALHO_EVENTOS = [
  'ID',
  'Título',
  'Tipo',
  'Dia',
  'Mês',
  'Emoji',
  'Mensagem',
  'Último Aviso Prévio',
  'Último Aviso No Dia',
  'Último Envio Mensal',
];
const COL_EVENTOS_ULTIMO_AVISO_PREVIO = 8;
const COL_EVENTOS_ULTIMO_AVISO_NO_DIA = 9;
const COL_EVENTOS_ULTIMO_ENVIO_MENSAL = 10;

function obterOuCriarAbaEventos_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA_EVENTOS);
  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA_EVENTOS);
    aba
      .getRange(1, 1, 1, CABECALHO_EVENTOS.length)
      .setValues([CABECALHO_EVENTOS])
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#ffffff');
    aba.setFrozenRows(1);
  }
  return aba;
}

function lerEventosBrutos_() {
  const aba = obterOuCriarAbaEventos_();
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return [];

  return aba
    .getRange(2, 1, ultimaLinha - 1, CABECALHO_EVENTOS.length)
    .getValues()
    .map(function (linha, indice) {
      return {
        linha: indice + 2,
        id: linha[0],
        titulo: linha[1],
        tipo: linha[2],
        dia: Number(linha[3]),
        mes: Number(linha[4]) || null,
        emoji: linha[5],
        mensagem: linha[6],
        ultimoAvisoPrevio: linha[7] instanceof Date ? linha[7] : null,
        ultimoAvisoNoDia: linha[8] instanceof Date ? linha[8] : null,
        ultimoEnvioMensal: linha[9] instanceof Date ? linha[9] : null,
      };
    })
    .filter(function (evento) {
      return evento.id;
    });
}

// Formato exposto ao Web App — sem os campos internos de controle de envio.
function obterEventos_() {
  return lerEventosBrutos_().map(function (evento) {
    return {
      id: evento.id,
      titulo: evento.titulo,
      tipo: evento.tipo,
      dia: evento.dia,
      mes: evento.mes,
      emoji: evento.emoji,
      mensagem: evento.mensagem,
    };
  });
}

function criarEvento_(dados) {
  const titulo = String((dados && dados.titulo) || '').trim();
  const tipo = dados && dados.tipo === 'Simples' ? 'Simples' : 'Personalizado';
  const dia = Number(dados && dados.dia);

  if (!titulo) throw new Error('Título do evento é obrigatório.');
  if (!dia || dia < 1 || dia > 31) throw new Error('Dia inválido.');

  let mes = '';
  let emoji = '';
  let mensagem = '';

  if (tipo === 'Personalizado') {
    mes = Number(dados && dados.mes);
    if (!mes || mes < 1 || mes > 12) throw new Error('Mês inválido.');
    emoji = String((dados && dados.emoji) || '').trim();
  } else {
    mensagem = String((dados && dados.mensagem) || '').trim();
    if (!mensagem) throw new Error('Mensagem é obrigatória para eventos do tipo Simples.');
  }

  const aba = obterOuCriarAbaEventos_();
  const id = Utilities.getUuid();
  aba.appendRow([id, titulo, tipo, dia, mes, emoji, mensagem, '', '', '']);

  return obterEventos_();
}

function removerEvento_(id) {
  const aba = obterOuCriarAbaEventos_();
  const evento = lerEventosBrutos_().filter(function (existente) {
    return existente.id === id;
  })[0];
  if (evento) aba.deleteRow(evento.linha);
  return obterEventos_();
}

// Handler do gatilho diário (ver Gatilhos.gs). Também pode ser rodada
// manualmente pelo menu "Essência do Brasil" pra testar sem esperar 15h.
function executarVerificacaoEventos() {
  const hoje = new Date();
  const hojeSemHora = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());

  lerEventosBrutos_().forEach(function (evento) {
    if (evento.tipo === 'Personalizado') {
      processarEventoPersonalizado_(evento, hojeSemHora);
    } else if (evento.tipo === 'Simples') {
      processarEventoSimples_(evento, hojeSemHora);
    }
  });
}

function processarEventoPersonalizado_(evento, hoje) {
  // Dia/mês 0 = "sem data definida ainda" (ver cadastrarEventosSugeridos_)
  // — sem esse corte, new Date(ano, mes-1, 0) rola pro último dia do mês
  // anterior (uma data REAL), que eventualmente bateria com "hoje" e
  // disparava e-mail por engano. Sai antes de calcular qualquer data.
  if (!evento.dia || !evento.mes) return;

  const dataEvento = new Date(hoje.getFullYear(), evento.mes - 1, evento.dia);
  const dataAviso = new Date(dataEvento.getFullYear(), dataEvento.getMonth(), dataEvento.getDate() - 7);

  if (mesmaData_(hoje, dataAviso) && !jaEnviadoEsteAno_(evento.ultimoAvisoPrevio, hoje)) {
    enviarEmailEventoPersonalizado_(evento, 7);
    atualizarUltimoEnvioEvento_(evento.linha, COL_EVENTOS_ULTIMO_AVISO_PREVIO, hoje);
  }

  if (mesmaData_(hoje, dataEvento) && !jaEnviadoEsteAno_(evento.ultimoAvisoNoDia, hoje)) {
    enviarEmailEventoPersonalizado_(evento, 0);
    atualizarUltimoEnvioEvento_(evento.linha, COL_EVENTOS_ULTIMO_AVISO_NO_DIA, hoje);
  }
}

function processarEventoSimples_(evento, hoje) {
  if (hoje.getDate() !== evento.dia) return;
  if (jaEnviadoEsteMes_(evento.ultimoEnvioMensal, hoje)) return;

  enviarEmailEventoSimples_(evento);
  atualizarUltimoEnvioEvento_(evento.linha, COL_EVENTOS_ULTIMO_ENVIO_MENSAL, hoje);
}

function mesmaData_(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function jaEnviadoEsteAno_(ultimoEnvio, hoje) {
  return ultimoEnvio instanceof Date && ultimoEnvio.getFullYear() === hoje.getFullYear();
}

function jaEnviadoEsteMes_(ultimoEnvio, hoje) {
  return (
    ultimoEnvio instanceof Date &&
    ultimoEnvio.getFullYear() === hoje.getFullYear() &&
    ultimoEnvio.getMonth() === hoje.getMonth()
  );
}

function atualizarUltimoEnvioEvento_(linha, coluna, data) {
  obterOuCriarAbaEventos_().getRange(linha, coluna).setValue(data);
}

function enviarEmailEventoPersonalizado_(evento, diasRestantes) {
  const emoji = evento.emoji || '🎉';
  const assunto =
    diasRestantes > 0
      ? emoji + ' Faltam ' + diasRestantes + ' dias para ' + evento.titulo + ' — Essência do Brasil'
      : emoji + ' Hoje é ' + evento.titulo + ' — Essência do Brasil';

  MailApp.sendEmail({
    to: EMAIL_DESTINATARIO_EVENTOS,
    subject: assunto,
    htmlBody: construirEmailEventoPersonalizado_(evento, diasRestantes),
  });
}

function construirEmailEventoPersonalizado_(evento, diasRestantes) {
  const emoji = evento.emoji || '🎉';
  const tituloMensagem =
    diasRestantes > 0 ? 'Faltam ' + diasRestantes + ' dias para ' + evento.titulo + '!' : 'Hoje é ' + evento.titulo + '!';
  const textoMensagem =
    diasRestantes > 0
      ? 'Prepare a campanha — é um evento que podemos considerar para marketing.'
      : 'O grande dia chegou — é um evento que podemos considerar para marketing.';

  return (
    '<div style="background:#0b1220;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#111a2e;border-radius:14px;overflow:hidden;border:1px solid #263353;">' +
    '<tr><td style="background:#1d4ed8;padding:36px 32px 32px 32px;text-align:center;">' +
    '<div style="font-size:44px;line-height:1.2;">' +
    emoji +
    '</div>' +
    '<div style="font-size:24px;font-weight:bold;color:#ffffff;margin-top:12px;">' +
    escaparHtml_(tituloMensagem) +
    '</div>' +
    '<div style="font-size:14px;color:#dbeafe;margin-top:6px;">Essência do Brasil &middot; Painel de Vendas &amp; Estoque</div>' +
    '</td></tr>' +
    '<tr><td style="padding:28px 32px 32px 32px;text-align:center;">' +
    '<div style="color:#e6ecff;font-size:15px;line-height:1.7;">' +
    escaparHtml_(textoMensagem) +
    '</div>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</div>'
  );
}

function enviarEmailEventoSimples_(evento) {
  MailApp.sendEmail({
    to: EMAIL_DESTINATARIO_EVENTOS,
    subject: '📌 Lembrete: ' + evento.titulo + ' — Essência do Brasil',
    htmlBody: construirEmailEventoSimples_(evento),
  });
}

function construirEmailEventoSimples_(evento) {
  return (
    '<div style="background:#0b1220;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#111a2e;border-radius:14px;overflow:hidden;border:1px solid #263353;">' +
    '<tr><td style="background:#15803d;padding:28px 32px;">' +
    '<div style="font-size:26px;">📌</div>' +
    '<div style="font-size:20px;font-weight:bold;color:#ffffff;margin-top:8px;">' +
    escaparHtml_(evento.titulo) +
    '</div>' +
    '<div style="font-size:13px;color:#dcfce7;margin-top:4px;">Essência do Brasil &middot; Lembrete mensal</div>' +
    '</td></tr>' +
    '<tr><td style="padding:28px 32px;">' +
    '<div style="color:#e6ecff;font-size:14.5px;line-height:1.7;white-space:pre-wrap;">' +
    escaparHtml_(evento.mensagem) +
    '</div>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</div>'
  );
}

// Utilitário manual (menu "Essência do Brasil"): cadastra as datas
// sugeridas de marketing + o lembrete de reajuste das datas móveis.
// Idempotente — pula qualquer título que já exista, então pode rodar de
// novo sem duplicar. Datas de Dia das Mães/Pais/Black Friday e as 4
// estações calculadas para 2026; precisam de ajuste manual anual (é
// literalmente o que o evento "Ajustar datas móveis" lembra de fazer).
function cadastrarEventosSugeridos() {
  const existentes = lerEventosBrutos_().map(function (evento) {
    return evento.titulo;
  });

  const sugeridos = [
    { titulo: 'Dia das Mães', dia: 10, mes: 5, emoji: '💐' },
    { titulo: 'Dia dos Pais', dia: 9, mes: 8, emoji: '👨‍👧‍👦' },
    { titulo: 'Dia dos Avós', dia: 26, mes: 7, emoji: '👴👵' },
    { titulo: 'Dia dos Namorados', dia: 12, mes: 6, emoji: '❤️' },
    { titulo: 'Dia Internacional da Mulher', dia: 8, mes: 3, emoji: '🌸' },
    { titulo: 'Black Friday', dia: 27, mes: 11, emoji: '🛍️' },
    { titulo: 'Natal', dia: 25, mes: 12, emoji: '🎄' },
    { titulo: 'Dia do Cliente', dia: 15, mes: 9, emoji: '🙏' },
    { titulo: 'Dia do Amigo', dia: 20, mes: 7, emoji: '🤝' },
    { titulo: 'Semana do Consumidor', dia: 15, mes: 3, emoji: '🛒' },
    { titulo: 'Dia Mundial do Meio Ambiente', dia: 5, mes: 6, emoji: '🌿' },
    { titulo: 'Início do Outono', dia: 20, mes: 3, emoji: '🍂' },
    { titulo: 'Início do Inverno', dia: 21, mes: 6, emoji: '❄️' },
    { titulo: 'Início da Primavera', dia: 23, mes: 9, emoji: '🌱' },
    { titulo: 'Início do Verão', dia: 21, mes: 12, emoji: '☀️' },
    { titulo: 'Ano Novo', dia: 1, mes: 1, emoji: '🎆' },
    { titulo: 'Dia da Terra', dia: 22, mes: 4, emoji: '🌎' },
    { titulo: 'Outubro Rosa', dia: 1, mes: 10, emoji: '🎗️' },
    { titulo: 'Dia Mundial da Saúde', dia: 7, mes: 4, emoji: '💚' },
    { titulo: 'Dia Internacional da Yoga', dia: 21, mes: 6, emoji: '🧘' },
    { titulo: 'Semana do Brasil', dia: 1, mes: 9, emoji: '🇧🇷' },
    { titulo: 'Dia do Frete Grátis', dia: 28, mes: 4, emoji: '📦' },
    { titulo: 'Dia de São Valentim', dia: 14, mes: 2, emoji: '💘' },
    { titulo: 'Dia do Artesão', dia: 19, mes: 3, emoji: '🎨' },
    { titulo: 'Dia do Solteiro', dia: 11, mes: 11, emoji: '🎁' },
    { titulo: 'Dia da Árvore', dia: 21, mes: 9, emoji: '🌳' },
    { titulo: 'Consciência Negra', dia: 20, mes: 11, emoji: '✊🏾' },
    { titulo: 'Dia Mundial do Veganismo', dia: 1, mes: 11, emoji: '🥑' },
    { titulo: 'Boxing Day', dia: 26, mes: 12, emoji: '🎁' },
    { titulo: 'Dia do Professor', dia: 15, mes: 10, emoji: '🍎' },
    { titulo: 'Dia Internacional da Família', dia: 15, mes: 5, emoji: '👨‍👩‍👧‍👦' },
    { titulo: 'Dia Internacional da Pessoa Idosa', dia: 1, mes: 10, emoji: '👵👴' },
    { titulo: 'Dia da Saudade', dia: 30, mes: 1, emoji: '🥺' },
    { titulo: 'Dia Nacional do Café', dia: 24, mes: 5, emoji: '☕' },
    // Pedra do mês (nascimento) — 12 eventos anuais, dia 1 de cada mês.
    { titulo: 'Pedra do Mês — Granada (Janeiro)', dia: 1, mes: 1, emoji: '💎' },
    { titulo: 'Pedra do Mês — Ametista (Fevereiro)', dia: 1, mes: 2, emoji: '💎' },
    { titulo: 'Pedra do Mês — Água-marinha (Março)', dia: 1, mes: 3, emoji: '💎' },
    { titulo: 'Pedra do Mês — Quartzo Cristal (Abril)', dia: 1, mes: 4, emoji: '💎' },
    { titulo: 'Pedra do Mês — Esmeralda (Maio)', dia: 1, mes: 5, emoji: '💎' },
    { titulo: 'Pedra do Mês — Pedra da Lua (Junho)', dia: 1, mes: 6, emoji: '💎' },
    { titulo: 'Pedra do Mês — Rubi (Julho)', dia: 1, mes: 7, emoji: '💎' },
    { titulo: 'Pedra do Mês — Peridoto (Agosto)', dia: 1, mes: 8, emoji: '💎' },
    { titulo: 'Pedra do Mês — Safira (Setembro)', dia: 1, mes: 9, emoji: '💎' },
    { titulo: 'Pedra do Mês — Turmalina (Outubro)', dia: 1, mes: 10, emoji: '💎' },
    { titulo: 'Pedra do Mês — Citrino (Novembro)', dia: 1, mes: 11, emoji: '💎' },
    { titulo: 'Pedra do Mês — Turquesa (Dezembro)', dia: 1, mes: 12, emoji: '💎' },
    // Data móvel (Páscoa - 47 dias) — seed calculado pra 2027 (9/2), ajusta
    // todo ano junto com o lembrete abaixo.
    { titulo: 'Carnaval — Perfumes para encontros', dia: 9, mes: 2, emoji: '🎭' },
    // Também data móvel (dia da Páscoa em si) — seed calculado pra 2027 (28/3).
    { titulo: 'Páscoa', dia: 28, mes: 3, emoji: '🐰' },
    { titulo: 'Festa Junina', dia: 24, mes: 6, emoji: '🌽' },
    { titulo: 'Dia do Livro', dia: 23, mes: 4, emoji: '📖' },
    { titulo: 'Proclamação da República', dia: 15, mes: 11, emoji: '🇧🇷' },
    // Data muda todo ano (calendário divulgado pelo INEP) — seed
    // aproximado (1ª quinzena de novembro), ajusta todo ano junto com o
    // lembrete abaixo, igual Carnaval/Páscoa.
    { titulo: 'Dia Nacional do Enem', dia: 8, mes: 11, emoji: '📝' },
    // Sem data oficial única amplamente confirmada — usando a referência
    // mais próxima encontrada (27/9); ajuste se souber uma data melhor.
    { titulo: 'Dia do E-commerce', dia: 27, mes: 9, emoji: '💻' },

    // Série de vendas espelhadas (1.1 a 12.12) — 1/1 e 11/11 já têm outro
    // evento (Ano Novo, Dia do Solteiro); ficam como eventos extras no
    // mesmo dia, de propósito (o calendário já suporta vários por dia).
    { titulo: 'Venda Espelhada 1.1', dia: 1, mes: 1, emoji: '🔥' },
    { titulo: 'Venda Espelhada 2.2', dia: 2, mes: 2, emoji: '🔥' },
    { titulo: 'Venda Espelhada 3.3', dia: 3, mes: 3, emoji: '🔥' },
    { titulo: 'Venda Espelhada 4.4', dia: 4, mes: 4, emoji: '🔥' },
    { titulo: 'Venda Espelhada 5.5', dia: 5, mes: 5, emoji: '🔥' },
    { titulo: 'Venda Espelhada 6.6', dia: 6, mes: 6, emoji: '🔥' },
    { titulo: 'Venda Espelhada 7.7', dia: 7, mes: 7, emoji: '🔥' },
    { titulo: 'Venda Espelhada 8.8', dia: 8, mes: 8, emoji: '🔥' },
    { titulo: 'Venda Espelhada 9.9', dia: 9, mes: 9, emoji: '🔥' },
    { titulo: 'Venda Espelhada 10.10', dia: 10, mes: 10, emoji: '🔥' },
    { titulo: 'Venda Espelhada 11.11', dia: 11, mes: 11, emoji: '🔥' },
    { titulo: 'Venda Espelhada 12.12', dia: 12, mes: 12, emoji: '🔥' },

    // "A definir" — dia/mês 0 (ver processarEventoPersonalizado_: nunca
    // dispara e-mail nem aparece no calendário mensal, só na lista "Todos
    // os eventos cadastrados", com o botão "Copiar Eventos Sem Data" pra
    // colar no chat e confirmar a data real quando você souber).
    { titulo: 'Copa do Mundo Feminina (a definir)', dia: 0, mes: 0, emoji: '⚽' },
    { titulo: 'Copa do Mundo Masculina (a definir)', dia: 0, mes: 0, emoji: '⚽' },
    { titulo: 'Olimpíadas (a definir)', dia: 0, mes: 0, emoji: '🏅' },
    { titulo: 'Rock in Rio (a definir)', dia: 0, mes: 0, emoji: '🎸' },
    { titulo: 'G20 (a definir)', dia: 0, mes: 0, emoji: '🌐' },
    { titulo: 'Todo Mundo em Copa — Show em Copacabana (a definir)', dia: 0, mes: 0, emoji: '🎤' },

    {
      titulo:
        'Ajustar datas móveis (Dia das Mães, Dia dos Pais, Black Friday, Semana do Brasil, Dia do Frete Grátis, Carnaval, Páscoa, Enem, 4 estações)',
      dia: 1,
      mes: 12,
      emoji: '🔄',
    },
  ];

  // Campanhas recorrentes mensais (tipo Simples) — dias escolhidos entre
  // os menos concorridos do calendário (contando quantos eventos já caem
  // em cada dia-do-mês, em qualquer mês), evitando dias sem 29/30/31 (que
  // falhariam silenciosamente em meses mais curtos).
  const simples = [
    { titulo: 'Lançamento do Mês', dia: 13, mensagem: 'Hora de planejar o lançamento do mês! 🚀' },
    { titulo: 'Pague 3 e Leve 4', dia: 18, mensagem: 'Não esqueça de rodar a campanha Pague 3 e Leve 4 este mês! 🛍️' },
  ];

  const aba = obterOuCriarAbaEventos_();
  sugeridos.forEach(function (sugestao) {
    if (existentes.indexOf(sugestao.titulo) !== -1) return;
    aba.appendRow([Utilities.getUuid(), sugestao.titulo, 'Personalizado', sugestao.dia, sugestao.mes, sugestao.emoji, '', '', '', '']);
  });
  simples.forEach(function (sugestao) {
    if (existentes.indexOf(sugestao.titulo) !== -1) return;
    aba.appendRow([Utilities.getUuid(), sugestao.titulo, 'Simples', sugestao.dia, '', '', sugestao.mensagem, '', '', '']);
  });
}
