PROMPT — Dashboard de Vendas e Estoque Inteligente | Essência do Brasil

Contexto: A Essência do Brasil vende perfumes, óleos essenciais, óleos vegetais, essências e roll-on em múltiplos marketplaces, muitas vezes agrupados em kits dentro de um único anúncio. O objetivo é construir automação diária (Google Apps Script) que alimenta uma base de dados no Google Sheets, e um dashboard em HTML, aberto no navegador, que lê essa base — sem nenhuma análise de preço/faturamento, e sem dados de cliente.

Planilha (só banco de dados, sem dashboard nativo dentro dela): a planilha Google Sheets já existente "Análise e Controle" (conta victor@gigaimports.com), que hoje só tem a aba Base_Dados (cópia exata do Análise e Controle.xlsx do repositório apps-script-work, 6.144 linhas: coluna A = título completo do anúncio, coluna B = produto(s) canônico(s) separados por ; quando é kit). O .xlsx do repositório passa a ser só histórico estático; a fonte viva do Base_Dados é essa planilha Google, mantida manualmente pelo cliente. O sistema deve sinalizar títulos de pedido sem correspondência no Base_Dados, e o matching é case-insensitive.

Fontes externas — só leitura, nunca escrever nelas:

"Teste Fase 2 Vendas Essência do Brasil" (victor@gigaimports.com) → aba "Mês Atual". Já alimentada por integração existente com Bling ERP (Apps Script já vinculado: getAccessToken() + importarPedidosUltimos3Dias(), mapeando 18 lojas/marketplaces). Colunas: Data - Venda (bug conhecido: na verdade contém código de produto do Bling, não uma data — ignorar), Número do pedido multiloja, Loja, Cliente, Quantidade, Descrição, Total dos Itens/Responsável/Valor Líquido (não usar, fora de escopo).
"Etiquetas" (victor@gigaimports.com) → aba "Inventário" (não confundir com a aba "Validade" da mesma planilha): indica se há rótulo/etiqueta disponível por produto (presença binária por pasta/caixa física). Fonte para identificar "produto vendido sem etiqueta". O repositório GitHub victor-bot-01/apps-script-temp contém o Apps Script/web app "Sistema Essência v2.0" por trás dessa planilha — usar só como referência de padrão visual, não alterar, e não reaproveitar o projeto (o dashboard novo é separado dele).
Abas novas a criar dentro de "Análise e Controle" (banco de dados puro):

Pedidos: Data | Pedido | Loja | Cliente | Quantidade | Descrição — alimentada automaticamente, cópia simplificada dos pedidos novos de "Mês Atual" (sem as colunas de valor).
Estoque: Produto (canônico) | Quantidade | Última Atualização — preenchimento manual pelo cliente, layout simples por ora (pode evoluir depois).
Ficha Técnica: uma linha por produto canônico com Gênero, Coleção, Família Olfativa, Categoria (perfume/óleo essencial/óleo vegetal/essência/roll-on), Volumetria — preenchimento manual.
Log de Execução: horário de início/fim, duração, qtd de pedidos importados, erros e avisos de cada rodada.
Regras de negócio:

Escopo só produtos Essência do Brasil — excluir joias da Ponte Vecchio mesmo que apareçam no Base_Dados.
Nenhuma análise de preço, faturamento ou ticket médio — só quantidade/unidades.
Sem dados nem análises de cliente.
Data de cada pedido novo = dia anterior ao dia da execução (não existe data confiável na fonte). Isso funciona porque a fonte externa só recebe pedidos do dia corrente a partir das 6h30, e a rotina sempre roda antes disso.
Deduplicação por pedido + produto (não por "último ID", pois os números de pedido não são um contador único entre os 18 marketplaces).
Automação:

Execução principal às 02:00; tentativas de segurança às 02:30, 03:00, 03:30 e 04:00, cada uma abortando se a importação do dia já foi concluída.
Se todas as tentativas de um dia falharem: processa o acumulado no dia seguinte, registra aviso explícito no Log de Execução de que o lote pode conter mais de 1 dia de pedidos (ajuste manual de data fica a cargo do cliente).
Alerta por e-mail em caso de falha, para victor@gigaimports.com.
Backup automático da base antes de cada atualização; validação contra pedidos duplicados; botão de atualização manual.
Dashboard (frontend):

Página HTML própria (Apps Script Web App, doGet()), aberta no navegador — não é dashboard dentro do Sheets, sem gráfico/tabela dinâmica nativa lá.
Projeto novo e separado do "Sistema Essência" (apps-script-temp) — visual no mesmo estilo (sidebar, cards, tema), mas independente.
Código do Apps Script vinculado à planilha "Análise e Controle" (Extensões > Apps Script de lá), sem espelho no GitHub.
Acesso restrito a um único usuário: victor@gigaimports.com.
"Filtros" e "formatação condicional" viram, na prática, dropdowns/seletores de período/marketplace/categoria na tela e cores/badges de alerta em HTML/CSS — não recursos nativos do Sheets.
Funcionalidades analíticas (todas por quantidade, nunca por valor):

KPIs: pedidos no período, itens vendidos, produtos ativos/sem venda, estoque total, produtos em risco de ruptura, kits vs. individuais, produto campeão, marketplace campeão.
Rankings: top 10 produtos, ranking de kits, crescimento vs. período anterior (semana/mês/mesmo mês ano anterior/últimos 12 meses), semanas do mês com mais vendas.
Kit vs. individual: produto vendeu mais sozinho ou em kit; produtos só-kit/só-individual/ambos; produto que mais participa de kits; kit mais e menos vendido.
Estoque inteligente: níveis mínimo/ideal/segurança/máximo (base 70% de probabilidade de venda semana/mês); dias restantes até acabar; produtos em risco de ruptura; produtos em excesso; produtos parados; dias médios para vender 1 unidade; giro médio mensal/semanal.
Previsão: próxima semana, 15 dias, próximo mês, 3 e 6 meses; intervalo de confiança; média móvel; tendência linear; crescimento percentual; sugestão de quantidade a produzir.
Marketplace: quantidade vendida e produtos/kits/perfumes/essências/óleos mais vendidos, por marketplace.
Categorias (via Ficha Técnica): ranking por categoria; perfumes por gênero/coleção/família olfativa/volumetria; óleo essencial x vegetal, tamanho e misturas mais vendidas; participação de cada categoria de essência (frutada, doce, amadeirada, floral, cítrica, gourmet).
Calendário/sazonalidade: heatmap de dias com mais vendas; vendas por dia da semana; produto sazonal/em crescimento/em queda/que voltou a vender; produtos sem venda em 30/60/90/180 dias; produtos com perda de vendas.
Alertas inteligentes automáticos (ex.: "produto X acima da média", "produto Y caiu N%", "estoque acaba em N dias", "produto sem etiqueta", "kit crescendo rapidamente").
Inventário rápido: separar produtos embalados/rotulados/sem rótulo/prontos/em produção/aguardando embalagem-etiqueta; gerar etiquetas pequenas tipo checklist (☐ Produto) na quantidade sugerida de produção; cruzar com a aba "Inventário" da planilha "Etiquetas".
Relatórios automáticos em PDF: resumo semanal, resumo mensal, produtos críticos, produtos para produzir, produtos sem etiqueta, estoque.
Visual do dashboard HTML: página inicial com cards de KPI e gráficos interativos; menu lateral; paleta escura com detalhes em azul/verde/laranja; ícones; gráficos de barra/linha/pizza/heatmap/gauge; formatação condicional para alertas/tendências/metas. Visual livre, desde que profissional e moderno.
