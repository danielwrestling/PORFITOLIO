# PORFITOLIO
Atividades em dados realizadas.
🏋️ Dashboard de Gestão de Academia — Power BI
Dashboard desenvolvido para a gestão completa de uma academia, consolidando indicadores de retenção, cancelamentos, vendas e engajamento em um único painel. O grande diferencial do projeto é a automação total do tratamento de dados: basta inserir as planilhas exportadas do sistema da academia, sem qualquer manipulação manual — todo o tratamento acontece via Power Query e DAX.
---
🎯 Objetivo
Eliminar o trabalho manual de tratar planilhas mês a mês, dando à gestão da academia uma visão rápida e confiável sobre:
Quantos alunos estão ativos e em dia com o pagamento
Por que os alunos estão cancelando
Quem está vendendo mais no balcão
Qual a taxa de churn do período
Como está o engajamento nos agregadores (Totalpass e Wellhub)
---
📊 Painéis do Dashboard
1. Ativos
Cartão com o total de alunos ativos, segmentados em Adimplentes e Inadimplentes.
2. Cancelamentos
Gráfico de pizza com os motivos de cancelamento, com destaque específico para o percentual de cancelamentos causados por inadimplência.
3. Vendas
Gráfico de colunas com o volume de vendas por consultor, considerando exclusivamente vendas realizadas no balcão (equipe de recepção).
4. Churn
Cartão com o cálculo da taxa de churn do período, cruzando ativos e cancelamentos.
5. Agregadores (Totalpass / Wellhub)
Total de check-ins por plataforma
Ranking de quem mais fez check-in em cada uma
Identificação de alunos que fizeram apenas 1 check-in no mês (sinal de baixo engajamento/risco de cancelamento)
---
🔧 Como funciona por trás
Todo o pipeline de dados foi construído para ser plug-and-play:
As planilhas são exportadas do sistema da academia (Alunos, Cancelamentos, Vendas, Check-ins)
O Power Query identifica e padroniza automaticamente: status dos alunos, motivos de cancelamento, filtro de vendas por canal/consultor e período de check-ins
As medidas DAX calculam os indicadores (churn, adimplência, rankings) em tempo real, sem necessidade de reprocessamento manual
O código completo das queries M e das medidas DAX está documentado em `dashboard-academia-powerbi.md`.
---
🛠️ Stack
`Power BI` · `Power Query (M)` · `DAX` · `Excel`
---
📌 Principais aprendizados do projeto
Estruturação de pipelines de ETL dentro do próprio Power BI, eliminando etapas manuais de tratamento
Modelagem de indicadores de negócio (churn, adimplência, benchmarking entre vendedores) a partir de dados brutos
Construção de lógicas de segmentação e ranking em DAX (ex: identificar o aluno com mais check-ins do mês)
Padronização de dados textuais inconsistentes (grafias diferentes de status/motivos) direto na camada de transformação
