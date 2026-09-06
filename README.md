Dashboard de Gestão de Academia — Power BI (100% Automatizado)
Projeto para consolidar indicadores de retenção, cancelamentos, vendas e engajamento de uma academia. O usuário final só precisa substituir as planilhas de origem — todo o tratamento e cálculo é automático via Power Query (M) e DAX.
---
📋 Indicadores entregues
Indicador	Descrição
Nº de Ativos	Total de alunos ativos
Nº de Inadimplentes	Alunos ativos em atraso
Nº de Adimplentes	Alunos ativos em dia
Vendas por Consultor	Vivian, Guilherme, Jonas, Erico (apenas vendas no balcão)
Nº de Cancelamentos	Total de cancelamentos no período
Motivos de Cancelamento	Distribuição por motivo
Nº de Check-ins Wellhub	Total de check-ins na plataforma
Top 3 Wellhub	Alunos com mais check-ins na Wellhub
Check-in único Wellhub	Qtd de alunos com apenas 1 check-in na Wellhub
Nº de Check-ins Totalpass	Total de check-ins na plataforma
Top 3 Totalpass	Alunos com mais check-ins na Totalpass
Check-in único Totalpass	Qtd de alunos com apenas 1 check-in na Totalpass
---
🗂️ Estrutura de pastas (fixa, para automação total)
```
C:\Dashboard Academia\
 ├── Alunos.xlsx
 ├── Cancelamentos.xlsx
 ├── Vendas.xlsx
 └── Checkins.xlsx
```
Regra de ouro: os nomes e o local dos arquivos nunca mudam. Todo mês, você só abre cada arquivo e substitui os dados (mesma aba, mesmas colunas), sem tocar em nada dentro do Power BI.
---
🔧 Parte 1 — Parâmetros do Power Query
No Power BI Desktop: Página Inicial → Transformar Dados → Gerenciar Parâmetros → Novo Parâmetro. Crie 4 parâmetros de texto:
Nome	Valor atual
`CaminhoArquivoAlunos`	`C:\Dashboard Academia\Alunos.xlsx`
`CaminhoArquivoCancelamentos`	`C:\Dashboard Academia\Cancelamentos.xlsx`
`CaminhoArquivoVendas`	`C:\Dashboard Academia\Vendas.xlsx`
`CaminhoArquivoCheckins`	`C:\Dashboard Academia\Checkins.xlsx`
Isso é o que permite ao Power Query sempre ler o arquivo certo, sem você reconfigurar nada.
---
🔧 Parte 2 — Queries Power Query (M)
Para cada tabela: Obter Dados → Excel → selecionar arquivo/aba → Editor Avançado → colar o código abaixo.
1. Alunos (Ativos / Adimplência)
```m
let
    Origem = Excel.Workbook(File.Contents(CaminhoArquivoAlunos), null, true),
    Base = Origem{[Item="Alunos",Kind="Sheet"]}[Data],
    CabecalhoPromovido = Table.PromoteHeaders(Base, [PromoteAllScalars=true]),
    TiposAlterados = Table.TransformColumnTypes(CabecalhoPromovido,{
        {"Aluno", type text}, {"Status", type text}, {"Plano", type text}
    }),
    StatusPadronizado = Table.TransformColumns(TiposAlterados,{
        {"Status", each Text.Proper(Text.Trim(_)), type text}
    }),
    ColunaSituacao = Table.AddColumn(StatusPadronizado, "Situacao", each 
        if Text.Contains(Text.Lower([Status]), "inadimpl") then "Inadimplente" else "Adimplente"
    )
in
    ColunaSituacao
```
2. Cancelamentos
```m
let
    Origem = Excel.Workbook(File.Contents(CaminhoArquivoCancelamentos), null, true),
    Base = Origem{[Item="Cancelamentos",Kind="Sheet"]}[Data],
    CabecalhoPromovido = Table.PromoteHeaders(Base, [PromoteAllScalars=true]),
    TiposAlterados = Table.TransformColumnTypes(CabecalhoPromovido,{
        {"Aluno", type text}, {"Motivo", type text}, {"Data Cancelamento", type date}
    }),
    MotivoPadronizado = Table.TransformColumns(TiposAlterados,{
        {"Motivo", each Text.Proper(Text.Trim(_)), type text}
    }),
    ColunaFlagInadimplencia = Table.AddColumn(MotivoPadronizado, "Motivo Inadimplencia", each 
        if Text.Contains(Text.Lower([Motivo]), "inadimpl") then "Sim" else "Não"
    )
in
    ColunaFlagInadimplencia
```
3. Vendas (filtrando Balcão + consultores válidos)
```m
let
    Origem = Excel.Workbook(File.Contents(CaminhoArquivoVendas), null, true),
    Base = Origem{[Item="Vendas",Kind="Sheet"]}[Data],
    CabecalhoPromovido = Table.PromoteHeaders(Base, [PromoteAllScalars=true]),
    TiposAlterados = Table.TransformColumnTypes(CabecalhoPromovido,{
        {"Consultor", type text}, {"Canal", type text}, {"Data Venda", type date}, {"Valor", type number}
    }),
    ConsultoresRecepcao = {"Vivian", "Guilherme", "Jonas", "Erico"},
    FiltroConsultor = Table.SelectRows(TiposAlterados, each List.Contains(ConsultoresRecepcao, [Consultor])),
    FiltroBalcao = Table.SelectRows(FiltroConsultor, each Text.Lower([Canal]) = "balcão" or Text.Lower([Canal]) = "balcao")
in
    FiltroBalcao
```
4. Check-ins (Totalpass / Wellhub)
```m
let
    Origem = Excel.Workbook(File.Contents(CaminhoArquivoCheckins), null, true),
    Base = Origem{[Item="Checkins",Kind="Sheet"]}[Data],
    CabecalhoPromovido = Table.PromoteHeaders(Base, [PromoteAllScalars=true]),
    TiposAlterados = Table.TransformColumnTypes(CabecalhoPromovido,{
        {"Aluno", type text}, {"Plataforma", type text}, {"Data Checkin", type date}
    }),
    PlataformaPadronizada = Table.TransformColumns(TiposAlterados,{
        {"Plataforma", each Text.Proper(Text.Trim(_)), type text}
    }),
    ColunaAnoMes = Table.AddColumn(PlataformaPadronizada, "AnoMes", each Date.ToText([Data Checkin], "yyyy-MM"))
in
    ColunaAnoMes
```
Depois de colar as 4 queries: Fechar e Aplicar.
---
📐 Parte 3 — Medidas DAX
Vá em Modelagem → Nova Medida e crie uma por uma:
Ativos
```dax
Total Ativos = COUNTROWS(Alunos)

Adimplentes = CALCULATE(COUNTROWS(Alunos), Alunos[Situacao] = "Adimplente")

Inadimplentes = CALCULATE(COUNTROWS(Alunos), Alunos[Situacao] = "Inadimplente")
```
Cancelamentos
```dax
Total Cancelamentos = COUNTROWS(Cancelamentos)
```
(Motivos de cancelamento não precisam de medida — o gráfico de pizza usa a coluna `Cancelamentos[Motivo]` direto, com `Total Cancelamentos` nos valores.)
Vendas por Consultor
```dax
Qtd Vendas por Consultor = COUNTROWS(Vendas)
```
(Usada em um gráfico de colunas com `Vendas[Consultor]` no eixo — já filtrado para Vivian, Guilherme, Jonas e Erico lá no Power Query.)
Check-ins — Totais
```dax
Checkins Wellhub = CALCULATE(COUNTROWS(Checkins), Checkins[Plataforma] = "Wellhub")

Checkins Totalpass = CALCULATE(COUNTROWS(Checkins), Checkins[Plataforma] = "Totalpass")
```
Check-ins — Contagem por Aluno (medida base para os rankings)
```dax
Checkins por Aluno = 
CALCULATE(
    COUNTROWS(Checkins),
    ALLEXCEPT(Checkins, Checkins[Aluno], Checkins[Plataforma])
)
```
Check-ins — Top 3 por plataforma
Crie uma tabela/matriz visual com `Checkins[Aluno]` e a medida `Checkins por Aluno`, filtrada por `Plataforma = Wellhub`. Aplique um filtro visual do tipo "Top N" na coluna Aluno:
Clique no visual → painel Filtros → arraste `Aluno` → escolha Filtro Top N → Top 3 → por valor de `Checkins por Aluno`
Repita o mesmo visual para `Plataforma = Totalpass`. Isso é mais simples e mais fácil de manter do que uma medida DAX de ranking, e atualiza sozinho conforme os dados mudam.
(Alternativa 100% em DAX, se preferir não usar filtro visual — crie uma medida de rank e filtre por ela em uma tabela calculada ou measure de exibição condicional; para a maioria dos casos o filtro Top N visual é suficiente e mais robusto.)
Check-ins — Alunos com apenas 1 check-in
```dax
Alunos 1 Checkin Wellhub = 
COUNTROWS(
    FILTER(
        SUMMARIZE(
            FILTER(Checkins, Checkins[Plataforma] = "Wellhub"),
            Checkins[Aluno],
            "QtdCheckins", [Checkins por Aluno]
        ),
        [QtdCheckins] = 1
    )
)

Alunos 1 Checkin Totalpass = 
COUNTROWS(
    FILTER(
        SUMMARIZE(
            FILTER(Checkins, Checkins[Plataforma] = "Totalpass"),
            Checkins[Aluno],
            "QtdCheckins", [Checkins por Aluno]
        ),
        [QtdCheckins] = 1
    )
)
```
---
🖼️ Parte 4 — Montagem dos visuais
Indicador	Visual sugerido
Nº Ativos / Adimplentes / Inadimplentes	3 Cartões lado a lado
Vendas por Consultor	Gráfico de Colunas (`Consultor` no eixo, `Qtd Vendas por Consultor` no valor)
Nº Cancelamentos	Cartão
Motivos de Cancelamento	Gráfico de Pizza (`Motivo` na legenda, `Total Cancelamentos` no valor)
Nº Check-ins Wellhub / Totalpass	2 Cartões
Top 3 Wellhub / Totalpass	2 Tabelas com filtro Top N (conforme explicado acima)
Check-in único Wellhub / Totalpass	2 Cartões (`Alunos 1 Checkin Wellhub` / `Alunos 1 Checkin Totalpass`)
---
🔄 Parte 5 — Automação total (zero edição de código, sempre)
Nível 1 — Atualização manual com 1 clique
Sempre salve as novas exportações substituindo os arquivos na mesma pasta, com os mesmos nomes
Abra o Power BI Desktop → clique em Atualizar na faixa Página Inicial
Todos os cartões, gráficos e Top N recalculam sozinhos — nenhuma query ou medida precisa ser tocada
Nível 2 — Atualização automática agendada (sem precisar abrir nada)
Publique o relatório: Página Inicial → Publicar → escolher workspace no Power BI Service (app.powerbi.com)
Instale o Gateway de Dados Local (On-premises Data Gateway) na máquina onde ficam os arquivos — [download oficial da Microsoft]
No Power BI Service: Configurações do Dataset → Conexões de Gateway → associe o gateway instalado
Em Configurações do Dataset → Atualização Agendada → ative e defina a frequência (ex: diariamente às 7h)
A partir daí, basta substituir os arquivos na pasta — o Power BI puxa e atualiza sozinho, sem qualquer intervenção
---
🛠️ Stack
`Power BI` · `Power Query (M)` · `DAX` · `Excel` · `On-premises Data Gateway`
