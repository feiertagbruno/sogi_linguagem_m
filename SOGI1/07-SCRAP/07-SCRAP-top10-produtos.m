let
    Fonte = #"07-SCRAP-perda-processo",
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([filtro_semana] <> 0)),
    
    // Agrupar as colunas e somar o CUSTO
    #"Linhas Agrupadas" = Table.Group(#"Linhas Filtradas", 
        {"semana_texto", "WK", "PRODUTO", "DESCRIÇÃO"}, 
        {{"CUSTO", each List.Sum([CUSTO]), type nullable number}}),

    // Ordenar as linhas por CUSTO em ordem decrescente
    #"Linhas Ordenadas" = Table.Sort(#"Linhas Agrupadas", {{"CUSTO", Order.Descending}}),

    // Agrupar novamente por semana e selecionar as 10 maiores
    #"Top 10 por Semana" = Table.Group(#"Linhas Ordenadas", 
        {"semana_texto", "WK"}, 
        {{"Top 10", each Table.FirstN(_, 10), type table [semana_texto=nullable text, WK=nullable number, PRODUTO=nullable text, DESCRIÇÃO=nullable text, CUSTO=nullable number]}}),

    // Expandir a tabela novamente
    #"Dados Expandidos" = Table.ExpandTableColumn(#"Top 10 por Semana", "Top 10", {"PRODUTO", "DESCRIÇÃO", "CUSTO"}, {"PRODUTO", "DESCRIÇÃO", "CUSTO"}),

    maior_semana = List.Max(Table.Column(#"Dados Expandidos", "WK")),
    filtro_ultima_semana = Table.AddColumn(#"Dados Expandidos", "filtro_ultima_semana", 
        each if [WK] = maior_semana then "Última Semana" else "Outras Semanas" , type text),
    rel_produto = Table.AddColumn(filtro_ultima_semana,"rel_produto", each Text.From([WK]) & "-" & Text.Trim([PRODUTO]), type text)
in
    rel_produto