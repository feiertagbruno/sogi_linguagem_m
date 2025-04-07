let
    Fonte = #"08-OEE-agrup-semana",
    Fonte_sort = Table.Sort(Fonte, {{"semana_num", Order.Ascending}, {"OEE", Order.Ascending}}),
    agrupamento_semana_maquina = Table.Group(
        Fonte_sort, {"semana_num", "semana_texto", "semana_filtro"}, {{"agrupamento", each _ }}
    ),
    indice_por_semana = Table.TransformColumns(agrupamento_semana_maquina, {
        "agrupamento", each Table.AddIndexColumn(_, "índice", 1, 1, Int64.Type)
    }),
    expande_agrupamento = Table.ExpandTableColumn(indice_por_semana, "agrupamento", 
        {"índice", "Máquina", "performance", "qualidade", "ocupação", "OEE", "target"}),
    filtra_6_primeiros = Table.SelectRows(expande_agrupamento, each List.Contains({1,2,3,4,5,6}, [índice])),
    #"Tipo Alterado" = Table.TransformColumnTypes(filtra_6_primeiros,{{"índice", Int64.Type}, {"Máquina", Int64.Type}, {"performance", type number}, {"qualidade", type number}, {"ocupação", type number}, {"OEE", type number}, {"target", Int64.Type}}),

    
    semana_maquina = Table.AddColumn(#"Tipo Alterado", "semana_maquina", 
        each Text.From([semana_num]) & "-" & Text.From([Máquina]), type text)

in
    semana_maquina