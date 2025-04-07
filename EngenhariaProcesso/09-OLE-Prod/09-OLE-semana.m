let
    Fonte = #"09-OLE-base",
    #"Linhas Agrupadas" = Table.Group(Fonte, {"semana_num", "Processo"}, {
        {"Produção Planejada", each List.Sum([Produção Planejada]), type nullable number}, 
        {"Produção Realizada", each List.Sum([Produção Realizada]), type nullable number}, 
        {"Defeitos", each List.Sum([Defeitos]), type nullable number}, 
        {"Diferença", each List.Sum([Diferença]), type number}, 
        {"horas_planejado", each List.Sum([horas_planejado]), type number}, 
        {"horas_realizado", each List.Sum([horas_realizado]), type number}}),
    ocupação = Table.AddColumn(#"Linhas Agrupadas", "ocupação", 
    each try
        if [horas_planejado] = 0 or [horas_realizado] = 0 then 0 else
        [horas_realizado] / [horas_planejado] * 100
    otherwise 0, type number
    ),
    desempenho = Table.AddColumn(ocupação, "desempenho", 
        each try
            if [Produção Planejada] = 0 or [Produção Realizada] = 0 then 0 else
            [Produção Realizada] / [Produção Planejada] * 100
        otherwise 0, type number
    ),
    qualidade = Table.AddColumn(desempenho, "qualidade",
        each try
            if [Defeitos] = 0 then 100 else
            if [Produção Realizada] = 0 then 0 else
            100 - ([Defeitos] / [Produção Realizada] * 100)
        otherwise 0, type number
    ),
    OLE = Table.AddColumn(qualidade, "OLE",
        each [ocupação] * [desempenho] * [qualidade] /10000
        , type number
    ),
    target = Table.AddColumn(OLE, "target", each 90, Int64.Type),

    semana_texto = Table.AddColumn(target, "semana_texto", 
        each "WK" & Text.Middle(Text.From([semana_num]),4, 2) & "/" & Text.Middle(Text.From([semana_num]),2,2)
    , type text ),

    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(semana_texto,"semana_num"))
        ,{"semana_num",Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int32.Type),
    add_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(semana_texto,"semana_num",index_semana,"semana_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"}
    ),

    semana_filtro = Table.AddColumn(add_index_semana, "semana_filtro", 
        each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores"
    , type text),

    processo_semana = Table.AddColumn(semana_filtro,"processo_semana",
        each Text.Upper(Text.From([Processo])) & "-" & Text.From([semana_num])
    , type text)
in
    processo_semana