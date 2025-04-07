let
    Fonte = #"09-OLE-base",
    mes_num = Table.AddColumn(Fonte, "mes_num", each Date.Year([Data]) * 100 + Date.Month([Data]), Int32.Type),
    #"Linhas Agrupadas" = Table.Group(mes_num, {"mes_num", "Processo"}, {
        {"Produção Planejada", each List.Sum([Produção Planejada]), type nullable number}, 
        {"Produção Realizada", each List.Sum([Produção Realizada]), type nullable number}, 
        {"Defeitos", each List.Sum([Defeitos]), type nullable number}, 
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

    ano_num = Table.AddColumn(target, "ano_num", each Number.FromText(Text.Middle(Text.From([mes_num]),0,4)), Int32.Type),

    #"Tipo Alterado" = Table.TransformColumnTypes(ano_num,{{"ano_num", Int64.Type}, {"OLE", type number}, {"target", type number}}),

    maior_ano = List.Max(Table.Column(#"Tipo Alterado","ano_num")),

    ano_texto = Table.AddColumn(#"Tipo Alterado", "ano_texto", each Text.From([ano_num]), type text),
    mes_texto = Table.AddColumn(ano_texto, "mes_texto", 
        each Date.MonthName( #date(#"00-ano-atual",  Number.FromText(Text.Middle(Text.From([mes_num]),4,2)  ),1) )
        & "/" & Text.Middle(Text.From([mes_num]),2,2)
    , type text),

    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(mes_texto,"mes_num"))
        ,{"mes_num",Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int32.Type),
    add_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(mes_texto,"mes_num",index_mes,"mes_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes"}
    ),

    mes_filtro = Table.AddColumn(add_index_mes, "mes_filtro", 
        each if [index_mes] <= 3 then "Últimos 3 Meses" else "Meses Anteriores"
        , type text
    )
in
    mes_filtro