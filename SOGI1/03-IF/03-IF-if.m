let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\3 - 4Q IFinal_novo.xlsm"), null, true),
    IF_BI_SOGI_Table = Fonte{[Item="IF_BI_SOGI",Kind="Table"]}[Data],
    remover_util_zeros = Table.SelectRows(IF_BI_SOGI_Table, each ([#"Util."] = 1)),

    //deixar somente ano atual e ano passado
    filtro_ano = Table.SelectRows(
        remover_util_zeros, each [Ano] >= (#"00-ano-atual" - 1)
    ),

    Ano_Texto = Table.DuplicateColumn(filtro_ano, "Ano", "Ano Texto"),


    //semana
    semana_ordem = Table.AddColumn(Ano_Texto,"semana_ordem",each Number.FromText( Text.From([Ano]) & Text.PadStart(Text.From([Semana]),2,"0") ), Int32.Type),

    tirar_semana_atual = Table.SelectRows(semana_ordem, each [semana_ordem] < #"00-semana-atual-completo"),

    maior_semana = List.Max(tirar_semana_atual[semana_ordem]),
    semanas_distintas = Table.Sort( Table.Distinct(Table.SelectColumns(tirar_semana_atual,"semana_ordem")), {"semana_ordem",Order.Descending} ),
    index_semana_ordem = Table.AddIndexColumn(semanas_distintas,"index_semana_ordem",1,1,Int32.Type),
    add_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(tirar_semana_atual,"semana_ordem",index_semana_ordem,"semana_ordem","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana_ordem"},{"index_semana_ordem"}
    ),
    filtro_semana = Table.AddColumn(add_index_semana, "filtro_semana", 
        each if [index_semana_ordem] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text
    ),
    semana_texto = Table.AddColumn(filtro_semana,"semana_texto", 
        each "WK" & Text.PadStart(Text.From([Semana]),2,"0") & "/" & Text.Middle(Text.From([Ano]),2,2) 
    , type text),


    //mês
    mes_ordem = Table.AddColumn(semana_texto,"mes_ordem", 
        each Number.FromText( Text.From([Ano]) & Text.PadStart(Text.From([Mês]),2,"0") )
        ,Int32.Type
    ),
    maior_mes = List.Max(mes_ordem[mes_ordem]),
    //filtro_mes
    meses_distintos = Table.Sort(Table.Distinct(Table.SelectColumns(mes_ordem,"mes_ordem"))
        , {"mes_ordem",Order.Descending}),
    indice_mes_ordem = Table.AddIndexColumn(meses_distintos,"indice_mes_ordem",1,1,Int32.Type),
    add_indice_mes = Table.ExpandTableColumn(
        Table.NestedJoin(mes_ordem,"mes_ordem",indice_mes_ordem,"mes_ordem","dados",JoinKind.LeftOuter)
        ,"dados",{"indice_mes_ordem"},{"indice_mes_ordem"}
    ),
    

    filtro_mes = Table.AddColumn(add_indice_mes, "filtro_mes", 
        each if [indice_mes_ordem] <= 3 then "Últimos 3 Meses" else "Meses Anteriores"
    ),
    mes_texto = Table.AddColumn(filtro_mes, "mes_texto", each Date.MonthName([Data]) & "/" & Text.Middle(Text.From([Ano]),2,2) ),
    //filtro_mes_num = Table.AddColumn(filtro_mes, "filtro_mes_num", each if [filtro_mes] = null then 0 else [Mês]),


    tipos_alterados = Table.TransformColumnTypes(mes_texto,{{"Util.", Int64.Type}, {"DDS", type text}, {"Data", type date}, {"Semana", Int64.Type}, {"Mês", Int64.Type}, {"Ano", Int64.Type}, {"Linha 1 - 1oT", Int64.Type}, {"Linha 2 - 1oT", Int64.Type}, {"Linha 3 - 1oT", Int64.Type}, {"Total", Int64.Type}, {"Quant. Insp. Prancha", Int64.Type}, {"Quant. Insp. Secador", Int64.Type}, {"Total Planta", Int64.Type}, {"IF Prancha", type number}, {"IF Secador", type number}, {"IF Planta", type number}, {"IF Target", type number}, {"Ano Texto", type text}, {"filtro_semana", type text}, {"filtro_mes", type text},{"mes_texto", type text}})
    
in
    tipos_alterados