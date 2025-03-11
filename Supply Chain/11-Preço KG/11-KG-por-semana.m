let
    Fonte = Table.Buffer(#"11-KG-base"),

    group_por_processo = Table.Group(Fonte,{"REFERENCE", "AGENTE DE CARGAS","MODAL","Data Real de Entrega"},{
        {"GROSS WEIGHT (kgs)", each List.Sum([#"GROSS WEIGHT (kgs)"]), type number},
        {"FRETE INTERN. (BRL)", each List.Sum([#"FRETE INTERN. (BRL)"]), type number}
    }),
    filtra_fretes_zerados = Table.SelectRows(group_por_processo, each ([#"FRETE INTERN. (BRL)"] <> 0)),

    semana_num = Table.AddColumn(filtra_fretes_zerados, "semana_num", 
        each Date.Year([Data Real de Entrega]) * 100 + Date.WeekOfYear([Data Real de Entrega])
    , Int32.Type),

    group_semana_ag_cargas = Table.Group(semana_num, {"AGENTE DE CARGAS","MODAL","semana_num"}, {
        {"GROSS WEIGHT (kgs)", each List.Sum([#"GROSS WEIGHT (kgs)"]), type number},
        {"FRETE INTERN. (BRL)", each List.Sum([#"FRETE INTERN. (BRL)"]), type number}
    }),

    calculo_preco_kg = Table.AddColumn(group_semana_ag_cargas,"preco_kg", each [#"FRETE INTERN. (BRL)"] / [#"GROSS WEIGHT (kgs)"], type number),

    ano_num = Table.AddColumn(calculo_preco_kg, "ano_num", each Number.From(Text.Middle(Text.From([semana_num]),2,2)), Int32.Type ),

    maior_semana = List.Max(Table.Column(ano_num,"semana_num")),

    ult_6_semanas = Table.Sort(Table.Distinct(Table.SelectColumns(ano_num, {"semana_num"})), {"semana_num", Order.Descending} ),
    ult_6_s_indice = Table.AddIndexColumn(ult_6_semanas,"indice", 1,1,Int16.Type),

    traz_indice = Table.ExpandTableColumn(
        Table.NestedJoin(ano_num,{"semana_num"},ult_6_s_indice,{"semana_num"},"dados",JoinKind.LeftOuter)
    , "dados",{"indice"},{"indice"}
    ),
    forma_filtro_pelo_indice =  Table.RenameColumns(
        Table.TransformColumns(traz_indice, {"indice", 
        each if _ = null then null else
        if _ <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores" })
    , {"indice", "filtro_6_semanas"}
    ),

    semana_texto = Table.AddColumn(forma_filtro_pelo_indice,"semana_texto", 
        each "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle(Text.From([semana_num]),2,2)
    , type text),

    coleta_semana = (semana) => (
        let
            resultado = Number.FromText(Text.Middle(Text.From(semana),4,2))
        in
            resultado
    ),

    mes_num = Table.AddColumn(semana_texto, "mes_num", 
        each Number.FromText(Text.Middle(Text.From([semana_num]),0,4)) * 100 +
            coleta_mes_pela_semana( coleta_semana([semana_num]) )
            , Int64.Type
    ),

    ult_3_meses = Table.Sort(Table.Distinct(Table.SelectColumns(mes_num, {"mes_num"})), {"mes_num", Order.Descending} ),
    ult_6_m_indice = Table.AddIndexColumn(ult_3_meses,"indice", 1,1,Int16.Type),

    traz_indice_m = Table.ExpandTableColumn(
        Table.NestedJoin(mes_num,{"mes_num"},ult_6_m_indice,{"mes_num"},"dados",JoinKind.LeftOuter)
    , "dados",{"indice"},{"indice"}
    ),
    forma_filtro_pelo_indice_m =  Table.RenameColumns(
        Table.TransformColumns(traz_indice_m, {"indice", 
        each if _ = null then null else
        if _ <= 3 then "Últimos 3 Meses" else "Meses Anteriores" })
    , {"indice", "filtro_3_meses"}
    ),


    mes_texto = Table.AddColumn(forma_filtro_pelo_indice_m, "mes_texto", 
        each Text.Start( Date.MonthName( #date(#"00-ano-atual",Number.FromText(Text.Middle(Text.From([mes_num]),4,2)),1) ) ,3)
        & "/" & Text.From([ano_num])
    , type text),

    traz_cor_ag_carga = Table.ExpandTableColumn(
        Table.NestedJoin(mes_texto,{"AGENTE DE CARGAS"},#"11-KG-cores-ag-cargas",{"AGENTE DE CARGAS"},"dados",JoinKind.FullOuter)
    , "dados", {"Cor"},{"Cor"}
    ),
    ano_texto = Table.AddColumn(traz_cor_ag_carga,"ano_texto",each "20" & Text.From([ano_num]), type text),

    relacao = Table.AddColumn(ano_texto, "relacao", 
        each [ano_texto] & "-" & Text.From([semana_num]) & "-" & [AGENTE DE CARGAS] & "-" & [MODAL]
    , type text ),
    #"Linhas Filtradas1" = Table.SelectRows(relacao, each ([ano_num] <> null))

in
    #"Linhas Filtradas1"