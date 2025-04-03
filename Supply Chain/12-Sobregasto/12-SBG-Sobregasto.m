let
    Fonte = #"11-KG-base",

    group_por_processo = Table.Group(Fonte,{"REFERENCE","AGENTE DE CARGAS","MODAL","MODAL ORIGINAL","Data Real de Entrega","tipo"},{
        {"FOB AMOUNT (BRL)", each List.Sum([#"FOB AMOUNT (BRL)"]), type number},
        {"FRETE INTERN. (BRL)", each List.Sum([#"FRETE INTERN. (BRL)"]), type number}
    }),

    semana_num = Table.AddColumn(group_por_processo, "semana_num", each Date.Year([Data Real de Entrega]) * 100 + Date.WeekOfYear([Data Real de Entrega]), Int32.Type),


    ano_num = Table.AddColumn(semana_num, "ano_num", each Date.Year([Data Real de Entrega]), Int32.Type ),

    traz_cor_ag_carga = Table.ExpandTableColumn(
        Table.NestedJoin(ano_num,{"AGENTE DE CARGAS"},#"11-KG-cores-ag-cargas",{"AGENTE DE CARGAS"},"dados",JoinKind.FullOuter)
    , "dados", {"Cor"},{"Cor"}
    ),
    ano_texto = Table.AddColumn(traz_cor_ag_carga,"ano_texto",each Text.From([ano_num]), type text),

    relacao = Table.AddColumn(ano_texto, "relacao", 
        each [ano_texto] & "-" & Text.From([semana_num]) & "-" & [AGENTE DE CARGAS] & "-" & [MODAL]
    , type text ),

    porcent_ideal = Table.AddColumn(relacao,"porcent_ideal", each 0.3, type number),

    frete_ideal = Table.AddColumn(porcent_ideal, "frete_ideal",each [#"FOB AMOUNT (BRL)"] * [porcent_ideal], type number),

    sobrecarga = Table.AddColumn(frete_ideal, "sobrecarga", each [#"FRETE INTERN. (BRL)"] - [frete_ideal], type number),

    filtrar_sobrecarga_positiva = Table.SelectRows(sobrecarga, each [sobrecarga] > 0),



    ////////
    maior_semana = List.Max(Table.Column(filtrar_sobrecarga_positiva,"semana_num")),
    
    ult_6_semanas = Table.Sort(
        Table.SelectRows(Table.Distinct(Table.SelectColumns(filtrar_sobrecarga_positiva, {"semana_num"})) ,
        each [semana_num] <> null )   
    , {"semana_num", Order.Descending} ),
    ult_6_s_indice = Table.AddIndexColumn(ult_6_semanas,"indice", 1,1,Int16.Type),

    traz_indice = Table.ExpandTableColumn(
        Table.NestedJoin(filtrar_sobrecarga_positiva,{"semana_num"},ult_6_s_indice,{"semana_num"},"dados",JoinKind.LeftOuter)
    , "dados",{"indice"},{"indice"}
    ),
    forma_filtro_pelo_indice =  Table.RenameColumns(
        Table.TransformColumns(traz_indice, {"indice", 
        each if _ = null then null else
        if _ <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores" })
    , {"indice", "filtro_6_semanas"}
    ),

    semana_texto = Table.AddColumn(forma_filtro_pelo_indice,"semana_texto", 
        each "W" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle([ano_texto],2,2)
    , type text),

    coleta_semana = (semana) => (
        let
            resultado = Number.FromText(Text.Middle(Text.From(semana),4,2))
        in
            resultado
    ),

    mes_num = Table.AddColumn(semana_texto, "mes_num", 
        each [ano_num] * 100 +
            (if List.Contains({1,2,3,4,5},coleta_semana([semana_num])) then 1 else 
            if List.Contains({6,7,8,9},coleta_semana([semana_num])) then 2 else
            if List.Contains({10,11,12,13}, coleta_semana([semana_num])) then 3 else 
            if List.Contains({14,15,16,17,18}, coleta_semana([semana_num])) then 4 else
            if List.Contains({19,20,21,22}, coleta_semana([semana_num])) then 5 else
            if List.Contains({23,24,25,26}, coleta_semana([semana_num])) then 6 else
            if List.Contains({27,28,29,30,31}, coleta_semana([semana_num])) then 7 else
            if List.Contains({32,33,34,35}, coleta_semana([semana_num])) then 8 else
            if List.Contains({36,37,38,39}, coleta_semana([semana_num])) then 9 else
            if List.Contains({40,41,42,43,44}, coleta_semana([semana_num])) then 10 else
            if List.Contains({45,46,47,48}, coleta_semana([semana_num])) then 11 else 12)
            , Int32.Type
    ),

    ult_3_meses = Table.Sort(
        Table.SelectRows(Table.Distinct(Table.SelectColumns(mes_num, {"mes_num"})),
        each [mes_num] <> null )
    , {"mes_num", Order.Descending} ),
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
        each Text.Start( Date.MonthName( #date(#"00-ano-atual",Number.FromText(Text.Middle(Text.From([mes_num]),4,2)),1) ) ,3) & "/" & Text.Middle([ano_texto],2,2)
    , type text),
    #"Outras Colunas Removidas" = Table.SelectColumns(mes_texto,{
        "REFERENCE","AGENTE DE CARGAS", "MODAL", "MODAL ORIGINAL",
        "FOB AMOUNT (BRL)", "FRETE INTERN. (BRL)","semana_num", 
        "ano_num", "Cor", "ano_texto", 
        "relacao", "sobrecarga", "filtro_6_semanas", 
        "semana_texto", "mes_num", "filtro_3_meses", 
        "mes_texto", "tipo"
    }),

    relacao_sbg = Table.AddColumn(#"Outras Colunas Removidas", "relacao_sbg", each [relacao] & "-" & [REFERENCE], type text )

in
    relacao_sbg