let
    Fonte = Sql.Database("172.16.10.9", "Protheus12", [Query="SELECT#(lf)#(tab)YEAR([DATA]) AS ANO, #(lf)#(tab)YEAR([DATA]) * 100 + MONTH([DATA]) AS MÊS, #(lf)#(tab)YEAR([DATA]) * 100 + DATEPART(WEEK, [DATA]) AS WK,#(lf)#(tab)PRODUTO, DESCRIÇÃO, ENDEREÇO, [DESCRIÇÃO MOTIVO], SUM(CUSTO) AS CUSTO#(lf)#(lf)FROM VW_MN_PERDAS_PROCESSO#(lf)#(lf)WHERE LEFT([DATA],4) >= YEAR( DATEADD( YEAR,-1,GETDATE() ) )#(lf)#(lf)GROUP BY #(lf)#(tab)YEAR([DATA]), #(lf)#(tab)YEAR([DATA]) * 100 + MONTH([DATA]), #(lf)#(tab)YEAR([DATA]) * 100 + DATEPART(WEEK, [DATA]),#(lf)#(tab)PRODUTO, DESCRIÇÃO, ENDEREÇO, [DESCRIÇÃO MOTIVO]"]),

    //REMOVER ETAPA EM 2026
    #"Valor Substituído" = Table.ReplaceValue(Fonte,202501,202502,Replacer.ReplaceValue,{"WK"}),

    remove_ultima_semana = Table.SelectRows(#"Valor Substituído", 
        each [WK] < #"00-semana-atual-completo"),

    une_linha3_linha4 = Table.TransformColumns(remove_ultima_semana, {
        {"ENDEREÇO", each if _ = "LINHA3" or _ = "LINHA4" then "LINHA3/LINHA4" else _}
    }),
    #"Personalização Adicionada" = Table.AddColumn(une_linha3_linha4, "TARGET", each 0.0077),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Personalização Adicionada",{{"TARGET", type number}}),

    //MÊS
    mes_texto = Table.AddColumn(#"Tipo Alterado1", "mes_texto", 
        each Date.MonthName(#date(#"00-ano-atual",Number.FromText(Text.Middle(Text.From([MÊS]),4,2)),1)) & "/" & Text.Middle(Text.From([ANO]),2,2) ,type text),
    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(mes_texto,"MÊS"))
        ,{"MÊS",Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int32.Type),
    add_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(mes_texto,"MÊS",index_mes,"MÊS","dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes"},{"index_mes"}
    ),
    filtro_mes = Table.AddColumn(add_index_mes, "filtro_mes", 
        each if [index_mes] <= 3 then "Últimos 3 Meses" else "Meses Anteriores" , type text ),
    
    //SEMANA
    semana_texto = Table.AddColumn(filtro_mes, "semana_texto", 
        each "WK" & Text.Middle(Text.From([WK]),4,2) & "/" & Text.Middle(Text.From([ANO]),2,2) , type text),
    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(semana_texto,"WK"))
        ,{"WK",Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int32.Type),
    add_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(semana_texto,"WK",index_semana,"WK","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"},{"index_semana"}
    ),
    filtro_semana = Table.AddColumn(add_index_semana, "filtro_semana", 
        each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text ),
    ano_texto = Table.AddColumn(filtro_semana, "ano_texto", each Text.From([ANO]), type text),
    rel_perda = Table.AddColumn(ano_texto, "rel_perda", 
        each Text.From([ANO]) & "-" & Text.From([MÊS]) & "-" & Text.From([WK])
        , type text ),
    endereço_defeito = Table.AddColumn(rel_perda,"endereço_defeito", each [ENDEREÇO] & " - " & [DESCRIÇÃO MOTIVO], type text)
in
    endereço_defeito