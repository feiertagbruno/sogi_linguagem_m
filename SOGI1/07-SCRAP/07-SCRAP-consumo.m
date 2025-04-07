let
    Fonte = Sql.Database("172.16.10.9", "Protheus12", [Query="SELECT#(lf)#(tab)YEAR([DT EMISSÃO]) AS ANO, #(lf)#(tab)YEAR([DT EMISSÃO]) * 100 + MONTH([DT EMISSÃO]) AS MÊS, #(lf)#(tab)YEAR([DT EMISSÃO]) * 100 + DATEPART(WEEK, [DT EMISSÃO]) AS SEMANA,#(lf)#(tab)PRODUTO, [DESCRI PROD SD3],#(lf)#(tab)SUM(CUSTO) AS CONSUMO#(lf)#(lf)FROM VW_MN_SCRAP_RETRABALHO_CONSUMO#(lf)#(lf)WHERE#(lf)#(tab)[TP MOVIMENTO] = '999'#(lf)#(tab)AND [TIPO RE/DE] = 'RE2'#(lf)#(tab)AND [TIPO] NOT IN ('PI','PA')#(lf)#(tab)AND [DT EMISSÃO] >= '20240101'#(lf)#(lf)GROUP BY YEAR([DT EMISSÃO]), #(lf)#(tab)YEAR([DT EMISSÃO]) * 100 + MONTH([DT EMISSÃO]), #(lf)#(tab)YEAR([DT EMISSÃO]) * 100 + DATEPART(WEEK, [DT EMISSÃO]),#(lf)#(tab)PRODUTO, [DESCRI PROD SD3]"]),

    remove_semana_atual = Table.SelectRows(Fonte, each [SEMANA] < #"00-semana-atual-completo"),

    ano_texto = Table.AddColumn(remove_semana_atual,"ano_texto", each Text.From([ANO]), type text),

    //MÊS
    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(ano_texto,"MÊS"))
        ,{"MÊS",Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int32.Type),
    add_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(ano_texto,"MÊS",index_mes,"MÊS","dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes"},{"index_mes"}
    ),
    mes_texto = Table.AddColumn(add_index_mes, "mes_texto", 
        each Date.MonthName(#date(#"00-ano-atual",Number.FromText(Text.Middle(Text.From([MÊS]),4,2)),1))
        & "/" & Text.Middle([ano_texto],2,2)
        , type text
    ),
    filtro_mes = Table.AddColumn(mes_texto, "filtro_mes", 
        each if [index_mes] <= 3 then "Últimos 3 Meses" else "Meses Anteriores" , type text ),

    //SEMANA
    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(filtro_mes,"SEMANA"))
        ,{"SEMANA", Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int32.Type),
    add_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(filtro_mes,"SEMANA",index_semana,"SEMANA","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"},{"index_semana"}
    ),
    filtro_semana = Table.AddColumn(add_index_semana, "filtro_semana", 
        each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores"
        , type text ),
    semana_texto = Table.AddColumn(filtro_semana, "semana_texto", 
        each "WK" & Text.Middle(Text.From([SEMANA]),4,2) & "/" & Text.Middle([ano_texto],2,2)
        , type text ),
    rel_consumo = Table.AddColumn(semana_texto, "rel_consumo", 
        each Text.From([ANO]) & "-" & Text.From([MÊS]) & "-" & Text.From([SEMANA])
        , type text)
    
in
    rel_consumo