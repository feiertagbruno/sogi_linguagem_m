let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\3 - 4Q IFinal_novo.xlsm"), null, true),
    RIAQ2024_Table = Fonte{[Item="RIAQTB",Kind="Table"]}[Data],
    mes_atual = Number.FromText( Text.From(#"00-ano-atual") & Text.PadStart(Text.From(#"00-mes-atual"),2,"0") ) ,
    soma_por_setor = Table.Group(
        RIAQ2024_Table, {"Setor"}, {
            {"SomaRIAQ", each List.Sum([RIAQ]), type nullable number}
    }),
    setores_filtrados = Table.SelectRows(soma_por_setor, each [SomaRIAQ] > 0),
    filtro_setores = Table.Join(RIAQ2024_Table,"Setor",setores_filtrados,"Setor"),
    null_to_zero = Table.ReplaceValue(filtro_setores,null,0,Replacer.ReplaceValue,{"RIAQ"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(null_to_zero,{{"Setor", type text}, {"Ano", type text}, {"Mês", Int64.Type}, 
    {"Status", type text}, {"RIAQ", Int64.Type}, {"TOTAL", Int64.Type}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{"SomaRIAQ"}),
    RIAQFechado = Table.AddColumn(#"Colunas Removidas", "RIAQFechado", each if [Status] = "Fechado" then [RIAQ] else 0),
    #"Tipo Alterado1" = Table.TransformColumnTypes(RIAQFechado,{{"RIAQFechado", Int64.Type}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Tipo Alterado1",{{"Status", Text.Upper, type text}}),

    mes_ordem = Table.AddColumn(#"Texto em Maiúscula","mes_ordem", 
        each Number.FromText(Text.From([Ano]) & Text.PadStart(Text.From([Mês]),2,"0") )),
	tirar_dezembro_2024 = Table.SelectRows(mes_ordem,each [mes_ordem] <> 202412),
    meses_distintos = Table.Sort(Table.Distinct(Table.SelectColumns(tirar_dezembro_2024,"mes_ordem")),
        {"mes_ordem", Order.Descending}),
    meses_distintos_filtados = Table.SelectRows(meses_distintos, each [mes_ordem] <= mes_atual),
    index_mes_ordem = Table.AddIndexColumn(meses_distintos_filtados,"index_mes_ordem",1,1,Int32.Type),
    add_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(mes_ordem,"mes_ordem",index_mes_ordem,"mes_ordem","dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes_ordem"},{"index_mes_ordem"}
    ),

    filtro_mes = Table.AddColumn(add_index_mes, "filtro_mes", 
        each if [index_mes_ordem] = null then null else
        if [index_mes_ordem] <= 3 then "Últimos 3 Meses" 
        else "Meses Anteriores", type text),
    mes_texto = Table.AddColumn(filtro_mes, "mes_texto",
        each Date.MonthName(#date(#"00-ano-atual",[Mês],1)) & "/" & Text.Middle(Text.From([Ano]),2,2)
        , type text )
in
    mes_texto