let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\1- 4Q ARM 98_novo.xlsm"), null, true),
    Tabela1_Table = Fonte{[Item = "Tabela1", Kind = "Table"]}[Data],
    remove_linhas_vazias = Table.SelectRows(Tabela1_Table, each [CODIGO] <> null),
    colunas_removidas = Table.RemoveColumns(
        remove_linhas_vazias, {"TP", "U.M.", "ARMZ", "PREVISÃO PARA REPOSIÇÃO", "ESTOQUE DISPONIVEL", "GRUPO", "FL"}
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(
        colunas_removidas,
        {
            {"CODIGO", type text},
            {"DESCRICAO", type text},
            {"VALOR EM ESTOQUE", type number},
            {"DESCRICAO DO ARMAZEM", type text},
            {"WK", Int64.Type},
            {"MÊS", type date},
            {"CLASSIF", type text}
        }
    ),


    ano_num = Table.AddColumn(#"Tipo Alterado", "ano_num", each Date.Year([MÊS]), Int32.Type),
    mes_num = Table.AddColumn(ano_num, "mes_num", each [ano_num] * 100 + Date.Month([MÊS]), Int32.Type),
    semana_num = Table.AddColumn(mes_num, "semana_num", each [ano_num] * 100 + [WK], Int32.Type),


	group_por_ano = Table.Group(semana_num,"ano_num",{"agrupamento", each _, type table}),
	add_ano_texto_em_cada_ano = Table.TransformColumns(group_por_ano,{
		{"agrupamento", each #"add_ano_texto_na_ultima_semana"(_)}
	}),
	expande_tabela_ano = Table.ExpandTableColumn(
		add_ano_texto_em_cada_ano ,"agrupamento",
		{"CODIGO","DESCRICAO","VALOR EM ESTOQUE","DESCRICAO DO ARMAZEM","WK","MÊS","CLASSIF","mes_num","semana_num","ano_texto"}
	),


	maior_semana = List.Max(Table.Column(expande_tabela_ano,"semana_num")),
    mes_texto = Table.AddColumn(
        expande_tabela_ano, "mes_texto", 
		each if #"valida_relacao_semana_mes_544"(maior_semana,[mes_num],[semana_num]) = 1 then
			Date.MonthName([MÊS]) & "/" & Text.Middle(Text.From([ano_num]), 2, 2) else
			null
		,type text
    ),

	//TRAZ INDEX MÊS
	meses_distintos = Table.Sort(
		Table.Distinct(Table.SelectColumns(mes_texto, "mes_num")), {"mes_num", Order.Descending}
	),
	index_mes = Table.AddIndexColumn(meses_distintos, "index_mes", 1, 1, Int32.Type),
	add_index_mes = Table.ExpandTableColumn(
		Table.NestedJoin(mes_texto, "mes_num", index_mes, "mes_num", "dados", JoinKind.LeftOuter), "dados", {
			"index_mes"
		}
	),


    filtro_mes = Table.AddColumn(
        add_index_mes, "filtro_mes", each 
			if [mes_texto] = null then null else
			if [index_mes] <= 3
			then "Últimos 3 Meses" 
			else "Meses Anteriores",
        type text
    ),
    //SEMANAS
    semana_texto = Table.AddColumn(
        filtro_mes,
        "semana_texto",
        each "WK" & Text.PadStart(Text.From([WK]), 2, "0") & "/" & Text.Middle(Text.From([ano_num]), 2, 2),
        type text
    ),
	semanas_distintas = Table.Sort(
		Table.Distinct(Table.SelectColumns(semana_texto,"semana_num")),
		{"semana_num",Order.Descending}
	),
	index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1, Int32.Type),
	add_index_semana = Table.ExpandTableColumn(
		Table.NestedJoin(semana_texto,"semana_num",index_semana,"semana_num","dados",JoinKind.LeftOuter)
		,"dados",{"index_semana"}
	),
	filtro_semana = Table.AddColumn(add_index_semana,"filtro_semana",
		each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text
	),
	ajusta_tipos = Table.TransformColumnTypes(
        filtro_semana,
        {
            {"CODIGO", type text},
            {"DESCRICAO", type text},
            {"VALOR EM ESTOQUE", type number},
            {"DESCRICAO DO ARMAZEM", type text},
            {"MÊS", type date},
            {"CLASSIF", type text},
			{"mes_num", Int32.Type},
			{"semana_num", Int32.Type},
            {"ano_texto", type text}
        }
    ),
    #"Outras Colunas Removidas" = Table.SelectColumns(ajusta_tipos,{"ano_num", "CODIGO", "DESCRICAO", "VALOR EM ESTOQUE", "DESCRICAO DO ARMAZEM", "CLASSIF", "mes_num", "semana_num", "ano_texto", "mes_texto", "filtro_mes", "semana_texto", "filtro_semana"}),
	ordem_mes = Table.AddColumn(#"Outras Colunas Removidas","ordem_mes", 
		each if [mes_texto] = null then null else [mes_num], Int32.Type
	),
	filtro_ultima_semana = Table.AddColumn(ordem_mes, "filtro_ultima_semana",
		each if [semana_num] = maior_semana then "Última Semana" else "Semanas Anteriores", type text
	),
	relacao_top3 = Table.AddColumn(filtro_ultima_semana,"relacao_top3",
		each Text.From([mes_num]) & "-" & Text.From([semana_num]) & "-" & [CLASSIF] & "-" & [CODIGO]
	,type text),
	buffer = Table.Buffer(relacao_top3),

    defeitos_distintos = Table.Sort(
        Table.Group(buffer,"CLASSIF",{"VALOR EM ESTOQUE", each List.Sum([VALOR EM ESTOQUE]), type number})
		,{"VALOR EM ESTOQUE", Order.Descending}
    ),

    index = Table.AddIndexColumn(defeitos_distintos,"index",1,1,Int8.Type),

    cor = Table.ExpandTableColumn(
        Table.NestedJoin(index,"index",#"99-cores","index","dados",JoinKind.LeftOuter)
        ,"dados",{"Cor"}
	),

	traz_cor = Table.ExpandTableColumn(
		Table.NestedJoin(relacao_top3,"CLASSIF",cor,"CLASSIF","dados",JoinKind.LeftOuter)
		,"dados",{"Cor"}
	)

in
    traz_cor