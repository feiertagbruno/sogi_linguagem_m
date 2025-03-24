let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Custo dos SKUs\Multi Estruturas Sem Frete.xlsx"), null, true),
    tabela_consolidada_Table = Fonte{[Item="tabela_consolidada",Kind="Table"]}[Data],
	
	ano_num = Table.AddColumn(tabela_consolidada_Table, "ano_num", each (
		Date.Year([Data de Referência])
	), Int16.Type),
	
	semana_num = Table.AddColumn(ano_num, "semana_num", each (
		[ano_num] * 100 + Date.WeekOfYear([Data de Referência])
	) , type text),

	menor_semana = #"00-semana-standard",
	maior_semana = List.Max(Table.Column(semana_texto,"semana_num")),
	semanas_considerar = List.Combine({
		{menor_semana}, 
		List.FirstN(List.Sort(List.Distinct(Table.Column(semana_num,"semana_num")), Order.Descending),3)
	}),

	filtro_semanas = Table.SelectRows(semana_num,each List.Contains(semanas_considerar,[semana_num])),

	ano_texto = Table.ExpandTableColumn(
		Table.TransformColumns(
			Table.Group(
				filtro_semanas,"ano_num",{"group", each _, type table}
			),{
				{"group",each #"add_ano_texto_na_ultima_semana"(_), type table}
			}
		)
		,"group",{"semana_num","Data de Referência",
			"Código","Tipo","Descrição","Últimas Entradas",
			"Último Fechamento","Custo Médio","ano_texto","Últ Compra com Premissa"}
	),

	rename_ult_compra = Table.RenameColumns((
		Table.RemoveColumns(ano_texto,"Últimas Entradas")
	),{{"Últ Compra com Premissa","Últimas Entradas"}}),

	semana_texto = Table.AddColumn(rename_ult_compra,"semana_texto",each (
		"WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle(Text.From([ano_num]),2,2)
	) , type text),

	// mes_num = Table.AddColumn(semana_texto,"mes_num", each (
	// 	[ano_num] * 100 + #"coleta_mes_pela_semana"(Number.FromText(Text.Middle(Text.From([semana_num]),4,2)))
	// ), Int32.Type),


	// mes_texto = Table.AddColumn(mes_num,"mes_texto", each (
	// 	if #"valida_relacao_semana_mes_544"(maior_semana,[mes_num],[semana_num]) = 1 then
	// 	Date.MonthName(#date(#"00-ano-atual",Number.FromText(Text.Middle(Text.From([mes_num]),4,2)),1)) & "/" &
	// 	Text.Middle(Text.From([ano_num]),2,2) else null
	// ), type text)

	filtro_categoria = Table.AddColumn(semana_texto,"filtro_categoria", each (
		if [semana_num] = menor_semana then "Standard" else
		if [semana_num] = maior_semana then "Última Semana" else "Semanas Anteriores"
	) , type text),

	colunas_necessarias = Table.SelectColumns(filtro_categoria,{
		"semana_num","Código","Descrição","Últimas Entradas","Último Fechamento","Custo Médio",
		"semana_texto","filtro_categoria"
	}),

    tipo_alterado = Table.TransformColumnTypes(colunas_necessarias,{
		{"semana_num",Int32.Type},{"Código",type text},{"Descrição", type text},
		{"Últimas Entradas",type number}, {"Último Fechamento", type number},
		{"Custo Médio", type number}
	}),

	custo_standard = Table.ExpandTableColumn(
		Table.NestedJoin(tipo_alterado,{"Código","semana_num"},#"18-SKU-custo-standard",{"CÓDIGO","semana_num"},
		"dados",JoinKind.LeftOuter)
		,"dados",{"custo_standard"}
	)

in
    custo_standard