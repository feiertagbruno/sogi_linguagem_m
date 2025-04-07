let
    Fonte = Excel.Workbook(File.Contents("M:\INJEÇÃO\RESTRITO\01 CONTROLE PRODUÇÃO INJETADOS\2025\01 Controle de Produção 2025.xlsm"), null, true),
    tbDiario24_Table = Fonte{[Item="tbDiario24",Kind="Table"]}[Data],
	selectionar_colunas = Table.SelectColumns(tbDiario24_Table,{
		"Data","Máquina","Turno","Código","Descrição","Meta/Hora",
		"Produção Planejada","Produção Realizada","Defeitos","Causa"
	}),
    tipo_alterado = Table.TransformColumnTypes(selectionar_colunas,{
		{"Data", type date},
		{"Máquina", type text}, {"Turno", type text}, {"Código", type text}, 
		{"Descrição", type text},
		{"Meta/Hora", type text}, {"Produção Planejada", Int64.Type}, 
		{"Produção Realizada", Int64.Type}, {"Defeitos", Int64.Type}, 
		{"Causa", type text}
	}),
	selecionar_linhas = Table.SelectRows(tipo_alterado, each (
		([Produção Planejada] <> 0 or [Produção Realizada] <> 0)
		and Date.Year([Data]) >= #"00-ano-atual" - 1
	)),

	ano_num = Table.AddColumn(selecionar_linhas,"ano_num", each (
		Date.Year([Data])
	), Int16.Type),
	mes_num = Table.AddColumn(ano_num,"mes_num", each (
		[ano_num] * 100 + Date.Month([Data])
	), Int32.Type),
	mes_num_buffer = Table.Buffer(mes_num),
	meses_distintos = Table.Sort(
		Table.Distinct(Table.SelectColumns(mes_num_buffer,"mes_num")),
		{"mes_num", Order.Descending}
	),
	add_mes_filtro = Table.TransformColumns((
		Table.AddIndexColumn(meses_distintos,"mes_filtro", 1,1,Int16.Type)
	), {
		{"mes_filtro", each if _ <= 3 then "Últimos 3 Meses" else "Meses Anteriores", type text}
	}),
	traz_mes_filtro = Table.ExpandTableColumn(
		(Table.NestedJoin(mes_num_buffer,"mes_num",add_mes_filtro,"mes_num","dados",JoinKind.LeftOuter)),
		"dados",{"mes_filtro"}
	),

	semana_num = Table.AddColumn(traz_mes_filtro,"semana_num", each (
		[ano_num] * 100 + Date.WeekOfYear([Data])
	), Int32.Type),
	semana_num_buffer = Table.Buffer(semana_num),
	semanas_distintas = Table.Sort((
		Table.SelectRows(
			Table.Distinct(Table.SelectColumns(
				semana_num_buffer,"semana_num"
			)), each [semana_num] <> #"00-semana-atual-completo"
		)
		
	), {"semana_num",Order.Descending}),
	add_semana_filtro = Table.TransformColumns((
		Table.AddIndexColumn(semanas_distintas,"semana_filtro",1,1,Int16.Type)
	), {
		{"semana_filtro", each if _ <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text}
	}),
	traz_semana_filtro = Table.ExpandTableColumn(
		Table.NestedJoin(semana_num_buffer,"semana_num",add_semana_filtro,"semana_num","dados",JoinKind.LeftOuter)
		,"dados",{"semana_filtro"}
	),
	preenche_semana_corrente = Table.TransformColumns(traz_semana_filtro,{
		{"semana_filtro", each if _ = null then "Semana Corrente" else _, type text}
	}),
	
	ano_texto = Table.AddColumn(preenche_semana_corrente,"ano_texto", each Text.From([ano_num]), type text),
	mes_texto = Table.AddColumn(ano_texto,"mes_texto",each (
		Text.Middle(Date.MonthName([Data]),0,3) & "/" & Text.Middle([ano_texto],2,2)
	), type text),
	semana_texto = Table.AddColumn(mes_texto,"semana_texto", each (
		"W" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle([ano_texto],2,2)
	), type text),

	descri_cod = Table.AddColumn(semana_texto,"descri_cod", each (
		Text.Trim(Text.From([Descrição])) & " " &  Text.Trim(Text.From([Código]))
	), type text),

	remover_colunas_desn = Table.RemoveColumns(descri_cod,{"Descrição"}),
	buffer = Table.Buffer(remover_colunas_desn),

	/* próximos passos:
		1. agrupar por semana_num e produto somando o planejado e o realizado
		2. criar coluna com a diferença
		3. Ordenar diferença decrescente e ordem alfabética crescente
		4. criar coluna índice e transformar os menores ou igual a 3 em Top 3 else "Outros"

		Pareto
		1. Salvar a soma da semana em uma variável
		2. adicionar pareto: firstn (index) / variavel_soma
	*/

	group_semana_produto = Table.Group(buffer,{"semana_num","Código"},{
		{"Produção Planejada",each List.Sum([Produção Planejada]), Int64.Type},
		{"Produção Realizada",each List.Sum([Produção Realizada]), Int64.Type}
	}),
	diferenca = Table.AddColumn(group_semana_produto,"diferenca", each (
		[Produção Realizada] - [Produção Planejada]
	), Int64.Type),
	group_semana_prod_table = Table.Group(diferenca,"semana_num",{
		{"group", each _ , type table}
	}),
	agressores_e_top_3 = Table.TransformColumns(group_semana_prod_table,{
		{"group", (tb) => (
			let
				sort = Table.Sort(tb,{{"diferenca", Order.Ascending}}),
				top_3 = Table.AddIndexColumn(sort,"agressores",1,1,Int16.Type),
				transform_agressores = Table.TransformColumns(top_3,{
					{"agressores", each if _ <= 3 then "Top 3" else "Outros", type text}
				})
			in
				transform_agressores
		), type table}
	}),
	expandir_agressores_e_top3 = Table.ExpandTableColumn(agressores_e_top_3,"group",{"Código","agressores"}),
	traz_agressores_e_top3 = Table.ExpandTableColumn(
		Table.NestedJoin(buffer,{"semana_num","Código"},expandir_agressores_e_top3,{"semana_num","Código"},
			"dados",JoinKind.LeftOuter
		),
		"dados", {"agressores"}
	),
    tipo_alterado_1 = Table.TransformColumnTypes(traz_agressores_e_top3,{
		{"Data", type date},{"agressores",type text}
	}),

	transform_maquina = Table.TransformColumns(tipo_alterado_1,{
		{"Máquina", each (
			try
				Text.PadStart(_,2,"0")
			otherwise
				_
		), type text}
	})
in
    transform_maquina