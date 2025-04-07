let
    Fonte = Excel.Workbook(File.Contents("M:\Produção\2 INFORMATIVO\01 - CONTROLE PRODUÇÃO\2025\Resumo Diário de Produção.xlsm"), null, true),
    tabelaOLEbase_Table = Fonte{[Item="tabelaOLEbase",Kind="Table"]}[Data],
	selecionar_cols = Table.SelectColumns(tabelaOLEbase_Table,{
		"Data","Processo","Máq/#(lf)Linha","Turno","Código","Descrição","Operação","Meta/Hr",
		"Produção Planejada","Produção Realizada","Defeitos","HC","Comentários"
	}),
    tipo_alterado = Table.TransformColumnTypes(selecionar_cols,{
		{"Data", type date},{"Processo", type text}, 
		{"Máq/#(lf)Linha", type text}, {"Turno", type text}, {"Código", type text}, 
		{"Descrição", type text}, {"Operação", type text}, {"Meta/Hr", type text}, 
		{"Produção Planejada", Int64.Type}, {"Produção Realizada", Int64.Type}, 
		{"Defeitos", Int64.Type},{"HC", Int64.Type}, {"Comentários", type text}
	}),
	processo_em_maiusculo = Table.TransformColumns(tipo_alterado,{
		{"Processo", each Text.Upper(Text.Clean(Text.Trim(Text.From(_)))), type text}
	}),
	selectionar_linhas = Table.SelectRows(processo_em_maiusculo, each 
		[Data] <> null and Date.Year([Data]) >= #"00-ano-atual" - 1 
		// and List.Contains({"SECADOR","PRANCHA"}, [Processo])
	),

	ano_num = Table.AddColumn(selectionar_linhas,"ano_num", each (
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

	descri_cod_buffer = Table.Buffer(descri_cod),
	descri_cod_distintos = Table.Distinct(Table.SelectColumns(descri_cod_buffer,{"descri_cod","Processo"})),
	datas_distintas = Table.Distinct(Table.SelectColumns(descri_cod_buffer,{
		"ano_num","ano_texto","mes_num","mes_texto","mes_filtro","semana_num","semana_texto","semana_filtro"
	})),
	cross_join = Table.ExpandTableColumn(
		Table.AddColumn(datas_distintas,"dados", each descri_cod_distintos, type table),
		"dados",{"descri_cod","Processo"}
	),
	todos_prod = Table.AddColumn(cross_join,"todos_prod", each "x", type text),

	combine = Table.Combine({descri_cod_buffer,todos_prod})

	
in
    combine