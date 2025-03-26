let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Custo dos SKUs\Multi Estruturas Sem Frete.xlsx"), null, true),
    detalhamento_standard_Table = Fonte{[Item="detalhamento_standard",Kind="Table"]}[Data],
	
	colunas_necessarias = Table.SelectColumns(detalhamento_standard_Table,{
		"COD PA","CODIGO","QTDE.NECESSARIA","Custo c/ Premissa","DESCRICAO"
	}),

    tipo_alterado = Table.TransformColumnTypes(colunas_necessarias,{
		{"COD PA", type text}, {"CODIGO", type text}, 
		{"QTDE.NECESSARIA", type number}, {"Custo c/ Premissa", type number},
		{"DESCRICAO", type text}
	}),
	buffer = Table.Buffer(tipo_alterado),

	group = Table.Group(buffer,{"COD PA","CODIGO"},{
		{"QTDE.NECESSARIA", each List.Sum([#"QTDE.NECESSARIA"]), type number},
		{"Custo c/ Premissa", each List.Sum([#"Custo c/ Premissa"]), type number}
	}),
	buffer_2 = Table.Buffer(group),

	group_por_CODIGO = Table.Group(buffer,"CODIGO",{"DESCRICAO", each List.First([DESCRICAO]), type text}),

	buffer_descricao = Table.Buffer(group_por_CODIGO),

	traz_descricoes = Table.ExpandTableColumn(
		Table.NestedJoin(buffer_2,{"CODIGO"}, buffer_descricao,{"CODIGO"},"dados", JoinKind.LeftOuter)
		,"dados",{"DESCRICAO"}
	)


in
    traz_descricoes