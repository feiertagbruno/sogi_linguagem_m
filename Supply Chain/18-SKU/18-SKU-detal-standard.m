let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Custo dos SKUs\Multi Estruturas Sem Frete.xlsx"), null, true),
    detalhamento_standard_Table = Fonte{[Item="detalhamento_standard",Kind="Table"]}[Data],
	
	colunas_necessarias = Table.SelectColumns(detalhamento_standard_Table,{
		"COD PA","CODIGO","QTDE.NECESSARIA","Custo c/ Premissa"
	}),

    tipo_alterado = Table.TransformColumnTypes(colunas_necessarias,{
		{"COD PA", type text}, {"CODIGO", type text}, 
		{"QTDE.NECESSARIA", type number}, {"Custo c/ Premissa", type number}
	}),

	relacao = Table.AddColumn(tipo_alterado, "relacao", each [COD PA] & "-" & [CODIGO], type text),

	group = Table.Group(relacao,{"COD PA","CODIGO","relacao"},{
		{"QTDE.NECESSARIA", each List.Sum([#"QTDE.NECESSARIA"]), type number},
		{"Custo c/ Premissa", each List.Sum([#"Custo c/ Premissa"]), type number}
	})

in
    group