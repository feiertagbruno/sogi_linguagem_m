let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2024\BI-SOGI\Bases\Custo do Estoque\Posições de Estoque 2024.xlsx"), null, true),
    posest_Table = Fonte{[Item="posest",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(posest_Table,{{"CODIGO", type text}, {"ARMAZEM", type text}, 
		{"QUANT", Int64.Type}, {"CUSTO", type number}, {"DATA", type date}}),
	tira_semanas_desnecessarias = Table.SelectRows(#"Tipo Alterado", 
		each not List.Contains({49,50,51,52},Date.WeekOfYear([DATA]))
	)
in
    tira_semanas_desnecessarias