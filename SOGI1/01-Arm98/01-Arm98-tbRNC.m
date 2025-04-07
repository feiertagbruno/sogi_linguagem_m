let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\1- 4Q ARM 98_novo.xlsm"), null, true),
    tbRNC_Table = Fonte{[Item="tbRNC",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(tbRNC_Table,{{"Índice", Int64.Type}, {"WK Exib", type text}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por quê", type text}, {"Quanto", type text}, {"1-Porque", type text}, {"2-Porque", type text}, {"3-Porque", type any}, {"4-Porque", type any}, {"5-Porque", type any}, {"Ação de Contenção", type text}, {"Responsável A", type text}, {"Data Fechamento A", type text}, {"Ação Corretiva", type text}, {"Responsável B", type text}, {"Data Fechamento B", type text}, {"Status", type text}, {"Foto", type any}}),
	busca_produto = Table.AddColumn(#"Tipo Alterado","produto", each Text.BeforeDelimiter(Text.BeforeDelimiter([#"Quê/Qual"]," "),"-"), type text),
	cor_status = Table.ExpandTableColumn(
		Table.NestedJoin(busca_produto,"Status",#"99-cores-status","Status_Title","dados",JoinKind.LeftOuter),
		"dados",{"Status_Cor"},{"cor_status"}
	),
	transforma_ocultos = Table.TransformColumns(cor_status,{
		{"Ocultar do SOGI (Fechados de 2024)", each if _ = null then "Ativos" else "Mostrar Ocultos", type text}
	})
in
    transforma_ocultos