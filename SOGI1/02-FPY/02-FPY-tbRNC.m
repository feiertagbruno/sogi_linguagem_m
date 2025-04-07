let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2025\6 - 4Q FPY.xlsx"), null, true),
    tab_rnc_Table = Fonte{[Item="tab_rnc",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(tab_rnc_Table,{{"Índice", Int64.Type}, {"Semana", type text}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por Quê", type text}, {"Quanto", type text}, {"1 Porque", type text}, {"2 Porque", type text}, {"3 Porque", type text}, {"4 Porque", type text}, {"5 Porque", type any}, {"Ação de Contenção", type text}, {"Responsável A", type text}, {"Fechamento A", type text}, {"Ação Corretiva", type text}, {"Responsável B", type text}, {"Fechamento B", type text}, {"Status", type text}, {"Imagem", type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Tipo Alterado",{{"Quem", Text.Trim, type text},{"Quê/Qual",Text.Trim, type text}}),
    add_column_categoria = Table.AddColumn(#"Texto Aparado", "categoria_fpy_rnc", each Text.Upper(Text.BeforeDelimiter([#"Quê/Qual"]," ")) & [Quem]),
    cor_status = Table.ExpandTableColumn(
        Table.NestedJoin(add_column_categoria,{"Status"},#"99-cores-status",{"Status_Title"},"dados",JoinKind.LeftOuter)
        ,"dados",{"Status_Cor"},{"status_cor"}
    ),
    captura_categoria = (nome_produto) => (
        let
            resultado = (if Text.Upper(Text.Middle(Text.From(nome_produto),0,2)) = "PR" then "PRANCHA" else
                if Text.Upper(Text.Middle(Text.From(nome_produto),0,2)) = "SE" then "SECADOR" else "")
        in
            resultado
    ),
    relacao_rnc = Table.AddColumn(cor_status,"relacao_rnc", each captura_categoria([#"Quê/Qual"]) & "-" & Text.From([Semana]) & "-" & [Quem], type text),
	ocultos = Table.TransformColumns(relacao_rnc,{
		{"Ocultar do SOGI", each if _ = null then "Ativos" else "Mostrar Ocultos", type text}
	})
in
    ocultos