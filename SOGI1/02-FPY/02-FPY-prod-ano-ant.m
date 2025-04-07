let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2024\2 - 4Q FPY.xlsx"), null, true),
    Produção_Sheet = Fonte{[Item="Produção",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Produção_Sheet, [PromoteAllScalars=true]),
		seleciona_colunas = Table.SelectColumns(#"Cabeçalhos Promovidos",{"Data","PRODUTO","SKU","Produção Real "}),
    #"Tipo Alterado" = Table.TransformColumnTypes(seleciona_colunas,{
			  {"Data", type date}, {"PRODUTO", type text}, {"SKU", type text}, {"Produção Real ", Int64.Type}
		})
in
    #"Tipo Alterado"