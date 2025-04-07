let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2025\6 - 4Q FPY.xlsx"), null, true),
    Produção_Sheet = Fonte{[Item="Produção",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Produção_Sheet, [PromoteAllScalars=true]),
    seleciona_colunas = Table.SelectColumns(#"Cabeçalhos Promovidos",{"Data","PRODUTO","SKU","Produção Real "}),
    #"Tipo Alterado" = Table.TransformColumnTypes(seleciona_colunas,{
			{"Data", type date}, {"PRODUTO", type text}, {"SKU", type text}, {"Produção Real ", Int64.Type}
    }),
    prod_ano_anterior = #"02-FPY-prod-ano-ant",
    combine = Table.Combine({prod_ano_anterior,#"Tipo Alterado"}),
    rename = Table.RenameColumns(combine,{
        {"SKU","Código"},{"Produção Real ","Quantidade"},{"PRODUTO","Grupo"}
    }),
    ano_num = Table.AddColumn(rename,"ano_num", each Date.Year([Data]), Int16.Type),
    mes_num = Table.AddColumn(ano_num,"mes_num", each [ano_num] * 100 + Date.Month([Data]), Int32.Type),
    semana_num = Table.AddColumn(mes_num,"semana_num", each [ano_num] * 100 + Date.WeekOfYear([Data]), Int32.Type),
    remove_data = Table.RemoveColumns(semana_num,"Data"),
    relacao = Table.AddColumn(remove_data,"relacao", each [Grupo] & "-" & Text.From([mes_num]) & "-" & Text.From([semana_num]), type text)
in
    relacao