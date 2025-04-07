let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\3 - 4Q IFinal_novo.xlsm"), null, true),
    tbRIAQ_Table = Fonte{[Item="tbRIAQ",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(tbRIAQ_Table,{{"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por quê", type text}, {"Quanto", type text}, {"1 Porque", type text}, {"2 Porque", type text}, {"3 Porque", type text}, {"4 Porque", type text}, {"5 Porque", type text}, {"Ação de Contenção", type text}, {"Ação Corretiva", type text}, {"Responsável", type text}, {"Fechamento", type text}, {"Status", type text}, {"Imagem", type text}, {"Semana", Int64.Type}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "ordem_riaqs", each if [Status] = "Pendente" then 1 else if [Status] = "Andamento" then 2 else if [Status] = "Fechado" then 3 else 99),
	ocultos = Table.TransformColumns(#"Personalização Adicionada",{{"Ocultar do SOGI", each if _ = null then "Ativos" else "Mostrar Ocultos", type text}}),
	tipos_alterados = Table.TransformColumnTypes(ocultos, {
		{"ordem_riaqs", Int16.Type}
	})
in
    tipos_alterados