let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\7 - IRC - Planos de Ação.xlsm"), null, true),
    tbPlanAcaoIRC_Table = Fonte{[Item="tbPlanAcaoIRC",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(tbPlanAcaoIRC_Table,{{"Índice", Int64.Type}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por Quê", type text}, {"Quanto", type text}, {"1 Porque", type text}, {"2 Porque", type text}, {"3 Porque", type text}, {"4 Porque", type text}, {"5 Porque", type any}, {"Ação de Contenção", type text}, {"Ação Corretiva", type text}, {"Responsável", type text}, {"Data Fechamento", type text}, {"Status", type text}}),
    cores_status = Table.ExpandTableColumn(
        Table.NestedJoin(#"Tipo Alterado", {"Status"}, #"99-cores-status",{"Status_Title"},"dados", JoinKind.LeftOuter)
    ,"dados",{"Status_Cor"},{"status_cores"}
    ),
	ocultos = Table.TransformColumns(cores_status,{{"Ocultar do SOGI", each if _ = null then "Ativos" else "Mostrar Ocultos", type text}})
in
    ocultos