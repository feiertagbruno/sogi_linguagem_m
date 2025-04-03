let
    Fonte = Excel.Workbook(File.Contents("M:\Manutenção\COMUM\1 - KPI Manutenção Ferramentaria Água e Energia\2025\Database_4Q KPI Manutenção.xlsx"), null, true),
    tbPlanAcaoDisp_Table = Fonte{[Item="tbPlanAcaoDisp",Kind="Table"]}[Data],
    #"Linhas Filtradas" = Table.SelectRows(tbPlanAcaoDisp_Table, each ([Índice] <> null)),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Linhas Filtradas",{{"Data Fechamento A", type text}, {"Data Fechamento B", type text}, {"Índice", type text}}),
    traz_cores_status = Table.ExpandTableColumn(
        Table.NestedJoin(#"Tipo Alterado",{"Status"},#"99-cores-status",{"Status_Title"},"dados",JoinKind.LeftOuter)
        ,"dados",{"Status_Cor"}
    )
in
    traz_cores_status