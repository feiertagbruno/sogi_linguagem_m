let
    Fonte = Excel.Workbook(File.Contents("M:\Manutenção\COMUM\1 - KPI Manutenção Ferramentaria Água e Energia\2025\Database_4Q KPI Ferramentaria.xlsm"), null, true),
    tbPlanAcaoDisp_Table = Fonte{[Item="tbPlanAcaoDisp",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(tbPlanAcaoDisp_Table,{{"Índice", Int64.Type}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por Quê", type text}, {"Quanto", type text}, {"1 Porque", type text}, {"2 Porque", type text}, {"3 Porque", type text}, {"4 Porque", type text}, {"5 Porque", type text}, {"Ação de Contenção", type text}, {"Responsável A", type text}, {"Data Fechamento A", type date}, {"Ação Corretiva", type text}, {"Responsável B", type text}, {"Data Fechamento B", type date}, {"Status", type text}, {"Ocultar do SOGI", type any}})
in
    #"Tipo Alterado"