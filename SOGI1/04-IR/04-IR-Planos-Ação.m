let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\4 - CQMF-21-23-04-Controle de Recebimento de Materiais.xlsm"), null, true),
    planos_de_acao_Table = Fonte{[Item="planos_de_acao",Kind="Table"]}[Data],
    #"Personalização Adicionada" = Table.AddColumn(planos_de_acao_Table, "ordem_status", each if [Status] = "Pendente" then 1 else 
if [Status] = "Andamento" then 2 else
if [Status] = "Fechado" then 3 else 99),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Personalização Adicionada",{{"Índice", Int64.Type}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por quê", type text}, {"Quanto", type any}, {"1-Porque", type text}, {"2-Porque", type text}, {"3-Porque", type text}, {"4-Porque", type text}, {"5-Porque", type text}, {"Ação de Contenção", type text}, {"Responsável A", type text}, {"Data Fechamento A", type text}, {"Ação Corretiva", type text}, {"Responsável B", type text}, {"Data Fechamento B", type text}, {"Status", type text}, {"ordem_status", Int64.Type}})
in
    #"Tipo Alterado"