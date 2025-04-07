let
    Fonte = Excel.Workbook(File.Contents("M:\INJEÇÃO\RESTRITO\01 CONTROLE PRODUÇÃO INJETADOS\2024\01 Controle de Produção 2024.xlsm"), null, true),
    planos_acao_OEE_Table = Fonte{[Item="planos_acao_OEE",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(planos_acao_OEE_Table,{{"Índice", Int64.Type}, {"Semana", Int64.Type}, {"Máquina", Int64.Type}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por quê", type text}, {"Quanto", type text}, {"1 Porque", type text}, {"2 Porque", type text}, {"3 Porque", type text}, {"4 Porque", type text}, {"5 Porque", type any}, {"Ação de Contenção", type text}, {"Responsável A", type text}, {"Fechamento A", type any}, {"Ação Corretiva", type text}, {"Responsável B", type text}, {"Fechamento B", type any}, {"Status", type text}}),
    semana_maquina = Table.AddColumn(#"Tipo Alterado", "semana_maquina", 
        each Text.From([Semana]) & "-" & Text.From([Máquina]), type text),
    cores_status = Table.ExpandTableColumn(
        Table.NestedJoin(semana_maquina, "Status", #"99-cores-status","Status_Title", "cores_status", JoinKind.LeftOuter),
        "cores_status", {"Status_Cor"},{"cores_status"}
    ),
    #"Tipo Alterado1" = Table.TransformColumnTypes(cores_status,{{"Fechamento A", type text}, {"Fechamento B", type text}})

in
    #"Tipo Alterado1"