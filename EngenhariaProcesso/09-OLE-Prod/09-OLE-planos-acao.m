let
    Fonte = Excel.Workbook(File.Contents("M:\Produção\2 INFORMATIVO\01 - CONTROLE PRODUÇÃO\2024\Resumo Diário de Produção.xlsm"), null, true),
    planos_de_acao_OLE_Table = Fonte{[Item="planos_de_acao_OLE",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(planos_de_acao_OLE_Table,{{"Índice", Int64.Type}, {"Semana", Int64.Type}, {"Processo", type text}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por quê", type text}, {"Quanto", type text}, {"1 Porque", type text}, {"2 Porque", type text}, {"3 Porque", type text}, {"4 Porque", type text}, {"5 Porque", type any}, {"Ação de Contenção", type text}, {"Responsável A", type text}, {"Fechamento A", type text}, {"Ação Corretiva", type text}, {"Responsável B", type text}, {"Fechamento B", type text}, {"Status", type text}}),
    processo_semana = Table.AddColumn(#"Tipo Alterado","processo_semana", each Text.Upper(Text.From([Processo])) & "-" & Text.From([Semana]) ),
    cores_status = Table.ExpandTableColumn(
        Table.NestedJoin(processo_semana,"Status", #"99-cores-status","Status_Title", "Dados", JoinKind.LeftOuter)
        , "Dados", {"Status_Cor"}, {"cores_status"}
    )
    
in
    cores_status