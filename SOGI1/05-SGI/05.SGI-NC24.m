let
    Fonte = #"05-SGI-planos-acao",
    #"Linhas Agrupadas" = Table.Group(Fonte, {"Evento", "Status", "Quem", "Onde","Mês", "Ano"}, {{"Qtd", each Table.RowCount(Table.Distinct(_)), Int64.Type}}),
    coluna_NC = Table.AddColumn(#"Linhas Agrupadas","Ocorrência", each "NC"),
    evento_ocorrencia = Table.AddColumn(coluna_NC, "evento_ocorrencia", each [Ocorrência] & " - " & [Evento]),
    #"Colunas Renomeadas" = Table.RenameColumns(evento_ocorrencia,{{"Quem", "Gestor"}, {"Onde", "Processo"}})
in
    #"Colunas Renomeadas"