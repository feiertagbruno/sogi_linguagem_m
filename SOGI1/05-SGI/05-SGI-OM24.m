let
    Fonte = #"05-SGI-Op-melhoria",
    renome_coluna_evento = Table.RenameColumns(Fonte,{{"Origem de OM", "Evento"}}),

	ano_num = Table.AddColumn(renome_coluna_evento, "Ano", each Date.Year([Data de emissão]), Int16.Type),
	mes_num = Table.AddColumn(ano_num,"Mês", each [Ano] * 100 + Date.Month([Data de emissão]), Int32.Type),

    group_by_eventos = Table.Group(mes_num, {"Evento", "STATUS", "Processo", "Gestor","Mês","Ano"}, {
        {"Qtd", each Table.RowCount(Table.Distinct(_)), Int64.Type}}),
    coluna_OM = Table.AddColumn(group_by_eventos,"Ocorrência", each "OM"),
    evento_ocorrencia = Table.AddColumn(coluna_OM, "evento_ocorrencia", each [Ocorrência] & " - " & [Evento]),
    
    #"Colunas Renomeadas" = Table.RenameColumns(evento_ocorrencia,{{"STATUS", "Status"}})
in
    #"Colunas Renomeadas"