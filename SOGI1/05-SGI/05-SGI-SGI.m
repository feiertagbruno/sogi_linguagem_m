let
    Fonte = Table.Combine({#"05-SGI-OM24", #"05-SGI-NC24"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Ano", Int16.Type}, {"Status", type text}, {"Processo", type text}, {"Gestor", type text}, {"Mês", Int32.Type}}),
    ano_texto = Table.AddColumn(#"Tipo Alterado", "ano_texto", each Text.From([Ano]), type text),
    mes_texto = Table.AddColumn(ano_texto, "mes_texto", 
        each if [Mês] = null then null else
        Date.MonthName(#date(2024,Number.FromText(Text.Middle(Text.From([Mês]),4,2)),1)) & "/" & Text.Middle([ano_texto],2,2)
    , type text),


    //adicionando as cores dos eventos

    traz_cor_evento = Table.ExpandTableColumn(
        Table.NestedJoin(mes_texto,{"evento_ocorrencia"},#"05-SGI-cores-eventos", {"evento"}, "dados", JoinKind.LeftOuter)
        ,"dados",{"cor"},{"Cor_Evento"}
    ),

    //adiionando as cores dos gestores
    gestores_distintos = Table.Distinct(Table.SelectColumns(traz_cor_evento,{"Gestor"})),
    #"Linhas Filtradas" = Table.SelectRows(gestores_distintos, each ([Gestor] <> null)),
    add_index_gestor = Table.AddIndexColumn(#"Linhas Filtradas","i",1,1,Int64.Type),
    #"Colunas Renomeadas3" = Table.RenameColumns(add_index_gestor,{{"Gestor", "Gestor_a"}}),
    join_index_gestor = Table.Join(traz_cor_evento,"Gestor", #"Colunas Renomeadas3","Gestor_a", JoinKind.LeftOuter),
    join_cores_gestor = Table.Join(join_index_gestor,"i",#"99-cores","index", JoinKind.LeftOuter),
    remover_i_gestor = Table.RemoveColumns(join_cores_gestor,{"i","Gestor_a","index"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(remover_i_gestor,{{"Cor", "Cor_Gestor"}}),

    //adicionado as cores dos processos
    processos_distintos = Table.Distinct(Table.SelectColumns(#"Colunas Renomeadas1",{"Processo"})),
    filtro_nao_vazias_processos = Table.SelectRows(processos_distintos, each ([Processo] <> null)),
    add_index_processo = Table.AddIndexColumn(filtro_nao_vazias_processos,"i",1,1,Int64.Type),
    #"Colunas Renomeadas4" = Table.RenameColumns(add_index_processo,{{"Processo", "Processo_a"}}),
    join_index_processo = Table.Join(#"Colunas Renomeadas1","Processo", #"Colunas Renomeadas4","Processo_a", JoinKind.LeftOuter),
    join_cores_processo = Table.Join(join_index_processo,"i",#"99-cores","index", JoinKind.LeftOuter),
    remover_i_processo = Table.RemoveColumns(join_cores_processo,{"i","Processo_a","index"}),
    #"Colunas Renomeadas5" = Table.RenameColumns(remover_i_processo,{{"Cor", "Cor_Processo"}}),
    #"Valor Substituído" = Table.ReplaceValue(#"Colunas Renomeadas5","EM ANDAMENTO","Andamento",Replacer.ReplaceText,{"Status"}),
    #"Valor Substituído1" = Table.ReplaceValue(#"Valor Substituído","ATRASADO","Pendente",Replacer.ReplaceText,{"Status"}),
    #"Colocar Cada Palavra Em Maiúscula" = Table.TransformColumns(#"Valor Substituído1",{{"Status", Text.Proper, type text}}),

    //CORES STATUS
    join_cores_status = Table.Join(#"Colocar Cada Palavra Em Maiúscula","Status",#"99-cores-status","Status_Title", JoinKind.LeftOuter),
    remover_col_cor_status = Table.RemoveColumns(join_cores_status,{"Status_Title"}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(remover_col_cor_status,{{"evento_ocorrencia", type text}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado1", "rel_sgi_sgi", each 
        Text.Trim([evento_ocorrencia]) & Text.Trim([Gestor]) & Text.Trim([Status])),

    ano_filtro = Table.AddColumn(#"Personalização Adicionada", "ano_filtro", each 
        if Text.Upper([Status]) = "ANDAMENTO" or [Ano] = #"00-ano-atual" then #"00-ano-atual" else [Ano]
        ,Int16.Type
    )


in
    ano_filtro