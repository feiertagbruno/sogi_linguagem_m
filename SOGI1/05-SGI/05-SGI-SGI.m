let
    Fonte = Table.Combine({#"05-SGI-SGI23", #"05-SGI-OM24", #"05-SGI-NC24"}),
    mes_texto = Table.AddColumn(Fonte, "mes_texto", 
        each if [Mês] = null then null else
        Date.MonthName(#date(2024,[Mês],1))
    , type text),
    #"Tipo Alterado" = Table.TransformColumnTypes(mes_texto,{{"Ano", Int64.Type}, {"Status", type text}, {"Processo", type text}, {"Gestor", type text}, {"Mês", Int64.Type}}),
    ano_texto = Table.AddColumn(#"Tipo Alterado", "ano_texto", each Text.From([Ano]), type text),

    //adicionando as cores dos eventos
    eventos_distintos = Table.Distinct(Table.SelectColumns(ano_texto,{"evento_ocorrencia"})),
    add_index = Table.AddIndexColumn(eventos_distintos,"i",1,1,Int64.Type),
    #"Colunas Renomeadas2" = Table.RenameColumns(add_index,{{"evento_ocorrencia", "Evento_a"}}),
    join_index = Table.Join(ano_texto,"evento_ocorrencia", #"Colunas Renomeadas2","Evento_a", JoinKind.LeftOuter),
    join_cores = Table.Join(join_index,"i",#"99-cores","index", JoinKind.LeftOuter),
    #"Colunas Removidas" = Table.RemoveColumns(join_cores,{"i","Evento_a", "index"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Colunas Removidas",{{"Cor", "Cor_Evento"}}),

    //adiionando as cores dos gestores
    gestores_distintos = Table.Distinct(Table.SelectColumns(#"Colunas Renomeadas",{"Gestor"})),
    #"Linhas Filtradas" = Table.SelectRows(gestores_distintos, each ([Gestor] <> null)),
    add_index_gestor = Table.AddIndexColumn(#"Linhas Filtradas","i",1,1,Int64.Type),
    #"Colunas Renomeadas3" = Table.RenameColumns(add_index_gestor,{{"Gestor", "Gestor_a"}}),
    join_index_gestor = Table.Join(#"Colunas Renomeadas","Gestor", #"Colunas Renomeadas3","Gestor_a", JoinKind.LeftOuter),
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
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado1", "rel_sgi_sgi", each Text.Trim([evento_ocorrencia]) & Text.Trim([Gestor]) & Text.Trim([Status]))


in
    #"Personalização Adicionada"