let
    caminho_planilha = "M:\Produção\2 INFORMATIVO\01 - CONTROLE PRODUÇÃO\2025\Resumo Diário de Produção.xlsm",
    Fonte = Excel.Workbook(File.Contents(caminho_planilha), null, true),
    tabelaOLEbase_Table = Fonte{[Item="tabelaOLEbase",Kind="Table"]}[Data],
    tabelaOLEbase_Buffer = Table.Buffer(tabelaOLEbase_Table),

    #"Outras Colunas Removidas" = Table.SelectColumns(tabelaOLEbase_Buffer,{
        "Data", "Processo", "Código", "Descrição", "Operação", "Meta/Hr", 
        "Produção Planejada", "Produção Realizada", "Defeitos", "Diferença", "Comentários", "HC"}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Outras Colunas Removidas",{{"Data", type date}}),
    filtr_ano_atual_e_anterior = Table.SelectRows(#"Tipo Alterado1", each (Date.Year([Data]) >= #"00-ano-atual"-1 and [Código] <> null) ),

    nulos_em_0 = Table.ReplaceValue(filtr_ano_atual_e_anterior,null,0,Replacer.ReplaceValue,{"Produção Planejada", "Produção Realizada"}),

    tratamento_planejamentos_zerados = Table.RemoveColumns(
        Table.AddColumn(
            Table.RenameColumns(nulos_em_0, {"Produção Planejada", "Produção Planejada_"})
            ,"Produção Planejada",
            each if ([Produção Planejada_] = 0 or [Produção Planejada_] = null) and [Produção Realizada] > 0
            then [Produção Realizada] else [Produção Planejada_]
        )
        , "Produção Planejada_"
    ),
    #"Colunas Reordenadas" = Table.ReorderColumns(tratamento_planejamentos_zerados,{"Data", "Processo", "Código", "Descrição", "Operação", "Meta/Hr", "Produção Planejada", "Produção Realizada", "Defeitos", "Diferença", "Comentários", "HC"}),

    #"Texto Aparado" = Table.TransformColumns(#"Colunas Reordenadas",{{"Código", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"Código", Text.Clean, type text}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Texto Limpo",{{"Código", Text.Upper, type text}}),

    carga_maquina_Table = Fonte{[Item="carga_maquina",Kind="Table"]}[Data],
    carga_maquina_Buffer = Table.Buffer(carga_maquina_Table),
    cm_colunas_removidas = Table.SelectColumns(carga_maquina_Buffer,{"Item", "HC Teorico"}),
    cm_duplicatas_removidas = Table.Distinct(cm_colunas_removidas, {"Item"}),
    cm_texto_aparado = Table.TransformColumns(cm_duplicatas_removidas,{{"Item", Text.Trim, type text}}),
    cm_texto_limpo = Table.TransformColumns(cm_texto_aparado,{{"Item", Text.Clean, type text}}),
    cm_tipo_alterado = Table.TransformColumnTypes(cm_texto_limpo,{{"HC Teorico", Int64.Type}}),
    cm_texto_em_maiuscula = Table.TransformColumns(cm_tipo_alterado,{{"Item", Text.Upper, type text}}),

    #"Consultas Mescladas" = Table.NestedJoin(#"Texto em Maiúscula", {"Código"}, cm_texto_em_maiuscula, {"Item"}, "09-OLE-carga-maquina", JoinKind.LeftOuter),
    #"09-OLE-carga-maquina Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "09-OLE-carga-maquina", {"HC Teorico"}, {"HC Teorico"}),

    horas_planejado = Table.AddColumn(#"09-OLE-carga-maquina Expandido", "horas_planejado", 
        each try [Produção Planejada] / [#"Meta/Hr"] otherwise 0, type number),
    horas_realizado = Table.AddColumn(horas_planejado, "horas_realizado", 
        each try [Produção Realizada] / [#"Meta/Hr"] otherwise 0, type number),

    tbExcecoes_Table = Fonte{[Item="tbExcecoes",Kind="Table"]}[Data],
    tbExcecoes_Buffer = Table.Buffer(tbExcecoes_Table),
    hr_trab_excecao = Table.ExpandTableColumn(
        Table.NestedJoin(horas_realizado,"Data", tbExcecoes_Buffer,"Data", "Horas de Trabalho", JoinKind.LeftOuter), 
        "Horas de Trabalho" ,{"Horas de Trabalho"}, {"hr_trab_excecao"}),

    tbHorasTrabalhadas_Table = Fonte{[Item="tbHorasTrabalhadas",Kind="Table"]}[Data],
    tbHorasTrabalhadas_Buffer = Table.Buffer(tbHorasTrabalhadas_Table),
    ordem_descrescente = Table.Sort(tbHorasTrabalhadas_Buffer, {"Data Início", Order.Descending}),

    agrupa_plan_horas_trabalhadas = Table.AddColumn(hr_trab_excecao, "hr_trab", each
        let
            data = [Data],
            dia_da_semana = Date.DayOfWeek([Data]),
            filtro = Table.SelectRows(ordem_descrescente, each [N] = dia_da_semana and [Data Início] <= data),
            hTrab = if Table.IsEmpty(filtro) then 0 else filtro{0}[Horas de Trabalho] 
        in hTrab, type number),

    substitui_hc_nulo = Table.ReplaceValue(agrupa_plan_horas_trabalhadas, null, 0, Replacer.ReplaceValue, {"HC", "HC Teorico"}),
    #"Colunas Renomeadas" = Table.RenameColumns(substitui_hc_nulo,{{"Meta/Hr", "meta_uso"}}),

    trazer_injeção = Table.Combine({#"Colunas Renomeadas", #"09-OLE-oee-semana"}),

    semana_num = Table.AddColumn(trazer_injeção, "semana_num", 
        each Date.Year([Data]) * 100 + Date.WeekOfYear([Data]), Int32.Type),
    tira_semana_correte = Table.SelectRows(semana_num, each [semana_num] < #"00-semana-atual-completo"),
    #"Linhas Classificadas" = Table.Sort(tira_semana_correte,{{"Data", Order.Ascending}})
in
    #"Linhas Classificadas"