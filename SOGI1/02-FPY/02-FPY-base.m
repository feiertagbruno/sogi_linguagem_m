let
    // ANO ATUAL
    Fonte_ano_atual = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2025\6 - 4Q FPY.xlsx"), null, true),
    Tabela2_Table_ano_atual = Fonte_ano_atual{[Item="Tabela2",Kind="Table"]}[Data],
    selecionar_colunas_ano_atual = Table.SelectColumns(Tabela2_Table_ano_atual,{"QNT.", "CLASSIFICAÇÃO", "CODPRODUTO", "DATADEF", "DESCRICAODEFEITO", "DESCRICAOMOTIVO", "DESCRICAOSOLUCAO", "LINHA"}),
    tipo_alterado_ano_atual = Table.TransformColumnTypes(selecionar_colunas_ano_atual,{
        {"QNT.", Int64.Type}, {"CLASSIFICAÇÃO", type text}, {"CODPRODUTO", type text},{"DATADEF", type date}, 
        {"DESCRICAODEFEITO", type text}, 
        {"DESCRICAOMOTIVO", type text}, {"DESCRICAOSOLUCAO", type text}, {"LINHA", Int64.Type}}
    ),

    // ANO ANTERIOR
    Fonte_ano_anterior = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2024\2 - 4Q FPY.xlsx"), null, true),
    Tabela2_Table_ano_anterior = Fonte_ano_anterior{[Item="Tabela2",Kind="Table"]}[Data],
    selecionar_colunas_ano_anterior = Table.SelectColumns(Tabela2_Table_ano_anterior,{"QNT.", "CLASSIFICAÇÃO", "CODPRODUTO", "DATADEF", "DESCRICAODEFEITO", "DESCRICAOMOTIVO", "DESCRICAOSOLUCAO", "LINHA"}),
    tipo_alterado_ano_anterior = Table.TransformColumnTypes(selecionar_colunas_ano_anterior,{
        {"QNT.", Int64.Type}, {"CLASSIFICAÇÃO", type text}, {"CODPRODUTO", type text},{"DATADEF", type date}, 
        {"DESCRICAODEFEITO", type text}, 
        {"DESCRICAOMOTIVO", type text}, {"DESCRICAOSOLUCAO", type text}, {"LINHA", Int64.Type}}
    ),

    unir_bases = Table.Combine({tipo_alterado_ano_atual,tipo_alterado_ano_anterior}),
    #"Texto Aparado" = Table.TransformColumns(unir_bases,{{"DESCRICAODEFEITO", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"DESCRICAODEFEITO", Text.Clean, type text}}),
    #"Colocar Cada Palavra Em Maiúscula" = Table.TransformColumns(#"Texto Limpo",{{"DESCRICAODEFEITO", Text.Proper, type text}}),

    ano_num = Table.AddColumn(#"Colocar Cada Palavra Em Maiúscula","ano_num", each Date.Year([DATADEF]), Int32.Type),
    ano_texto = Table.AddColumn(ano_num,"ano_texto", each Text.From([ano_num]), type text),

    // SEMANA
    semana_num = Table.AddColumn(ano_texto,"semana_num", each [ano_num] * 100 + Date.WeekOfYear(Date.AddDays([DATADEF],2)), Int32.Type),
    remove_semana_atual = Table.SelectRows(semana_num, each [semana_num] < #"00-semana-atual-completo" and not List.Contains(#"00-semanas-desconsiderar", [semana_num] )),
    ///////////////
    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(remove_semana_atual,"semana_num")),
        {"semana_num", Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int32.Type),
    traz_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(remove_semana_atual,"semana_num",index_semana,"semana_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"}
    ),
    filtro_semana = Table.AddColumn(traz_index_semana,"filtro_semana", each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text),
    ///////////////
    semana_texto = Table.AddColumn(filtro_semana,"semana_texto",each "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle([ano_texto],2,2), type text ),

    // MÊS
    mes_num = Table.AddColumn(semana_texto,"mes_num", each [ano_num] * 100 + Date.Month([DATADEF]), Int32.Type),
    ///////////////
    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(mes_num,"mes_num"))
        ,{"mes_num", Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int32.Type),
    add_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(mes_num,"mes_num",index_mes,"mes_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes"}
    ),
    filtro_mes = Table.AddColumn(add_index_mes,"filtro_mes",each if [index_mes] <= 3 then "Últimos 3 Meses" else "Meses Anteriores", type text),
    ///////////////
    mes_texto = Table.AddColumn(filtro_mes,"mes_texto", each Text.Middle(Date.MonthName([DATADEF]),0,3) & "/" & Text.Middle([ano_texto],2,2), type text),

    group_para_reduzir_linhas = Table.Group(mes_texto,
        {"CLASSIFICAÇÃO", "CODPRODUTO", "DESCRICAODEFEITO", "DESCRICAOMOTIVO", "DESCRICAOSOLUCAO", "LINHA",
        "ano_num", "ano_texto", "semana_num","semana_texto","filtro_semana","mes_num","mes_texto", "filtro_mes"},
        {{"QNT.", each List.Sum([#"QNT."]), type number }}
    ),


    relacao = Table.AddColumn(group_para_reduzir_linhas,"relacao",each [CLASSIFICAÇÃO] & "-" & Text.From([mes_num]) & "-" & Text.From([semana_num]), type text),

    categoria_defeito = Table.AddColumn(relacao,"categoria_defeito", each [DESCRICAODEFEITO] & " - " & Text.Middle([CLASSIFICAÇÃO],0,1) , type text),

    planta = Table.AddColumn(categoria_defeito,"planta",each "PLANTA",type text),

    // TRAZER 6 MEIORES DEFEITOS BASEADO NA ULTIMA SEMANA, DEPOIS OS MAIORES DAS SEMANAS ANTERIORES

    // tabela criada para solucionar quando não voltam índices vazios na função identificar_maiores_defeitos na etapa agrupa_por_defeito
    defeito_vazia =             Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i44FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [DESCRICAODEFEITO = _t]),
    
    identificar_maiores_defeitos = (tabela as table) => (
        let
            maior_semana = List.Max(Table.Column(tabela,"semana_num")),
            selecionar_maior_semana = Table.Sort(
                Table.Group( 
                    Table.SelectRows(tabela, each [semana_num] = maior_semana)
                , {"DESCRICAODEFEITO"} , {{"QNT.", each List.Sum([#"QNT."]), type number}}
                )
            , {{"QNT.", Order.Descending}}
            ),
            adicionar_indice = Table.AddIndexColumn(selecionar_maior_semana, "indice",1 ,1 ,Int64.Type),

            traz_indices_maiores_defeitos = Table.ExpandTableColumn(
                Table.NestedJoin(tabela, "DESCRICAODEFEITO", adicionar_indice, "DESCRICAODEFEITO", "dados", JoinKind.LeftOuter)
                ,"dados", {"indice"}, {"indice1"}
            ),
            seleciona_indices_nulos = Table.Sort(
                Table.SelectRows(traz_indices_maiores_defeitos, each [indice1] = null)
                , {{"semana_num", Order.Descending}, {"QNT.", Order.Descending}}
            ),
            agrupa_por_defeito =
            try 
            Table.RenameColumns(
                Table.FromList(List.Distinct(Table.Column(seleciona_indices_nulos,"DESCRICAODEFEITO")))
                , {"Column1", "DESCRICAODEFEITO"}
            )
            otherwise
            defeito_vazia
            ,
            
            maior_indice = List.Max(Table.Column(traz_indices_maiores_defeitos,"indice1")),
            add_indice_defeito = Table.AddIndexColumn(agrupa_por_defeito, "indice", maior_indice + 1 ,1 ,Int64.Type),

            traz_indices_outros_defeitos = Table.ExpandTableColumn(
                Table.NestedJoin(traz_indices_maiores_defeitos, "DESCRICAODEFEITO", add_indice_defeito, "DESCRICAODEFEITO", "dados", JoinKind.LeftOuter)
                ,"dados", {"indice"}, {"indice2"}
            ),
            unifica_col_indice = Table.RemoveColumns(
                Table.AddColumn(traz_indices_outros_defeitos, "indice", each if [indice1] <> null then [indice1] else [indice2], Int64.Type)
                , {"indice1", "indice2"}
            )
        in
            unifica_col_indice),
    
    agrupamento_por_categoria = Table.Group(
        planta, {"CLASSIFICAÇÃO"}, {{"agrupamento", each _}}
    ),
    adicionar_indices_defeitos = Table.TransformColumns(agrupamento_por_categoria, {
        "agrupamento", each identificar_maiores_defeitos(_)
    }),
    expande_agrupamento = Table.ExpandTableColumn(adicionar_indices_defeitos, "agrupamento",
        {"CODPRODUTO", "DESCRICAODEFEITO", "DESCRICAOMOTIVO", "DESCRICAOSOLUCAO", "LINHA",
        "ano_num", "ano_texto", "semana_num","semana_texto","filtro_semana","mes_num","mes_texto", "filtro_mes","QNT.",
		"relacao","categoria_defeito","planta","indice"}
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(expande_agrupamento,{{"QNT.", Int64.Type}, {"CODPRODUTO", type text}, {"DESCRICAODEFEITO", type text}, {"DESCRICAOMOTIVO", type text}, {"DESCRICAOSOLUCAO", type text}, {"LINHA", type text}, {"ano_num", Int64.Type}, {"ano_texto", type text}, {"semana_num", Int64.Type}, {"semana_texto", type text}, {"filtro_semana", type text}, {"mes_num", Int64.Type}, {"mes_texto", type text}, {"filtro_mes", type text}, {"relacao", type text}, {"categoria_defeito", type text}, {"planta", type text}, {"indice", Int64.Type}}),

    traz_cores_por_defeito = Table.ExpandTableColumn(
        Table.NestedJoin(#"Tipo Alterado","indice", #"99-cores","index", "dados", JoinKind.LeftOuter)
        , "dados", {"Cor"}, {"cor_defeito"}
    ),
    multiplica_indice_secadores_por_mil = Table.RenameColumns(
        Table.RemoveColumns(
            Table.AddColumn(
                traz_cores_por_defeito, "indice_temp",
                each if [CLASSIFICAÇÃO] = "SECADOR" then [indice] + 1000 else [indice]
                ,Int32.Type
            )//AddColumn
        ,"indice")//removeColumns
    ,{"indice_temp","indice"}),//rename

    relacao_pareto = Table.AddColumn(multiplica_indice_secadores_por_mil,"relacao_pareto",
        each [CLASSIFICAÇÃO] & "-" & Text.From([semana_num]) & "-" & [DESCRICAODEFEITO], type text),

    relacao_rnc = Table.AddColumn(relacao_pareto,"relacao_rnc", each [CLASSIFICAÇÃO] & "-" & [semana_texto] & "-" & [DESCRICAODEFEITO], type text),
    #"Linhas Filtradas" = Table.SelectRows(relacao_rnc, each true)
    
in
    #"Linhas Filtradas"