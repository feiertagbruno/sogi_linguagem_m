let
    Fonte = Table.Buffer(Table.Combine({#"04-IR-tbRegional", #"04-IR-tbNacional", #"04-IR-tbImportado"})),
    subst_na_por_0 = Table.ReplaceValue(Fonte,"N/A",0,Replacer.ReplaceValue,{"Tamanho da Amostra", "QTD Reprovada"}),
    nulos_por_0 = Table.TransformColumns(subst_na_por_0, {
        {"QTD RECEBIDA", each if _ = null then 0 else _},
        {"QTD Reprovada", each if _ = null then 0 else _},
        {"Tamanho da Amostra", each if _ = null then 0 else _}
    }),

    filtra_amostras_zeradas = Table.SelectRows(nulos_por_0, each ([Tamanho da Amostra] <> 0)),
    ano_num_IR = Table.AddColumn(filtra_amostras_zeradas, "ano_num_IR", each Date.Year([Data])),
    qtds_em_numeros = Table.TransformColumnTypes(ano_num_IR,{{"QTD RECEBIDA", type number},{"Tamanho da Amostra", type number}, {"QTD Reprovada", type number}}),

    ano_texto_IR = Table.AddColumn(qtds_em_numeros, "ano_texto_IR", each Text.From([ano_num_IR]), type text),
    
    //SEMANA
    semana_num_IR = Table.AddColumn(ano_texto_IR, "semana_num", each 
        Number.FromText([ano_texto_IR] & Text.PadStart(Text.From(Date.WeekOfYear([Data])),2,"0") )
    , Int32.Type),
    //REMOVER SEMANA ATUAL
    semana_atual = #"00-semana-atual-completo",
    remover_semana_atual = Table.SelectRows(semana_num_IR,each [semana_num] < semana_atual),
    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(remover_semana_atual,"semana_num"))
        ,{"semana_num",Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int32.Type),
    add_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(remover_semana_atual,"semana_num",index_semana,"semana_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"},{"index_semana"}
    ),

    filtro_semana_num_IR = Table.AddColumn(
        add_index_semana,"filtro_semana", 
        each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores"
        , type text
    ),
    filtro_semana_texto_IR = Table.AddColumn(filtro_semana_num_IR, "semana_texto",
        each "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle([ano_texto_IR],2,2) , type text),

    //MÊS
    mes_num_IR = Table.AddColumn(filtro_semana_texto_IR, "mes_num_IR", 
        each Number.FromText([ano_texto_IR] & Text.PadStart( Text.From(Date.Month([Data])),2,"0"))
        ,Int32.Type
    ),
    meses_distintos = Table.Sort( Table.Distinct(Table.SelectColumns(mes_num_IR,"mes_num_IR")),
        {"mes_num_IR", Order.Descending}),
    index_mes_num_IR = Table.AddIndexColumn(meses_distintos,"index_mes_num_IR",1,1,Int32.Type),
    add_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(mes_num_IR,"mes_num_IR",index_mes_num_IR,"mes_num_IR","dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes_num_IR"},{"index_mes_num_IR"}
    ),

    filtro_mes_num_IR = Table.AddColumn(
        add_index_mes,"filtro_mes",
        each if [index_mes_num_IR] <= 3 then "Últimos 3 Meses" else "Meses Anteriores", type text
    ),
    mes_texto_IR = Table.AddColumn(filtro_mes_num_IR,"mes_texto_IR",  
        each Date.MonthName([Data]) & "/" & Text.Middle([ano_texto_IR],2,2)
    ),


    #"Colunas Removidas" = Table.RemoveColumns(mes_texto_IR,{"Data","Origem "}),
    #"Linhas Agrupadas" = Table.Group(#"Colunas Removidas", 
        {"Descrição do Produto", "Fornecedor", "ano_num_IR", "ano_texto_IR", "mes_num_IR", "mes_texto_IR", "filtro_mes", "semana_num",
        "filtro_semana", "semana_texto"}, {
            {"QTD RECEBIDA", each List.Sum([QTD RECEBIDA]), type number},
            {"Tamanho da Amostra", each List.Sum([Tamanho da Amostra]), type number}, 
            {"QTD Reprovada", each List.Sum([QTD Reprovada]), type number}
        }),
    /////////
    indice_rejeicao = Table.AddColumn(#"Linhas Agrupadas", "indice_rejeicao", each [QTD Reprovada] / [Tamanho da Amostra]),
    target = Table.AddColumn(indice_rejeicao, "target", each 0.01),
    #"Tipo Alterado" = Table.TransformColumnTypes(target,{{"Descrição do Produto", type text}, {"Fornecedor", type text}, {"QTD RECEBIDA", type number},
        {"Tamanho da Amostra", type number}, {"QTD Reprovada", type number}, {"ano_num_IR", Int64.Type}, {"ano_texto_IR", type text}, 
        {"mes_num_IR", Int64.Type}, {"mes_texto_IR", type text}, {"indice_rejeicao", Percentage.Type}, {"target", Percentage.Type}}),
    
    filtro_ultimas_6_wk_com_defeitos = Table.SelectRows(Table.SelectColumns(#"Tipo Alterado",{"semana_num","QTD Reprovada"}), 
        each [QTD Reprovada] > 0 and [semana_num] <> null),
    semanas_distintas_com_defeitos = Table.Sort(
        Table.Distinct(Table.SelectColumns(filtro_ultimas_6_wk_com_defeitos,"semana_num"))
        , {"semana_num",Order.Descending}
    ),

    add_indice = Table.AddIndexColumn(semanas_distintas_com_defeitos,"indice", 1,1,Int32.Type),

    transforma_indice = Table.TransformColumns(add_indice, {{"indice", each if _ <= 6 then "semanas_dinâmicas" else null}}),

    traz_filtro_semanas_dinamicas = Table.ExpandTableColumn(
        Table.NestedJoin(#"Tipo Alterado", {"semana_num"},transforma_indice,{"semana_num"},"dados",JoinKind.LeftOuter)
    , "dados",{"indice"},{"filtro_semanas_dinamicas"}
    ),

    // ESSA PARTE DO CÓDIGO SERVE PARA FILTRAR NO 4 QUADRANTE SOMENTE FORNECEDORES QUE TIVERAM DEFEITOS
    filtrar_forn_com_defeitos = Table.SelectRows(traz_filtro_semanas_dinamicas,
        each [QTD Reprovada] > 0
    ),
    fornecedores_distintos = Table.Distinct(Table.SelectColumns(filtrar_forn_com_defeitos,"Fornecedor")),
    coluna_fornecedor_com_defeito = Table.AddColumn(fornecedores_distintos,"fornecedor_com_defeito",each "Fornecedor com defeito", type text),
    traz_filtro_fornecedor_com_defeito = Table.ExpandTableColumn(
        Table.NestedJoin(traz_filtro_semanas_dinamicas,"Fornecedor",coluna_fornecedor_com_defeito,"Fornecedor","dados",JoinKind.LeftOuter)
        ,"dados",{"fornecedor_com_defeito"},{"fornecedor_com_defeito"}
    )

in
    traz_filtro_fornecedor_com_defeito