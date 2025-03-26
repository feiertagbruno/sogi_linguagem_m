let
    Fonte_tab = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Custo dos SKUs\Multi Estruturas Sem Frete.xlsx"), null, true),
    det_table = Fonte_tab{[Item="detalhamento",Kind="Table"]}[Data],

    transf_data = Table.TransformColumnTypes(det_table,{"Data de Referência", type date}),

    colunas_necessarias = Table.SelectColumns(transf_data,{
        "Cód Original","Desc Orig","Insumo","Descrição Insumo","Quant Utilizada",
        "Data de Referência","Últ Compra com Premissa"
    }),

    transforma_tipos = Table.TransformColumnTypes(colunas_necessarias,{
        {"Data de Referência", type date},{"Cód Original", type text},{"Desc Orig",type text},
        {"Insumo", type text},{"Descrição Insumo", type text}, {"Quant Utilizada",type number},
        {"Últ Compra com Premissa", type number}
    }),

    ano_num = Table.AddColumn(transforma_tipos,"ano_num", each Date.Year([#"Data de Referência"]), Int16.Type),

	mes_num = Table.AddColumn(ano_num,"mes_num", each [ano_num] * 100 + Date.Month([Data de Referência]), Int32.Type),

    remover_cols = Table.RemoveColumns(mes_num,{"Data de Referência","Descrição Insumo"}),

    group_por_mes_num = Table.Group(remover_cols,{"mes_num","ano_num"},{
        "group",each _, type table[
            #"Cód Original" = Text.Type, #"Desc Orig" = Text.Type, Insumo = Text.Type,
            #"Quant Utilizada" = Number.Type,
            #"Últ Compra com Premissa" = Number.Type
        ]
    }),

    trazer_comparacao_standard = Table.TransformColumns(group_por_mes_num,{
        {"group", (tb) => (
            let
                traz_std = Table.ExpandTableColumn(
                    Table.NestedJoin(tb,{"Cód Original","Insumo"},#"18-SKU-detal-standard",{"COD PA","CODIGO"},"dados",JoinKind.FullOuter)
                    ,"dados", {"COD PA","CODIGO","QTDE.NECESSARIA", "Custo c/ Premissa","DESCRICAO"}, 
                    {"COD PA","CODIGO","qtd_necess_std","custo_std","DESCRICAO"}
                )
            in
                traz_std
        )}
    }),
    expandir = Table.ExpandTableColumn(
        trazer_comparacao_standard,"group",{
            "Cód Original","Desc Orig","Insumo","Quant Utilizada",
            "Últ Compra com Premissa",
            "COD PA","CODIGO","qtd_necess_std","custo_std","DESCRICAO"
        }
    ),

    cols_mes_num = Table.Column(expandir,"mes_num"),
    maior_mes = List.Max(cols_mes_num),
    menor_mes = List.Min(cols_mes_num),

    mes_filtro = Table.AddColumn(expandir,"mes_filtro", each (
        if [mes_num] = menor_mes then "Real WK42/24" else
        if [mes_num] = maior_mes then "Último Mês" else "Mês Anterior"
    ), type text),

    mes_filtro_buffer = Table.Buffer(mes_filtro),

    group_desc_orig = Table.Group(mes_filtro_buffer,"Cód Original",{"Desc Orig", each List.First([Desc Orig]), type text}),
    group_desc_orig_buffer = Table.Buffer(group_desc_orig),

    completa_col_PA = Table.RenameColumns((
        Table.RemoveColumns((
            Table.AddColumn(mes_filtro_buffer,"Cód Original temp",each (
                if [Cód Original] = null then [COD PA] else [Cód Original]
            ), type text)
        ),{"Cód Original"})
    ),{"Cód Original temp","Cód Original"}),

    // completa_descricao_PA = Table.ExpandTableColumn(
    //     Table.NestedJoin((
    //         Table.RemoveColumns(completa_col_PA,{"Desc Orig"})
    //     ),"Cód Original",group_desc_orig_buffer,"Cód Original", "dados", JoinKind.LeftOuter)
    //     ,"dados",{"Desc Orig"}
    // ),

    completa_insumo = Table.RenameColumns((
        Table.RemoveColumns((
            Table.AddColumn(completa_col_PA,"Insumo temp", each (
                if [Insumo] = null then [CODIGO] else [Insumo]
            ), type text)
        ),{"Insumo"})
    ),{"Insumo temp", "Insumo"}),

    remover_colunas_desn = Table.RemoveColumns(completa_insumo,{
        "ano_num","COD PA","CODIGO","DESCRICAO","Desc Orig"
    }),

    // group feito para unificar linhas que poderiam estar duplicadas (somente para garantia)
    group_pre_pareto = Table.Group(remover_colunas_desn,
        {"Cód Original","Insumo","mes_num","mes_filtro"},{
        {"Quant Utilizada", each Number.Round(List.Sum([Quant Utilizada]),6), type number},
        {"Últ Compra com Premissa", each Number.Round(List.Sum([Últ Compra com Premissa]),5), type number},
        {"qtd_necess_std", each Number.Round(List.Sum([qtd_necess_std]),6), type number},
        {"custo_std", each Number.Round(List.Sum([custo_std]),5), type number}
    }),

    group_pareto = Table.Group(group_pre_pareto,{"Cód Original","mes_num","mes_filtro"}, {
        {"group", each _, type table}
    }),

    calcula_paretos = Table.TransformColumns(group_pareto,{
        {"group", (tb) => (
            let
                
            in
        ), type table}
    })



in
    group_pareto