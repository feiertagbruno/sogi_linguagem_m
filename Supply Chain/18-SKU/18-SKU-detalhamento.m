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

	mes_num = Table.AddColumn(transforma_tipos,"mes_num", each Date.Year([#"Data de Referência"]) * 100 + 
        Date.Month([Data de Referência]), Int32.Type),

    remover_cols = Table.RemoveColumns(mes_num,{"Data de Referência","Descrição Insumo"}),

    /////

    cols_mes_num = Table.Column(remover_cols,"mes_num"),
    maior_mes = List.Max(cols_mes_num),
    menor_mes = List.Min(cols_mes_num),

    nome_filtro = "Real WK42/24",

    mes_filtro = Table.AddColumn(remover_cols,"mes_filtro", each (
        if [mes_num] = menor_mes then nome_filtro else
        if [mes_num] = maior_mes then "Último Mês" else "Mês Anterior"
    ), type text),

    mes_filtro_buffer = Table.Buffer(mes_filtro),

    // DESCRIÇÃO ORIGINAL BUFFER
    group_desc_orig = Table.Group(mes_filtro_buffer,"Cód Original",{"Desc Orig", each List.First([Desc Orig]), type text}),
    group_desc_orig_buffer = Table.Buffer(group_desc_orig),

    /////

    remover_colunas_desn = Table.RemoveColumns(mes_filtro,{"Desc Orig","mes_num"}),
    rename = Table.RenameColumns(remover_colunas_desn,{
        {"Cód Original","sku"},{"Insumo","insumo"},{"Quant Utilizada","uso"},{"Últ Compra com Premissa","custo"}
    }),

    combine = Table.Combine({rename,#"18-SKU-detal-standard"}),
    combine_buffer = Table.Buffer(combine),

    remove_uso = Table.RemoveColumns(combine_buffer,"uso"),

    ///// PIVOT
    lista_filtros = List.Distinct(Table.Column(remove_uso,"mes_filtro")),

    pivot = Table.Pivot(remove_uso,lista_filtros,"mes_filtro","custo",List.Sum),

    traz_descricao_insumo = Table.ExpandTableColumn(
        Table.NestedJoin(pivot,"insumo",#"97-SB1","codigo","dados",JoinKind.LeftOuter)
        ,"dados",{"descricao"},{"descricao_insumo"}
    ),

    // STANDARD X ULTIMO MES
    diff_std_x_ult = Table.AddColumn(traz_descricao_insumo,"diff_std_x_ult", each Number.Round((
        (if [Standard] = null then 0 else [Standard]) - 
        (if [Último Mês] = null then 0 else [Último Mês])
    ), 2), type number),

    cat_std_x_ult = Table.AddColumn(diff_std_x_ult,"categoria",each (
        if [diff_std_x_ult] > 0 then "positivo" else
        if [diff_std_x_ult] < 0 then "negativo" else "zero"
    ), type text),

    grp_cat_std_x_ult = Table.Group(cat_std_x_ult,{"categoria","sku"},{
        {"group", each _, type table}
    }),
    pareto_std_x_ult = Table.TransformColumns(grp_cat_std_x_ult,{
        {"group", (tb) => (
            let
                categoria = Record.Field(tb{0},"categoria")
            in
                if categoria = "zero" then
                    let
                        add_pareto = Table.AddColumn(tb,"pareto_std_x_ult", each 1, type number)
                    in
                        add_pareto
                else
                    let
                        ordem = if categoria = "positivo" then {
                            {"diff_std_x_ult",Order.Descending},{"descricao_insumo",Order.Ascending}
                        } else {
                            {"diff_std_x_ult",Order.Ascending},{"descricao_insumo",Order.Descending}
                        },
                        somatoria = List.Sum(Table.Column(tb,"diff_std_x_ult")),
                        sort = Table.Sort(tb,ordem),
                        add_index = Table.AddIndexColumn(sort,"index",1,1,Int32.Type),
                        pareto_std_x_ult = Table.AddColumn(add_index,"pareto_std_x_ult", each Number.Round((
                            List.Sum(List.FirstN(Table.Column(add_index,"diff_std_x_ult"),[index])) / somatoria
                        ),2),type number),
                        remove_index = Table.RemoveColumns(pareto_std_x_ult,"index")
                    in
                        remove_index
        ), type table}
    }),
    exp_std_x_ult = Table.ExpandTableColumn(pareto_std_x_ult,"group",
        List.Combine({{"insumo","descricao_insumo"},lista_filtros,{"diff_std_x_ult","pareto_std_x_ult"}})),

    remove_categoria = Table.RemoveColumns(exp_std_x_ult,{"categoria"}),
    
    // WK42 x ÚLTIMO MÊS
    diff_wk42_x_ult = Table.AddColumn(remove_categoria,"diff_wk42_x_ult", each Number.Round((
        (if [#"Real WK42/24"] = null then 0 else [#"Real WK42/24"]) - 
        (if [Último Mês] = null then 0 else [Último Mês])
    ), 2), type number),

    cat_wk42_x_ult = Table.AddColumn(diff_wk42_x_ult,"categoria",each (
        if [diff_wk42_x_ult] > 0 then "positivo" else
        if [diff_wk42_x_ult] < 0 then "negativo" else "zero"
    ), type text),

    grp_cat_wk42_x_ult = Table.Group(cat_wk42_x_ult,{"categoria","sku"},{
        {"group", each _, type table}
    }),
    pareto_wk42_x_ult = Table.TransformColumns(grp_cat_wk42_x_ult,{
        {"group", (tb) => (
            let
                categoria = Record.Field(tb{0},"categoria")
            in
                if categoria = "zero" then
                    let
                        add_pareto = Table.AddColumn(tb,"pareto_wk42_x_ult", each 1, type number)
                    in
                        add_pareto
                else
                    let
                        ordem = if categoria = "positivo" then {
                            {"diff_wk42_x_ult",Order.Descending},{"descricao_insumo",Order.Ascending}
                        } else {
                            {"diff_wk42_x_ult",Order.Ascending},{"descricao_insumo",Order.Descending}
                        },
                        somatoria = List.Sum(Table.Column(tb,"diff_wk42_x_ult")),
                        sort = Table.Sort(tb,ordem),
                        add_index = Table.AddIndexColumn(sort,"index",1,1,Int32.Type),
                        pareto_wk42_x_ult = Table.AddColumn(add_index,"pareto_wk42_x_ult", each Number.Round((
                            List.Sum(List.FirstN(Table.Column(add_index,"diff_wk42_x_ult"),[index])) / somatoria
                        ),2),type number),
                        remove_index = Table.RemoveColumns(pareto_wk42_x_ult,"index")
                    in
                        remove_index
        ), type table}
    }),

    exp_wk42_x_ult = Table.ExpandTableColumn(pareto_wk42_x_ult,"group",
        List.Combine({{"insumo","descricao_insumo"},lista_filtros,{
            "diff_std_x_ult","pareto_std_x_ult","diff_wk42_x_ult","pareto_wk42_x_ult"
        }})),

    remove_cat_wk42_x_ult = Table.RemoveColumns(exp_wk42_x_ult,{"categoria"}),

    // MÊS ANTERIOR X ÚLTIMO MÊS
    diff_ant_x_ult = Table.AddColumn(remove_cat_wk42_x_ult,"diff_ant_x_ult", each Number.Round((
        (if [Mês Anterior] = null then 0 else [Mês Anterior]) - 
        (if [Último Mês] = null then 0 else [Último Mês])
    ), 2), type number),

    cat_ant_x_ult = Table.AddColumn(diff_ant_x_ult,"categoria",each (
        if [diff_ant_x_ult] > 0 then "positivo" else
        if [diff_ant_x_ult] < 0 then "negativo" else "zero"
    ), type text),

    grp_cat_ant_x_ult = Table.Group(cat_ant_x_ult,{"categoria","sku"},{
        {"group", each _, type table}
    }),

    pareto_ant_x_ult = Table.TransformColumns(grp_cat_ant_x_ult,{
        {"group", (tb) => (
            let
                categoria = Record.Field(tb{0},"categoria")
            in
                if categoria = "zero" then
                    let
                        add_pareto = Table.AddColumn(tb,"pareto_ant_x_ult", each 1, type number)
                    in
                        add_pareto
                else
                    let
                        ordem = if categoria = "positivo" then {
                            {"diff_ant_x_ult",Order.Descending},{"descricao_insumo",Order.Ascending}
                        } else {
                            {"diff_ant_x_ult",Order.Ascending},{"descricao_insumo",Order.Descending}
                        },
                        somatoria = List.Sum(Table.Column(tb,"diff_ant_x_ult")),
                        sort = Table.Sort(tb,ordem),
                        add_index = Table.AddIndexColumn(sort,"index",1,1,Int32.Type),
                        pareto_ant_x_ult = Table.AddColumn(add_index,"pareto_ant_x_ult", each Number.Round((
                            List.Sum(List.FirstN(Table.Column(add_index,"diff_ant_x_ult"),[index])) / somatoria
                        ),2),type number),
                        remove_index = Table.RemoveColumns(pareto_ant_x_ult,"index")
                    in
                        remove_index
        ), type table}
    }),

    exp_ant_x_ult = Table.ExpandTableColumn(pareto_ant_x_ult,"group",
        List.Combine({{"insumo","descricao_insumo"},lista_filtros,{
            "diff_std_x_ult","pareto_std_x_ult","diff_wk42_x_ult","pareto_wk42_x_ult",
            "diff_ant_x_ult","pareto_ant_x_ult"
        }})),

    remove_cat_ant_x_ult = Table.RemoveColumns(exp_ant_x_ult,{"categoria"}),

    // STD X WK42/24
    diff_std_x_wk42 = Table.AddColumn(remove_cat_ant_x_ult,"diff_std_x_wk42", each Number.Round((
        (if [Standard] = null then 0 else [Standard]) - 
        (if [#"Real WK42/24"] = null then 0 else [#"Real WK42/24"])
    ), 2), type number),

    cat_std_x_wk42 = Table.AddColumn(diff_std_x_wk42,"categoria",each (
        if [diff_std_x_wk42] > 0 then "positivo" else
        if [diff_std_x_wk42] < 0 then "negativo" else "zero"
    ), type text),

    grp_cat_std_x_wk42 = Table.Group(cat_std_x_wk42,{"categoria","sku"},{
        {"group", each _, type table}
    }),

    pareto_std_x_wk42 = Table.TransformColumns(grp_cat_std_x_wk42,{
        {"group", (tb) => (
            let
                categoria = Record.Field(tb{0},"categoria")
            in
                if categoria = "zero" then
                    let
                        add_pareto = Table.AddColumn(tb,"pareto_std_x_wk42", each 1, type number)
                    in
                        add_pareto
                else
                    let
                        ordem = if categoria = "positivo" then {
                            {"diff_std_x_wk42",Order.Descending},{"descricao_insumo",Order.Ascending}
                        } else {
                            {"diff_std_x_wk42",Order.Ascending},{"descricao_insumo",Order.Descending}
                        },
                        somatoria = List.Sum(Table.Column(tb,"diff_std_x_wk42")),
                        sort = Table.Sort(tb,ordem),
                        add_index = Table.AddIndexColumn(sort,"index",1,1,Int32.Type),
                        pareto_std_x_wk42 = Table.AddColumn(add_index,"pareto_std_x_wk42", each Number.Round((
                            List.Sum(List.FirstN(Table.Column(add_index,"diff_std_x_wk42"),[index])) / somatoria
                        ),2),type number),
                        remove_index = Table.RemoveColumns(pareto_std_x_wk42,"index")
                    in
                        remove_index
        ), type table}
    }),

    exp_std_x_wk42 = Table.ExpandTableColumn(pareto_std_x_wk42,"group",
        List.Combine({{"insumo","descricao_insumo"},lista_filtros,{
            "diff_std_x_ult","pareto_std_x_ult","diff_wk42_x_ult","pareto_wk42_x_ult",
            "diff_ant_x_ult","pareto_ant_x_ult", "diff_std_x_wk42","pareto_std_x_wk42"
        }})),

    remove_cat_std_x_wk42 = Table.RemoveColumns(exp_std_x_wk42,{"categoria"}),


    /////
    
    alterar_tipo = Table.TransformColumnTypes(remove_cat_std_x_wk42,{
        {"sku", type text}, {"insumo",type text},{nome_filtro,type number},
        {"Último Mês", type number},{"Mês Anterior",type number},{"Standard",type number},
        {"diff_std_x_ult", type number}, {"pareto_std_x_ult", type number},
        {"diff_wk42_x_ult", type number}, {"pareto_wk42_x_ult", type number},
        {"diff_ant_x_ult", type number}, {"pareto_ant_x_ult", type number},
        {"diff_std_x_wk42", type number}, {"pareto_std_x_wk42", type number}
    }),

    trazer_descricao_sku = Table.ExpandTableColumn(
        Table.NestedJoin(alterar_tipo,"sku",group_desc_orig_buffer,"Cód Original","dados",JoinKind.LeftOuter)
        ,"dados",{"Desc Orig"},{"descricao_sku_temp"}
    ),
    descricao_sku = Table.RemoveColumns((
        Table.AddColumn(trazer_descricao_sku,"descricao_sku", each (
            Text.From([descricao_sku_temp]) & " - " & Text.From([sku])
        ), type text)
    ),"descricao_sku_temp"),
    concat_descri_insumo = Table.RenameColumns((
        Table.RemoveColumns((
            Table.AddColumn(descricao_sku,"descricao_insumo_temp", each (
                Text.From([descricao_insumo]) & " " & [insumo]
            ), type text)
        ),"descricao_insumo")
    ),{"descricao_insumo_temp","descricao_insumo"})

in
    concat_descri_insumo