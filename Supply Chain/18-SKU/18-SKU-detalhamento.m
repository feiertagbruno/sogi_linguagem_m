let
    Fonte_tab = Excel.Workbook(File.Contents("C:\Users\bruno.martini\Desktop\Multi Estruturas Sem Frete.xlsx"), null, true),
    det_table = Fonte_tab{[Item="detalhamento",Kind="Table"]}[Data],

    transf_data = Table.TransformColumnTypes(det_table,{"Data de Referência", type date}),

    menor_data = List.Min(Table.Column(transf_data,"Data de Referência")),
    maior_data = Date.AddMonths(List.Max(Table.Column(transf_data,"Data de Referência")),-2),

    reduz_colunas = Table.SelectRows(transf_data, each [Data de Referência] >= maior_data or [Data de Referência] = menor_data),

    colunas_necessarias = Table.SelectColumns(reduz_colunas,{
        "Cód Original","Desc Orig","Insumo","Descrição Insumo","Quant Utilizada",
        "Últ Fechamento","Data de Referência","Últ Compra com Premissa"
    }),

    transforma_tipos = Table.TransformColumnTypes(colunas_necessarias,{
        {"Data de Referência", type date},{"Cód Original", type text},{"Desc Orig",type text},
        {"Insumo", type text},{"Descrição Insumo", type text}, {"Quant Utilizada",type number},
        {"Últ Fechamento", type number},{"Últ Compra com Premissa", type number}
    }),

    ano_num = Table.AddColumn(transforma_tipos,"ano_num", each Date.Year([#"Data de Referência"]), Int16.Type),

	mes_num = Table.AddColumn(ano_num,"mes_num", each [ano_num] * 100 + Date.Month([Data de Referência]), Int32.Type),

    col_meses = Table.Column(mes_num, "mes_num"),
    
    menor_mes = List.Min(col_meses),
    maiores_3_meses = List.FirstN( List.Sort( List.Distinct( col_meses ) ,Order.Descending) ,3),
    maior_mes = List.Max(col_meses),

    meses_a_considerar = Table.TransformColumnTypes(
        Table.FromList(
            List.Combine({{menor_mes}, maiores_3_meses})
            ,Splitter.SplitByNothing(),{"meses_a_considerar"}
        )
        ,{"meses_a_considerar", Int32.Type}
    ),

    filtra_meses_a_considerar = Table.RemoveColumns(
        Table.NestedJoin(
            mes_num,"mes_num",meses_a_considerar,"meses_a_considerar","dados",JoinKind.Inner
        ), "dados"
    ),

    relacao = Table.AddColumn(filtra_meses_a_considerar,"relacao",each [Cód Original] & "-" & [Insumo], type text),

    mes_filtro = Table.AddColumn(relacao,"mes_filtro", each (
        if [mes_num] = menor_mes then "Standard" else
        if [mes_num] = maior_mes then "Último Mês" else Date.ToText([Data de Referência], "MMM/yy")
    ), type text)

in
    mes_filtro