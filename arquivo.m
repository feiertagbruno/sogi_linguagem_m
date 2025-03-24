let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Custo dos SKUs\Multi Estruturas Sem Frete.xlsx"), null, true),
    detalhamento_Table = Fonte{[Item="detalhamento",Kind="Table"]}[Data],

    ano_num = Table.AddColumn(detalhamento_Table,"ano_num", each Date.Year([Data de Referência]), Int16.Type),

	mes_num = Table.AddColumn(ano_num,"mes_num", each [ano_num] * 100 + Date.Month([Data de Referência]), Int32.Type),

    menor_mes = List.Min(Table.Column(mes_num,"mes_num")),
    maiores_3_meses = List.FirstN( List.Sort( List.Distinct( Table.Column(mes_num, "mes_num") ) ,Order.Descending) ,3),
    meses_a_considerar = List.Combine({{menor_mes}, maiores_3_meses}),
    filtro = Table.SelectRows(mes_num, each List.Contains(meses_a_considerar,[mes_num]) ),

    filtra_maior_semana_do_mes = Table.Group(filtro,),

    #"Tipo Alterado" = Table.TransformColumnTypes(detalhamento_Table,{
        {"Cód Original", type text}, {"Tipo Cód Orig", type text}, {"Desc Orig", type text}, 
        {"Código Pai", type text}, {"Desc Pai", type text}, {"Tipo Pai", type text}, 
        {"Insumo", type text}, {"Descrição Insumo", type text}, {"Quant Utilizada", type number}, 
        {"Tipo Insumo", type text}, {"Origem", type text}, {"Alternativos", type text}, 
        {"Últ Compra", type number}, {"Últ Compra Alt", type text}, {"Últ Fechamento", type number}, 
        {"Últ Fecham Alt", type text}, {"Médio Atual", type number}, {"Médio Atual Alt", type text}, 
        {"Data de Referência", type date}, {"Últ Compra com Premissa", type number}
    })
in
    #"Tipo Alterado"