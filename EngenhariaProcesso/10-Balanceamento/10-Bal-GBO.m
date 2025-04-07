let
    Fonte = Excel.Workbook(File.Contents("M:\Engenharia\COMUM\04. SOGI\GBO\4Q BALANCEAMENTO.xlsx"), null, true),
    tb_GBO_Table = Fonte{[Item="tb_GBO",Kind="Table"]}[Data],
    #"Colunas Removidas" = Table.RemoveColumns(tb_GBO_Table,{"2021"}),
    ano_atual = Date.Year(DateTime.LocalNow()),
    unpivot = Table.UnpivotOtherColumns(#"Colunas Removidas", 
        {"Produto","Modelo"}
    ,"periodo","balanceamento"),
    filtrar_zerados = Table.SelectRows(unpivot, each [balanceamento] <> "-" and [balanceamento] <> 0),

    table_ano = Table.SelectRows(filtrar_zerados, each 
        List.AnyTrue( List.Transform({ano_atual}, (ano) => _[periodo] = Text.From(ano) ) ) 
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(table_ano,{{"periodo", Int64.Type}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado",{{"periodo", "ano_num"}}),

    table_meses = Table.SelectRows(filtrar_zerados, 
        each Number.FromText([periodo]) <= 12
    ),
    #"Tipo Alterado1" = Table.TransformColumnTypes(table_meses,{{"periodo", Int64.Type}}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Tipo Alterado1",{{"periodo", "mes_num"}}),

    unir_tabelas_mes_ano = Table.Combine({#"Colunas Renomeadas",#"Colunas Renomeadas1"}),
    #"Tipo Alterado2" = Table.TransformColumnTypes(unir_tabelas_mes_ano,{{"balanceamento", type number}, {"Produto", type text}, {"Modelo", type text}}),

    preencher_ano_atual = Table.TransformColumns(#"Tipo Alterado2", {"ano_num", each if _ = null then #"00-ano-atual" else _, Int64.Type}),

    ano_texto = Table.AddColumn(preencher_ano_atual, "ano_texto", each Text.From([ano_num]), type text),

    mes_texto = Table.AddColumn(ano_texto, "mes_texto", each 
        if [mes_num] = null then null else Date.MonthName(#date(2024,[mes_num],1)), type text ),
    
    target = Table.AddColumn(mes_texto,"target", each 0.85, type number),

    maior_mes = List.Max(Table.Column(target,"mes_num")),
    ultimos_6_meses = Table.AddColumn(target,"ultimos_6_meses", 
        each if [mes_num] = null then null else
        if [mes_num] > maior_mes - 6 then "Últimos 6 Meses" else "Meses Anteriores"
        , type text)

in
    ultimos_6_meses