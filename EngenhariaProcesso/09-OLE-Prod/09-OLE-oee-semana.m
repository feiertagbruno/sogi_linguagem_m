let
    Fonte = #"08-OEE-por-produto",
    #"Outras Colunas Removidas" = Table.SelectColumns(Fonte,{"Data", "Código", "Descrição", "Causa", "meta_uso", "Produção Planejada", "Produção Realizada", "Defeitos", "Diferença", "paradas", "horas_utilizadas"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Outras Colunas Removidas",{{"Causa", "Comentários"}}),
    horas_realizado = Table.RemoveColumns(
        Table.AddColumn(#"Colunas Renomeadas", "horas_realizado", each [paradas] + [horas_utilizadas], type number ),
        {"paradas", "horas_utilizadas"}
    ) ,
    horas_planejado = Table.AddColumn(horas_realizado, "horas_planejado",
        each [Produção Planejada] / [meta_uso] * 60
    ) ,
    Processo = Table.AddColumn(horas_planejado, "Processo", each "Injeção", type text),
    Operação = Table.AddColumn(Processo, "Operação", each "Produção", type text)
in
    Operação