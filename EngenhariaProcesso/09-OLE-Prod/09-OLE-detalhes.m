let
    Fonte = #"09-OLE-base",
    #"Linhas Agrupadas" = Table.Group(Fonte, {"semana_num", "Processo", "Código", "Descrição", "Operação", "meta_uso", "Comentários", "HC", "HC Teorico"}, {{"Produção Planejada", each List.Sum([Produção Planejada]), type nullable number}, {"Produção Realizada", each List.Sum([Produção Realizada]), type nullable number}, {"horas_planejado", each List.Sum([horas_planejado]), type nullable number}, {"horas_realizado", each List.Sum([horas_realizado]), type number}, {"Defeitos", each List.Sum([Defeitos]), type nullable number}, {"Diferença", each List.Sum([Diferença]), type number}}),
    processo_semana = Table.AddColumn(#"Linhas Agrupadas","processo_semana",
        each Text.Upper(Text.From([Processo])) & "-" & Text.From([semana_num])
    , type text),
    semana_texto = Table.AddColumn(processo_semana,"semana_texto",
        each "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle(Text.From([semana_num]),2,2)
    ,type text)
in
    semana_texto