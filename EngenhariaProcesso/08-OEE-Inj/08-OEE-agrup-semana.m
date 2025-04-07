let
    Fonte = #"08-OEE-Diário24",
    agrupamento_por_semana = Table.Group(Fonte, {"ano_num", "semana_num", "semana_texto", "semana_filtro", "Máquina"}, {
        {"Produção Planejada", each List.Sum([Produção Planejada]), type number}, 
        {"Produção Realizada", each List.Sum([Produção Realizada]), type number}, 
        {"Diferença", each List.Sum([Diferença]), type number}, 
        {"Defeitos", each List.Sum([Defeitos]), type number},
        {"paradas", each List.Sum([paradas]), type number},
        {"horas_utilizadas", each List.Sum([horas_utilizadas]), type number},
        {"disponiveis", each List.Sum([disponiveis]), type number},
        {"performance", each List.Average([performance]), type number},
        {"qualidade", each List.Average([qualidade]), type number}
    }),

    ocupação = Table.AddColumn(agrupamento_por_semana, "ocupação", 
        each
            try 
            let resultado = [horas_utilizadas] / [disponiveis]
            in if Number.IsNaN(resultado) then 0 else resultado * 100
            otherwise 0
            , type number
    ),
    OEE = Table.AddColumn(ocupação, "OEE", 
        each [performance] * [qualidade] * [ocupação] / 10000, type number),
    target = Table.AddColumn(OEE, "target", each 31, type number),
    
    #"Tipo Alterado" = Table.TransformColumnTypes(target,{
        {"ano_num", Int64.Type}, {"semana_num", Int64.Type}, {"Máquina", Int64.Type}}),

    semana_maquina = Table.AddColumn(#"Tipo Alterado", "semana_maquina", 
        each Text.From([semana_num]) & "-" & Text.From([Máquina]), type text)

in
    semana_maquina