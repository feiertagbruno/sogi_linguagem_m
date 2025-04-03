let
    // ano atual
    Fonte_ano_atual = Excel.Workbook(
        File.Contents(
            "M:\Manutenção\COMUM\1 - KPI Manutenção Ferramentaria Água e Energia\2025\Database_4Q KPI Manutenção.xlsx"
        ),
        null,
        true
    ),
    dados_ano_atual = Fonte_ano_atual{[Item = "DBFull25", Kind = "Sheet"]}[Data],
    cabecalho_ano_atual = Table.PromoteHeaders(dados_ano_atual, [PromoteAllScalars = true]),
    selecionar_colunas_ano_atual = Table.SelectColumns(cabecalho_ano_atual,{"Data","Máquinas","Turno","Horas Disp.","Horas Manut.",
        "Qtd. Paradas","Número protocolo/S.S/O.S","Motivo parada","Observação"}),


    combinar_anos = Table.Combine({#"14-Man-Base-ano-anterior",selecionar_colunas_ano_atual}),

    #"Tipo Alterado" = Table.TransformColumnTypes(
        combinar_anos,
        {
            {"Data", type date},
            {"Máquinas", type text},
            {"Turno", type text},
            {"Horas Disp.", type number},
            {"Horas Manut.", type number},
            {"Qtd. Paradas", Int16.Type},
            {"Número protocolo/S.S/O.S", type text},
            {"Motivo parada", type text},
            {"Observação", type text}
        }
    ),

    ano_num = Table.AddColumn(#"Tipo Alterado","ano_num", each Date.Year([Data]),Int16.Type),
    semana_num = Table.AddColumn(ano_num, "semana_num", each [ano_num] * 100 + Date.WeekOfYear([Data]), Int32.Type),
    mes_num = Table.AddColumn(semana_num, "mes_num", each [ano_num] * 100 + Date.Month([Data]), Int32.Type),

    ano_texto = Table.AddColumn(mes_num, "ano_texto", each Text.From([ano_num]), type text),
    mes_texto = Table.AddColumn(ano_texto,"mes_texto", each (
        Date.MonthName([Data]) & "/" & Text.Middle([ano_texto],2,2)
    ), type text),
    semana_texto = Table.AddColumn(mes_texto, "semana_texto", each (
        "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle([ano_texto],2,2)
    ), type text),

    remove_data = Table.RemoveColumns(semana_texto,"Data"),

    group_para_reduz_linhas = Table.Group(remove_data,{
        "Máquinas","Turno","Número protocolo/S.S/O.S","Motivo parada","Observação",
        "ano_num","ano_texto","mes_num","mes_texto","semana_num","semana_texto"
    },{
        {"Horas Disp.", each List.Sum([#"Horas Disp."]), type number},
        {"Horas Manut.", each List.Sum([#"Horas Manut."]), type number},
        {"Qtd. Paradas", each List.Sum([#"Qtd. Paradas"]), type number}
    }),

    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(group_para_reduz_linhas,"semana_num"))
        ,{"semana_num",Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int16.Type),
    traz_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(group_para_reduz_linhas,{"semana_num"},index_semana,{"semana_num"},"dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"}
    ),
    filtro_semana = Table.RemoveColumns((
        Table.AddColumn(traz_index_semana,"filtro_semana",each (
            if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores"
        ),type text)
    ),{"index_semana"}),

    meses_distintos = Table.Sort((
        Table.Distinct(Table.SelectColumns(filtro_semana,{"mes_num"}))
    ),{"mes_num",Order.Descending}),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int16.Type),
    traz_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(filtro_semana,{"mes_num"},index_mes,{"mes_num"},"dados",JoinKind.LeftOuter)
        , "dados", {"index_mes"}
    ),
    filtro_mes = Table.RemoveColumns((
        Table.AddColumn(traz_index_mes,"filtro_mes", each (
            if [index_mes] <= 3 then "Últimos 3 Meses" else "Meses Anteriores"
        ), type text)
    ),{"index_mes"}),
    
    // ÚLTIMAS 6 SEMANAS COM HORAS MANUTEÇÃO (PARA O 2º QUADRANTE)
    filtra_hr_manut_maior_que_zero = Table.SelectRows(filtro_mes,each (
        [#"Horas Manut."] <> null and [#"Horas Manut."] <> 0
    )),
    semanas_distintas_hr_manut = Table.Sort((
        Table.Distinct(Table.SelectColumns(filtra_hr_manut_maior_que_zero,{"semana_num"}))
    ),{"semana_num",Order.Descending}),
    index_semana_hr_manut = Table.AddIndexColumn(semanas_distintas_hr_manut,"index_semana_hr_manut",1,1,Int16.Type),
    traz_index_semana_hr_manut = Table.ExpandTableColumn(
        Table.NestedJoin(filtro_mes,{"semana_num"},index_semana_hr_manut,{"semana_num"},"dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana_hr_manut"}
    ),
    filtro_semana_hr_manut = Table.RemoveColumns(
        Table.AddColumn(traz_index_semana_hr_manut,"filtro_semana_hr_manut", each (
            if [index_semana_hr_manut] = null then null else
            if [index_semana_hr_manut] <= 6 then "Últimas 6 Semanas"
            else "Semanas Anteriores"
        ), type text)
        ,{"index_semana_hr_manut"}
    )

in
    filtro_semana_hr_manut