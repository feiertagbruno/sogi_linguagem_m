let

    Fonte = Excel.Workbook(File.Contents("M:\INJEÇÃO\RESTRITO\01 CONTROLE PRODUÇÃO INJETADOS\2025\01 Controle de Produção 2025.xlsm"), null, true),
    Diario_Sheet = Fonte{[Item="tbDiario24",Kind="Table"]}[Data],
    #"Outras Colunas Removidas" = Table.SelectColumns(Diario_Sheet,{"Data", "Wk", "Máquina", "Código", "Meta/Hora", "Produção Planejada", "Produção Realizada", "Defeitos", "Diferença", "Causa", "Uso"}),
    fitro_data = Table.SelectRows(#"Outras Colunas Removidas", each [Data] <> null 
        and ( (Date.Year([Data]) * 100) + Date.WeekOfYear([Data]) ) < #"00-semana-atual-completo" //menor que a semana atual
        and ( (Date.Year([Data]) * 100) + Date.WeekOfYear([Data]) ) > ( (#"00-ano-atual" - 2) * 100 ) + 99 //maior que dois anos atrás
    ),
    remover_meta_com_asterisco = Table.SelectRows(fitro_data, 
        each not Text.Contains(Text.From([#"Meta/Hora"]), "*")
    ),
    multiplicar_meta_uso = Table.AddColumn(remover_meta_com_asterisco,"meta_uso", each [#"Meta/Hora"] / [Uso], type number),
    //modif_nome_meta_uso = Table.RenameColumns(remover_meta_com_asterisco, {{"Meta/Hora", "meta_uso"}}),

    RemoverEspacosDuplos = (texto as text) as text =>
    let
        textoCorrigido = Text.Replace(texto, "  ", " ")
    in
        if Text.Contains(textoCorrigido, "  ") then 
            @RemoverEspacosDuplos(textoCorrigido)
        else 
            textoCorrigido,

    tratamento_col_causa = Table.TransformColumns(multiplicar_meta_uso, {{"Causa", 
        each if _ = null then null else RemoverEspacosDuplos(Text.Trim(_)), type text}}),

    coluna_verificacao_ss = Table.AddColumn(tratamento_col_causa, "Verificação SS", 
        each if [Causa] = null then 0 else
        if Text.Contains(Text.Upper([Causa]), "SS: ") 
        and Text.Middle(Text.AfterDelimiter(Text.Upper([Causa]), "SS: "), 0, 4) = 
            Text.Select(Text.Middle(Text.AfterDelimiter(Text.Upper([Causa]), "SS: "), 0, 4), {"0".."9"})
        then 1 else 0),

    ano_atual = #"00-ano-atual",
    ano_num = Table.AddColumn(coluna_verificacao_ss, "ano_num", each Date.Year([Data]), Int32.Type),
    mes_num = Table.AddColumn(ano_num, "mes_num", each [ano_num] * 100 + Date.Month([Data]), Int32.Type),
    semana_num = Table.AddColumn(mes_num, "semana_num", each [ano_num] * 100 + Date.WeekOfYear([Data]), Int32.Type ),

        
    paradas = Table.AddColumn(semana_num, "paradas", each 
        if [Verificação SS] = 0 then 0 else
        try
            (if [Diferença] < 0 then -[Diferença] / [meta_uso] 
            else [Diferença] / [meta_uso]) * 60 
        otherwise 0 ),
    horas_utilizadas = Table.AddColumn(paradas, "horas_utilizadas", 
        each (([Produção Realizada] + [Defeitos]) / [meta_uso]) * 60
    ),

    //SOMA PRODUÇÃO
    soma_producao_group = Table.Group(horas_utilizadas, {"ano_num", "semana_num", "Máquina"}, {
        {"soma_producao", each List.Sum([Produção Realizada]), type number}
    }),
    soma_producao = Table.Join(horas_utilizadas, {"ano_num", "semana_num", "Máquina"}, 
        soma_producao_group,{"ano_num", "semana_num", "Máquina"}),

    //DIAS TRABALHADOS
    dias_trabalhados = Table.Distinct(Table.SelectColumns(soma_producao,{"Data","semana_num"})),
    dias_trabalhados_count = Table.Group(dias_trabalhados, "semana_num", {{"dias_trabalhados", each List.Count(_)}}),
    mescla_dias_trabalhados = Table.Join(
        soma_producao, "semana_num", dias_trabalhados_count, "semana_num"
    ),

    disponiveis = Table.AddColumn(mescla_dias_trabalhados, "disponiveis", 
        each try let
            resultado = ((24*60*[dias_trabalhados])*([Produção Realizada]/[soma_producao])) - [paradas]  
            in if Number.IsNaN(resultado) then 0 else resultado
        otherwise 0
    ),
    agrupamento_por_dia = Table.Group(disponiveis, {"Data", "Máquina"}, {
        {"Produção Planejada", each List.Sum([Produção Planejada]), type number}, 
        {"Produção Realizada", each List.Sum([Produção Realizada]), type number}, 
        {"Defeitos", each List.Sum([Defeitos]), type number}, 
        {"Diferença", each List.Sum([Diferença]), type number}, 
        {"paradas", each List.Sum([paradas]), type number}, 
        {"horas_utilizadas", each List.Sum([horas_utilizadas]), type number}, 
        {"disponiveis", each List.Sum([disponiveis]), type number}
    }),


    performance = Table.AddColumn(agrupamento_por_dia, "performance", 
        each if [Produção Realizada] = 0 then 0 else
        if [Produção Planejada] = 0 then 100 else
        [Produção Realizada] / [Produção Planejada] * 100 , type number ),
    qualidade = Table.AddColumn(performance, "qualidade", 
        each try
			if [Produção Realizada] = 0 then 0 else
                let resultado = 1 - ([Defeitos] / [Produção Realizada])
                in if Number.IsNaN(resultado) then 0 else resultado * 100
            otherwise 0
            , type number
    ),

    ano_num_2 = Table.AddColumn(qualidade, "ano_num", each Date.Year([Data]), Int64.Type),
    mes_num_2 = Table.AddColumn(ano_num_2, "mes_num", each [ano_num] * 100 + Date.Month([Data]), Int32.Type),
    semana_num_2 = Table.AddColumn(mes_num_2, "semana_num", each [ano_num] * 100 + Date.WeekOfYear([Data]), Int32.Type),

    maior_mes = List.Max(Table.Column(semana_num_2, "mes_num")),

    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(semana_num_2,"mes_num"))
        ,{"mes_num",Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int32.Type),
    add_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(semana_num_2,"mes_num",index_mes,"mes_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes"},{"index_mes"}
    ),

    mes_filtro = Table.AddColumn(add_index_mes, "mes_filtro", 
        each if [index_mes] <= 3 then "Últimos 3 Meses" else "Meses Anteriores"
        , type text
    ),

    ano_texto = Table.AddColumn(mes_filtro , "ano_texto", each Text.From([ano_num]), type text),

    mes_texto = Table.AddColumn(ano_texto, "mes_texto", 
        each Date.MonthName(#date(#"00-ano-atual",Number.From(Text.Middle(Text.From([mes_num]),4,2)),1)) & "/" & Text.Middle([ano_texto],2,2)
    , type text ),


    maior_semana = List.Max(Table.Column(mes_texto, "semana_num")),

    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(mes_texto,"semana_num"))
        ,{"semana_num",Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int32.Type),
    add_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(mes_texto,"semana_num",index_semana,"semana_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"},{"index_semana"}
    ),

    semana_filtro = Table.AddColumn(add_index_semana, "semana_filtro", 
        each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores"
        , type text
    ),
    semana_texto = Table.AddColumn(semana_filtro, "semana_texto", 
        each if [semana_num] = null then null else 
        "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle([ano_texto],2,2)
    , type text )
in
    semana_texto