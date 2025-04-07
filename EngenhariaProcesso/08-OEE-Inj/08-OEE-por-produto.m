let
    Fonte = Excel.Workbook(File.Contents("M:\INJEÇÃO\RESTRITO\01 CONTROLE PRODUÇÃO INJETADOS\2025\01 Controle de Produção 2025.xlsm"), null, true),
    Diario_Sheet = Fonte{[Item="Diario",Kind="Sheet"]}[Data],
    cabeçalho = Table.PromoteHeaders(Table.Skip(Diario_Sheet,3)),
    #"Outras Colunas Removidas" = Table.SelectColumns(cabeçalho,{"Data", "Wk", "Máquina", "Código", "Descrição", "Meta/Hora", "Produção Planejada", "Produção Realizada", "Defeitos", "Diferença", "Causa", "Uso"}),
    remover_data_em_branco = Table.SelectRows(#"Outras Colunas Removidas", each ([Data] <> null) ),
    remover_meta_com_asterisco = Table.SelectRows(remover_data_em_branco, 
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
        each if _ = null then null else RemoverEspacosDuplos(Text.Trim(Text.From(_)))}}),

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
    agrupamento_por_semana = Table.Group(disponiveis, {"Data", "semana_num", "Máquina", "Código", "Descrição", "Causa", "meta_uso"}, {
        {"Produção Planejada", each List.Sum([Produção Planejada]), type number}, 
        {"Produção Realizada", each List.Sum([Produção Realizada]), type number}, 
        {"Defeitos", each List.Sum([Defeitos]), type number}, 
        {"Diferença", each List.Sum([Diferença]), type number}, 
        {"paradas", each List.Sum([paradas]), type number}, 
        {"horas_utilizadas", each List.Sum([horas_utilizadas]), type number},
        {"disponiveis", each List.Sum([disponiveis]), type number}
    }),
    sort = Table.Sort(agrupamento_por_semana, {{"Data", Order.Ascending}, {"Máquina", Order.Ascending}}),
    #"Tipo Alterado" = Table.TransformColumnTypes(sort,{{"Máquina", Int64.Type}, {"semana_num", Int64.Type}, {"Data", type date}, {"Código", type text}, {"Descrição", type text}}),

    semana_maquina = Table.AddColumn(#"Tipo Alterado", "semana_maquina", 
        each Text.From([semana_num]) & "-" & Text.From([Máquina]), type text),
    
    semana_texto = Table.AddColumn(semana_maquina,"semana_texto", 
        each "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle(Text.From([semana_num]),2,2)
    , type text),
    maquina_texto = Table.AddColumn(semana_texto, "maquina_texto", each Text.From([Máquina]), type text)

in
    maquina_texto