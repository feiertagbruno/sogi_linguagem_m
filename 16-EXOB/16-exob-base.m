let

    Fonte_ano_atual = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\EXOB\Base_EXOB_2025.xlsx"), null, true),
    tbEXOB_Table_ano_atual = Fonte_ano_atual{[Item="tbEXOB",Kind="Table"]}[Data],
    tipo_alterado_ano_atual = Table.TransformColumnTypes(tbEXOB_Table_ano_atual,{{"Ano", Int64.Type}, {"Semana", Int64.Type}, {"Código", type text}, {"Obsoletos", type number}, {"Excessos", type number}}),
    filtra_linhas_zeradas_ano_atual = Table.SelectRows(tipo_alterado_ano_atual,
        each ([Obsoletos] <> 0 and [Obsoletos] <> "") or ([Excessos] <> 0 and [Excessos] <> "")
    ),

    combina_fontes = Table.Buffer(Table.Combine({filtra_linhas_zeradas_ano_atual, #"16-exob-base-ano-anterior"})),

    filtra_PA_e_PI = 
    // Table.RemoveColumns((
        Table.SelectRows((
            Table.ExpandTableColumn(
                Table.NestedJoin(combina_fontes,"Código",#"97-SB1","codigo","dados",JoinKind.LeftOuter)
                ,"dados",{"tipo"}
            )
        ), each not List.Contains({"PA","PI"},[tipo])),
    // ),"tipo"),

    semana_num = Table.AddColumn(filtra_PA_e_PI, "semana_num", each [Ano] * 100 + [Semana], Int32.Type),
    ano_num = Table.RenameColumns(semana_num,{{"Ano", "ano_num"}}),
    mes_num = Table.AddColumn(ano_num, "mes_num", 
        each [ano_num] * 100 + #"coleta_mes_pela_semana"([Semana]), Int32.Type
    ),

    // ano_texto
	group_por_ano = Table.Group(mes_num,"ano_num",{"agrupamento", each _, type table}),
	add_ano_texto_em_cada_ano = Table.TransformColumns(group_por_ano,{
		{"agrupamento", each #"add_ano_texto_na_ultima_semana"(_)}
	}),
	expande_tabela_ano = Table.ExpandTableColumn(
		add_ano_texto_em_cada_ano ,"agrupamento",
		{"ano_texto","mes_num","semana_num","Código","Obsoletos","Excessos"}
	),
    alterar_tipos = Table.TransformColumnTypes(expande_tabela_ano,{
        {"ano_texto", type text},{"mes_num",Int32.Type},{"semana_num", Int32.Type},{"Código",type text},
        {"Obsoletos",type number}, {"Excessos",type number}
    }),


    semana_texto = Table.AddColumn(alterar_tipos, "semana_texto", 
        each "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle(Text.From([ano_num]),2,2), type text
    ),


	maior_semana = List.Max(Table.Column(semana_texto,"semana_num")),
    mes_texto = Table.AddColumn(
        semana_texto, "mes_texto", 
		each if #"valida_relacao_semana_mes_544"(maior_semana,[mes_num],[semana_num]) = 1 then
			Date.MonthName(#date(#"00-ano-atual", Number.FromText(Text.Middle( Text.From([mes_num]),4,2)), 1 )) & "/" &
            Text.Middle(Text.From([ano_num]),2,2) else null
		,type text
    ),


    semanas_distintas = Table.Sort( Table.Distinct(
        Table.SelectColumns(mes_texto,"semana_num")
    ) , {"semana_num",Order.Descending} ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"index_semana",1,1,Int16.Type),
    traz_index_semana = Table.ExpandTableColumn(
        Table.NestedJoin(mes_texto,"semana_num",index_semana, "semana_num","dados",JoinKind.LeftOuter)
        ,"dados",{"index_semana"}
    ),
    filtro_semana = Table.AddColumn(traz_index_semana,"filtro_semana",
        each if [index_semana] <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text
    ),

    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(filtro_semana,"mes_num"))
        ,{"mes_num", Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int16.Type),
    traz_index_mes = Table.ExpandTableColumn(
        Table.NestedJoin(filtro_semana,"mes_num",index_mes,"mes_num","dados",JoinKind.LeftOuter)
        ,"dados", {"index_mes"}
    ),
    filtro_mes = Table.AddColumn(traz_index_mes,"filtro_mes", 
        each if [mes_texto] = null then null else
        if [index_mes] <= 3 then "Últimos 3 Meses" else "Meses Anteriores", type text
     ),

    ////////////////////////////////////////
    // trazer ordem maiores obsoletos e excessos
    // estas etapas serão feitas 2 vezes, uma para obsoletos outra para excessos
    // 1 agrupar por semana
    // 2 ordena por Obsoleto descending e add o index
    // 3 adiciona coluna com o filtro de top 10
    // 4 agrupa novamente por filtro do top 10
    // 5 fazer o transformColumn usando if filtro = "Top 10"
    // 6 adicionar a coluna de pareto -> List.Sum Table.FirstN(tb,index) / List.Sum(colunaObsoleto)
    // 7 fazer isso para os dois e expandir a tabela

    // 1
    group_por_semana = Table.Group(filtro_mes, "semana_num", {"agrupamento", each _, type table}),

    processamentos_semana = Table.TransformColumns(group_por_semana, {"agrupamento", (tb_semana) => (
        let
            //Obsoleto
            //2 e 3
            traz_indices_por_obsoletos = Table.AddColumn((
                Table.AddIndexColumn((
                    Table.Sort(tb_semana,{"Obsoletos",Order.Descending})
                ),"index_obsoletos",1,1,Int16.Type)
            ), "filtro_top10_obsol", each if [index_obsoletos] <= 10 then "Top 10" else "Todos", type text),

            // 5 e 6
            soma_obsoletos = List.Sum(Table.Column(traz_indices_por_obsoletos,"Obsoletos")),

            pareto_obsoleto = Table.AddColumn(traz_indices_por_obsoletos,"pareto_obsoletos", (r) => (
                List.Sum(Table.Column(Table.FirstN(traz_indices_por_obsoletos,r[index_obsoletos]),"Obsoletos")) / soma_obsoletos
            ), type number),

            //Excessos
            //2 e 3
            traz_indices_por_excessos = Table.AddColumn((
                Table.AddIndexColumn((
                    Table.Sort(pareto_obsoleto,{"Excessos",Order.Descending})
                ),"index_excessos",1,1,Int16.Type)
            ), "filtro_top10_excessos", each if [index_excessos] <= 10 then "Top 10" else "Todos", type text),

            // 5 e 6
            soma_excessos = List.Sum(Table.Column(traz_indices_por_excessos,"Excessos")),
            pareto_excessos = Table.AddColumn(traz_indices_por_excessos,"pareto_excessos", (r) => (
                List.Sum(Table.Column(Table.FirstN(traz_indices_por_excessos,r[index_excessos]),"Excessos")) / soma_excessos
            ), type number)

        in
            pareto_excessos
    )}),

    expande_tabelas = Table.ExpandTableColumn(processamentos_semana,"agrupamento",
        {"ano_num","ano_texto","mes_num","mes_texto","semana_texto","index_semana",
        "Código","Obsoletos","Excessos",
        "filtro_mes","filtro_semana",
        "filtro_top10_obsol","index_obsoletos","pareto_obsoletos",
        "filtro_top10_excessos","index_excessos","pareto_excessos"}),


    //////////////////////////////////////// fim

    
    alterar_tipos_2 = Table.TransformColumnTypes(expande_tabelas,{
        {"ano_texto", type text},{"ano_num",Int16.Type},{"mes_texto",type text},{"mes_num",Int32.Type},
        {"semana_texto",type text},{"Código",type text},
        {"Obsoletos",type number}, {"Excessos",type number},
        {"filtro_mes",type text},{"filtro_semana",type text},
        {"filtro_top10_obsol",type text}, {"pareto_obsoletos",type number},
        {"filtro_top10_excessos", type text},{"pareto_excessos",type number}
    }),
    filtro_ult_semana = Table.AddColumn(alterar_tipos_2,"filtro_ult_semana",
        each if [semana_num] = maior_semana then "Última Semana" else "Todas", type text ),

    semana_menos_1 = Table.TransformColumns(traz_index_semana,{
        {"index_semana", each _ - 1, Int16.Type}
    }),
    traz_diferencas = Table.ExpandTableColumn(
        Table.NestedJoin(
            filtro_ult_semana,{"Código","index_semana"},semana_menos_1,{"Código","index_semana"},"dados",JoinKind.LeftOuter
        )
        , "dados",{"Obsoletos","Excessos"},{"obs_semana_passada","exc_semana_passada"}
    ),

    muda_nulos_por_zero = Table.TransformColumns(traz_diferencas,{
        {"obs_semana_passada", each Replacer.ReplaceValue(_,null,0), type number},
        {"exc_semana_passada", each Replacer.ReplaceValue(_,null,0), type number}
    }),

    diferencas_obsoletos = Table.AddColumn(muda_nulos_por_zero,"diferencas_obsoletos", 
        each [Obsoletos] - [obs_semana_passada], type number ),
    diferencas_excessos = Table.AddColumn(diferencas_obsoletos,"diferencas_excessos",
        each [Excessos] - [exc_semana_passada], type number )


in
    diferencas_excessos