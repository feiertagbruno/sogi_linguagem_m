let

    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2025\Base IRC_SOGI.xlsx"), null, true),
    tbIRC_Table = Fonte{[Item="tbIRC",Kind="Table"]}[Data],

    #"Outras Colunas Removidas" = Table.SelectColumns(tbIRC_Table,
        {"FC", "Referência Produto", "Descrição Produto", "Defeito Constatado", "Família"}),
    ajuste_familia = Table.TransformColumns(#"Outras Colunas Removidas", {{"Família", 
        each if Text.Contains(Text.Upper(_), "PRANCHA") then "PRANCHA" else
        if Text.Contains(Text.Upper(_), "SECADOR") then "SECADOR" else "" }}),

    ano_num = Table.AddColumn(ajuste_familia, "ano_num", each Date.Year([FC]), Int16.Type),
    quartil_num = Table.AddColumn(ano_num, "quartil_num", each [ano_num] * 100 + Date.QuarterOfYear([FC]), Int32.Type),
    mes_num = Table.AddColumn(quartil_num, "mes_num", each [ano_num] * 100 + Date.Month([FC]), Int64.Type),

    // agrupamento para adicionar a coluna com a quantidade de defeitos
    group_por_mes = Table.Group(mes_num, 
        {"ano_num","quartil_num","mes_num","Referência Produto", "Defeito Constatado", "Família"},
        {"quant", each List.Count([Defeito Constatado]), Int64.Type} ),

    ////////////////////////
    // FILTRO ÚLTIMOS 6 MESES
    meses_distintos = Table.Sort((
        Table.Distinct(Table.SelectColumns(group_por_mes, "mes_num"))
    ),{"mes_num",Order.Descending}),
    index_mes = Table.AddIndexColumn(meses_distintos,"index_mes",1,1,Int16.Type),
    transforma_index_em_filtro = Table.TransformColumns(index_mes,{
        {"index_mes",each (if _ <= 6 then "Últimos 6 Meses" else "Meses Anteriores"), type text}
    }),
    traz_filtro = Table.ExpandTableColumn(
        Table.NestedJoin(group_por_mes,{"mes_num"},transforma_index_em_filtro,{"mes_num"},"dados",JoinKind.LeftOuter)
        ,"dados",{"index_mes"},{"filtro_mes"}
    ),
    ////////////////////////

    ///////////////////////
    // FILTRO 4 ULTIMOS QUADRANTES
    quadrantes_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(traz_filtro,"quartil_num"))
        ,{"quartil_num",Order.Descending}
    ),
    index_quartil = Table.AddIndexColumn(quadrantes_distintos,"index_quartil",1,1,Int16.Type),
    transforma_index_em_filtro_quartil = Table.TransformColumns(index_quartil,{
        {"index_quartil", each if _ <= 4 then "Últimos 4 Quartis" else "Quartis Anteriores", type text}
    }),
    traz_filtro_quartil = Table.ExpandTableColumn((
        Table.NestedJoin(traz_filtro,"quartil_num",transforma_index_em_filtro_quartil,"quartil_num","dados",JoinKind.LeftOuter)
    ),"dados",{"index_quartil"},{"filtro_quartil"}),
    ///////////////////////


    ano_texto = Table.AddColumn(traz_filtro_quartil, "ano_texto", each Text.From([ano_num]), type text),
    mes_texto = Table.AddColumn(ano_texto, "mes_texto", 
        each Date.MonthName(#date(#"00-ano-atual",Number.FromText(Text.Middle(Text.From([mes_num]),4,2)),1)) & "/" & Text.Middle([ano_texto],2,2)
    , type text),
    extrai_quartil = (mes_num as number) as number => (
        let
            quartil = Date.QuarterOfYear(#date(2000,Number.FromText(Text.Middle(Text.From(mes_num),4,2)),1))
        in
            quartil
    ),
    quartil_texto = Table.AddColumn(mes_texto, "quartil_texto", 
        each "Q" & Text.From(extrai_quartil([mes_num])) & "/" & Text.Middle([ano_texto],2,2)
    , type text),




    reduz_categ_defeitos = Table.TransformColumns(quartil_texto,{
        {"Defeito Constatado", 
            each
                if List.AnyTrue(List.Transform({"CARCACA","CARCAÇA"}, (p) => Text.Contains(Text.Upper(_), p) )) then "CARCAÇA" else
                if List.AnyTrue(List.Transform( {"RESISTÊNCIA", "RESISTENCIA"}, (p) => Text.Contains(Text.Upper(_),p) )) then "RESISTÊNCIA" else
                if Text.Contains( Text.Upper(_), "PLACA" ) then "PLACA DEFEITUOSA" else
                if Text.Contains( Text.Upper(_), "PATIM" ) then "PATIN" else
                if List.AnyTrue( List.Transform({"CABO", "CAIXA DE CONTATO"}, (p) => Text.Contains( Text.Upper(_),p ) )) then "CABO | CAIXA DE CONTATO" else
				if Text.Contains( Text.Upper(_), "CHAVE") then "CHAVE" else
                if List.AnyTrue( List.Transform({"HELICE","HÉLICE"}, (p) => Text.Contains(Text.Upper(_), p) )) then "HÉLICE" else
                if Text.Contains(Text.Upper(_), "MOTOR") then "MOTOR" else "Z | OUTROS"
        }
    }),

    #"Tipo Alterado" = Table.TransformColumnTypes(reduz_categ_defeitos,{
        {"Referência Produto", type text}, {"Defeito Constatado", type text}, 
        {"Família", type text}, {"quant", Int64.Type}
    }),

    relação = Table.AddColumn(#"Tipo Alterado", "relação", 
        each [mes_texto] & "-" & [Referência Produto]
    , type text),
    
    relacao_defeito = Table.AddColumn(relação,"relacao_defeito", 
        each [mes_texto] & "-" & [Família] & "-" & [Defeito Constatado], type text),
    

    add_produtos_em_linha = Table.Buffer(Table.TransformColumns((Table.ExpandTableColumn((
        Table.NestedJoin(relacao_defeito,"Referência Produto", #"06-IRC-produtos_em_linha", "Código","dados",JoinKind.LeftOuter)
    ),"dados",{"Código"},{"em_linha"})),{
        {"em_linha", each (
            if _ <> null then "Em Linha" else "Fora de Linha"
        ), type text}
    })),

    /* //////////////////////////////////// inicio TRAZER PARETO
    1. AGRUPAR POR UNIDADE DE NEGÓCIO
    2. agrupar por mês
        em paralelo:
    3.  a. agrupar por defeito criando uma coluna com a soma das quantidades
        b. ordenar por defeito decrescente e alfabética crescente
        c. criar coluna de índice
    4. trazer o indice
    5. fazer a divisão da quant pelo total(List.Sum) para trazer o pareto
    6. transformar o índice em filtro_defeito (top 5 e Outros)
    7. abrir groups
    */
    // 1
    group_undidade_negocio = Table.Group(add_produtos_em_linha,{"Família","em_linha","mes_num"},{
        {"group", each _, type table}
    }),
    // 2
    // 3 4 5 6
    index_mes_pareto = Table.TransformColumns(group_undidade_negocio, {"group",(tb_familia) => (
            let
                // cria o indice dos defeitos em ordem
                tb_index_mes = Table.AddIndexColumn((
                    Table.Sort((
                            Table.Group(tb_familia,"Defeito Constatado",{"quant", each List.Sum([quant]), type number})
                    ),{{"quant",Order.Descending},{"Defeito Constatado", Order.Ascending}})
                ),"index_mes_pareto",1,1,Int32.Type),
                // cria o pareto
                pareto = Table.AddColumn(tb_index_mes,"pareto", each (
                    List.Sum(List.FirstN(Table.Column(tb_index_mes,"quant"),[index_mes_pareto])) / 
                        List.Sum(Table.Column(tb_index_mes,"quant"))
                ), type number),
                // Nested Join no pareto e no indice de volta à tabela tb_familia
                traz_index_mes_pareto = Table.ExpandTableColumn(
                    Table.NestedJoin(tb_familia,"Defeito Constatado",pareto,"Defeito Constatado","dados",JoinKind.LeftOuter)
                    ,"dados",{"pareto", "index_mes_pareto"}
                ),
                // faz o filtro Top 5
                transforma_indice_em_filtro = Table.TransformColumns(traz_index_mes_pareto,{
                    {"index_mes_pareto", each if _ <= 5 then "Top 5" else "Outros", type text}
                })
            in
                transforma_indice_em_filtro
    ), type table}),
    // 7 abrir os groups
    abrir_os_groups = Table.ExpandTableColumn(index_mes_pareto, "group",{
        "ano_num","quartil_num","Referência Produto","Defeito Constatado",
        "quant","filtro_mes","filtro_quartil","ano_texto","mes_texto",
        "quartil_texto","relação","relacao_defeito","pareto","index_mes_pareto"
    }),
    alterar_tipos = Table.TransformColumnTypes(abrir_os_groups, {
        {"mes_num",Int32.Type},{"ano_num",Int16.Type},{"quartil_num",Int32.Type},
        {"Referência Produto",type text},{"Defeito Constatado", type text},{"quant", type number},
        {"filtro_mes",type text},{"filtro_quartil", type text},{"ano_texto",type text},
        {"mes_texto",type text},{"quartil_texto",type text},{"relação",type text},
        {"relacao_defeito",type text},{"em_linha", type text},{"pareto", type number},
        {"index_mes_pareto",type text}
    })
    //////////////////////////////////// fim TRAZER PARETO


in
    alterar_tipos