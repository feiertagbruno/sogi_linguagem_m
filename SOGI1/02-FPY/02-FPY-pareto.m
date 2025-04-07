let
    Fonte = Table.Buffer(#"02-FPY-base"),

    // AGRUPAR PARA REDUZIR LINHAS E COLUNAS
    group_por_semana = Table.Group(
        Fonte,{"CLASSIFICAÇÃO","semana_num","semana_texto","categoria_defeito","DESCRICAODEFEITO"},
        {"QNT.", each List.Sum([#"QNT."]), Int32.Type}
    ),

    // AGRUPAR POR CATEGORIA
    group_por_categoria = Table.Group(
        group_por_semana,{"CLASSIFICAÇÃO"},
        {"group_categoria", each _, type table [#"QNT."=Int32.Type, categoria_defeito=nullable text, semana_num=Int32.Type, semana_text=text]}
    ),

    //DENTRO DO GROUP DA CATEGORIA, AGRUPA POR SEMANA
    group_por_semana_por_categoria = Table.TransformColumns(
        group_por_categoria, {"group_categoria", each Table.Group(_,
            {"semana_num", "semana_texto"},
            {"group_semana", each _, type table [categoria_defeito=nullable text, #"QNT."=Int32.Type]}
        )}
    ),

    /////////////////// inicio
    // 1 ORDENA DESCRESCENTE A QTD POR SEMANA
    // 2 ADICIONA O INDICE
    // 3 FILTRA OS INDICES MAIORES QUE 10
    // 4 ADICIONA A QTD ACUMULADA ATÉ A LINHA USANDO O ÍNDICE COM A FUNÇÃO firstn
    // 5 ADICIONA A QUANTIDADE SOMADA (SOMASES) PARA DEPOIS CALCULAR A PORCENTAGEM ACUMULADA

    // 1, 2 e 3
    index_defeito = Table.TransformColumns(group_por_semana_por_categoria,{"group_categoria", (tabela_categoria) =>
        Table.TransformColumns(tabela_categoria,{"group_semana", (tabela_semana) => 
            Table.SelectRows(

                Table.AddIndexColumn(
                    Table.Sort(tabela_semana,{{"QNT.",Order.Descending},{"categoria_defeito",Order.Ascending}})
                ,"index_defeito",1,1,Int32.Type)
                
            , each [index_defeito] <= 10)
        })
    }),

    // 4
    quantidade_acumulada = Table.TransformColumns(index_defeito,{"group_categoria", (tabela_categoria) =>
        Table.TransformColumns(tabela_categoria,{"group_semana", (tabela_semana) => 
            Table.AddColumn(
                tabela_semana,"qtd_acumulada",
                (r) => List.Sum(Table.Column(Table.FirstN(tabela_semana,r[index_defeito]),"QNT.")), Int32.Type
            )
        })
    }),

    // 5
    quantidade_somada_semana = Table.TransformColumns(quantidade_acumulada,{"group_categoria", (tabela_categoria) =>
        Table.TransformColumns(tabela_categoria,{"group_semana", (tabela_semana) => 
            Table.AddColumn(tabela_semana,"qtd_somada",
                each List.Sum(Table.Column(tabela_semana,"QNT.")), Int32.Type
            )
        })
    }),

    /////////////////// fim


    //DENTRO DO GROUP DA CATEGORIA, EXPANDE AS SEMANAS
    expandir_semanas = Table.TransformColumns(
        quantidade_somada_semana,
        {"group_categoria", each 
            Table.ExpandTableColumn(
                _, "group_semana",{"DESCRICAODEFEITO","categoria_defeito","QNT.","index_defeito","qtd_acumulada","qtd_somada"}
            )
        }
    ),

    // EXPANDE AS CATEGORIAS
    expandir_tabelas = Table.ExpandTableColumn(
        expandir_semanas
        ,"group_categoria", {"semana_num","semana_texto","DESCRICAODEFEITO","categoria_defeito","QNT.","index_defeito","qtd_acumulada","qtd_somada"}
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(expandir_tabelas,{
        {"semana_num", Int64.Type},{"semana_texto", type text},{"DESCRICAODEFEITO",type text},{"categoria_defeito",type text},{"QNT.",Int32.Type},
        {"index_defeito",Int32.Type},{"qtd_acumulada",Int32.Type},{"qtd_somada",Int32.Type}
    }),

    pareto = Table.AddColumn(#"Tipo Alterado","pareto", each Number.Round([qtd_acumulada] / [qtd_somada],2), type number),

    ajustar_index_para_ordem_no_BI = Table.RenameColumns(
        Table.RemoveColumns(
            Table.AddColumn(pareto,"index_temp",
                each if [CLASSIFICAÇÃO] = "SECADOR" then [index_defeito] + 1000 else [index_defeito]
                , Int32.Type
            ) //AddColumn
        ,"index_defeito") //Remove
    ,{"index_temp","index_defeito"}), //Rename

    maior_semana = List.Max(Table.Column(ajustar_index_para_ordem_no_BI,"semana_num")),

    filtro_ultima_semana = Table.AddColumn(ajustar_index_para_ordem_no_BI,"filtro_ultima_semana",
    each if [semana_num] = maior_semana then "Última Semana" else "Semanas Anteriores", type text ),

    relacao_pareto = Table.AddColumn(filtro_ultima_semana,"relacao_pareto",each 
        [CLASSIFICAÇÃO] & "-" & Text.From([semana_num]) & "-" & [DESCRICAODEFEITO], type text)
in
    relacao_pareto