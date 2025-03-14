let
    Fonte = #"01-Arm98-Base",
    #"Outras Colunas Removidas" = Table.SelectColumns(Fonte,
		{"CODIGO", "DESCRICAO", "VALOR EM ESTOQUE", "CLASSIF", 
		"mes_num", "semana_num", "semana_texto", "filtro_ultima_semana", 
		"relacao_top3","Cor"}),

	group_por_semana = Table.Group(#"Outras Colunas Removidas","semana_num",{
		{"group_por_semana", each _, type table}
	}),

	// etapa feita para que fique uma linha por produto por classif (que é o status do bloqueio)
	soma_valores_por_produto = Table.TransformColumns(group_por_semana, {
		{"group_por_semana", each Table.Group( _, 
			{"CODIGO","DESCRICAO","CLASSIF","Cor","mes_num","semana_texto","filtro_ultima_semana","relacao_top3"},
			{"VALOR EM ESTOQUE", each List.Sum([VALOR EM ESTOQUE]), type number}
		)}
	}),

	group_por_classif = Table.TransformColumns(soma_valores_por_produto,{
		{"group_por_semana", each
			Table.Group(_,{"CLASSIF","Cor"},{"group_por_classif", each _, type table})
		}
	}),

	sort_e_index_produto = Table.TransformColumns(group_por_classif,{{"group_por_semana", (tabela_semana) => (
			Table.TransformColumns(tabela_semana,{{"group_por_classif", (tabela_classif) => (
				Table.AddIndexColumn((
					Table.Sort(tabela_classif,{"VALOR EM ESTOQUE", Order.Descending})
				),"index_produto",1,1,Int16.Type)
			)}})
		)}
	}),

	filtro_top3_produtos = Table.TransformColumns(sort_e_index_produto,{{"group_por_semana", (tabela_semana) => (
			Table.TransformColumns(tabela_semana,{{"group_por_classif", (tabela_classif) => (
				Table.AddColumn(tabela_classif,"filtro_top3_produtos",each (
					if [index_produto] <= 3 then "Top 3 Produtos" else "Todos os Produtos"
				), type text)
			)}})
		)}
	}),

	expande_tabelas = Table.ExpandTableColumn((
		Table.ExpandTableColumn(filtro_top3_produtos,"group_por_semana",{
			"CLASSIF","Cor","group_por_classif"
		})
	),"group_por_classif",{
		"CODIGO","DESCRICAO","VALOR EM ESTOQUE","mes_num","semana_texto","filtro_ultima_semana","relacao_top3","index_produto","filtro_top3_produtos"
	}),
    #"Tipo Alterado" = Table.TransformColumnTypes(expande_tabelas,{{"CLASSIF", type text}, {"CODIGO", type text}, {"DESCRICAO", type text}, {"VALOR EM ESTOQUE", type number}, {"semana_texto", type text}, {"filtro_ultima_semana", type text}, {"relacao_top3", type text}, {"index_produto", Int16.Type}, {"filtro_top3_produtos", type text}})

in
    #"Tipo Alterado"