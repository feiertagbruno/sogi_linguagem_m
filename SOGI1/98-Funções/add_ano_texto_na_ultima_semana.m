(tabela_ano as table) as table => (
		let
			maior_semana = List.Max(Table.Column(tabela_ano,"semana_num")),
			ano_texto = Table.AddColumn(tabela_ano,"ano_texto",
				each if [semana_num] = maior_semana then Text.From([ano_num])
				else null
			, type text)
		in ano_texto
	)