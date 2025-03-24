let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2025\5 - 4Q SGI_novo.xlsm"), null, true),
    tbPlanAcaoSGI_Table = Fonte{[Item="tbPlanAcaoSGI",Kind="Table"]}[Data],
    #"Linhas Filtradas" = Table.SelectRows(tbPlanAcaoSGI_Table, each ([Evento] <> null)),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Linhas Filtradas",{{"Evento", Text.Upper, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto em Maiúscula",{{"Evento", Text.Trim, type text}}),

    ano_do_plano_sgi = Table.AddColumn(#"Texto Aparado", "Ano", 
        each if Value.Is([Quando], type date) then Date.Year([Quando])
        else Date.Year(Date.FromText(Text.Start(Text.Trim([Quando]), 10))), Int16.Type),

    filtra_maiores_ano_anterior = Table.SelectRows(ano_do_plano_sgi, each [Ano] >= #"00-ano-atual" - 1),

    mes_do_plano_sgi = Table.AddColumn(filtra_maiores_ano_anterior, "Mês", 
        each [Ano] * 100 +
        (if not Value.Is([Quando], type date) then 
        Date.Month(Date.FromText(Text.Start(Text.Trim([Quando]), 10))) else
        Date.Month([Quando])) , Int32.Type),

    ordem_status_sgi = Table.AddColumn(mes_do_plano_sgi, "ordem_status_sgi", each if [Status] = "Pendente" then 1 else
        if [Status] = "Andamento" then 2 else
        if [Status] = "Fechado" then 3 else 99),


    //CORES STATUS
    join_cores_status = Table.Join(ordem_status_sgi,"Status",#"99-cores-status","Status_Title", JoinKind.LeftOuter),
    remover_col_cor_status = Table.RemoveColumns(join_cores_status,{"Status_Title"}),
    #"Personalização Adicionada" = Table.AddColumn(remover_col_cor_status, "rel_sgi_planos", each "NC - " & Text.Trim([Evento]) & Text.Trim([Quem]) & Text.Trim([Status])),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Personalização Adicionada",{{"Fechamento", type text}}),

	ocultos = Table.TransformColumns(#"Tipo Alterado",{{"Ocultar do SOGI", each if _ = null then "Ativos" else "Mostrar Ocultos", type text}})

in
    ocultos