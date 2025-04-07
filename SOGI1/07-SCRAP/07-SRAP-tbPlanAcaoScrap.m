let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\6 - SCRAP - Planos de Ação.xlsm"), null, true),
    tbPlanAcaoScrap_Table = Fonte{[Item="tbPlanAcaoScrap",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(tbPlanAcaoScrap_Table,{{"Índice", Int64.Type}, {"Quê/Qual", type text}, {"Onde", type text}, {"Quando", type text}, {"Quem", type text}, {"Por Quê", type text}, {"Turno", type text}, {"Quanto", type number}, {"1 Porque", type text}, {"2 Porque", type text}, {"3 Porque", type text}, {"4 Porque", type text}, {"5 Porque", type text}, {"Ação de Contenção", type text}, {"Responsável A", type text}, {"Data Fechamento A", type text}, {"Ação Corretiva", type text}, {"Responsável B", type text}, {"Data Fechamento B", type text}, {"Status", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([#"Quê/Qual"] <> null)),
    semana = Table.AddColumn(#"Linhas Filtradas","Semana", 
        each Number.FromText(
            "20" & Text.AfterDelimiter([Quando],"/")
            & Text.BetweenDelimiters([Quando],"WK","/")
        )
    , Int32.Type),
    cor_status = Table.Join(semana, "Status", #"99-cores-status", "Status_Title"),
    status_ordem = Table.AddColumn(cor_status, "status_ordem", 
        each if Text.Lower([Status]) = "pendente" then 1 else
        if Text.Lower( [Status] ) = "andamento" then 2 else
        if Text.Lower([Status]) = "fechado" then 3 else 99 ),
    #"Tipo Alterado1" = Table.TransformColumnTypes(status_ordem,{{"Data Fechamento A", type text}, {"Data Fechamento B", type text}}),
	ocultos = Table.TransformColumns(#"Tipo Alterado1",{{"Ocultar do SOGI", each if _ = null then "Ativos" else "Mostrar Ocultos", type text}}),
    rel_produto = Table.AddColumn(ocultos,"rel_produto", each Text.From([Semana]) & "-" & Text.Trim([Código]), type text)
in
    rel_produto