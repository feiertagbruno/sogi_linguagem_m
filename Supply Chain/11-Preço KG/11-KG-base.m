let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Mapa de Acompanhamento de Embarques Importados\Database_Logística de Transporte_Custo_Sobregasto e Landed 2025.xlsx"), null, true),
    tbMapa_Table = Fonte{[Item="base_SOGI",Kind="Table"]}[Data],
    #"Erros Removidos" = Table.RemoveRowsWithErrors(tbMapa_Table, {"Mês de Entrada"}),
    
    filtra_ano_atual_e_anterior = Table.SelectRows(#"Erros Removidos",each Date.Year([Mês de Entrada]) >= #"00-ano-atual" - 1 ),

    seleciona_colunas = Table.SelectColumns(filtra_ano_atual_e_anterior,
        {"Reference","ITEM","Description","Forwarder","Modalidade ","Gross Weight",
            "FOB (BRL) AMOUNT","FRETE INTERN. (BRL)","Mês de Entrada"}),
    
    renomeia_colunas = Table.RenameColumns(seleciona_colunas,{
        {"FOB (BRL) AMOUNT","FOB AMOUNT (BRL)"},
        {"Mês de Entrada","Data Real de Entrega"},
        {"Reference","REFERENCE"},
        {"ITEM","PART NUMBER"},
        {"Forwarder","AGENTE DE CARGAS"},
        {"Modalidade ", "MODAL"},
        {"Gross Weight", "GROSS WEIGHT (kgs)"}
    }),

    selecionar_valores = Table.TransformColumns(renomeia_colunas, {
        {"FOB AMOUNT (BRL)", each Text.Select(Text.From(_), {"0".."9",","})},
        {"FRETE INTERN. (BRL)", each Text.Select(Text.From(_),{"0".."9",","})}
    }),
    #"Tipo Alterado" = Table.TransformColumnTypes(selecionar_valores,{
        {"Data Real de Entrega", type date}, 
        {"REFERENCE", type text}, {"PART NUMBER", type text}, 
        {"AGENTE DE CARGAS", type text}, {"MODAL", type text}, {"GROSS WEIGHT (kgs)", type number}, 
        {"FOB AMOUNT (BRL)", type number}, {"FRETE INTERN. (BRL)", type number}
    }),

	//essa cópia serve para mostrar no detalhamento o modal original do processo
	copia_da_coluna_modal = Table.DuplicateColumn(#"Tipo Alterado","MODAL","MODAL ORIGINAL",type text),

    categoria_sea_air = Table.TransformColumns(copia_da_coluna_modal, {"MODAL", 
        each if Text.Contains(Text.Upper(_),"AIR") then "AIR & SEA/AIR" else _, type text}),

    semana_num = Table.AddColumn(categoria_sea_air,"semana_num", 
        each Date.Year([Data Real de Entrega]) * 100 + Date.WeekOfYear([Data Real de Entrega])
    ,Int32.Type ),

    relacao = Table.AddColumn(semana_num, "relacao", 
        each Text.From(Date.Year([Data Real de Entrega])) & "-" & Text.From([semana_num]) & "-" & [AGENTE DE CARGAS] & "-" & [MODAL]
    , type text ),

    relacao_sbg = Table.AddColumn(relacao, "relacao_sbg", each [relacao] & "-" & [REFERENCE], type text ),

    semana_texto = Table.AddColumn(relacao_sbg,"semana_texto", 
        each "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle(Text.From([semana_num]),2,2)
    , type text),

    sobregasto_por_processo = Table.Group(semana_texto,{"REFERENCE"},{
        {"FOB AMOUNT (BRL)", each List.Sum([#"FOB AMOUNT (BRL)"]), type number},
        {"FRETE INTERN. (BRL)", each List.Sum([#"FRETE INTERN. (BRL)"]), type number}
    }),

    frete_ideal = Table.AddColumn(sobregasto_por_processo,"frete_ideal_processo", each [#"FOB AMOUNT (BRL)"] * 0.3, type number),

    sobregasto = Table.AddColumn(frete_ideal, "sobregasto", 
        each if [#"FRETE INTERN. (BRL)"] - [frete_ideal_processo] > 0 then
            [#"FRETE INTERN. (BRL)"] - [frete_ideal_processo] else 0
        , type number
    ),

    porcent_sobregasto_processo = Table.AddColumn(sobregasto, "porcent_sobregasto_processo",
        each if [sobregasto] = 0 then 0 else
        [sobregasto] / [#"FOB AMOUNT (BRL)"]
        , type number
    ),

    traz_sobregasto_e_porcent = Table.ExpandTableColumn(
        Table.NestedJoin(semana_texto,{"REFERENCE"}, porcent_sobregasto_processo,{"REFERENCE"},"dados",JoinKind.LeftOuter)
        ,"dados",{"frete_ideal_processo","sobregasto","porcent_sobregasto_processo"}
    ),

    ano_num = Table.AddColumn(traz_sobregasto_e_porcent,"ano_num", each (
        Date.Year([Data Real de Entrega])
    ), Int32.Type),
    ano_texto = Table.AddColumn(ano_num,"ano_texto", each Text.From([ano_num]), type text),

    mes_num = Table.AddColumn(ano_texto,"mes_num", each (
        [ano_num] * 100 + Date.Month([Data Real de Entrega])
    ), Int32.Type),
    mes_texto = Table.AddColumn(mes_num,"mes_texto", each (
        Date.MonthName([Data Real de Entrega]) & "/" & Text.Middle([ano_texto],2,2) 
    ), type text),

    // FILTRO ÚLTIMAS 6 SEMANAS
    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(mes_texto,"semana_num")),
        {"semana_num",Order.Descending}
    ),
    ult_6_s_indice = Table.AddIndexColumn(semanas_distintas,"filtro_semana", 1,1,Int16.Type),
    filtro_semana = Table.TransformColumns(ult_6_s_indice,{
        {"filtro_semana",each if _ <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text}
    }),

    traz_indice = Table.ExpandTableColumn(
        Table.NestedJoin(mes_texto,{"semana_num"},filtro_semana,{"semana_num"},"dados",JoinKind.LeftOuter)
    , "dados",{"filtro_semana"}
    ),

    // FILTRO ÚLTIMOS 3 MESES
    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(traz_indice,"mes_num")),
        {"mes_num",Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"filtro_mes",1,1,Int16.Type),
    filtro_mes = Table.TransformColumns(index_mes,{
        {"filtro_mes", each if _ <= 3 then "Últimos 3 Meses" else "Meses Anteriores", type text}
    }),
    traz_filtro_mes = Table.ExpandTableColumn(
        Table.NestedJoin(traz_indice,"mes_num",filtro_mes,"mes_num","dados", JoinKind.LeftOuter)
        ,"dados",{"filtro_mes"}
    ),

    cores_agentes_de_carga = Table.ExpandTableColumn(
        Table.NestedJoin(traz_filtro_mes,"AGENTE DE CARGAS",#"11-KG-cores-ag-cargas","AGENTE DE CARGAS","dados",JoinKind.LeftOuter)
        ,"dados",{"Cor"}
    )

in
    cores_agentes_de_carga