let
    // ano anterior
    Fonte_ano_anterior = Excel.Workbook(
        File.Contents(
            "M:\Manutenção\COMUM\1 - KPI Manutenção Ferramentaria Água e Energia\2024\Database_4Q KPI Ferramentaria.xlsx"
        ),
        null,
        true
    ),
    dados_ano_anterior = Fonte_ano_anterior{[Item = "DBFull24", Kind = "Sheet"]}[Data],
    cabecalho_ano_anterior = Table.PromoteHeaders(dados_ano_anterior, [PromoteAllScalars = true]),
    selecionar_colunas_ano_anterior = Table.SelectColumns(cabecalho_ano_anterior,{"Data","Molde","Horas Disp.","Tempo Parado",
        "Qtd. Paradas","Número protocolo/S.S","Motivo parada","Observação"}),
    tipo_alterado = Table.TransformColumnTypes(selecionar_colunas_ano_anterior,{{"Data", type date}}),
    filtra_ano_anterior = Table.SelectRows(tipo_alterado,each Date.Year([Data]) = #"00-ano-atual" - 1)

in
    filtra_ano_anterior