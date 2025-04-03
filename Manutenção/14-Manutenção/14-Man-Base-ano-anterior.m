let
    // ano anterior
    Fonte_ano_anterior = Excel.Workbook(
        File.Contents(
            "M:\Manutenção\COMUM\1 - KPI Manutenção Ferramentaria Água e Energia\2024\Database_4Q KPI Manutenção.xlsx"
        ),
        null,
        true
    ),
    dados_ano_anterior = Fonte_ano_anterior{[Item = "DBFull24", Kind = "Sheet"]}[Data],
    cabecalho_ano_anterior = Table.PromoteHeaders(dados_ano_anterior, [PromoteAllScalars = true]),
    selecionar_colunas_ano_anterior = Table.SelectColumns(cabecalho_ano_anterior,{"Data","Máquinas","Turno","Horas Disp.","Horas Manut.",
        "Qtd. Paradas","Número protocolo/S.S/O.S","Motivo parada","Observação"})

in
    selecionar_colunas_ano_anterior