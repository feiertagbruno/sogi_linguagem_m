let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Mapa de Acompanhamento de Embarques Importados\Database_Logística de Transporte_Custo_Sobregasto e Landed 2025.xlsx"), null, true),
    tbMapa_Table = Fonte{[Item="base_SOGI",Kind="Table"]}[Data],
    #"Erros Removidos" = Table.RemoveRowsWithErrors(tbMapa_Table, {"Mês de Entrada"}),
    
    filtra_ano_atual_e_anterior = Table.SelectRows(#"Erros Removidos",each Date.Year([Mês de Entrada]) >= #"00-ano-atual" - 1 ),

    seleciona_colunas = Table.SelectColumns(filtra_ano_atual_e_anterior,
        {"Reference","ITEM","Forwarder","Modalidade ","Gross Weight","FOB (BRL) AMOUNT","FRETE INTERN. (BRL)","Mês de Entrada"})
    
in
    seleciona_colunas