let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\4 - CQMF-21-23-04-Controle de Recebimento de Materiais.xlsm"), null, true),
    tbNacional_Table = Fonte{[Item="tbNacional",Kind="Table"]}[Data],
    #"Outras Colunas Removidas" = Table.SelectColumns(tbNacional_Table,{"DATA", "Descrição do Produto", "Fornecedor", "Origem ", "QTD RECEBIDA","Tamanho da Amostra", "QTD Reprovada"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Outras Colunas Removidas",{{"DATA", "Data"}}),
    filtro_ano_atual_e_anterior = Table.SelectRows(#"Colunas Renomeadas",each Date.Year([Data]) >= #"00-ano-atual"-1)
in
    filtro_ano_atual_e_anterior