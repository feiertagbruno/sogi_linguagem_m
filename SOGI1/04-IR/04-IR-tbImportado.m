let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\COMUM\22.BI-SOGI\2025\4 - CQMF-21-23-04-Controle de Recebimento de Materiais.xlsm"), null, true),
    tbImportado_Table = Fonte{[Item="tbImportado",Kind="Table"]}[Data],
    #"Outras Colunas Removidas" = Table.SelectColumns(tbImportado_Table,{"Data", "Descrição do Produto", "Fornecedor", "Origem ", "QTD RECEBIDA","Tamanho da Amostra", "QTD Reprovada"}),
    filtro_ano_atual_e_anterior = Table.SelectRows(#"Outras Colunas Removidas",each Date.Year([Data]) >= #"00-ano-atual"-1)
in
    filtro_ano_atual_e_anterior