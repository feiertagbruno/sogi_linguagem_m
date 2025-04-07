let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2025\Base IRC_SOGI.xlsx"), null, true),
    produtos_em_linha_Table = Fonte{[Item="produtos_em_linha",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(produtos_em_linha_Table,{{"Código", type text}, {"Descrição", type text}, {"Familia", type text}, {"Grupo de Produto", type text}})
in
    #"Tipo Alterado"