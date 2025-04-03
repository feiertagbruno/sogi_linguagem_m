let
    Fonte = Excel.Workbook(File.Contents("M:\Manutenção\COMUM\1 - KPI Manutenção Ferramentaria Água e Energia\2025\Database_4Q KPI Ferramentaria.xlsm"), null, true),
    descricoes_moldes_Table = Fonte{[Item="descricoes_moldes",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(descricoes_moldes_Table,{{"Descrição", type text}, {"Molde", type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Tipo Alterado",{{"Molde", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"Molde", Text.Clean, type text}})
in
    #"Texto Limpo"