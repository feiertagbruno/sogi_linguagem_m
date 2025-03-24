let
    Fonte = Excel.Workbook(File.Contents("M:\SOGI\2025\BI-SOGI\Bases\Custo dos SKUs\Multi Estruturas Sem Frete.xlsx"), null, true),
    custo_stantard_Table = Fonte{[Item="custo_stantard",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(custo_stantard_Table,{{"CÓDIGO", type text}, {"MATERIA-PRIMA", type number}}),

    menor_semana = #"00-semana-standard",
    
	semana_num = Table.AddColumn(#"Tipo Alterado", "semana_num", each menor_semana, Int32.Type ),

    rename_custo_standard = Table.RenameColumns(semana_num,{
        {"MATERIA-PRIMA","custo_standard"}
    })
in
    rename_custo_standard