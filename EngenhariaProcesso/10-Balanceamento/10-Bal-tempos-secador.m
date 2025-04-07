let
    Fonte = Excel.Workbook(File.Contents("M:\Engenharia\COMUM\04. SOGI\GBO\4Q BALANCEAMENTO.xlsx"), null, true),
    tempos_secador_Table = Fonte{[Item="tempos_secador",Kind="Table"]}[Data],

    subs_nulos_por_0 = Table.TransformColumns(tempos_secador_Table, 
        List.Transform(
            Table.ColumnNames(tempos_secador_Table)
            , each {_, (x) => if x = null then 0 else x} )
    ),

    unpivot = Table.UnpivotOtherColumns(subs_nulos_por_0, {"Família#(lf)Secador"}, "processo","tempo"),

    group_para_indice = Table.Group(unpivot, {"Família#(lf)Secador"}, {"dados", each _} ),
    add_indice = Table.TransformColumns(group_para_indice, {"dados", 
        each Table.AddIndexColumn(_, "indice", 1,1,Int64.Type)
    }),
    expande_indice = Table.ExpandTableColumn(add_indice, "dados", {"processo","tempo","indice"}, {"processo","tempo","indice"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(expande_indice,{{"Família#(lf)Secador", type text}, {"processo", type text}, {"tempo", type number}, {"indice", Int64.Type}})

in
    #"Tipo Alterado"