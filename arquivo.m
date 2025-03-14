let
    Fonte = #"01-Arm98-Base",

    defeitos_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(Fonte,"CLASSIF"))
        ,{"CLASSIF", Order.Ascending}
    ),

    index = Table.AddIndexColumn(defeitos_distintos,"index",1,1,Int8.Type),

    cor = Table.ExpandTableColumn(
        Table.NestedJoin(index,"index",#"99-cores","index","dados",JoinKind.LeftOuter)
        ,"dados",{"Cor"}
    )

in
    cor