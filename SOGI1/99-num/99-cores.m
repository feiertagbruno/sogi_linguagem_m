let
    Fonte = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("TY4xDkQhCAXv4rZbIKBoiaKX+PH+1/huxGRDM8XwMs8TPjEWmzOs748RqY7DA5hVDucGIM2dKijFWbUxuJ9zJ/eFeXQ8bLURuGOZU3KOqefJ7g+e0X2wVOX2/LVpIZR6GA2V1X8bzbtPe5/d591cfJMLWbn9us/CWi8=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Cor = _t]),
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Cor", type text}}),
    add_index = Table.AddIndexColumn(#"Tipo Alterado","index",1,1,Int64.Type)
in
    add_index