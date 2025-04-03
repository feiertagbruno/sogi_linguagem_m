let
    Fonte = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMrRUitWJVjKCUMamEMpMKTYWAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [#"99num" = _t]),
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"99num", Int64.Type}})
in
    #"Tipo Alterado"