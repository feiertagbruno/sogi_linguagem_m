let
    Fonte = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("VdK5DcNADEXBXjZWIB7roxZB/bdhRQYnfNmAn9e18sxex4p1H//IGTWjZ+wZrxnvGZ8Z3xlxUhgCRKAIGIEjgASSgBJYEkt6DyyJJbEklsSSWBJLYikshaUcB0thKSyFpbAUlsLSWBpLY2k/BUtjaSyNpbE0lo1lY9lYNpb9WO4f", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ano_num = _t, semana_num = _t]),
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"ano_num", Int64.Type}, {"semana_num", Int64.Type}})
in
    #"Tipo Alterado"