let
    Fonte = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("hdIxDsQwCATAv7jeAgPG8JYo///GXXJdiPYaI1saAVofx5gDoxbSrxrw/atyvxfkuqd8DxW1ceIYykTCujAmFKuacN4jsonFhEO6CCLSEN7EZiJRfY9kwjH7VMWEvk01hTexLljoWbDohKW+BN43mSz2LdAXwnKPjXvwB6HBF3YPZf5J/vlXzg8=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Máquina = _t, performance = _t, qualidade = _t, ocupação = _t, OEE = _t, target = _t, ano_num = _t]),
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Máquina", Int64.Type}, {"performance", type number}, {"qualidade", type number}, {"ocupação", type number}, {"OEE", type number}, {"target", Int64.Type}})
in
    #"Tipo Alterado"