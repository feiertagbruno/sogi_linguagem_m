let
    Fonte = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("fY7RCoIwFEB/ZaxXg23OTR/vtisMnAsxCMSv6aF/6TPyx0oLZ5A9n8PhDANtLTkS/rgTODvfx84D8W2PXQs0owenZFEwOmYfUeyJXIi8wiRCMB7bHpoZolI21xu4FvCyFpRhTJskdWgbCDDdpmskLhLbzEFcekxKePdi+LevpUQrkri77yqTM5bEr33OS1fXG/hrn1da6DJJ/gQQlgeUNX89jE8=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [evento = _t, cor = _t]),
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"evento", type text}, {"cor", type text}})
in
    #"Tipo Alterado"