let
    Fonte = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("dc2xCoAgEIDhV4lrbUl7gUMvWrKoNmmQFFqypfensII4aP2G/7cW+hB9iEeAAnIqpZAC5sICRu+2y/fblZAyeR2W1fmkolJYlR91THsymsxEPG00tpd3PE2qQd3xdFJkSm32l8FpwPHpvNf5BA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Status_Title = _t, Status_Cor = _t]),
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Status_Title", type text}, {"Status_Cor", type text}})
in
    #"Tipo Alterado"