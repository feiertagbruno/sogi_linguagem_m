let
    Fonte = Excel.Workbook(File.Contents("M:\Qualidade\RESTRITO\01-Gestão da Qualidade\01. 4 Quadrantes\01-INDICADORES\2025\5 - 4Q SGI_novo.xlsm"), null, true),
    tbOMSGI_Table = Fonte{[Item="tbOMSGI",Kind="Table"]}[Data],
    #"Linhas Filtradas" = Table.SelectRows(tbOMSGI_Table, each ([Número da OM] <> null)),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Linhas Filtradas",{{"Número da OM", type text}, {"Origem de OM", type text}, {"Data de emissão", type date}, {"Requisito #(lf)NBRISO 9001/14001", type text}, {"Descrição do Problema", type text}, {"Emissor", type text}, {"Data do envio", type date}, {"Processo", type text}, {"Receptor", type text}, {"Gestor                ", type text}, {"Class.", type text}, {"HOJE", type date}, {"DATA DO RETORNO#(lf)REAL", type date}, {"STATUS #(lf)7 DIAS", type date}, {"RESP.", type text}, {"Nessário a emissão de nova OM?", type any}, {"Observações, justificativas, comentários, etc.", type date}, {"STATUS", type text}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Tipo Alterado",{{"Origem de OM", Text.Upper, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto em Maiúscula",{{"Origem de OM", Text.Trim, type text}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Texto Aparado",{{"Gestor                ", "Gestor"}}),

    ocultas_no_sogi = Table.SelectRows(
        #"Colunas Renomeadas", each [Ocultar no SOGI] = null
    ),
    
    trata_status = (status) as text => 
        let
            status_ok = Text.Upper(Text.Trim(Text.From(status))),
            resultado = if status_ok = "ANDAMENTO" then "Andamento" else
                if status_ok = "FECHADO" or status_ok = "FECHADA" then "Fechado" else
                if status_ok = "PENDENTE" then "Pendente" else ""
        in
            resultado,
    arruma_status = Table.TransformColumns(ocultas_no_sogi,{
        {"STATUS", each trata_status(_), type text}
    }),

    relacao = Table.AddColumn(arruma_status,"relacao", each "OM - " & Text.Trim(Text.From([Origem de OM])) & 
        Text.Trim(Text.From([Gestor])) & [STATUS]
    )

in
    relacao