let
    Fonte = Sql.Database("172.16.10.9", "Protheus12", [Query="DECLARE @ULT_DIA_MES_ANTERIOR DATE#(lf)SET @ULT_DIA_MES_ANTERIOR = DATEADD(DAY, -DAY(GETDATE()), GETDATE());#(lf)#(lf)IF OBJECT_ID('tempdb..#CONSULTA') IS NOT NULL DROP TABLE #CONSULTA;#(lf)#(lf)WITH DIAS_MES AS (#(lf)#(tab)SELECT DATEADD(MONTH,-36, @ULT_DIA_MES_ANTERIOR) AS DATA#(lf)#(lf)#(tab)UNION ALL#(lf)#(lf)#(tab)SELECT DATEADD(MONTH,1,DATA) FROM DIAS_MES#(lf)#(tab)WHERE DATA <= GETDATE()#(lf)),#(lf)TODOS_OS_MESES AS (#(lf)#(tab)SELECT YEAR(DATA) AS ANO,#(lf)#(tab)#(tab)MONTH(DATA) AS MES#(lf)#(tab)FROM DIAS_MES#(lf)),#(lf)#(lf)MESES AS (#(lf)#(tab)SELECT DISTINCT#(lf)#(tab)#(tab)DATEPART(YEAR, [Data Emissão]) AS ANO,#(lf)#(tab)#(tab)DATEPART(MONTH, [Data Emissão]) AS MES#(lf)#(tab)FROM VW_MN_AUXILIAR_DECLARACAO_DCI_MENSAL#(lf)#(tab)WHERE [Data Emissão] BETWEEN DATEADD(MONTH,-36, @ULT_DIA_MES_ANTERIOR)#(lf)#(tab)#(tab)AND @ULT_DIA_MES_ANTERIOR#(lf)),#(lf)VENDAS_MENSAIS AS (#(lf)#(tab)SELECT Código AS CODIGO, #(lf)#(tab)#(tab)DATEPART(YEAR, [Data Emissão]) AS ANO,#(lf)#(tab)#(tab)DATEPART(MONTH, [Data Emissão]) AS MES,#(lf)#(tab)#(tab)SUM(QUANT) AS QUANT, Mix#(lf)#(tab)FROM VW_MN_AUXILIAR_DECLARACAO_DCI_MENSAL#(lf)#(tab)WHERE#(lf)#(tab)#(tab)[Descricao CFOP] LIKE 'VENDA%'#(lf)#(tab)#(tab)AND [Data Emissão] BETWEEN DATEADD(MONTH,-36, @ULT_DIA_MES_ANTERIOR)#(lf)#(tab)#(tab)#(tab)AND @ULT_DIA_MES_ANTERIOR#(lf)#(tab)#(tab)AND [Descr. Produto] NOT LIKE '%SUCATA%'#(lf)#(tab)GROUP BY Código, DATEPART(YEAR, [Data Emissão]),#(lf)#(tab)#(tab)DATEPART(MONTH, [Data Emissão]), Mix#(lf)),#(lf)PRODUTOS AS (#(lf)#(tab)SELECT DISTINCT CODIGO#(lf)#(tab)FROM VENDAS_MENSAIS#(lf)),#(lf)PRODUTOS_MESES AS (#(lf)#(tab)SELECT CODIGO, ANO, MES#(lf)#(tab)FROM TODOS_OS_MESES#(lf)#(tab)CROSS JOIN PRODUTOS#(lf)),#(lf)DEVOLUCOES AS (#(lf)#(tab)SELECT #(lf)#(tab)#(tab)Produto,#(lf)#(tab)#(tab)DATEPART(YEAR, [DT Emissao]) AS ANO,#(lf)#(tab)#(tab)DATEPART(MONTH, [DT Emissao]) AS MES,#(lf)#(tab)#(tab)SUM(Quantidade) AS QUANT_DEV#(lf)#(tab)FROM VW_MN_NOTAS_LANCADAS#(lf)#(tab)WHERE [Descricao CFOP] LIKE '%DEVOLUCAO%'#(lf)#(tab)#(tab)AND [Tipo Produto] = 'PA'#(lf)#(tab)#(tab)AND [DT Emissao] BETWEEN DATEADD(MONTH,-36, @ULT_DIA_MES_ANTERIOR)#(lf)#(tab)#(tab)#(tab)AND @ULT_DIA_MES_ANTERIOR#(lf)#(tab)GROUP BY #(lf)#(tab)#(tab)Produto,#(lf)#(tab)#(tab)DATEPART(YEAR, [DT Emissao]),#(lf)#(tab)#(tab)DATEPART(MONTH, [DT Emissao])#(lf)),#(lf)#(lf)CONSULTA AS (#(lf)#(tab)SELECT#(lf)#(tab)#(tab)M.CODIGO as codigo, M.ANO as ano_num, M.MES as mes_num, #(lf)#(tab)#(tab)ISNULL(VM.QUANT,0) AS venda, ISNULL(D.QUANT_DEV,0) AS devolucao,#(lf)#(tab)#(tab)AVG((ISNULL(VM.QUANT,0) - ISNULL(D.QUANT_DEV,0) )) OVER(#(lf)#(tab)#(tab)#(tab)PARTITION BY M.CODIGO#(lf)#(tab)#(tab)#(tab)ORDER BY M.CODIGO, M.ANO, M.MES#(lf)#(tab)#(tab)#(tab)ROWS BETWEEN 11 PRECEDING AND CURRENT ROW#(lf)#(tab)#(tab)) AS media_quant, #(lf)#(tab)#(tab)CASE WHEN Mix = '001' THEN 'PRANCHA'#(lf)#(tab)#(tab)WHEN Mix = '002' THEN 'SECADOR' END AS mix#(lf)#(lf)#(tab)FROM PRODUTOS_MESES M#(lf)#(tab)LEFT JOIN VENDAS_MENSAIS VM ON M.ANO = VM.ANO AND M.MES = VM.MES AND M.CODIGO = VM.CODIGO#(lf)#(tab)LEFT JOIN DEVOLUCOES D ON M.ANO = D.ANO AND M.MES = D.MES AND VM.CODIGO = D.Produto AND M.CODIGO = D.Produto#(lf))#(lf)#(lf)SELECT * INTO #CONSULTA FROM CONSULTA;#(lf)#(lf)WITH MIX AS (#(lf)#(tab)SELECT codigo, mix#(lf)#(tab)FROM #CONSULTA#(lf)#(tab)WHERE mix IS NOT NULL#(lf)#(tab)GROUP BY codigo, mix#(lf))#(lf)#(lf)SELECT  #(lf)#(tab)A.codigo, A.ano_num, A.mes_num, A.venda, A.devolucao, A.media_quant, B.mix#(lf)#(lf)FROM #CONSULTA A#(lf)LEFT JOIN MIX B#(lf)#(tab)ON A.codigo = B.codigo#(lf)"]),
    filtra_ano_atual_e_anterior = Table.SelectRows(Fonte, 
        each [ano_num] >= #"00-ano-atual"-1 ),
    
    //produtos_em_linha = Excel.Workbook(File.Contents("X:\Base IRC_SOGI.xlsx"), null, true){[Item="produtos_em_linha",Kind="Table"]}[Data][Código],
    //filtra_produtos = Table.SelectRows(filtra_ano_atual_e_anterior, each List.Contains(produtos_em_linha, [codigo]) ),

    relação = Table.AddColumn(filtra_ano_atual_e_anterior, "relação", each 
        Date.MonthName(#date(#"00-ano-atual",[mes_num],1)) & "/" & Text.Middle(Text.From([ano_num]),2,2) & "-" & [codigo]
    , type text),
    
    quartil_num = Table.AddColumn(relação, "quartil_num", 
        each if [mes_num] = null then null else
        if List.Contains({1,2,3}, [mes_num]) then 1 else
        if List.Contains({4,5,6}, [mes_num]) then 2 else
        if List.Contains({7,8,9}, [mes_num]) then 3 else 4
        , Int32.Type
    ),
    quartil_texto = Table.AddColumn(quartil_num, "quartil_texto", 
        each "Q" & Text.From([quartil_num]) & "/" & Text.Middle(Text.From([ano_num]),2,2)
    , type text)
    
in
    quartil_texto