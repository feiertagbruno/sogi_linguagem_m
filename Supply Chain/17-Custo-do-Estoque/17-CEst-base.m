let
    Fonte = Sql.Database("172.16.10.9", "Protheus12", [Query="DECLARE @DATA_FECHAMENTO VARCHAR(10) = '20241231';#(lf)DECLARE @DATA_FINAL VARCHAR(10) = CONVERT( varchar, GETDATE(),112 );#(lf)DECLARE @MAIOR_DATA_FECHAMENTO VARCHAR(10) = (SELECT MAX(B9_DATA) FROM VW_MN_SB9 B9 WHERE B9.D_E_L_E_T_ <> '*');#(lf)DECLARE @PRIMEIRO_DIA_MES VARCHAR(10) = CONVERT(VARCHAR,DATEADD(DAY,1,@DATA_FECHAMENTO),112);#(lf)DECLARE @ULTIMO_DIA_MES VARCHAR(10) = #(lf)#(tab)CASE #(lf)#(tab)#(tab)WHEN @DATA_FECHAMENTO = @MAIOR_DATA_FECHAMENTO THEN#(lf)#(tab)#(tab)#(tab)CONVERT(VARCHAR,GETDATE(),112)#(lf)#(tab)#(tab)ELSE #(lf)#(tab)#(tab)#(tab)CONVERT(VARCHAR,EOMONTH(@PRIMEIRO_DIA_MES),112)#(lf)#(tab)END;#(lf)DECLARE @HOJE VARCHAR(10) = CONVERT(VARCHAR,GETDATE(),112);#(lf)DECLARE @ARMAZENS_STRING VARCHAR(MAX) = '11,14,17,20,89,01,80,91,95,85,93,98'#(lf)-- DECLARAR O @ARMAZENS COMO 'TD' PARA TRAZER TODOS OS ARMAZENS#(lf)DECLARE @ARMAZENS TABLE (ARMAZEM VARCHAR(2));#(lf)INSERT INTO @ARMAZENS SELECT value FROM string_split(@ARMAZENS_STRING,',')#(lf)#(lf)IF OBJECT_ID('tempdb..#RESULTADO_FINAL') IS NOT NULL DROP TABLE #RESULTADO_FINAL;#(lf)#(lf)CREATE TABLE #RESULTADO_FINAL (#(lf)#(tab)CODIGO VARCHAR(15),#(lf)#(tab)ARMAZEM VARCHAR(2),#(lf)#(tab)QUANT DECIMAL(18,5),#(lf)#(tab)CUSTO DECIMAL(18,5),#(lf)#(tab)DATA VARCHAR(10)#(lf));#(lf)#(lf)WHILE @DATA_FECHAMENTO <= @MAIOR_DATA_FECHAMENTO AND @DATA_FECHAMENTO < @DATA_FINAL#(lf)BEGIN#(lf)#(lf)#(tab)IF OBJECT_ID('tempdb..#QUANTIDADES') IS NOT NULL DROP TABLE #QUANTIDADES#(lf)#(tab)IF OBJECT_ID('tempdb..#PRECOS') IS NOT NULL DROP TABLE #PRECOS#(lf)#(tab)IF OBJECT_ID('tempdb..#D1D2DB') IS NOT NULL DROP TABLE #D1D2DB#(lf)#(tab)IF OBJECT_ID('tempdb..#QTDS_SOMADAS_CUSTOS_NULOS') IS NOT NULL DROP TABLE #QTDS_SOMADAS_CUSTOS_NULOS#(lf)#(tab)IF OBJECT_ID('tempdb..#PRODUTOS_DIAS') IS NOT NULL DROP TABLE #PRODUTOS_DIAS#(lf)#(tab)IF OBJECT_ID('tempdb..#SB9') IS NOT NULL DROP TABLE #SB9#(lf)#(tab)IF OBJECT_ID('tempdb..#D1D2DB') IS NOT NULL DROP TABLE #D1D2DB#(lf)#(tab)IF OBJECT_ID('tempdb..#POSEST_POR_DIA') IS NOT NULL DROP TABLE #POSEST_POR_DIA;#(lf)#(lf)#(tab)-----------------------------------------#(lf)#(tab)WITH CONSULTA_PRECOS AS (#(lf)#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)TRIM(B9_COD) CODIGO,#(lf)#(tab)#(tab)#(tab)B9_CM1 CUSTO,#(lf)#(tab)#(tab)#(tab)B9_DATA DATA#(lf)#(tab)#(tab)FROM VW_MN_SB9 B9#(lf)#(tab)#(tab)INNER JOIN VW_MN_SB1 B1#(lf)#(tab)#(tab)#(tab)ON B1.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND B1_COD = B9_COD#(lf)#(tab)#(tab)#(tab)AND B1_TIPO NOT IN ('MO','SV')#(lf)#(tab)#(tab)WHERE B9.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND B9_QINI <> 0#(lf)#(tab)#(tab)#(tab)AND B9_DATA = @DATA_FECHAMENTO#(lf)#(tab)#(tab)#(tab)AND (B9_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(lf)#(tab)#(tab)UNION ALL#(lf)#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)TRIM(D3_COD),#(lf)#(tab)#(tab)#(tab)ROUND(D3_CUSTO1 / D3_QUANT,5) CUSTO,#(lf)#(tab)#(tab)#(tab)D3_EMISSAO#(lf)#(tab)#(tab)FROM VW_MN_SD3 D3#(lf)#(tab)#(tab)INNER JOIN VW_MN_SB1 B1#(lf)#(tab)#(tab)#(tab)ON B1.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND B1_COD = D3_COD#(lf)#(tab)#(tab)#(tab)AND B1_TIPO NOT IN ('MO','SV')#(lf)#(tab)#(tab)WHERE D3.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D3_ESTORNO <> 'S'#(lf)#(tab)#(tab)#(tab)AND D3_QUANT <> 0#(lf)#(tab)#(tab)#(tab)AND (D3_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(tab)#(tab)#(tab)AND D3_EMISSAO BETWEEN @PRIMEIRO_DIA_MES AND @ULTIMO_DIA_MES#(lf)#(tab)#(tab)#(tab)AND LEFT(D3_CF,2) <> 'PR'#(lf)#(tab)#(tab)#(tab)AND D3.R_E_C_N_O_ = (#(lf)#(tab)#(tab)#(tab)#(tab)SELECT MAX(R_E_C_N_O_) FROM VW_MN_SD3 D3A#(lf)#(tab)#(tab)#(tab)#(tab)WHERE D3A.D_E_L_E_T_ <> '*' AND D3A.D3_QUANT <> 0#(lf)#(tab)#(tab)#(tab)#(tab)#(tab)AND D3A.D3_EMISSAO = D3.D3_EMISSAO#(lf)#(tab)#(tab)#(tab)#(tab)#(tab)AND D3A.D3_COD = D3.D3_COD#(lf)#(tab)#(tab)#(tab)#(tab)#(tab)AND D3A.D3_ESTORNO <> 'S'#(lf)#(tab)#(tab)#(tab)#(tab)#(tab)AND LEFT(D3A.D3_CF,2) <> 'PR'#(lf)#(tab)#(tab)#(tab))#(lf)#(tab)#(tab)#(lf)#(tab)#(tab)UNION ALL#(lf)#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)TRIM(B2_COD),#(lf)#(tab)#(tab)#(tab)B2_CM1,#(lf)#(tab)#(tab)#(tab)CONVERT(VARCHAR,GETDATE(),112)#(lf)#(tab)#(tab)FROM VW_MN_SB2 B2#(lf)#(tab)#(tab)INNER JOIN VW_MN_SB1 B1#(lf)#(tab)#(tab)#(tab)ON B1.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND B1_COD = B2_COD#(lf)#(tab)#(tab)#(tab)AND B1_TIPO NOT IN ('MO','SV')#(lf)#(tab)#(tab)WHERE B2.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND B2_QATU <> 0#(lf)#(tab)#(tab)#(tab)AND @DATA_FECHAMENTO = @MAIOR_DATA_FECHAMENTO#(lf)#(tab)#(tab)#(tab)AND (B2_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(tab)#(lf)#(tab)),#(lf)#(lf)#(tab) GROUP_PRECOS AS (#(lf)#(tab)#(tab)SELECT #(lf)#(tab)#(tab)#(tab)CODIGO, ROUND(MAX(CUSTO),5) CUSTO, DATA#(lf)#(tab)#(tab)FROM CONSULTA_PRECOS#(lf)#(tab)#(tab)GROUP BY CODIGO, DATA#(lf)#(tab))#(lf)#(lf)#(tab)SELECT * INTO #PRECOS FROM GROUP_PRECOS;#(lf)#(lf)#(tab)-----------------------------------------#(lf)#(tab)WITH DIAS AS (#(lf)#(tab)#(tab)SELECT CONVERT(DATE,@DATA_FECHAMENTO) AS DATA#(lf)#(tab)#(tab)UNION ALL#(lf)#(tab)#(tab)SELECT DATEADD(DAY, 1, DATA) FROM DIAS WHERE DATEADD(DAY, 1, DATA) <= CONVERT(DATE,@ULTIMO_DIA_MES)#(lf)#(tab)),#(lf)#(lf)#(tab)TODOS_OS_PRODUTOS AS (#(lf)#(tab)#(tab)SELECT TRIM(B9_COD) CODIGO,#(lf)#(lf)#(tab)#(tab)#(tab)B9_LOCAL ARMAZEM#(lf)#(tab)#(tab)FROM VW_MN_SB9 B9#(lf)#(tab)#(tab)WHERE B9.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND B9_DATA >= @DATA_FECHAMENTO#(lf)#(tab)#(tab)#(tab)AND B9_QINI <> 0#(lf)#(tab)#(tab)#(tab)AND (B9_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(tab)#(lf)#(tab)#(tab)UNION#(lf)#(tab)#(lf)#(tab)#(tab)SELECT DISTINCT TRIM(D1_COD), D1_LOCAL#(lf)#(tab)#(tab)FROM VW_MN_SD1 D1#(lf)#(tab)#(tab)INNER JOIN VW_MN_SF4 F4#(lf)#(tab)#(tab)#(tab)ON F4.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D1.D1_TES = F4.F4_CODIGO#(lf)#(tab)#(tab)#(tab)AND F4.F4_ESTOQUE = 'S'#(lf)#(tab)#(tab)WHERE D1.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D1_DTDIGIT > @DATA_FECHAMENTO#(lf)#(tab)#(tab)#(tab)AND (D1_LOCAL IN(SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(lf)#(tab)#(tab)UNION#(lf)#(lf)#(tab)#(tab)SELECT DISTINCT TRIM(D2_COD), D2_LOCAL#(lf)#(tab)#(tab)FROM VW_MN_SD2 D2#(lf)#(tab)#(tab)INNER JOIN VW_MN_SF4 F4#(lf)#(tab)#(tab)#(tab)ON F4.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D2.D2_TES = F4.F4_CODIGO#(lf)#(tab)#(tab)#(tab)AND F4.F4_ESTOQUE = 'S'#(lf)#(tab)#(tab)WHERE D2.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D2_EMISSAO > @DATA_FECHAMENTO#(lf)#(tab)#(tab)#(tab)AND (D2_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(lf)#(tab)#(tab)UNION#(lf)#(lf)#(tab)#(tab)SELECT DISTINCT TRIM(DB_PRODUTO), DB_LOCAL#(lf)#(tab)#(tab)FROM VW_MN_SDB DB#(lf)#(tab)#(tab)WHERE DB.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND DB_ESTORNO <> 'S'#(lf)#(tab)#(tab)#(tab)--AND DB_QUANT <> 0#(lf)#(tab)#(tab)#(tab)AND DB_DATA > @DATA_FECHAMENTO#(lf)#(tab)#(tab)#(tab)AND DB_ORIGEM IN ('SD3')#(lf)#(tab)#(tab)#(tab)AND (DB_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(lf)#(tab)),#(lf)#(lf)#(tab)PRODUTOS_DIAS AS (#(lf)#(tab)#(tab)SELECT * FROM DIAS#(lf)#(tab)#(tab)CROSS JOIN TODOS_OS_PRODUTOS#(lf)#(tab))#(lf)#(lf)#(tab)SELECT * INTO #PRODUTOS_DIAS FROM PRODUTOS_DIAS;#(lf)#(lf)#(tab)-----------------------------------------#(lf)#(tab)WITH PRE_SB9 AS (#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)B9_LOCAL ARMAZEM,#(lf)#(tab)#(tab)#(tab)TRIM(B9_COD) CODIGO,#(lf)#(tab)#(tab)#(tab)B9_QINI QUANT,#(lf)#(tab)#(tab)#(tab)B9_CM1 CUSTO,#(lf)#(tab)#(tab)#(tab)B9_DATA DATA#(lf)#(tab)#(tab)FROM VW_MN_SB9 B9#(lf)#(tab)#(tab)WHERE B9.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND B9_QINI <> 0#(lf)#(tab)#(tab)#(tab)AND B9_DATA = @DATA_FECHAMENTO#(lf)#(tab)#(tab)#(tab)AND (B9_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(tab)),#(lf)#(lf)#(tab)SB9 AS (#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)PS.ARMAZEM,#(lf)#(tab)#(tab)#(tab)PS.CODIGO,#(lf)#(tab)#(tab)#(tab)ISNULL(B9.QUANT,0) QUANT,#(lf)#(tab)#(tab)#(tab)CUSTO,#(lf)#(tab)#(tab)#(tab)PS.DATA#(lf)#(lf)#(tab)#(tab)FROM #PRODUTOS_DIAS PS#(lf)#(tab)#(tab)LEFT JOIN PRE_SB9 B9#(lf)#(tab)#(tab)#(tab)ON PS.CODIGO = B9.CODIGO#(lf)#(tab)#(tab)#(tab)AND PS.ARMAZEM = B9.ARMAZEM#(lf)#(tab)#(tab)#(tab)AND PS.DATA = B9.DATA#(lf)#(tab)#(tab)#(tab)AND B9.QUANT <> 0#(lf)#(tab))#(lf)#(lf)#(tab)SELECT * INTO #SB9 FROM SB9;#(lf)#(lf)#(tab)-----------------------------------------#(lf)#(tab)WITH D1D2DB AS (#(lf)#(tab)#(tab)--SD1#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)D1_LOCAL ARMAZEM,#(lf)#(tab)#(tab)#(tab)TRIM(D1_COD) CODIGO,#(lf)#(tab)#(tab)#(tab)SUM(D1_QUANT) QUANT,#(lf)#(tab)#(tab)#(tab)D1_DTDIGIT DATA#(lf)#(tab)#(tab)FROM VW_MN_SD1 D1#(lf)#(tab)#(tab)INNER JOIN VW_MN_SF4 F4#(lf)#(tab)#(tab)#(tab)ON F4.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D1.D1_TES = F4.F4_CODIGO#(lf)#(tab)#(tab)#(tab)AND F4.F4_ESTOQUE = 'S'#(lf)#(tab)#(tab)WHERE#(lf)#(tab)#(tab)#(tab)D1.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D1_DTDIGIT BETWEEN @PRIMEIRO_DIA_MES AND @ULTIMO_DIA_MES#(lf)#(tab)#(tab)#(tab)AND (D1_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(tab)#(tab)GROUP BY D1_LOCAL, D1_COD, D1_DTDIGIT#(lf)#(lf)#(tab)#(tab)UNION ALL#(lf)#(lf)#(tab)#(tab)--SD2#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)D2_LOCAL,#(lf)#(tab)#(tab)#(tab)TRIM(D2_COD) CODIGO,#(lf)#(tab)#(tab)#(tab)-SUM(D2_QUANT),#(lf)#(tab)#(tab)#(tab)D2_EMISSAO#(lf)#(tab)#(tab)FROM VW_MN_SD2 D2#(lf)#(tab)#(tab)INNER JOIN VW_MN_SF4 F4#(lf)#(tab)#(tab)#(tab)ON F4.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D2.D2_TES = F4.F4_CODIGO#(lf)#(tab)#(tab)#(tab)AND F4.F4_ESTOQUE = 'S'#(lf)#(tab)#(tab)WHERE D2.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND D2_EMISSAO BETWEEN @PRIMEIRO_DIA_MES AND @ULTIMO_DIA_MES#(lf)#(tab)#(tab)#(tab)AND (D2_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(tab)#(tab)GROUP BY D2_LOCAL, D2_COD, D2_QUANT, D2_EMISSAO#(lf)#(lf)#(tab)#(tab)UNION ALL#(lf)#(lf)#(tab)#(tab)--SDB#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)DB_LOCAL,#(lf)#(tab)#(tab)#(tab)TRIM(DB_PRODUTO),#(lf)#(tab)#(tab)#(tab)SUM(CASE#(lf)#(tab)#(tab)#(tab)#(tab)WHEN DB_TM > 500 THEN -DB_QUANT#(lf)#(tab)#(tab)#(tab)#(tab)WHEN DB_TM < 500 THEN DB_QUANT#(lf)#(tab)#(tab)#(tab)#(tab)ELSE 0#(lf)#(tab)#(tab)#(tab)END),#(lf)#(tab)#(tab)#(tab)DB_DATA#(lf)#(tab)#(tab)#(tab)#(lf)#(tab)#(tab)FROM VW_MN_SDB DB#(lf)#(tab)#(tab)WHERE DB.D_E_L_E_T_ <> '*'#(lf)#(tab)#(tab)#(tab)AND DB_ESTORNO <> 'S'#(lf)#(tab)#(tab)#(tab)AND DB_DATA BETWEEN @PRIMEIRO_DIA_MES AND @ULTIMO_DIA_MES#(lf)#(tab)#(tab)#(tab)AND DB_ORIGEM IN ('SD3')#(lf)#(tab)#(tab)#(tab)AND (DB_LOCAL IN (SELECT ARMAZEM FROM @ARMAZENS)#(lf)#(tab)#(tab)#(tab)#(tab)OR (SELECT TOP 1 ARMAZEM FROM @ARMAZENS) = 'TD' )#(lf)#(tab)#(tab)GROUP BY DB_LOCAL, DB_PRODUTO, DB_DATA#(lf)#(tab))#(lf)#(lf)#(tab)SELECT * INTO #D1D2DB FROM D1D2DB;#(lf)#(lf)#(tab)-----------------------------------------#(lf)#(tab)WITH D1D2DB_PRECOS AS (#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)Q.CODIGO, ARMAZEM,#(lf)#(tab)#(tab)#(tab)QUANT, P.CUSTO, Q.DATA#(lf)#(tab)#(tab)FROM #D1D2DB Q#(lf)#(tab)#(tab)LEFT JOIN #PRECOS P#(lf)#(tab)#(tab)#(tab)ON Q.CODIGO = P.CODIGO#(lf)#(tab)#(tab)#(tab)AND P.DATA = (#(lf)#(tab)#(tab)#(tab)#(tab)SELECT MIN(DATA) FROM #PRECOS P2#(lf)#(tab)#(tab)#(tab)#(tab)WHERE P2.DATA >= Q.DATA#(lf)#(tab)#(tab)#(tab)#(tab)#(tab)AND P2.CODIGO = P.CODIGO#(lf)#(tab)#(tab)#(tab))#(lf)#(tab)),#(lf)#(lf)#(tab)D1D2DB_UNION_B9 AS (#(lf)#(tab)#(lf)#(tab)#(tab)SELECT #(lf)#(tab)#(tab)#(tab)CODIGO, ARMAZEM, QUANT, CUSTO, DATA#(lf)#(tab)#(tab)FROM #SB9#(lf)#(lf)#(tab)#(tab)UNION ALL#(lf)#(lf)#(tab)#(tab)SELECT #(lf)#(tab)#(tab)#(tab)CODIGO,#(lf)#(tab)#(tab)#(tab)ARMAZEM,#(lf)#(tab)#(tab)#(tab)QUANT, CUSTO,#(lf)#(tab)#(tab)#(tab)DATA #(lf)#(tab)#(tab)FROM D1D2DB_PRECOS#(lf)#(tab)),#(lf)#(lf)#(tab)GROUP_POR_DIA AS (#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)ARMAZEM, CODIGO, SUM(QUANT) QUANT, MAX(CUSTO) CUSTO, DATA#(lf)#(tab)#(tab)FROM D1D2DB_UNION_B9#(lf)#(tab)#(tab)GROUP BY ARMAZEM, CODIGO, DATA#(lf)#(tab)),#(lf)#(lf)#(tab)QTD_SOMADAS_COM_PRECOS_NULOS AS (#(lf)#(tab)#(tab)SELECT #(lf)#(tab)#(tab)#(tab)ARMAZEM,#(lf)#(tab)#(tab)#(tab)CODIGO,#(lf)#(tab)#(tab)#(tab)SUM(QUANT) OVER (#(lf)#(tab)#(tab)#(tab)#(tab)PARTITION BY CODIGO, ARMAZEM#(lf)#(tab)#(tab)#(tab)#(tab)ORDER BY ARMAZEM, CODIGO, DATA#(lf)#(tab)#(tab)#(tab)#(tab)ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW#(lf)#(tab)#(tab)#(tab)) QUANT,#(lf)#(tab)#(tab)#(tab)CUSTO,#(lf)#(tab)#(tab)#(tab)CONVERT(VARCHAR,DATA,112) DATA#(lf)#(tab)#(tab)FROM GROUP_POR_DIA#(lf)#(tab))#(lf)#(lf)#(tab)SELECT * INTO #QTDS_SOMADAS_CUSTOS_NULOS FROM QTD_SOMADAS_COM_PRECOS_NULOS;#(lf)#(lf)#(tab)-----------------------------------------#(lf)#(tab)WITH PREENCHE_CUSTOS_NULOS AS (#(lf)#(tab)#(tab)SELECT#(lf)#(tab)#(tab)#(tab)A.CODIGO, A.ARMAZEM,#(lf)#(tab)#(tab)#(tab)A.QUANT, #(lf)#(tab)#(tab)#(tab)ISNULL(B.CUSTO,C.CUSTO) CUSTO, #(lf)#(tab)#(tab)#(tab)A.DATA#(lf)#(tab)#(tab)FROM #QTDS_SOMADAS_CUSTOS_NULOS A#(lf)#(tab)#(tab)LEFT JOIN #PRECOS B#(lf)#(tab)#(tab)#(tab)ON A.CODIGO = B.CODIGO#(lf)#(tab)#(tab)#(tab)AND B.DATA = (#(lf)#(tab)#(tab)#(tab)#(tab)SELECT MAX(DATA)#(lf)#(tab)#(tab)#(tab)#(tab)FROM #PRECOS B2#(lf)#(tab)#(tab)#(tab)#(tab)WHERE B.CODIGO = B2.CODIGO#(lf)#(tab)#(tab)#(tab)#(tab)AND B2.DATA <= A.DATA#(lf)#(tab)#(tab)#(tab)#(tab)AND B2.CUSTO IS NOT NULL#(lf)#(tab)#(tab)#(tab))#(lf)#(tab)#(tab)LEFT JOIN #PRECOS C#(lf)#(tab)#(tab)#(tab)ON A.CODIGO = C.CODIGO#(lf)#(tab)#(tab)#(tab)AND C.DATA = (#(lf)#(tab)#(tab)#(tab)#(tab)SELECT MAX(DATA)#(lf)#(tab)#(tab)#(tab)#(tab)FROM #PRECOS B2#(lf)#(tab)#(tab)#(tab)#(tab)WHERE C.CODIGO = B2.CODIGO#(lf)#(tab)#(tab)#(tab)#(tab)AND B2.DATA > A.DATA#(lf)#(tab)#(tab)#(tab)#(tab)AND B2.CUSTO IS NOT NULL#(lf)#(tab)#(tab)#(tab))#(lf)#(tab))#(lf)#(lf)#(tab)SELECT * INTO #POSEST_POR_DIA FROM PREENCHE_CUSTOS_NULOS;#(lf)#(lf)#(tab)-----------------------------------------#(lf)#(tab)INSERT INTO #RESULTADO_FINAL#(lf)#(lf)#(tab)SELECT * FROM #POSEST_POR_DIA#(lf)#(tab)WHERE (DATA < @ULTIMO_DIA_MES OR DATA = @DATA_FINAL)#(lf)#(tab)#(tab)AND DATEPART(WEEKDAY,DATA) = 7#(lf)#(lf)#(tab)SET @DATA_FECHAMENTO = CONVERT(VARCHAR,EOMONTH(DATEADD(MONTH,1,@DATA_FECHAMENTO)),112);#(lf)#(tab)SET @PRIMEIRO_DIA_MES = CONVERT(VARCHAR,DATEADD(MONTH,1,@PRIMEIRO_DIA_MES),112);#(lf)#(tab)SET @ULTIMO_DIA_MES = #(lf)#(tab)#(tab)CASE #(lf)#(tab)#(tab)#(tab)WHEN @DATA_FECHAMENTO = @MAIOR_DATA_FECHAMENTO THEN#(lf)#(tab)#(tab)#(tab)#(tab)CONVERT(VARCHAR,GETDATE(),112)#(lf)#(tab)#(tab)#(tab)ELSE #(lf)#(tab)#(tab)#(tab)#(tab)CONVERT(VARCHAR,EOMONTH(@PRIMEIRO_DIA_MES),112)#(lf)#(tab)#(tab)END;#(lf)#(lf)END --end para o while do começo#(lf)#(lf)SELECT * FROM #RESULTADO_FINAL#(lf)WHERE QUANT <> 0"]),
    combina_fontes = Table.Combine({Fonte,#"17-CEst-base-ano-ant"}),
    remove_qtd_zero = Table.SelectRows(combina_fontes,each [QUANT] <> 0),
    fontes_buffered = Table.Buffer(remove_qtd_zero),
    traz_info_SB1 = Table.ExpandTableColumn(
        Table.NestedJoin(fontes_buffered,"CODIGO",#"97-SB1","codigo","dados",JoinKind.LeftOuter)
        ,"dados",{"tipo"}
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(traz_info_SB1,{{"DATA", type date}}),
    estoque_va_e_nva = Table.AddColumn(#"Tipo Alterado","categoria",each (
        if List.Contains({"11","14","17","20","89"},[ARMAZEM]) then "Estoque VA" else "Estoque NVA"
    ) , type text),

    ano_num = Table.AddColumn(estoque_va_e_nva,"ano_num",each Date.Year([DATA]), Int16.Type),
    ano_texto_temp = Table.AddColumn(ano_num,"ano_texto",each Text.From([ano_num]), type text),

    semana_num = Table.AddColumn(ano_texto_temp,"semana_num", each (
        [ano_num] * 100 + Date.WeekOfYear([DATA])
    ) , Int32.Type),
    
    mes_num = Table.AddColumn(semana_num,"mes_num",
        each [ano_num] * 100 + #"coleta_mes_pela_semana"(Number.FromText(Text.Middle(Text.From([semana_num]),4,2))), 
    Int32.Type),

    semana_texto = Table.AddColumn(mes_num,"semana_texto", each (
        "WK" & Text.Middle(Text.From([semana_num]),4,2) & "/" & Text.Middle([ano_texto],2,2)
    ), type text),

    maior_semana = List.Max(Table.Column(semana_texto,"semana_num")),

    mes_texto = Table.AddColumn(semana_texto,"mes_texto",each (
        if #"valida_relacao_semana_mes_544"(maior_semana,[mes_num],[semana_num]) = 0 then "" else
        Date.MonthName(#date(#"00-ano-atual",Number.FromText(Text.Middle(Text.From([mes_num]),4,2)),1)) 
        & "/" & Text.Middle([ano_texto],2,2)
    ), type text),

    altera_mes_num = Table.RenameColumns(
        Table.RemoveColumns(
            Table.AddColumn(mes_texto,"mes_num_a",each if [mes_texto] = "" then null else [mes_num], Int32.Type),
            "mes_num"
        ),
        {{"mes_num_a","mes_num"}}
    ),
    remove_ano_texto = Table.RemoveColumns(altera_mes_num,{"ano_texto"}),
    buffer = Table.Buffer(remove_ano_texto),

    group_por_ano = Table.Group(buffer,"ano_num",{"group",each _, type table}),
    add_ano_texto = Table.TransformColumns(group_por_ano,{"group", each (
        #"add_ano_texto_na_ultima_semana"(_)
    ), type table}),
    expand_ano_texto = Table.ExpandTableColumn(add_ano_texto,"group",{
        "CODIGO","ARMAZEM","QUANT","CUSTO","DATA","tipo","categoria","ano_texto","semana_num","semana_texto","mes_texto","mes_num"
    }),

    total = Table.AddColumn(expand_ano_texto,"total", each [QUANT] * [CUSTO], type number),

    semanas_distintas = Table.Sort(
        Table.Distinct(Table.SelectColumns(total,"semana_num"))
        ,{"semana_num",Order.Descending}
    ),
    index_semana = Table.AddIndexColumn(semanas_distintas,"filtro_semana",1,1,Int16.Type),
    filtro_semana = Table.TransformColumns(index_semana,{
        {"filtro_semana", each if _ <= 6 then "Últimas 6 Semanas" else "Semanas Anteriores", type text}
    }),
    traz_filtro_semana = Table.ExpandTableColumn(
        Table.NestedJoin(total,"semana_num",filtro_semana,"semana_num","dados",JoinKind.LeftOuter)
        , "dados",{"filtro_semana"}
    ),

    meses_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(traz_filtro_semana,"mes_num"))
        ,{"mes_num",Order.Descending}
    ),
    index_mes = Table.AddIndexColumn(meses_distintos,"filtro_mes",1,1,Int16.Type),
    filtro_mes = Table.TransformColumns(index_mes,{
        {"filtro_mes", each if _ <= 3 then "Últimos 3 Meses" else "Meses Anteriores", type text}
    }),
    traz_filtro_mes = Table.ExpandTableColumn(
        Table.NestedJoin(traz_filtro_semana,"mes_num",filtro_mes,"mes_num","dados",JoinKind.LeftOuter)
        , "dados",{"filtro_mes"}
    ),
    
    arruma_ano_num = Table.RenameColumns(
        (Table.RemoveColumns((
            Table.AddColumn(traz_filtro_mes,"ano_num_a", each if [ano_texto] = null then null else [ano_num], Int16.Type)
        ),{"ano_num"})),
        {"ano_num_a","ano_num"}
    ),

    armazens_distintos = Table.Sort(
        Table.Distinct(Table.SelectColumns(arruma_ano_num,"ARMAZEM"))
        ,{"ARMAZEM",Order.Ascending}
    ),
    index_armazem = Table.AddIndexColumn(armazens_distintos,"index_armazem",1,1,Int16.Type),
    cor_armazem = Table.ExpandTableColumn(
        Table.NestedJoin(index_armazem,"index_armazem",#"99-cores","index","dados",JoinKind.LeftOuter)
        ,"dados",{"Cor"}
    ),

    traz_cor = Table.ExpandTableColumn(
        Table.NestedJoin(arruma_ano_num,"ARMAZEM",cor_armazem,"ARMAZEM","dados",JoinKind.LeftOuter)
        ,"dados",{"Cor"}
    )

in
    traz_cor