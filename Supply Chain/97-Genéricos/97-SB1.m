let
    Fonte = Sql.Database("172.16.10.9", "Protheus12", [Query="SELECT#(lf)#(tab)TRIM(B1_COD) codigo,#(lf)#(tab)TRIM(B1_DESC) descricao,#(lf)#(tab)TRIM(B1_TIPO) tipo#(lf)FROM VW_MN_SB1 B1#(lf)WHERE B1.D_E_L_E_T_ <> '*'#(lf)#(tab)AND B1_TIPO IN ('PA','PI','MP','EM','BN','MC')"]),
    
	concat_desci_cod = Table.AddColumn(Fonte,"descricao_cod", each (
		[descricao] & " " & [codigo]
	), type text)

in
    concat_desci_cod