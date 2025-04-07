(maior_semana, mes, semana) as number => (
		let
			mes_em_texto = Text.Middle(Text.From(mes),4,2),
			semana_em_texto = Text.Middle(Text.From(semana),4,2),
			resultado = if semana = maior_semana then 1 else
				if mes_em_texto = "01" and semana_em_texto = "05" then 1 else
				if mes_em_texto = "02" and semana_em_texto = "09" then 1 else
				if mes_em_texto = "03" and semana_em_texto = "13" then 1 else
				if mes_em_texto = "04" and semana_em_texto = "18" then 1 else
				if mes_em_texto = "05" and semana_em_texto = "22" then 1 else
				if mes_em_texto = "06" and semana_em_texto = "26" then 1 else
				if mes_em_texto = "07" and semana_em_texto = "31" then 1 else
				if mes_em_texto = "08" and semana_em_texto = "35" then 1 else
				if mes_em_texto = "09" and semana_em_texto = "39" then 1 else
				if mes_em_texto = "10" and semana_em_texto = "44" then 1 else
				if mes_em_texto = "11" and semana_em_texto = "48" then 1 else
				if mes_em_texto = "12" and semana_em_texto = "52" then 1 else
				0
		in resultado
	)