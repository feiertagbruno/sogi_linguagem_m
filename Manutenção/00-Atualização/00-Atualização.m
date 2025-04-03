let
    agora = DateTime.LocalNow(),
    horario = Text.From(DateTime.Time(agora)),
    agora_calculado = Text.Combine({
        Text.PadStart(Text.From(Date.Day(agora)),2,"0"),
        "/",
        Text.PadStart(Text.From(Date.Month(agora)),2,"0"),
        "/",
        Text.From(Date.Year(agora)),
        " ",
        Text.PadStart(Text.From(Number.FromText(
            Text.Split(horario,":"){0}
        )-4),2,"0"),
        ":",
        Text.PadStart(Text.From(Number.FromText(
            Text.Split(horario,":"){1}
            )
        ),2,"0")
    }),
    acrescentar_manaus = Text.Combine({"Atualizado em ", agora_calculado," Horário Manaus"})

in
    acrescentar_manaus