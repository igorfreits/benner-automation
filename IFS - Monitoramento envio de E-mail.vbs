Public Sub Main

Dim qSql As BPesquisa
Set qSql = NewQuery

qSql.Add("SELECT                                                                                ")
qSql.Add("    CASE                                                                              ")
qSql.Add("        WHEN STATUS = 1 THEN 'Cadastrado'                                             ")
qSql.Add("        WHEN STATUS = 2 THEN 'Aguardando envio'                                       ")
qSql.Add("        WHEN STATUS = 3 THEN 'Enviando...'                                            ")
qSql.Add("        WHEN STATUS = 4 THEN 'Enviado'                                                ")
qSql.Add("        WHEN STATUS = 5 THEN 'Envio cancelado'                                        ")
qSql.Add("        WHEN STATUS = 6 THEN 'Erro de envio'                                          ")
qSql.Add("    END AS NOME_STATUS,                                                               ")
qSql.Add("    CASE                                                                              ")
qSql.Add("        WHEN DATAINCLUSAO >= DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()), 0) THEN 'Hoje' ")
qSql.Add("        ELSE 'Ontem'                                                                  ")
qSql.Add("    END AS DIA,                                                                       ")
qSql.Add("    COUNT(*) AS QUANTIDADE                                                            ")
qSql.Add("FROM Z_EMAILS                                                                         ")
qSql.Add("WHERE                                                                                 ")
qSql.Add("    DATAINCLUSAO >= DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()) - 1, 0)                  ")
qSql.Add("    AND DATAINCLUSAO < DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()) + 1, 0)               ")
qSql.Add("GROUP BY                                                                              ")
qSql.Add("    STATUS,                                                                           ")
qSql.Add("    CASE                                                                              ")
qSql.Add("        WHEN DATAINCLUSAO >= DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()), 0) THEN 'Hoje' ")
qSql.Add("        ELSE 'Ontem'                                                                  ")
qSql.Add("    END                                                                               ")
qSql.Active = True

	'Envio de e-mail
	Dim email As Mail
	Set email = NewMail
	email.SendTo = ""
	email.Subject = "Monitoramento envio de E-mail - " + Format(Now,"DD/MM/YYYY")

	Dim Corpo_Email As Variant

	Corpo_Email = "SITUAÇÃO - DIA - QUANTIDADE" + vbNewLine
	Corpo_Email = Corpo_Email + vbNewLine + vbNewLine

	While Not qSql.EOF()

        Corpo_Email = Corpo_Email + _
            qSql.FieldByName("NOME_STATUS").AsString + " - " + _
            qSql.FieldByName("DIA").AsString + " - " + _
			qSql.FieldByName("QUANTIDADE").AsString + vbNewLine

        qSql.Next
	Wend


	Corpo_Email=Corpo_Email + "---------------x---------------x---------------x---------------x---------------" + vbNewLine
    Corpo_Email=Corpo_Email + "Monitoramento envio de E-mail - " + Format(Now,"DD/MM/YYYY") + vbNewLine
	Corpo_Email=Corpo_Email + vbNewLine + "Kontik Viagens"+ vbNewLine
	Corpo_Email=Corpo_Email + "Suporte Benner" + vbNewLine

	email.Text.Add Corpo_Email

	email.Send

	Set email = Nothing

	Corpo_Email = Empty

End Sub
