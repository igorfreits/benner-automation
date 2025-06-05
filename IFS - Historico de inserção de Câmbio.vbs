Public Sub Main

Dim qSql As BPesquisa
Set qSql = NewQuery

qSql.Add("SELECT DATA, ROUND(VALOR, 3) AS VALOR, DATAINCLUSAO,                 ")
qSql.Add("CASE                                                                 ")
qSql.Add("WHEN DATAINCLUSAO IS NULL THEN 'ROBO'                                ")
qSql.Add("ELSE 'MANUAL'                                                        ")
qSql.Add("END AS INCLUSAO                                                      ")
qSql.Add("FROM GN_MOEDACOTACOES                                                ")
qSql.Add("WHERE MONTH(DATA) = MONTH(GETDATE())AND YEAR(DATA) = YEAR(GETDATE()) ")
qSql.Add("AND MOEDA = 12                                                       ")
qSql.Add("ORDER BY 1                                                           ")
qSql.Active = True

	'Envio de e-mail
	Dim email As Mail
	Set email = NewMail
	email.SendTo = ""
	email.Subject = "Historico de inserção de Câmbio- " + Format(Now,"DD/MM/YYYY")

	Dim Corpo_Email As Variant

	Corpo_Email = "DATA - VALOR - DATA INCLUSÃO - INCLUSÃO"
	Corpo_Email = Corpo_Email + vbNewLine + vbNewLine

	While Not qSql.EOF()

        Corpo_Email = Corpo_Email + _
            qSql.FieldByName("DATA").AsString + " - " + _
            qSql.FieldByName("VALOR").AsString + " - " + _
			qSql.FieldByName("DATAINCLUSAO").AsString + " - " + _
			qSql.FieldByName("INCLUSAO").AsString + _
			" - " + vbNewLine

        qSql.Next
	Wend


	Corpo_Email=Corpo_Email + "---------------x---------------x---------------x---------------x---------------" + vbNewLine
    Corpo_Email=Corpo_Email + "Historico de inserção de Câmbio - " + Format(Now,"DD/MM/YYYY") + vbNewLine
	Corpo_Email=Corpo_Email + vbNewLine + "Kontik Viagens"+ vbNewLine
	Corpo_Email=Corpo_Email + "Suporte Benner" + vbNewLine

	email.Text.Add Corpo_Email

	email.Send

	Set email = Nothing

	Corpo_Email = Empty

End Sub
