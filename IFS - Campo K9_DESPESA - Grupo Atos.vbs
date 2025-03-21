Public Sub MAIN

    Dim QRY_VENDAS As Object
    Set QRY_VENDAS = NewQuery

    QRY_VENDAS.Add("SELECT PNR.Handle AS HANDLE_PNR, ACC.Handle AS HANDLE_ACC, PNR.LOCALIZADORA, PNR.DATAINCLUSAO, ACC.REQUISICAO, ACC.K9_DESPESA AS DESPESA_BENNER, ACC.K9_BREAK1, " + _
                   "CASE WHEN CHARINDEX('<Valor>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > 0 " + _
                   "AND CHARINDEX('</Valor>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > 0 " + _
                   "THEN SUBSTRING(CONVERT(VARCHAR(MAX), LOG.XMLRESERVA), CHARINDEX('<Valor>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<Valor>'), " + _
                   "CHARINDEX('</Valor>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) - CHARINDEX('<Valor>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) - LEN('<Valor>')) " + _
                   "ELSE NULL END AS DESPESA_XML " + _
                   "FROM VM_PNRS PNR " + _
                   "LEFT JOIN VM_PNRACCOUNTINGS ACC ON ACC.PNR = PNR.Handle " + _
                   "LEFT JOIN BB_LOGINTEGRACOES LOG ON PNR.LOGINTEGRACAO = LOG.Handle " + _
                   "WHERE PNR.CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (68539)) " + _
                   "AND CAST(PNR.DATAINCLUSAO AS DATE) = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE) " + _
                   "AND PNR.TIPORESERVA IN (27)")

    QRY_VENDAS.Active = True

    While Not QRY_VENDAS.EOF

        ' Atualizar K9_BREAK1 se houver valor válido em DESPESA_XML
        If Not IsNull(QRY_VENDAS.FieldByName("DESPESA_XML").AsString) Then
            Dim UPT_BREAK As Object
            Set UPT_BREAK = NewQuery
            UPT_BREAK.Add("UPDATE VM_PNRACCOUNTINGS SET K9_BREAK1 = '" & QRY_VENDAS.FieldByName("DESPESA_XML").AsString & "' WHERE HANDLE = " & CStr(QRY_VENDAS.FieldByName("HANDLE_ACC").AsInteger))
            UPT_BREAK.ExecSQL
        End If

        QRY_VENDAS.Next
    Wend

    Dim QRY_DESPESA As Object
    Set QRY_DESPESA = NewQuery

	QRY_DESPESA.Add("SELECT PNR.Handle AS HANDLE_PNR, ACC.Handle AS HANDLE_ACC, PNR.LOCALIZADORA, PNR.DATAINCLUSAO, ACC.REQUISICAO, ACC.K9_DESPESA AS DESPESA_BENNER, ACC.K9_BREAK1 " + _
	                "FROM VM_PNRS PNR " + _
	                "LEFT JOIN VM_PNRACCOUNTINGS ACC ON ACC.PNR = PNR.HANDLE " + _
	                "LEFT JOIN BB_LOGINTEGRACOES LOG ON PNR.LOGINTEGRACAO = LOG.HANDLE " + _
	                "WHERE PNR.CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (68539)) " + _
	                "AND CAST(PNR.DATAINCLUSAO AS DATE) = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE) " + _
	                "AND PNR.TIPORESERVA IN (27)")


    QRY_DESPESA.Active = True

    While Not QRY_DESPESA.EOF

        ' Atualizar K9_DESPESA com base no valor de K9_BREAK1
        If QRY_DESPESA.FieldByName("K9_BREAK1").AsString = "Billable" Then
            Dim UPT_DESPESA As Object
            Set UPT_DESPESA = NewQuery
            UPT_DESPESA.Add("UPDATE VM_PNRACCOUNTINGS SET K9_DESPESA = 92 WHERE HANDLE = " & CStr(QRY_DESPESA.FieldByName("HANDLE_ACC").AsInteger))
            UPT_DESPESA.ExecSQL
        ElseIf QRY_DESPESA.FieldByName("K9_BREAK1").AsString = "Non Billable" Then
            Dim UPT_DESPESA2 As Object
            Set UPT_DESPESA2 = NewQuery
            UPT_DESPESA2.Add("UPDATE VM_PNRACCOUNTINGS SET K9_DESPESA = 93 WHERE HANDLE = " & CStr(QRY_DESPESA.FieldByName("HANDLE_ACC").AsInteger))
            UPT_DESPESA2.ExecSQL
        End If

        Dim DLT_BREAK As Object
        Set DLT_BREAK = NewQuery
        DLT_BREAK.Add("UPDATE VM_PNRACCOUNTINGS SET K9_BREAK1 = NULL WHERE HANDLE = " & CStr(QRY_DESPESA.FieldByName("HANDLE_ACC").AsInteger))
        DLT_BREAK.ExecSQL

        ' Atualizar PNR
        Dim UPT_VENDA As Object
        Set UPT_VENDA = NewQuery
        UPT_VENDA.Add("UPDATE VM_PNRS SET SITUACAO = 1, CONCLUIDO = 'S', EXPORTADO = 'N', AGUARDANDOEMISSAO = 'N' WHERE HANDLE = " & CStr(QRY_DESPESA.FieldByName("HANDLE_PNR").AsInteger))
        UPT_VENDA.ExecSQL

        QRY_DESPESA.Next
    Wend

End Sub
