Public Sub Main()

    Dim QRY_BRASKEM As Object
    Dim QRY_PROJETO_CC As Object
    Dim UPT_PROJETO_CC As Object
    Dim QRY_CC_CCPAX As Object
    Dim UPT_CC_CCPAX As Object

    Set QRY_BRASKEM = NewQuery
    QRY_BRASKEM.Active = False
    QRY_BRASKEM.Clear

    QRY_BRASKEM.Add("SELECT DISTINCT PNR.Handle HANDLE_PNR, " + _
                        "ACC.Handle HANDLE_ACC, PNR.LOCALIZADORA, PNR.DATAINCLUSAO, ACC.CENTRODECUSTO, ACC.K9_CENTROCUSTOPAX, " + _
                        "CASE WHEN CHARINDEX('<CodCentroCusto>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > 0 " + _
                        "AND CHARINDEX('</CodCentroCusto>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > " + _
                        "CHARINDEX('<CodCentroCusto>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) THEN " + _
                        "SUBSTRING(CONVERT(VARCHAR(MAX), Log.XMLRESERVA), " + _
                        "CHARINDEX('<CodCentroCusto>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CodCentroCusto>'), " + _
                        "CHARINDEX('</CodCentroCusto>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) - " + _
                        "(CHARINDEX('<CodCentroCusto>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CodCentroCusto>'))) " + _
                        "ELSE NULL END AS CENTRO_CUSTO, " + _
                        "CASE WHEN CHARINDEX('<ProjetoCod>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > 0 " + _
                        "AND CHARINDEX('</ProjetoCod>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > " + _
                        "CHARINDEX('<ProjetoCod>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) THEN " + _
                        "SUBSTRING(CONVERT(VARCHAR(MAX), Log.XMLRESERVA), " + _
                        "CHARINDEX('<ProjetoCod>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<ProjetoCod>'), " + _
                        "CHARINDEX('</ProjetoCod>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) - " + _
                        "(CHARINDEX('<ProjetoCod>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<ProjetoCod>'))) " + _
                        "ELSE NULL END AS PROJETO_COD " + _
                            "FROM VM_PNRS PNR " + _
                            "LEFT JOIN VM_PNRACCOUNTINGS ACC ON (ACC.PNR = PNR.Handle) " + _
                            "LEFT JOIN BB_LOGINTEGRACOES Log ON Log.Handle = PNR.LOGINTEGRACAO " + _
                            "WHERE PNR.DATAINCLUSAO >= DATEADD(DAY, -1, CAST(GETDATE() AS DATETIME)) And PNR.DATAINCLUSAO < CAST(GETDATE() As DATETIME) " + _
                            "AND PNR.TIPORESERVA IN (11) AND PNR.CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL = 35934)")

    QRY_BRASKEM.Active = True

    While Not QRY_BRASKEM.EOF
        Set QRY_PROJETO_CC = NewQuery
        QRY_PROJETO_CC.Active = False
        QRY_PROJETO_CC.Clear

        QRY_PROJETO_CC.Add("SELECT CENTROCUSTO, DESCRICAO, HANDLE, CODIGOINTERNO, VALIDO " + _
            "FROM BB_CLIENTECC " + _
            "WHERE CENTROCUSTO = '" & QRY_BRASKEM.FieldByName("PROJETO_COD").AsString & "' " + _
            "AND VALIDO = 'S' " + _
            "AND BB_CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (35934))")

        QRY_PROJETO_CC.Active = True

        Set UPT_PROJETO_CC = NewQuery
        UPT_PROJETO_CC.Active = False
        UPT_PROJETO_CC.Clear

        UPT_PROJETO_CC.Add("UPDATE VM_PNRACCOUNTINGS SET CENTRODECUSTO = '" & QRY_PROJETO_CC.FieldByName("CODIGOINTERNO").AsString & "' WHERE HANDLE = " & QRY_BRASKEM.FieldByName("HANDLE_ACC").AsString)
        UPT_PROJETO_CC.ExecSQL

        Set QRY_CC_CCPAX = NewQuery
        QRY_CC_CCPAX.Active = False
        QRY_CC_CCPAX.Clear

        QRY_CC_CCPAX.Add("SELECT CENTROCUSTO, DESCRICAO, HANDLE, CODIGOINTERNO, VALIDO " + _
            "FROM BB_CLIENTECC " + _
            "WHERE CENTROCUSTO = '" & QRY_BRASKEM.FieldByName("CENTRO_CUSTO").AsString & "' " + _
            "AND VALIDO = 'S' " + _
            "AND BB_CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (35934))")

        QRY_CC_CCPAX.Active = True

        Set UPT_CC_CCPAX = NewQuery
        UPT_CC_CCPAX.Active = False
        UPT_CC_CCPAX.Clear

        UPT_CC_CCPAX.Add("UPDATE VM_PNRACCOUNTINGS SET K9_CENTROCUSTOPAX = '" & QRY_CC_CCPAX.FieldByName("CODIGOINTERNO").AsString & "' WHERE HANDLE = " & QRY_BRASKEM.FieldByName("HANDLE_ACC").AsString)
        UPT_CC_CCPAX.ExecSQL

        QRY_BRASKEM.Next

        Dim UPT_CONCLUIDO As Object
        Set UPT_CONCLUIDO = NewQuery
        UPT_CONCLUIDO.Active = False
        UPT_CONCLUIDO.Clear

        UPT_CONCLUIDO.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE =" & QRY_BRASKEM.FieldByName("HANDLE_PNR").AsInteger)
        UPT_CONCLUIDO.ExecSQL

        QRY_BRASKEM.Next

    Wend

End Sub
