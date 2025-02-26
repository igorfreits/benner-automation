Public Sub MAIN()

    Dim QRY_RLOC_GOVER As Object
    Set QRY_RLOC_GOVER = NewQuery
    QRY_RLOC_GOVER.Active = False
    QRY_RLOC_GOVER.Clear

    QRY_RLOC_GOVER.Add("SELECT PNR.HANDLE HANDLE_PNR, ACC.HANDLE HANDLE_ACC, PNR.DATAINCLUSAO, PNR.LOCALIZADORA, " & _
                       "ACC.REQUISICAO, PNR.TIPORESERVA, ACC.TIPOACC " & _
                       "FROM VM_PNRS PNR " & _
                       "LEFT JOIN VM_PNRACCOUNTINGS ACC ON ACC.PNR = PNR.Handle " & _
                       "WHERE PNR.TIPORESERVA IN (27) AND PNR.MENSAGEM LIKE '%A duplicidade de rloc é permitida%' " & _
                       "AND ACC.TIPOACC IN (2,3) AND PNR.SITUACAO = 3")

    QRY_RLOC_GOVER.Active = True

    ' Criando a query de update condicionando o tipo de produto para recber o H pra Hotel ou C para Carro
    Dim UPT_RLOC_GOVER As Object
    Set UPT_RLOC_GOVER = NewQuery

    While Not QRY_RLOC_GOVER.EOF()

        Dim novoValor As String
        If QRY_RLOC_GOVER.FieldByName("TIPOACC").AsInteger = 3 Then
            novoValor = "H" & QRY_RLOC_GOVER.FieldByName("REQUISICAO").AsString
        ElseIf QRY_RLOC_GOVER.FieldByName("TIPOACC").AsInteger = 2 Then
            novoValor = "C" & QRY_RLOC_GOVER.FieldByName("REQUISICAO").AsString
        Else
            novoValor = QRY_RLOC_GOVER.FieldByName("REQUISICAO").AsString
        End If

        ' Atualizando a LOCALIZADORA com o novo valor
        UPT_RLOC_GOVER.Active = False
        UPT_RLOC_GOVER.Clear
        UPT_RLOC_GOVER.Add("UPDATE VM_PNRS SET LOCALIZADORA = '" & novoValor & "' WHERE HANDLE = " & QRY_RLOC_GOVER.FieldByName("HANDLE_PNR").AsInteger)
        UPT_RLOC_GOVER.ExecSQL

        ' Atualizando a situação do PNR
        UPT_RLOC_GOVER.Active = False
        UPT_RLOC_GOVER.Clear
        UPT_RLOC_GOVER.Add("UPDATE VM_PNRS SET SITUACAO = 1, CONCLUIDO = 'S', EXPORTADO = 'N', AGUARDANDOEMISSAO = 'N' WHERE HANDLE = " & QRY_RLOC_GOVER.FieldByName("HANDLE_PNR").AsInteger)
        UPT_RLOC_GOVER.ExecSQL

        QRY_RLOC_GOVER.Next

    Wend

    ' Limpando memória
    Set QRY_RLOC_GOVER = Nothing
    Set UPT_RLOC_GOVER = Nothing

End Sub
