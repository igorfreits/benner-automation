Public Sub Main()
    ' Codifique aqui o método principal
    Dim qrySQL As Object

    Set qrySQL = NewQuery
    qrySQL.Active = False
    qrySQL.Clear

    qrySQL.Add("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, PNR.CLIENTE, " + _
               "ACC.empenho AS DEPT, LOCALIZADORA, PNR.DATAINCLUSAO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.Handle = ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (71847)) " + _
               "AND ACC.EMPENHO IS NULL")

    qrySQL.Active = True

    While Not qrySQL.EOF
        Dim empenhoValue As String

        ' Verifica o valor de CLIENTE e atribui o EMPENHO correspondente
        If qrySQL.FieldByName("CLIENTE").AsInteger = 71846 Then
            empenhoValue = "4075"
        ElseIf qrySQL.FieldByName("CLIENTE").AsInteger = 92300 Then
            empenhoValue = "4014"
        ElseIf qrySQL.FieldByName("CLIENTE").AsInteger = 71848 Then
            empenhoValue = "4012"
        ElseIf qrySQL.FieldByName("CLIENTE").AsInteger = 71849 Then
            empenhoValue = "4086"
        Else
            empenhoValue = "" ' Caso não encontre correspondência
        End If

        If empenhoValue <> "" Then
            Dim qryUPDT As Object
            Set qryUPDT = NewQuery

            On Error GoTo ER1

            If Not InTransaction Then StartTransaction

            qryUPDT.Active = False
            qryUPDT.Clear

            qryUPDT.Add("UPDATE VM_PNRACCOUNTINGS " + _
                        "SET EMPENHO = '" + empenhoValue + "' " + _
                        "WHERE HANDLE = " + CStr(qrySQL.FieldByName("HANDLE_ACC").AsInteger))

            qryUPDT.ExecSQL

            If InTransaction Then Commit
            Set qryUPDT = Nothing

ER1:
            If InTransaction Then Rollback
            Set qryUPDT = Nothing
        End If

        Dim qryUPDT_1 As Object
        Set qryUPDT_1 = NewQuery

        On Error GoTo ER2

        If Not InTransaction Then StartTransaction

        qryUPDT_1.Active = False
        qryUPDT_1.Clear

        qryUPDT_1.Add("UPDATE VM_PNRS " + _
                      "SET SITUACAO = 1, CONCLUIDO = 'S', EXPORTADO = 'S', AGUARDANDOEMISSAO = 'S' " + _
                      "WHERE HANDLE = " + CStr(qrySQL.FieldByName("HANDLE_PNR").AsInteger))

        qryUPDT_1.ExecSQL

        If InTransaction Then Commit
        Set qryUPDT_1 = Nothing

ER2:
        If InTransaction Then Rollback
        Set qryUPDT_1 = Nothing

        qrySQL.Next
    Wend

    Set qrySQL = Nothing
End Sub
