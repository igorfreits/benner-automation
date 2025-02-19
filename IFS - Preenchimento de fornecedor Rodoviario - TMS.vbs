Public Sub Main()
    ' Declaração das variáveis
    Dim qry_email As Object
    Dim qryUPDT As Object
    Dim Email As Mail
    Dim Corpo_Email As String
    Dim Registros_Processados As String
    Dim qrySQL As Object
    Dim qryUPDT_1 As Object
    Dim qryUPDT_2 As Object
    Dim qryUPDT_3 As Object


    Set qrySQL_2 = NewQuery
    qrySQL_2.Active = False
    qrySQL_2.Clear

    qrySQL_2.Add("SELECT DISTINCT PNR.HANDLE HANDLE_PNR, " & _
                  "ACC.HANDLE HANDLE_ACC, " & _
                  "PNR.LOCALIZADORA, PNR.DATAINCLUSAO, " & _
                  "CASE WHEN CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > 0 " & _
                  "AND CHARINDEX('</CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > " & _
                  "CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) THEN " & _
                  "TRANSLATE(SUBSTRING(CONVERT(VARCHAR(MAX), LOG.XMLRESERVA), " & _
                  "CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CiaOperadora>'), " & _
                  "CHARINDEX('</CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) - " & _
                  "(CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CiaOperadora>'))), " & _
                  "'áéíóúâêîôûãõäëïöüç', 'aeiouaeiouaoaeiouc') " & _
                  "Else Null End As EMPRESA_VIACAO " & _
                  "FROM VM_PNRS PNR " & _
                  "LEFT JOIN VM_PNRACCOUNTINGS ACC ON (ACC.PNR = PNR.HANDLE) " & _
                  "Left Join BB_LOGINTEGRACOES LOG On LOG.Handle = PNR.LOGINTEGRACAO " & _
                  "WHERE YEAR(PNR.DATAINCLUSAO) = YEAR(GETDATE()) " & _
                  "AND PNR.TIPORESERVA IN (11) AND ACC.TIPOMISCELANIO = 23 " & _
                  "AND ACC.TIPOACC IN (4) " & _
                  "AND (ACC.FORNECEDOR IS NULL OR ACC.FORNECEDOR IN (0))")

    qrySQL_2.Active = True

    Do While Not qrySQL_2.EOF
        ' Prossegue com a lógica quando o EMPRESA_VIACAO não está vazio
        Set qrySQL_3 = NewQuery
        qrySQL_3.Active = False
        qrySQL_3.Clear

        qrySQL_3.Add("SELECT CODIGOGDS, CONTRATO "+ _
                        "FROM BB_FORNECEDORCONTRATOCODIGOS  "+ _
                        "WHERE TIPORESERVA IN (11) AND SISTEMARESERVA IN (69) AND CODIGOGDS IN ('"+ qrySQL_2.FieldByName("EMPRESA_VIACAO").AsString +"')")

        qrySQL_3.Active = True

        Set qryUPDT_2 = NewQuery
        qryUPDT_2.Active = False
        qryUPDT_2.Clear

        qryUPDT_2.Add("UPDATE VM_PNRACCOUNTINGS SET FORNECEDOR = '" + qrySQL_3.FieldByName("CONTRATO").AsString + "' WHERE HANDLE = " + qrySQL_2.FieldByName("HANDLE_ACC").AsString)

        qryUPDT_2.ExecSQL

        Set qryUPDT_3 = NewQuery
        qryUPDT_3.Active = False
        qryUPDT_3.Clear

        qryUPDT_3.Add("UPDATE VM_PNRS SET SITUACAO = 1, CONCLUIDO = 'S', EXPORTADO = 'N', AGUARDANDOEMISSAO = 'N' WHERE HANDLE = " + qrySQL_2.FieldByName("HANDLE_PNR").AsString)

        qryUPDT_3.ExecSQL

        qrySQL_2.Next
    Loop

    ' Inicializa a consulta para buscar vendas sem fornecedor
    Set qry_email = NewQuery
    qry_email.Active = False
    qry_email.Clear

    qry_email.Add("SELECT DISTINCT PNR.HANDLE HANDLE_PNR, " & _
                  "ACC.HANDLE HANDLE_ACC, " & _
                  "PNR.LOCALIZADORA, PNR.DATAINCLUSAO, " & _
                  "CASE WHEN CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > 0 " & _
                  "AND CHARINDEX('</CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > " & _
                  "CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) THEN " & _
                  "TRANSLATE(SUBSTRING(CONVERT(VARCHAR(MAX), LOG.XMLRESERVA), " & _
                  "CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CiaOperadora>'), " & _
                  "CHARINDEX('</CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) - " & _
                  "(CHARINDEX('<CiaOperadora>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CiaOperadora>'))), " & _
                  "'áéíóúâêîôûãõäëïöüç', 'aeiouaeiouaoaeiouc') " & _
                  "Else Null End As EMPRESA_VIACAO " & _
                  "FROM VM_PNRS PNR " & _
                  "LEFT JOIN VM_PNRACCOUNTINGS ACC ON (ACC.PNR = PNR.HANDLE) " & _
                  "Left Join BB_LOGINTEGRACOES LOG On LOG.Handle = PNR.LOGINTEGRACAO " & _
                  "WHERE YEAR(PNR.DATAINCLUSAO) = YEAR(GETDATE()) " & _
                  "AND PNR.TIPORESERVA IN (11) AND ACC.TIPOMISCELANIO = 23 " & _
                  "AND ACC.TIPOACC IN (4) " & _
                  "AND (ACC.FORNECEDOR IS NULL OR ACC.FORNECEDOR IN (0))")

    qry_email.Active = True

    ' Configura o e-mail
    Set Email = NewMail
    Email.SendTo = "suporte.benner@kontik.com.br"
    Email.Subject = "Portal Benner - Processado Erro - Vendas Rodoviarias TMS sem Fornecedor - Posição - " & Format(Now, "DD/MM/YYYY")

    ' Inicializa o corpo do e-mail
    Corpo_Email = "Portal Benner - Processado Erro - Vendas Rodoviarias Lemontech sem Fornecedor - Posição - " & Format(Now, "DD/MM/YYYY") & vbNewLine & vbNewLine
    Corpo_Email = Corpo_Email & "Rodoviarias para Cadastrar Codigo" & vbNewLine & vbNewLine

    ' Processa todos os registros encontrados
    Do While Not qry_email.EOF
        ' Verifica se o campo EMPRESA_VIACAO não é nulo
        If Not IsNull(qry_email.FieldByName("EMPRESA_VIACAO").AsString) Then
            ' Acumula as informações de cada registro
            Registros_Processados = qry_email.FieldByName("LOCALIZADORA").AsString & " - " & _
                                    Format(qry_email.FieldByName("DATAINCLUSAO").AsDateTime, "DD/MM/YYYY") & " - " & _
                                    qry_email.FieldByName("EMPRESA_VIACAO").AsString

            Corpo_Email = Corpo_Email & Registros_Processados & vbNewLine
        End If

        ' Avança para o próximo registro
        qry_email.Next
    Loop

    ' Finaliza o corpo do e-mail
    Corpo_Email = Corpo_Email & vbNewLine & "Cadastrar os códigos das Rodoviarias acima no Benner." & vbNewLine
    Corpo_Email = Corpo_Email & "Cadastrar da seguinte forma: TMS - Empresa viação - QUEROPASSAGEM" & vbNewLine
    Corpo_Email = Corpo_Email & "Ao fazer o cadastro, as vendas no processado erro serão regularizadas." & vbNewLine
    Corpo_Email = Corpo_Email & "---------------x---------------x---------------x---------------x---------------" & vbNewLine
    Corpo_Email = Corpo_Email & vbNewLine & "Kontik Viagens" & vbNewLine & "Equipe TI"

    ' Envia o e-mail
    Email.Text.Add Corpo_Email
    Email.Send

    ' Limpeza
    Set Email = Nothing
    Set qry_email = Nothing
End Sub
