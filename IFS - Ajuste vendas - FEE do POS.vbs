Public Sub Main()
    Dim qrySQL As Object
    Set qrySQL = NewQuery
    qrySQL.Active = False
    qrySQL.Clear


	'Query para validação

	' SELECT PNR.HANDLE AS HANDLE_PNR,
'        ACC.HANDLE AS HANDLE_ACC,
'        PNR.LOCALIZADORA AS RLOC,
'        CASE WHEN ACC.TIPOACC = 1 THEN 'Aéreo' WHEN ACC.TIPOACC = 2 THEN 'Carro' WHEN ACC.TIPOACC = 3 THEN 'Hotel' ELSE 'Serviço' END AS TIPO_ACC_DESCRICAO,
'        ACC.REQUISICAO AS REQ,
'        ACC.PASSAGEIRONAOCAD,
'        CASE WHEN ACC.K9_SERVICO = 91 THEN 'Online' ELSE 'Offline' END AS K9_SERVICO_DESCRICAO,
'        CASE WHEN ACC.K9_REQUISICAOCUSTOM = 91 THEN 'Online' WHEN ACC.K9_REQUISICAOCUSTOM = 92 THEN 'Offline' ELSE 'Outro' END AS K9_REQUISICAOCUSTOM_DESCRICAO,
'        CASE WHEN ACC.K9_DESPESA = 91 THEN 'Nenhum' WHEN ACC.K9_DESPESA = 92 THEN 'Billable' ELSE 'No Billable' END AS K9_DESPESA_DESCRICAO,
'        ACC.ORIGEMPEDIDO,
'        CASE WHEN ACC.FORMARECEBIMENTO = 3 THEN 'Cartão de Crédito' WHEN ACC.FORMARECEBIMENTO = 5 THEN 'Faturado' WHEN ACC.FORMARECEBIMENTO = 6 THEN 'Cartão Amex' WHEN ACC.FORMARECEBIMENTO = 8 THEN 'Cartão Agência' ELSE 'Outro' END AS FORMARECEBIMENTO_DESCRICAO,
'        ACC.TIPOCCRC,
'        ACC.NUMEROCCRC,
'        ACC.VENCCCRC,
'        ACC.AUTORIZACAOCCRC,
'        ACC.TITULARCCRC,
'        CASE WHEN ACC.FORMAPAGAMENTO = 2 THEN 'Cartão de Crédito' WHEN ACC.FORMAPAGAMENTO = 3 THEN 'Faturado' ELSE 'Outro' END AS FORMARPAGAMENTO_DESCRICAO,
'        ACC.TIPOCCPG,
'        ACC.NUMEROCCPG,
'        ACC.VENCCCPG,
'        ACC.AUTORIZACAOCCPG,
'        ACC.TITULARCCPG,
'        CASE WHEN ACC.TIPOCOBRANCAFEE = 1 THEN 'Padrão' WHEN ACC.TIPOCOBRANCAFEE = 3 THEN 'Cartão de Crédito' ELSE 'Outro' END AS TIPOCOBRANCAFEE_DESCRICAO,
'        ACC.TIPOCCFEE,
'        ACC.NUMEROCCFEE,
'        ACC.VENCCCFEE,
'        ACC.AUTORIZACAOCCFEE,
'        ACC.TITULARCCFEE
'   FROM VM_PNRS PNR
'        LEFT JOIN VM_PNRACCOUNTINGS ACC ON ACC.PNR = PNR.HANDLE
'  WHERE PNR.LOCALIZADORA IN ()



    qrySQL.Add "SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, PNR.LOCALIZADORA AS RLOC, " & _
               "ACC.TIPOACC, ACC.REQUISICAO AS REQ, ACC.PASSAGEIRONAOCAD, " & _
               "ACC.K9_SERVICO, ACC.K9_REQUISICAOCUSTOM, ACC.K9_DESPESA, ACC.ORIGEMPEDIDO, " & _
               "ACC.TIPOCCRC, ACC.NUMEROCCRC, ACC.VENCCCRC, ACC.AUTORIZACAOCCRC, ACC.TITULARCCRC, " & _
               "ACC.NUMEROCCPG, ACC.AUTORIZACAOCCPG, ACC.TITULARCCPG, ACC.TIPOCCPG, ACC.VENCCCPG " & _
               "FROM VM_PNRS PNR " & _
               "LEFT JOIN VM_PNRACCOUNTINGS ACC ON ACC.PNR = PNR.HANDLE " & _
               "WHERE PNR.LOCALIZADORA IN ('CISTLOT','ILDFEWT','KIVKTYT','NXTEXQT','SCBSZJT','SNOQFST','SNRSPHT','AALRFIT','GTTXWHT','NZBLRKT','DLDGOXT','QCJUQOT')"

    qrySQL.Active = True

    On Error GoTo ErrorHandler
    While Not qrySQL.EOF
        Dim qryUPDT As Object
        Set qryUPDT = NewQuery

        If Not InTransaction Then StartTransaction

        ' Verifica e preenche o campo TITULARCCRC
        If Trim(qrySQL.FieldByName("TITULARCCRC").AsString) = "" Then
            qryUPDT.Clear
            qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET TITULARCCRC = '" & qrySQL.FieldByName("PASSAGEIRONAOCAD").AsString & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
            qryUPDT.ExecSQL
        End If

        ' Verifica e preenche o campo TITULARCCPG
        If Trim(qrySQL.FieldByName("TITULARCCPG").AsString) = "" Then
            qryUPDT.Clear
            qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET TITULARCCPG = '" & qrySQL.FieldByName("PASSAGEIRONAOCAD").AsString & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
            qryUPDT.ExecSQL
        End If

        ' Verifica e preenche o campo TIPOCCRC com TIPOCCPG
        If Trim(qrySQL.FieldByName("TIPOCCRC").AsString) = "" Then
            qryUPDT.Clear
            qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET TIPOCCRC = '" & qrySQL.FieldByName("TIPOCCPG").AsString & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
            qryUPDT.ExecSQL
        End If

        ' Verifica se VENCCCRC está vazio e VENCCCPG não está vazio antes de fazer o update
		If Trim(qrySQL.FieldByName("VENCCCRC").AsString) = "" And Trim(qrySQL.FieldByName("VENCCCPG").AsString) <> "" Then
		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET VENCCCRC = '" & qrySQL.FieldByName("VENCCCPG").AsString & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If

        ' Gera um número aleatório de 6 dígitos para AUTORIZACAOCCRC se TIPOACC for 3
        If qrySQL.FieldByName("TIPOACC").AsInteger = 3 And Trim(qrySQL.FieldByName("AUTORIZACAOCCRC").AsString) = "" Then
            Dim numAutorizacao As String
            numAutorizacao = Format(Int((999999 - 100000 + 1) * Rnd + 100000), "000000") ' Gera número aleatório
            qryUPDT.Clear
            qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET AUTORIZACAOCCRC = '" & numAutorizacao & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
            qryUPDT.ExecSQL
        End If

        ' Verifica se TIPOCCPG é "CA" e altera para "MC"
		If Trim(qrySQL.FieldByName("TIPOCCPG").AsString) = "CA" Then
		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET TIPOCCPG = 'MC' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If

		 ' Verifica se TIPOCCRC é "CA" e altera para "MC"
		If Trim(qrySQL.FieldByName("TIPOCCRC").AsString) = "CA" Then
		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET TIPOCCRC = 'MC' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If

		' Verifica se ORIGEMPEDIDO está preenchido e deixa como NULL
		If Trim(qrySQL.FieldByName("ORIGEMPEDIDO").AsString) <> "" Then
		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET ORIGEMPEDIDO = NULL WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If

		' Verifica se TIPOCCPG é "I4"
		If Trim(qrySQL.FieldByName("TIPOCCPG").AsString) = "I4" Then
		    ' Atualiza NUMEROCCPG concatenando "4" no início
		    Dim novoNumeroCCPG As String
		    novoNumeroCCPG = "4" & qrySQL.FieldByName("NUMEROCCPG").AsString ' Concatenando "4" com o valor existente

		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET NUMEROCCPG = '" & novoNumeroCCPG & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If

		' Verifica se TIPOCCRC é "I4"
		If Trim(qrySQL.FieldByName("TIPOCCRC").AsString) = "I4" Then
		    ' Atualiza TIPOCCPG para "VI"
		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET TIPOCCRC = 'VI' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If

		' Verifica se TIPOCCRC é "I4"
		If Trim(qrySQL.FieldByName("TIPOCCRC").AsString) = "I4" Then
		    ' Atualiza NUMEROCCPG concatenando "4" no início
		    Dim novoNumeroCC As String
		    novoNumeroCCRC = "4" & qrySQL.FieldByName("NUMEROCCRC").AsString ' Concatenando "4" com o valor existente

		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET NUMEROCCRC = '" & novoNumeroCCRC & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If

		' Verifica se TIPOCCPG é "I4"
		If Trim(qrySQL.FieldByName("TIPOCCPG").AsString) = "I4" Then
		    ' Atualiza TIPOCCPG para "VI"
		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET TIPOCCPG = 'VI' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If


		' Verifica se NUMEROCCRC está vazio e se NUMEROCCPG é numérico antes de fazer o update
		If Trim(qrySQL.FieldByName("NUMEROCCRC").AsString) = "" Then
		    If IsNumeric(qrySQL.FieldByName("NUMEROCCPG").AsString) Then
		        qryUPDT.Clear
		        qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET NUMEROCCRC = '" & qrySQL.FieldByName("NUMEROCCPG").AsString & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		        qryUPDT.ExecSQL
		    End If
		End If

		' Atualiza o campo REQUISICAO com o valor de LOCALIZADORA
		qryUPDT.Clear
		qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET REQUISICAO = '" & qrySQL.FieldByName("RLOC").AsString & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		qryUPDT.ExecSQL

		' Verifica se AUTORIZACAOCCPG está vazio e preenche com o valor de AUTORIZACAOCCRC
		If Trim(qrySQL.FieldByName("AUTORIZACAOCCPG").AsString) = "" Then
		    qryUPDT.Clear
		    qryUPDT.Add "UPDATE VM_PNRACCOUNTINGS SET AUTORIZACAOCCPG = '" & qrySQL.FieldByName("AUTORIZACAOCCRC").AsString & "' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_ACC").AsInteger
		    qryUPDT.ExecSQL
		End If


        If InTransaction Then Commit
        Set qryUPDT = Nothing

        ' Atualiza a situação do PNR
        Dim qryUPDT_1 As Object
        Set qryUPDT_1 = NewQuery

        qryUPDT_1.Clear
        qryUPDT_1.Add "UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='S', AGUARDANDOEMISSAO='S' WHERE HANDLE = " & qrySQL.FieldByName("HANDLE_PNR").AsInteger
        qryUPDT_1.ExecSQL

        If InTransaction Then Commit
        Set qryUPDT_1 = Nothing

        qrySQL.Next
    Wend

    Exit Sub

ErrorHandler:
    If InTransaction Then Rollback
    MsgBox "Erro ao executar a rotina: " & Err.Description, vbCritical
    Resume Next
End Sub
