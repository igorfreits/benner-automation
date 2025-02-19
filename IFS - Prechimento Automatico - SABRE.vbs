Public Sub Main
	' Codifique aqui o método principal
	Dim Contador As Integer
	Dim qry_vendas As Object
	Set qry_vendas = NewQuery
	qry_vendas.Active = False
	qry_vendas.Clear

	qry_vendas.Add ("SELECT PNR.HANDLE HANDLE_PNR, ACC.HANDLE HANDLE_ACC, ACC.TIPOACC, PNR.DATAEMISSAO PNR_DATAEMISSAO, GRU.HANDLE HANDLE_GRU, GRU.NOME GRUPO_EMPRESARIAL, CLI.HANDLE HANDLE_CLI, CLI.NOME CLIENTE, "+ _
				"ACC.INFPOLITICA, ACC.CENTRODECUSTO, ACC.K9_REASONCODE,ACC.K9_JUSTIFICATIVA, ACC.K9_FINALIDADEPREDEFINIDA,ACC.INFCFI, acc.K9_DIVISAOPAX, ACC.EMAILSOLICITANTE, ACC.PASSAGEIROEMAIL, PNR.LOCALIZADORA, ACC.INFFINALIDADE, ACC.INFNIVELFUNCIONARIO, ACC.K9_CENTROCUSTOPAX,ACC.K9_MATRICULAAPROVADOR,ACC.K9_MATRICULAAPROVADOR, "+ _
				"ACC.INFDIVISAO, ACC.INFAPROVADOR, ACC.MATRICULA,acc.K9_DEPARTAMENTOPAX,ACC.PROJETO, ACC.REFERENCIAL2, ACC.REFERENCIAL1, ACC.INFOS, ACC.K9_CLIENTEPAX, ACC.K9_CARGO,ACC.K9_UNIDADEPAX,ACC.K9_UNIDADE, ACC.K9_UDID1,ACC.K9_CPFPASSAGEIRO,ACC.K9_UDID3,ACC.K9_UDID2, ACC.K9_BREAK3, acc.K9_BREAK2,  ACC.K9_BREAK1, LEFT(PNR.LOCALIZADORA,6) RLOC, "+ _
				"ACC.EMPENHO, ACC.PASSAGEIRONAOCAD, ACC.INFMOTIVO,ACC.INFDATASOLICITACAO,ACC.K9_SERVICO,ACC.K9_REQUISICAOCUSTOM,ACC.REQUISICAO, ACC.K9_NUMEROINTERNOCLIENTE,ACC.INFDATASOLICITACAO,ACC.K9_DATAAPROVACAO, ACC.CANALVENDA, ACC.K9_JUSTIFICATIVAPREDEFINIDA,ACC.K9_MATRICULAAPROVADOR, ACC.INFJUSTIFICATIVA,ACC.K9_MATRICULASOLICITANTE, ACC.CONVIDADO,ACC.INFCONTROLE "+ _
				"FROM VM_PNRS PNR "+ _
				"LEFT JOIN VM_PNRACCOUNTINGS ACC ON (ACC.PNR=PNR.HANDLE) "+ _
				"LEFT JOIN BB_TIPORESERVA TRV ON (TRV.HANDLE=PNR.TIPORESERVA) "+ _
				"LEFT JOIN GN_PESSOAS CLI ON (CLI.HANDLE=PNR.CLIENTE) "+ _
				"LEFT JOIN GN_PESSOAS GRU ON (GRU.HANDLE=CLI.GRUPOEMPRESARIAL) "+ _
				"WHERE PNR.SITUACAO = 3 AND PNR.TIPORESERVA = 1 "+ _
				"ORDER BY GRU.HANDLE, CLI.HANDLE, ACC.TIPOACC, PNR.HANDLE, ACC.HANDLE")


	qry_vendas.Active = True

	Contador=0
	NovaTaxa=Empty

	While Not qry_vendas.EOF()

		'Verifica os campos obrigatórios
		Dim qry_campos_obg As Object
		Set qry_campos_obg = NewQuery
		qry_campos_obg.Active = False
		qry_campos_obg.Clear

		qry_campos_obg.Add ("SELECT OBG.CENTROCUSTO,OBG.CONVIDADO,OBG.REQUISICAO, OBG.K9_REQUISICAOCUSTOM,OBG.K9_SERVICO, OBG.POLITICA,OBG.CANALDEVENDA, OBG.MOTIVO,OBG.K9_JUSTIFICATIVAPREDEFINIDA,OBG.K9_JUSTIFICATIVA,OBG.JUSTIFICATIVA,OBG.CENTROCUSTOTER, OBG.K9_BREAK1,OBG.CFI,OBG.OS,OBG.OSTER,OBG.PROJETO,OBG.PROJETOTER, OBG.K9_CLIENTEPAX,OBG.K9_CLIENTEPAXTER, OBG.CFI, OBG.K9_BREAK1TER, OBG.K9_BREAK2, OBG.K9_BREAK2TER, OBG.K9_BREAK3, OBG.K9_BREAK3TER, OBG.K9_CARGO, OBG.K9_CARGOTER, OBG.FINALIDADE,OBG.FINALIDADETER, "+ _
					  "OBG.K9_UDID1, OBG.K9_UDID1TER, OBG.K9_UDID2,OBG.K9_NUMEROINTERNOCLIENTE,OBG.K9_NUMEROINTERNOCLIENTETER, OBG.K9_UDID2TER, OBG.K9_UDID3, OBG.K9_UDID3TER, OBG.K9_REASONCODE, OBG.K9_REASONCODETER, OBG.DIVISAO, OBG.K9_CENTROCUSTOPAX, "+ _
					  "OBG.DIVISAOTER, OBG.EMPENHO, OBG.EMPENHOTER, OBG.APROVADOR,OBG.DATASOLICITACAO, OBG.K9_UNIDADE,OBG.K9_UNIDADEPAX,OBG.K9_UNIDADEPAXTER,OBG.INFORMACAOREFERENCIAL2,OBG.INFORMACAOREFERENCIAL1, OBG.K9_UNIDADETER,OBG.K9_CPFPASSAGEIRO,OBG.K9_CPFPASSAGEIROTER, obg.K9_DEPARTAMENTOPAX,OBG.K9_DEPARTAMENTOPAXTER, "+ _
                      "OBG.K9_MATRICULASOLICITANTE, OBG.INFORMACAODECONTROLE, OBG.K9_MATRICULASOLICITANTETER, OBG.K9_DIVISAOPAX,obg.K9_DIVISAOPAXter,OBG.APROVADORTER, OBG.MATRICULA, OBG.MATRICULATER,OBG.K9_MATRICULAAPROVADOR, OBG.K9_FINALIDADEPREDEFINIDA, OBG.K9_FINALIDADEPREDEFINIDATER, OBG.NIVELFUNCIONARIOTER,OBG.NIVELFUNCIONARIO, OBG.K9_CENTROCUSTOPAXTER "+ _
					  "FROM BB_CLIENTECONTRATOSCAMPOSOBRIG OBG "+ _
					  "LEFT JOIN BB_CLIENTECONTRATOS CCT ON (CCT.HANDLE=OBG.CONTRATO) "+ _
					  "LEFT JOIN GN_PESSOAS CLI ON (CLI.HANDLE=CCT.PESSOA) "+ _
					  "WHERE CLI.HANDLE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
					  "AND CCT.TIPO=1")

		qry_campos_obg.Active = True

		If qry_vendas.FieldByName("CENTRODECUSTO").AsInteger = 0 Then
			If qry_campos_obg.FieldByName("CENTROCUSTO").AsInteger = 3 Then

				Dim qry_cc As Object
				Set qry_cc = NewQuery
				qry_cc.Active = False
				qry_cc.Clear

					qry_cc.Add ("SELECT TOP 1 ACC.CENTRODECUSTO "+ _
							"FROM VM_PNRACCOUNTINGS ACC "+ _
							"LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) "+ _
							"WHERE PNR.DATAINCLUSAO>=GETDATE()-45 "+ _
							"AND ACC.PASSAGEIRONAOCAD='"+qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString+"' "+ _
							"AND CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							"AND ACC.CENTRODECUSTO IS NOT NULL "+ _
							"ORDER BY PNR.HANDLE DESC")

					qry_cc.Active = True

				If Not qry_cc.EOF () Then

					Contador=Contador+1
					Dim updt_cc As Object
					Set updt_cc = NewQuery

					On Error GoTo ER0_01
					If Not InTransaction Then StartTransaction

				 updt_cc.Active = False
				 updt_cc.Clear

				 updt_cc.Add("UPDATE VM_PNRACCOUNTINGS SET CENTRODECUSTO="+ qry_cc.FieldByName("CENTRODECUSTO").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

				 updt_cc.ExecSQL

					If InTransaction Then Commit
					Set updt_cc = Nothing

					ER0_01:
					  If InTransaction Then Rollback
					  Set updt_cc = Nothing
				Else
					Dim qrySQL02 As Object
					Set qrySQL02 = NewQuery
					qrySQL02.Active = False
					qrySQL02.Clear

					qrySQL02.Add ("SELECT TOP 1 HANDLE "+ _
							"FROM BB_CLIENTECC "+ _
							"WHERE BB_CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							"AND VALIDO='S' "+ _
							"ORDER BY HANDLE DESC")

					qrySQL02.Active = True

					If Not qrySQL02.EOF () Then

						Contador=Contador+1
						Dim qryUPDT_02 As Object
						Set qryUPDT_02 = NewQuery

						On Error GoTo ER0_02
						If Not InTransaction Then StartTransaction

						qryUPDT_02.Active = False
						qryUPDT_02.Clear

						qryUPDT_02.Add("UPDATE VM_PNRACCOUNTINGS SET CENTRODECUSTO="+ qrySQL02.FieldByName("HANDLE").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

						qryUPDT_02.ExecSQL

						If InTransaction Then Commit
						Set qryUPDT_02 = Nothing

						ER0_02:
						  If InTransaction Then Rollback
						  Set qryUPDT_02 = Nothing
					End If
				End If
			End If
		End If

        'K9_UNIDADEPAX
        If qry_vendas.FieldByName("K9_UNIDADEPAX").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_UNIDADEPAX").AsInteger = 93 Then
                Dim QRY_K9_UNIDADEPAX As Object
                Set QRY_K9_UNIDADEPAX = NewQuery
                QRY_K9_UNIDADEPAX.Active = False
                QRY_K9_UNIDADEPAX.Clear

                QRY_K9_UNIDADEPAX.Add "SELECT TOP 1 ACC.K9_UNIDADEPAX FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_UNIDADEPAX IS NOT NULL " & _
                             "GROUP BY ACC.K9_UNIDADEPAX"

                QRY_K9_UNIDADEPAX.Active = True

                If Not QRY_K9_UNIDADEPAX.EOF Then
                    Dim UPT_K9_UNIDADEPAX As Object
                    Set UPT_K9_UNIDADEPAX = NewQuery
                    UPT_K9_UNIDADEPAX.Active = False
                    UPT_K9_UNIDADEPAX.Clear

                    UPT_K9_UNIDADEPAX.Add "UPDATE VM_PNRACCOUNTINGS SET K9_UNIDADEPAX='" & QRY_K9_UNIDADEPAX.FieldByName("K9_UNIDADEPAX").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_UNIDADEPAX.ExecSQL
                    Set UPT_K9_UNIDADEPAX = Nothing
                End If

                Set QRY_K9_UNIDADEPAX = Nothing
            End If
        End If

        'K9_DEPARTAMENTOPAX
        If qry_vendas.FieldByName("K9_DEPARTAMENTOPAX").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_DEPARTAMENTOPAX").AsInteger = 93 Then
                Dim QRY_K9_DEPARTAMENTOPAX As Object
                Set QRY_K9_DEPARTAMENTOPAX = NewQuery
                QRY_K9_DEPARTAMENTOPAX.Active = False
                QRY_K9_DEPARTAMENTOPAX.Clear

                QRY_K9_DEPARTAMENTOPAX.Add "SELECT TOP 1 ACC.K9_DEPARTAMENTOPAX FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_DEPARTAMENTOPAX IS NOT NULL " & _
                             "GROUP BY ACC.K9_DEPARTAMENTOPAX"

                QRY_K9_DEPARTAMENTOPAX.Active = True

                If Not QRY_K9_DEPARTAMENTOPAX.EOF Then
                    Dim UPT_K9_DEPARTAMENTOPAX As Object
                    Set UPT_K9_DEPARTAMENTOPAX = NewQuery
                    UPT_K9_DEPARTAMENTOPAX.Active = False
                    UPT_K9_DEPARTAMENTOPAX.Clear

                    UPT_K9_DEPARTAMENTOPAX.Add "UPDATE VM_PNRACCOUNTINGS SET K9_DEPARTAMENTOPAX='" & QRY_K9_DEPARTAMENTOPAX.FieldByName("K9_DEPARTAMENTOPAX").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_DEPARTAMENTOPAX.ExecSQL
                    Set UPT_K9_DEPARTAMENTOPAX = Nothing
                End If

                Set QRY_K9_DEPARTAMENTOPAX = Nothing
            End If
        End If

		'Ajusta DEPARTAMENTO
		If qry_vendas.FieldByName("EMPENHO").AsString = "" Then
            If qry_campos_obg.FieldByName("EMPENHO").AsInteger = 3 Then
                Dim qry_dept As Object
                Set qry_dept = NewQuery
                qry_dept.Active = False
                qry_dept.Clear

                qry_dept.Add "SELECT TOP 1 ACC.EMPENHO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.EMPENHO IS NOT NULL " & _
                             "GROUP BY ACC.EMPENHO"

                qry_dept.Active = True

                If Not qry_dept.EOF Then
                    Dim upt_dept As Object
                    Set upt_dept = NewQuery
                    upt_dept.Active = False
                    upt_dept.Clear

                    upt_dept.Add "UPDATE VM_PNRACCOUNTINGS SET EMPENHO='" & qry_dept.FieldByName("EMPENHO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_dept.ExecSQL
                    Set upt_dept = Nothing
                End If

                Set qry_dept = Nothing
            End If
        End If


		'K9_CENTROCUSTOPAX
		If qry_vendas.FieldByName("K9_CENTROCUSTOPAX").AsInteger = 0 Then
			If qry_campos_obg.FieldByName("K9_CENTROCUSTOPAX").AsInteger = 93 Then

				Dim qry_cc_pax As Object
				Set qry_cc_pax = NewQuery
				qry_cc_pax.Active = False
				qry_cc_pax.Clear

					qry_cc_pax.Add ("SELECT TOP 1 ACC.K9_CENTROCUSTOPAX "+ _
							"FROM VM_PNRACCOUNTINGS ACC "+ _
							"LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) "+ _
							"WHERE PNR.DATAINCLUSAO>=GETDATE()-45 "+ _
							"AND ACC.PASSAGEIRONAOCAD='"+qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString+"' "+ _
							"AND CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							"AND ACC.K9_CENTROCUSTOPAX IS NOT NULL "+ _
							"ORDER BY PNR.HANDLE DESC")

					qry_cc_pax.Active = True

				If Not qry_cc_pax.EOF () Then

					Contador=Contador+1
					Dim updt_cc_pax As Object
					Set updt_cc_pax = NewQuery

					On Error GoTo ER0_01
					If Not InTransaction Then StartTransaction

				 updt_cc_pax.Active = False
				 updt_cc_pax.Clear

				 updt_cc_pax.Add("UPDATE VM_PNRACCOUNTINGS SET K9_CENTROCUSTOPAX="+ qry_cc_pax.FieldByName("K9_CENTROCUSTOPAX").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

				 updt_cc_pax.ExecSQL

					If InTransaction Then Commit
					Set updt_cc_pax = Nothing

					ER0_0100:
					  If InTransaction Then Rollback
					  Set updt_cc_pax = Nothing
				Else
					Dim qry_cc_pax02 As Object
					Set qry_cc_pax02 = NewQuery
					qry_cc_pax02.Active = False
					qry_cc_pax02.Clear

					qry_cc_pax02.Add ("SELECT TOP 1 HANDLE "+ _
							"FROM BB_CLIENTECC "+ _
							"WHERE BB_CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							"AND VALIDO='S' "+ _
							"ORDER BY HANDLE DESC")

					qry_cc_pax02.Active = True

					If Not qry_cc_pax02.EOF () Then

						Contador=Contador+1
						Dim updt_cc_pax02 As Object
						Set updt_cc_pax02 = NewQuery

						On Error GoTo ER0_0200
						If Not InTransaction Then StartTransaction

						updt_cc_pax02.Active = False
						updt_cc_pax02.Clear

						updt_cc_pax02.Add("UPDATE VM_PNRACCOUNTINGS SET K9_CENTROCUSTOPAX="+ qry_cc_pax02.FieldByName("HANDLE").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

						updt_cc_pax02.ExecSQL

						If InTransaction Then Commit
						Set updt_cc_pax02 = Nothing

						ER0_0200:
						  If InTransaction Then Rollback
						  Set updt_cc_pax02 = Nothing
					End If
				End If
			End If
		End If

		If qry_vendas.FieldByName("EMAILSOLICITANTE").AsString = "" And qry_vendas.FieldByName("PASSAGEIROEMAIL").AsString = "" Then
			Dim qry_mail_solic As Object
			Set qry_mail_solic = NewQuery
			qry_mail_solic.Active = False
			qry_mail_solic.Clear

			qry_mail_solic.Add("SELECT TOP 1 ACC.EMAILSOLICITANTE "+ _
						"FROM VM_PNRACCOUNTINGS ACC "+ _
						"LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) "+ _
						"WHERE PNR.DATAINCLUSAO>=GETDATE()-45 "+ _
						"AND ACC.PASSAGEIRONAOCAD='"+qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString+"' "+ _
						"AND CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
						"AND ACC.EMAILSOLICITANTE IS NOT NULL "+ _
						"ORDER BY PNR.HANDLE DESC")

			qry_mail_solic.Active = True

			If Not qry_mail_solic.EOF () Then
				Contador=Contador+1
				Dim upt_mail_solic As Object
				Set upt_mail_solic = NewQuery

				On Error GoTo ER03_0
				If Not InTransaction Then StartTransaction

				upt_mail_solic.Active = False
				upt_mail_solic.Clear

				upt_mail_solic.Add("UPDATE VM_PNRACCOUNTINGS SET EMAILSOLICITANTE='"+ qry_mail_solic.FieldByName("EMAILSOLICITANTE").AsString +"', PASSAGEIROEMAIL='"+ qry_mail_solic.FieldByName("EMAILSOLICITANTE").AsString +"' WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

				upt_mail_solic.ExecSQL

				If InTransaction Then Commit
				Set upt_mail_solic = Nothing

				ER03_0:
				  If InTransaction Then Rollback
				  Set upt_mail_solic = Nothing
			Else
				Contador=Contador+1
				Dim qryUPDT03_1 As Object
				Set qryUPDT03_1 = NewQuery

				On Error GoTo ER03_1
				If Not InTransaction Then StartTransaction

				qryUPDT03_1.Active = False
				qryUPDT03_1.Clear

				qryUPDT03_1.Add("UPDATE VM_PNRACCOUNTINGS SET EMAILSOLICITANTE='"+ qry_vendas.FieldByName("K_EMAILNOTADEBITO").AsString +"', PASSAGEIROEMAIL='"+ qry_vendas.FieldByName("K_EMAILNOTADEBITO").AsString +"' WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

				qryUPDT03_1.ExecSQL

				If InTransaction Then Commit
				Set qryUPDT03_1 = Nothing

				ER03_1:
				  If InTransaction Then Rollback
				  Set qryUPDT03_1 = Nothing
			End If
		End If

		'Ajusta EMAILSOLICITANTE
		If qry_vendas.FieldByName("EMAILSOLICITANTE").AsString = "" And qry_vendas.FieldByName("PASSAGEIROEMAIL").AsString <> "" Then
			Contador=Contador+1

			Dim qryUPDT03 As Object
			Set qryUPDT03 = NewQuery

			On Error GoTo ER03
			If Not InTransaction Then StartTransaction

			qryUPDT03.Active = False
			qryUPDT03.Clear

			qryUPDT03.Add("UPDATE VM_PNRACCOUNTINGS SET EMAILSOLICITANTE='"+ qry_vendas.FieldByName("PASSAGEIROEMAIL").AsString +"' WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

			qryUPDT03.ExecSQL

			If InTransaction Then Commit
			Set qryUPDT03 = Nothing

			ER03:
			  If InTransaction Then Rollback
			  Set qryUPDT03 = Nothing
		End If

		'Ajusta PASSAGEIROEMAIL
		If qry_vendas.FieldByName("PASSAGEIROEMAIL").AsString = "" And qry_vendas.FieldByName("EMAILSOLICITANTE").AsString <> "" Then
			Contador=Contador+1

			Dim qryUPDT04 As Object
			Set qryUPDT04 = NewQuery

			On Error GoTo ER04
			If Not InTransaction Then StartTransaction

			qryUPDT04.Active = False
			qryUPDT04.Clear

			qryUPDT04.Add("UPDATE VM_PNRACCOUNTINGS SET PASSAGEIROEMAIL='"+ qry_vendas.FieldByName("EMAILSOLICITANTE").AsString +"' WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

			qryUPDT04.ExecSQL

			If InTransaction Then Commit
			Set qryUPDT04 = Nothing

			ER04:
			  If InTransaction Then Rollback
			  Set qryUPDT04 = Nothing
		End If

        'INFCONTROLE
        If qry_vendas.FieldByName("INFCONTROLE").AsString = "" Then
            If qry_campos_obg.FieldByName("INFORMACAODECONTROLE").AsInteger = 3 Then
                Dim QRY_INFCONTROLE As Object
                Set QRY_INFCONTROLE = NewQuery
                QRY_INFCONTROLE.Active = False
                QRY_INFCONTROLE.Clear

                QRY_INFCONTROLE.Add "SELECT TOP 1 ACC.INFCONTROLE FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFCONTROLE IS NOT NULL " & _
                             "GROUP BY ACC.INFCONTROLE"

                QRY_INFCONTROLE.Active = True

                If Not QRY_INFCONTROLE.EOF Then
                    Dim UPT_INFCONTROLE As Object
                    Set UPT_INFCONTROLE = NewQuery
                    UPT_INFCONTROLE.Active = False
                    UPT_INFCONTROLE.Clear

                    UPT_INFCONTROLE.Add "UPDATE VM_PNRACCOUNTINGS SET INFCONTROLE='" & QRY_INFCONTROLE.FieldByName("INFCONTROLE").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_INFCONTROLE.ExecSQL
                    Set UPT_INFCONTROLE = Nothing
                End If

                Set QRY_INFCONTROLE = Nothing
            End If
        End If

        'PROJETO
        If qry_vendas.FieldByName("PROJETO").AsString = "" Then
            If qry_campos_obg.FieldByName("PROJETO").AsInteger = 3 Then
                Dim QRY_PROJETO As Object
                Set QRY_PROJETO = NewQuery
                QRY_PROJETO.Active = False
                QRY_PROJETO.Clear

                QRY_PROJETO.Add "SELECT TOP 1 ACC.PROJETO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.PROJETO IS NOT NULL " & _
                             "GROUP BY ACC.PROJETO"

                QRY_PROJETO.Active = True

                If Not QRY_PROJETO.EOF Then
                    Dim UPT_PROJETO As Object
                    Set UPT_PROJETO = NewQuery
                    UPT_PROJETO.Active = False
                    UPT_PROJETO.Clear

                    UPT_PROJETO.Add "UPDATE VM_PNRACCOUNTINGS SET PROJETO='" & QRY_PROJETO.FieldByName("PROJETO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_PROJETO.ExecSQL
                    Set UPT_PROJETO = Nothing
                End If

                Set QRY_PROJETO = Nothing
            End If
        End If
        'REFERENCIAL1
        If qry_vendas.FieldByName("REFERENCIAL1").AsString = "" Then
            If qry_campos_obg.FieldByName("INFORMACAOREFERENCIAL1").AsInteger = 3 Then
                Dim QRY_REFERENCIAL1 As Object
                Set QRY_REFERENCIAL1 = NewQuery
                QRY_REFERENCIAL1.Active = False
                QRY_REFERENCIAL1.Clear

                QRY_REFERENCIAL1.Add "SELECT TOP 1 ACC.REFERENCIAL1 FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.REFERENCIAL1 IS NOT NULL " & _
                             "GROUP BY ACC.REFERENCIAL1"

                QRY_REFERENCIAL1.Active = True

                If Not QRY_REFERENCIAL1.EOF Then
                    Dim UPT_REFERENCIAL1 As Object
                    Set UPT_REFERENCIAL1 = NewQuery
                    UPT_REFERENCIAL1.Active = False
                    UPT_REFERENCIAL1.Clear

                    UPT_REFERENCIAL1.Add "UPDATE VM_PNRACCOUNTINGS SET REFERENCIAL1='" & QRY_REFERENCIAL1.FieldByName("REFERENCIAL1").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_REFERENCIAL1.ExecSQL
                    Set UPT_REFERENCIAL1 = Nothing
                End If

                Set QRY_REFERENCIAL1 = Nothing
            End If
        End If

        'REFERENCIAL2
        If qry_vendas.FieldByName("REFERENCIAL2").AsString = "" Then
            If qry_campos_obg.FieldByName("INFORMACAOREFERENCIAL2").AsInteger = 3 Then
                Dim QRY_REFERENCIAL2 As Object
                Set QRY_REFERENCIAL2 = NewQuery
                QRY_REFERENCIAL2.Active = False
                QRY_REFERENCIAL2.Clear

                QRY_REFERENCIAL2.Add "SELECT TOP 1 ACC.REFERENCIAL2 FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.REFERENCIAL2 IS NOT NULL " & _
                             "GROUP BY ACC.REFERENCIAL2"

                QRY_REFERENCIAL2.Active = True

                If Not QRY_REFERENCIAL2.EOF Then
                    Dim UPT_REFERENCIAL2 As Object
                    Set UPT_REFERENCIAL2 = NewQuery
                    UPT_REFERENCIAL2.Active = False
                    UPT_REFERENCIAL2.Clear

                    UPT_REFERENCIAL2.Add "UPDATE VM_PNRACCOUNTINGS SET REFERENCIAL2='" & QRY_REFERENCIAL2.FieldByName("REFERENCIAL2").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_REFERENCIAL2.ExecSQL
                    Set UPT_REFERENCIAL2 = Nothing
                End If

                Set QRY_REFERENCIAL2 = Nothing
            End If
        End If
        'REQUISICAO
        If qry_vendas.FieldByName("REQUISICAO").AsString = "" Or qry_vendas.FieldByName("REQUISICAO").AsString <> qry_vendas.FieldByName("RLOC").AsString Then
				Contador=Contador+1

				Dim QRY_REQUISICAO As Object
				Set QRY_REQUISICAO = NewQuery

				On Error GoTo ER00
				If Not InTransaction Then StartTransaction

				QRY_REQUISICAO.Active = False
				QRY_REQUISICAO.Clear

				QRY_REQUISICAO.Add("UPDATE VM_PNRACCOUNTINGS SET REQUISICAO='"+ qry_vendas.FieldByName("RLOC").AsString +"' WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

				QRY_REQUISICAO.ExecSQL

				If InTransaction Then Commit
				Set QRY_REQUISICAO = Nothing

				ER00:
				  If InTransaction Then Rollback
				  Set QRY_REQUISICAO = Nothing
			End If

        'DATASOLICITACAO
            If qry_vendas.FieldByName("INFDATASOLICITACAO").AsString = "" Then
                If qry_campos_obg.FieldByName("DATASOLICITACAO").AsInteger = 3 Then
				Contador=Contador+1

				Dim QRY_DATASOLICITACAO As Object
				Set QRY_DATASOLICITACAO = NewQuery

				On Error GoTo ER00
				If Not InTransaction Then StartTransaction

				QRY_DATASOLICITACAO.Active = False
				QRY_DATASOLICITACAO.Clear

				QRY_DATASOLICITACAO.Add("UPDATE VM_PNRACCOUNTINGS SET INFDATASOLICITACAO='" + Format(qry_vendas.FieldByName("PNR_DATAEMISSAO").AsDateTime,"MM/DD/YYYY") + "' WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

				QRY_DATASOLICITACAO.ExecSQL

				If InTransaction Then Commit
				Set QRY_DATASOLICITACAO = Nothing

				ER00789:
				  If InTransaction Then Rollback
				  Set QRY_DATASOLICITACAO = Nothing
			End If
        End If


        'INFMOTIVO
        If qry_vendas.FieldByName("INFMOTIVO").AsString = "" Then
            If qry_campos_obg.FieldByName("MOTIVO").AsInteger = 3 Then
                Dim QRY_INFMOTIVO As Object
                Set QRY_INFMOTIVO = NewQuery
                QRY_INFMOTIVO.Active = False
                QRY_INFMOTIVO.Clear

                QRY_INFMOTIVO.Add "SELECT TOP 1 ACC.INFMOTIVO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFMOTIVO IS NOT NULL " & _
                             "GROUP BY ACC.INFMOTIVO"

                QRY_INFMOTIVO.Active = True

                If Not QRY_INFMOTIVO.EOF Then
                    Dim UPT_INFMOTIVO As Object
                    Set UPT_INFMOTIVO = NewQuery
                    UPT_INFMOTIVO.Active = False
                    UPT_INFMOTIVO.Clear

                    UPT_INFMOTIVO.Add "UPDATE VM_PNRACCOUNTINGS SET INFMOTIVO='" & QRY_INFMOTIVO.FieldByName("INFMOTIVO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_INFMOTIVO.ExecSQL
                    Set UPT_INFMOTIVO = Nothing
                End If

                Set QRY_INFMOTIVO = Nothing
            End If
        End If

        'NUMEROINTERNOCLIENTE
        If qry_vendas.FieldByName("K9_NUMEROINTERNOCLIENTE").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_NUMEROINTERNOCLIENTE").AsInteger = 93 Then
                Dim QRY_K9_NUMEROINTERNOCLIENTE As Object
                Set QRY_K9_NUMEROINTERNOCLIENTE = NewQuery
                QRY_K9_NUMEROINTERNOCLIENTE.Active = False
                QRY_K9_NUMEROINTERNOCLIENTE.Clear

                QRY_K9_NUMEROINTERNOCLIENTE.Add "SELECT TOP 1 ACC.K9_NUMEROINTERNOCLIENTE FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_NUMEROINTERNOCLIENTE IS NOT NULL " & _
                             "GROUP BY ACC.K9_NUMEROINTERNOCLIENTE"

                QRY_K9_NUMEROINTERNOCLIENTE.Active = True

                If Not QRY_K9_NUMEROINTERNOCLIENTE.EOF Then
                    Dim UPT_NUMEROINTERNOCLIENTE As Object
                    Set UPT_NUMEROINTERNOCLIENTE = NewQuery
                    UPT_NUMEROINTERNOCLIENTE.Active = False
                    UPT_NUMEROINTERNOCLIENTE.Clear

                    UPT_NUMEROINTERNOCLIENTE.Add "UPDATE VM_PNRACCOUNTINGS SET K9_NUMEROINTERNOCLIENTE='" & QRY_K9_NUMEROINTERNOCLIENTE.FieldByName("K9_NUMEROINTERNOCLIENTE").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_NUMEROINTERNOCLIENTE.ExecSQL
                    Set UPT_NUMEROINTERNOCLIENTE = Nothing
                End If

                Set QRY_K9_NUMEROINTERNOCLIENTE = Nothing
            End If
        End If
        'CONVIDADO
        If qry_vendas.FieldByName("CONVIDADO").AsString = "" Then
            If qry_campos_obg.FieldByName("CONVIDADO").AsInteger = 3 Then
                Dim QRY_CONVIDADO As Object
                Set QRY_CONVIDADO = NewQuery
                QRY_CONVIDADO.Active = False
                QRY_CONVIDADO.Clear

                QRY_CONVIDADO.Add "SELECT TOP 1 ACC.CONVIDADO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.CONVIDADO IS NOT NULL " & _
                             "GROUP BY ACC.CONVIDADO"

                QRY_CONVIDADO.Active = True

                If Not QRY_CONVIDADO.EOF Then
                    Dim UPT_CONVIDADO As Object
                    Set UPT_CONVIDADO = NewQuery
                    UPT_CONVIDADO.Active = False
                    UPT_CONVIDADO.Clear

                    UPT_CONVIDADO.Add "UPDATE VM_PNRACCOUNTINGS SET CONVIDADO='" & QRY_CONVIDADO.FieldByName("CONVIDADO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_CONVIDADO.ExecSQL
                    Set UPT_CONVIDADO = Nothing
                End If

                Set QRY_CONVIDADO = Nothing
            End If
        End If
        'K9_JUSTIFICATIVAPREDEFINIDA
        If qry_vendas.FieldByName("K9_JUSTIFICATIVAPREDEFINIDA").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_JUSTIFICATIVAPREDEFINIDA").AsInteger = 93 Then
                Dim QRY_K9_JUSTIFICATIVAPREDEFINIDA As Object
                Set QRY_K9_JUSTIFICATIVAPREDEFINIDA = NewQuery
                QRY_K9_JUSTIFICATIVAPREDEFINIDA.Active = False
                QRY_K9_JUSTIFICATIVAPREDEFINIDA.Clear

                QRY_K9_JUSTIFICATIVAPREDEFINIDA.Add "SELECT TOP 1 ACC.K9_JUSTIFICATIVAPREDEFINIDA FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_JUSTIFICATIVAPREDEFINIDA IS NOT NULL " & _
                             "GROUP BY ACC.K9_JUSTIFICATIVAPREDEFINIDA"

                QRY_K9_JUSTIFICATIVAPREDEFINIDA.Active = True

                If Not QRY_K9_JUSTIFICATIVAPREDEFINIDA.EOF Then
                    Dim UPT_K9_JUSTIFICATIVAPREDEFINIDA As Object
                    Set UPT_K9_JUSTIFICATIVAPREDEFINIDA = NewQuery
                    UPT_K9_JUSTIFICATIVAPREDEFINIDA.Active = False
                    UPT_K9_JUSTIFICATIVAPREDEFINIDA.Clear

                    UPT_K9_JUSTIFICATIVAPREDEFINIDA.Add "UPDATE VM_PNRACCOUNTINGS SET K9_JUSTIFICATIVAPREDEFINIDA='" & QRY_K9_JUSTIFICATIVAPREDEFINIDA.FieldByName("K9_JUSTIFICATIVAPREDEFINIDA").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_JUSTIFICATIVAPREDEFINIDA.ExecSQL
                    Set UPT_K9_JUSTIFICATIVAPREDEFINIDA = Nothing
                End If

                Set QRY_K9_JUSTIFICATIVAPREDEFINIDA = Nothing
            End If
        End If
        'CANALVENDA
        If qry_vendas.FieldByName("CANALVENDA").AsString = "" Then
            If qry_campos_obg.FieldByName("CANALDEVENDA").AsInteger = 3 Then
                Dim QRY_CANALVENDA As Object
                Set QRY_CANALVENDA = NewQuery
                QRY_CANALVENDA.Active = False
                QRY_CANALVENDA.Clear

                QRY_CANALVENDA.Add "SELECT TOP 1 ACC.CANALVENDA FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.CANALVENDA IS NOT NULL " & _
                             "GROUP BY ACC.CANALVENDA"

                QRY_CANALVENDA.Active = True

                If Not QRY_CANALVENDA.EOF Then
                    Dim UPT_CANALVENDA As Object
                    Set UPT_CANALVENDA = NewQuery
                    UPT_CANALVENDA.Active = False
                    UPT_CANALVENDA.Clear

                    UPT_CANALVENDA.Add "UPDATE VM_PNRACCOUNTINGS SET CANALVENDA='" & QRY_CANALVENDA.FieldByName("CANALVENDA").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_CANALVENDA.ExecSQL
                    Set UPT_CANALVENDA = Nothing
                End If

                Set QRY_CANALVENDA = Nothing
            End If
        End If
        'STATUSREQUISICAO
        If qry_vendas.FieldByName("K9_REQUISICAOCUSTOM").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_REQUISICAOCUSTOM").AsInteger = 93 Then
                Dim QRY_K9_REQUISICAOCUSTOM As Object
                Set QRY_K9_REQUISICAOCUSTOM = NewQuery
                QRY_K9_REQUISICAOCUSTOM.Active = False
                QRY_K9_REQUISICAOCUSTOM.Clear

                QRY_K9_REQUISICAOCUSTOM.Add "SELECT TOP 1 ACC.K9_REQUISICAOCUSTOM FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_REQUISICAOCUSTOM IS NOT NULL " & _
                             "GROUP BY ACC.K9_REQUISICAOCUSTOM"

                QRY_K9_REQUISICAOCUSTOM.Active = True

                If Not QRY_K9_REQUISICAOCUSTOM.EOF Then
                    Dim UPT_K9_REQUISICAOCUSTOM As Object
                    Set UPT_K9_REQUISICAOCUSTOM = NewQuery
                    UPT_K9_REQUISICAOCUSTOM.Active = False
                    UPT_K9_REQUISICAOCUSTOM.Clear

                    UPT_K9_REQUISICAOCUSTOM.Add "UPDATE VM_PNRACCOUNTINGS SET K9_REQUISICAOCUSTOM='" & QRY_K9_REQUISICAOCUSTOM.FieldByName("K9_REQUISICAOCUSTOM").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_REQUISICAOCUSTOM.ExecSQL
                    Set UPT_K9_REQUISICAOCUSTOM = Nothing
                End If

                Set QRY_K9_REQUISICAOCUSTOM = Nothing
            End If
        End If
        'STATUSSERVIÇO
        If qry_vendas.FieldByName("K9_SERVICO").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_SERVICO").AsInteger = 93 Then
                Dim QRY_K9_SERVICO As Object
                Set QRY_K9_SERVICO = NewQuery
                QRY_K9_SERVICO.Active = False
                QRY_K9_SERVICO.Clear

                QRY_K9_SERVICO.Add "SELECT TOP 1 ACC.K9_SERVICO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_SERVICO IS NOT NULL " & _
                             "GROUP BY ACC.K9_SERVICO"

                QRY_K9_SERVICO.Active = True

                If Not QRY_K9_SERVICO.EOF Then
                    Dim UPT_K9_SERVICO As Object
                    Set UPT_K9_SERVICO = NewQuery
                    UPT_K9_SERVICO.Active = False
                    UPT_K9_SERVICO.Clear

                    UPT_K9_SERVICO.Add "UPDATE VM_PNRACCOUNTINGS SET K9_SERVICO='" & QRY_K9_SERVICO.FieldByName("K9_SERVICO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_SERVICO.ExecSQL
                    Set UPT_K9_SERVICO = Nothing
                End If

                Set QRY_K9_SERVICO = Nothing
            End If
        End If

        'INFJUSTIFICATIVA
        If qry_vendas.FieldByName("INFJUSTIFICATIVA").AsString = "" Then
            If qry_campos_obg.FieldByName("JUSTIFICATIVA").AsInteger = 3 Then
                Dim QRY_INFJUSTIFICATIVA As Object
                Set QRY_INFJUSTIFICATIVA = NewQuery
                QRY_INFJUSTIFICATIVA.Active = False
                QRY_INFJUSTIFICATIVA.Clear

                QRY_INFJUSTIFICATIVA.Add "SELECT TOP 1 ACC.INFJUSTIFICATIVA FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFJUSTIFICATIVA IS NOT NULL " & _
                             "GROUP BY ACC.INFJUSTIFICATIVA"

                QRY_INFJUSTIFICATIVA.Active = True

                If Not QRY_INFJUSTIFICATIVA.EOF Then
                    Dim UPT_INFJUSTIFICATIVA As Object
                    Set UPT_INFJUSTIFICATIVA = NewQuery
                    UPT_INFJUSTIFICATIVA.Active = False
                    UPT_INFJUSTIFICATIVA.Clear

                    UPT_INFJUSTIFICATIVA.Add "UPDATE VM_PNRACCOUNTINGS SET INFJUSTIFICATIVA='" & QRY_INFJUSTIFICATIVA.FieldByName("INFJUSTIFICATIVA").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_INFJUSTIFICATIVA.ExecSQL
                    Set UPT_INFJUSTIFICATIVA = Nothing
                End If

                Set QRY_INFJUSTIFICATIVA = Nothing
            End If
        End If
        'K9_JUSTIFICATIVA
        If qry_vendas.FieldByName("K9_JUSTIFICATIVA").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_JUSTIFICATIVA").AsInteger = 93 Then
                Dim QRY_K9_JUSTIFICATIVA As Object
                Set QRY_K9_JUSTIFICATIVA = NewQuery
                QRY_K9_JUSTIFICATIVA.Active = False
                QRY_K9_JUSTIFICATIVA.Clear

                QRY_K9_JUSTIFICATIVA.Add "SELECT TOP 1 CAST(ACC.K9_JUSTIFICATIVA AS VARCHAR(MAX)) AS K9_JUSTIFICATIVA " & _
                            "FROM VM_PNRACCOUNTINGS ACC " & _
                            "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE = ACC.PNR) " & _
                            "WHERE PNR.DATAINCLUSAO >= GETDATE() - 45 " & _
                            "AND ACC.PASSAGEIRONAOCAD = '" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                            "AND PNR.CLIENTE = " & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                            "AND CAST(ACC.K9_JUSTIFICATIVA AS VARCHAR(MAX)) IS NOT NULL " & _
                            "GROUP BY CAST(ACC.K9_JUSTIFICATIVA AS VARCHAR(MAX))"

                QRY_K9_JUSTIFICATIVA.Active = True

                If Not QRY_K9_JUSTIFICATIVA.EOF Then
                    Dim UPT_K9_JUSTIFICATIVA As Object
                    Set UPT_K9_JUSTIFICATIVA = NewQuery
                    UPT_K9_JUSTIFICATIVA.Active = False
                    UPT_K9_JUSTIFICATIVA.Clear

                    ' Atualiza a justificativa convertida para VARCHAR(MAX) para evitar erro de tipos
                    UPT_K9_JUSTIFICATIVA.Add "UPDATE VM_PNRACCOUNTINGS " & _
                        "SET K9_JUSTIFICATIVA = CAST('" & QRY_K9_JUSTIFICATIVA.FieldByName("K9_JUSTIFICATIVA").AsString & "' AS VARCHAR(MAX)) " & _
                        "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_JUSTIFICATIVA.ExecSQL
                    Set UPT_K9_JUSTIFICATIVA = Nothing
                End If

                Set QRY_K9_JUSTIFICATIVA = Nothing
            End If
        End If



        'INFPOLITICA
        If qry_vendas.FieldByName("INFPOLITICA").AsString = "" Then
            If qry_campos_obg.FieldByName("POLITICA").AsInteger = 3 Then
                Dim QRY_INFPOLITICA As Object
                Set QRY_INFPOLITICA = NewQuery
                QRY_INFPOLITICA.Active = False
                QRY_INFPOLITICA.Clear

                QRY_INFPOLITICA.Add "SELECT TOP 1 ACC.INFPOLITICA FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFPOLITICA IS NOT NULL " & _
                             "GROUP BY ACC.INFPOLITICA"

                QRY_INFPOLITICA.Active = True

                If Not QRY_INFPOLITICA.EOF Then
                    Dim UPT_INFPOLITICA As Object
                    Set UPT_INFPOLITICA = NewQuery
                    UPT_INFPOLITICA.Active = False
                    UPT_INFPOLITICA.Clear

                    UPT_INFPOLITICA.Add "UPDATE VM_PNRACCOUNTINGS SET INFPOLITICA='" & QRY_INFPOLITICA.FieldByName("INFPOLITICA").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_INFPOLITICA.ExecSQL
                    Set UPT_INFPOLITICA = Nothing
                End If

                Set QRY_INFPOLITICA = Nothing
            End If
        End If

        'K9_MATRICULASOLICITANTE
        If qry_vendas.FieldByName("K9_MATRICULASOLICITANTE").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_MATRICULASOLICITANTE").AsInteger = 93 Then
                Dim QRY_K9_MATRICULASOLICITANTE As Object
                Set QRY_K9_MATRICULASOLICITANTE = NewQuery
                QRY_K9_MATRICULASOLICITANTE.Active = False
                QRY_K9_MATRICULASOLICITANTE.Clear

                QRY_K9_MATRICULASOLICITANTE.Add "SELECT TOP 1 ACC.K9_MATRICULASOLICITANTE FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_MATRICULASOLICITANTE IS NOT NULL " & _
                             "GROUP BY ACC.K9_MATRICULASOLICITANTE"

                QRY_K9_MATRICULASOLICITANTE.Active = True

                If Not QRY_K9_MATRICULASOLICITANTE.EOF Then
                    Dim UPT_K9_MATRICULASOLICITANTE As Object
                    Set UPT_K9_MATRICULASOLICITANTE = NewQuery
                    UPT_K9_MATRICULASOLICITANTE.Active = False
                    UPT_K9_MATRICULASOLICITANTE.Clear

                    UPT_K9_MATRICULASOLICITANTE.Add "UPDATE VM_PNRACCOUNTINGS SET K9_MATRICULASOLICITANTE='" & QRY_K9_MATRICULASOLICITANTE.FieldByName("K9_MATRICULASOLICITANTE").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_MATRICULASOLICITANTE.ExecSQL
                    Set UPT_K9_MATRICULASOLICITANTE = Nothing
                End If

                Set QRY_K9_MATRICULASOLICITANTE = Nothing
            End If
        End If

        'K9_MATRICULAAPROVADOR
        If qry_vendas.FieldByName("K9_MATRICULAAPROVADOR").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_MATRICULAAPROVADOR").AsInteger = 93 Then
                Dim QRY_K9_MATRICULAAPROVADOR As Object
                Set QRY_K9_MATRICULAAPROVADOR = NewQuery
                QRY_K9_MATRICULAAPROVADOR.Active = False
                QRY_K9_MATRICULAAPROVADOR.Clear

                QRY_K9_MATRICULAAPROVADOR.Add "SELECT TOP 1 ACC.K9_MATRICULAAPROVADOR FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_MATRICULAAPROVADOR IS NOT NULL " & _
                             "GROUP BY ACC.K9_MATRICULAAPROVADOR"

                QRY_K9_MATRICULAAPROVADOR.Active = True

                If Not QRY_K9_MATRICULAAPROVADOR.EOF Then
                    Dim UPT_K9_MATRICULAAPROVADOR As Object
                    Set UPT_K9_MATRICULAAPROVADOR = NewQuery
                    UPT_K9_MATRICULAAPROVADOR.Active = False
                    UPT_K9_MATRICULAAPROVADOR.Clear

                    UPT_K9_MATRICULAAPROVADOR.Add "UPDATE VM_PNRACCOUNTINGS SET K9_MATRICULAAPROVADOR='" & QRY_K9_MATRICULAAPROVADOR.FieldByName("K9_MATRICULAAPROVADOR").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_MATRICULAAPROVADOR.ExecSQL
                    Set UPT_K9_MATRICULAAPROVADOR = Nothing
                End If

                Set QRY_K9_MATRICULAAPROVADOR = Nothing
            End If
        End If

		'Finalidade
		If qry_vendas.FieldByName("INFFINALIDADE").AsString = "" Then
            If qry_campos_obg.FieldByName("FINALIDADE").AsInteger = 3 Then
                Dim qry_fin As Object
                Set qry_fin = NewQuery
                qry_fin.Active = False
                qry_fin.Clear

                qry_fin.Add "SELECT TOP 1 ACC.INFFINALIDADE FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFFINALIDADE IS NOT NULL " & _
                             "GROUP BY ACC.INFFINALIDADE"

                qry_fin.Active = True

                If Not qry_fin.EOF Then
                    Dim upd_fin As Object
                    Set upd_fin = NewQuery
                    upd_fin.Active = False
                    upd_fin.Clear

                    upd_fin.Add "UPDATE VM_PNRACCOUNTINGS SET INFFINALIDADE='" & qry_fin.FieldByName("INFFINALIDADE").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upd_fin.ExecSQL
                    Set upd_fin = Nothing
                End If

                Set qry_fin = Nothing
            End If
        End If

		'Nivel Funcionario
		If qry_vendas.FieldByName("INFNIVELFUNCIONARIO").AsString = "" Then
            If qry_campos_obg.FieldByName("NIVELFUNCIONARIO").AsInteger = 3 Then
                Dim qry_nv_func As Object
                Set qry_nv_func = NewQuery
                qry_nv_func.Active = False
                qry_nv_func.Clear

                qry_nv_func.Add "SELECT TOP 1 ACC.INFNIVELFUNCIONARIO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFNIVELFUNCIONARIO IS NOT NULL " & _
                             "GROUP BY ACC.INFNIVELFUNCIONARIO"

                qry_nv_func.Active = True

                If Not qry_nv_func.EOF Then
                    Dim upt_nv_func As Object
                    Set upt_nv_func = NewQuery
                    upt_nv_func.Active = False
                    upt_nv_func.Clear

                    upt_nv_func.Add "UPDATE VM_PNRACCOUNTINGS SET INFNIVELFUNCIONARIO='" & qry_nv_func.FieldByName("INFNIVELFUNCIONARIO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_nv_func.ExecSQL
                    Set upt_nv_func = Nothing
                End If

                Set qry_nv_func = Nothing
            End If
        End If

        'K9_CPFPASSAGEIRO
        If qry_vendas.FieldByName("K9_CPFPASSAGEIRO").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_CPFPASSAGEIRO").AsInteger = 93 Then
                Dim QRY_K9_CPFPASSAGEIRO As Object
                Set QRY_K9_CPFPASSAGEIRO = NewQuery
                QRY_K9_CPFPASSAGEIRO.Active = False
                QRY_K9_CPFPASSAGEIRO.Clear

                QRY_K9_CPFPASSAGEIRO.Add "SELECT TOP 1 ACC.K9_CPFPASSAGEIRO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_CPFPASSAGEIRO IS NOT NULL " & _
                             "GROUP BY ACC.K9_CPFPASSAGEIRO"

                QRY_K9_CPFPASSAGEIRO.Active = True

                If Not QRY_K9_CPFPASSAGEIRO.EOF Then
                    Dim UPT_K9_CPFPASSAGEIRO As Object
                    Set UPT_K9_CPFPASSAGEIRO = NewQuery
                    UPT_K9_CPFPASSAGEIRO.Active = False
                    UPT_K9_CPFPASSAGEIRO.Clear

                    UPT_K9_CPFPASSAGEIRO.Add "UPDATE VM_PNRACCOUNTINGS SET K9_CPFPASSAGEIRO='" & QRY_K9_CPFPASSAGEIRO.FieldByName("K9_CPFPASSAGEIRO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_CPFPASSAGEIRO.ExecSQL
                    Set UPT_K9_CPFPASSAGEIRO = Nothing
                End If

                Set QRY_K9_CPFPASSAGEIRO = Nothing
            End If
        End If
        'K9_CLIENTEPAX
        If qry_vendas.FieldByName("K9_CLIENTEPAX").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_CLIENTEPAX").AsInteger = 93 Then
                Dim QRY_K9_CLIENTEPAX As Object
                Set QRY_K9_CLIENTEPAX = NewQuery
                QRY_K9_CLIENTEPAX.Active = False
                QRY_K9_CLIENTEPAX.Clear

                QRY_K9_CLIENTEPAX.Add "SELECT TOP 1 ACC.K9_CLIENTEPAX FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_CLIENTEPAX IS NOT NULL " & _
                             "GROUP BY ACC.K9_CLIENTEPAX"

                QRY_K9_CLIENTEPAX.Active = True

                If Not QRY_K9_CLIENTEPAX.EOF Then
                    Dim UPT_K9_CLIENTEPAX As Object
                    Set UPT_K9_CLIENTEPAX = NewQuery
                    UPT_K9_CLIENTEPAX.Active = False
                    UPT_K9_CLIENTEPAX.Clear

                    UPT_K9_CLIENTEPAX.Add "UPDATE VM_PNRACCOUNTINGS SET K9_CLIENTEPAX='" & QRY_K9_CLIENTEPAX.FieldByName("K9_CLIENTEPAX").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_CLIENTEPAX.ExecSQL
                    Set UPT_K9_CLIENTEPAX = Nothing
                End If

                Set QRY_K9_CLIENTEPAX = Nothing
            End If
        End If
        'INFOS
        If qry_vendas.FieldByName("INFOS").AsString = "" Then
            If qry_campos_obg.FieldByName("OS").AsInteger = 3 Then
                Dim QRY_INFOS As Object
                Set QRY_INFOS = NewQuery
                QRY_INFOS.Active = False
                QRY_INFOS.Clear

                QRY_INFOS.Add "SELECT TOP 1 ACC.INFOS FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFOS IS NOT NULL " & _
                             "GROUP BY ACC.INFOS"

                QRY_INFOS.Active = True

                If Not QRY_INFOS.EOF Then
                    Dim UPT_INFOS As Object
                    Set UPT_INFOS = NewQuery
                    UPT_INFOS.Active = False
                    UPT_INFOS.Clear

                    UPT_INFOS.Add "UPDATE VM_PNRACCOUNTINGS SET INFOS='" & QRY_INFOS.FieldByName("INFOS").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_INFOS.ExecSQL
                    Set UPT_INFOS = Nothing
                End If

                Set QRY_INFOS = Nothing
            End If
        End If

        'K9_UNIDADE
        If qry_vendas.FieldByName("K9_UNIDADE").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_UNIDADE").AsInteger = 93 Then
                Dim QRY_K9_UNIDADE As Object
                Set QRY_K9_UNIDADE = NewQuery
                QRY_K9_UNIDADE.Active = False
                QRY_K9_UNIDADE.Clear

                QRY_K9_UNIDADE.Add "SELECT TOP 1 ACC.K9_UNIDADE FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_UNIDADE IS NOT NULL " & _
                             "GROUP BY ACC.K9_UNIDADE"

                QRY_K9_UNIDADE.Active = True

                If Not QRY_K9_UNIDADE.EOF Then
                    Dim UPT_K9_UNIDADE As Object
                    Set UPT_K9_UNIDADE = NewQuery
                    UPT_K9_UNIDADE.Active = False
                    UPT_K9_UNIDADE.Clear

                    UPT_K9_UNIDADE.Add "UPDATE VM_PNRACCOUNTINGS SET K9_UNIDADE='" & QRY_K9_UNIDADE.FieldByName("K9_UNIDADE").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_K9_UNIDADE.ExecSQL
                    Set UPT_K9_UNIDADE = Nothing
                End If

                Set QRY_K9_UNIDADE = Nothing
            End If
        End If


		'Ajusta BRAK1
		If qry_vendas.FieldByName("K9_BREAK1").AsString = "" Then
			If qry_campos_obg.FieldByName("K9_BREAK1").AsInteger = 93 Then
				Contador=Contador+1
				Dim qry_bk1 As Object
				Set qry_bk1 = NewQuery
				qry_bk1.Active = False
				qry_bk1.Clear

				qry_bk1.Add ("SELECT TOP 1 ACC.K9_BREAK1, COUNT(ACC.HANDLE) AS QTDE "+ _
							  "FROM VM_PNRACCOUNTINGS ACC "+ _
							  "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) "+ _
							  "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 AND ACC.PASSAGEIRONAOCAD='"+qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString+"'"+ _
							  "AND PNR.CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							  "AND ACC.K9_BREAK1 IS NOT NULL "+ _
							  "GROUP BY ACC.K9_BREAK1 "+ _
							  "ORDER BY QTDE DESC")

				qry_bk1.Active = True

				If Not qry_bk1.EOF () Then

					Contador=Contador+1
					Dim upt_bk1 As Object
					Set upt_bk1 = NewQuery

					On Error GoTo ER0_1578
					If Not InTransaction Then StartTransaction

					upt_bk1.Active = False
					upt_bk1.Clear

					upt_bk1.Add("UPDATE VM_PNRACCOUNTINGS SET K9_BREAK1="+ qry_bk1.FieldByName("K9_BREAK1").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

					upt_bk1.ExecSQL

					If InTransaction Then Commit
					Set upt_bk1 = Nothing

					ER0_1578:
					  If InTransaction Then Rollback
					  Set upt_bk1 = Nothing
				End If
			End If
		End If

		'Break2
		If qry_vendas.FieldByName("K9_BREAK2").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_BREAK2").AsInteger = 93 Then
                Dim qry_bk2 As Object
                Set qry_bk2 = NewQuery
                qry_bk2.Active = False
                qry_bk2.Clear

                qry_bk2.Add "SELECT TOP 1 ACC.K9_BREAK2 FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_BREAK2 IS NOT NULL " & _
                             "GROUP BY ACC.K9_BREAK2"

                qry_bk2.Active = True

                If Not qry_bk2.EOF Then
                    Dim upt_bk2 As Object
                    Set upt_bk2 = NewQuery
                    upt_bk2.Active = False
                    upt_bk2.Clear

                    upt_bk2.Add "UPDATE VM_PNRACCOUNTINGS SET K9_BREAK2='" & qry_bk2.FieldByName("K9_BREAK2").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_bk2.ExecSQL
                    Set upt_bk2 = Nothing
                End If

                Set upt_bk2 = Nothing
            End If
        End If


		'BRAK3
		If qry_vendas.FieldByName("K9_BREAK3").AsString = "" Then
			If qry_campos_obg.FieldByName("K9_BREAK3").AsInteger = 93 Then
				Contador=Contador+1

				Dim qry_bk3 As Object
				Set qry_bk3 = NewQuery
				qry_bk3.Active = False
				qry_bk3.Clear

				qry_bk3.Add ("SELECT TOP 1 ACC.K9_BREAK3, COUNT(ACC.HANDLE) AS QTDE "+ _
							  "FROM VM_PNRACCOUNTINGS ACC "+ _
							  "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) "+ _
							  "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 AND ACC.PASSAGEIRONAOCAD='"+qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString+"' "+ _
							  "AND PNR.CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							  "AND ACC.K9_BREAK3 IS NOT NULL "+ _
							  "GROUP BY ACC.K9_BREAK3 "+ _
							  "ORDER BY QTDE DESC")

				qry_bk3.Active = True

				If Not qry_bk3.EOF () Then

					Contador=Contador+1
					Dim upt_bk3 As Object
					Set upt_bk3 = NewQuery

					On Error GoTo ER0_151
					If Not InTransaction Then StartTransaction

					upt_bk3.Active = False
					upt_bk3.Clear

					upt_bk3.Add("UPDATE VM_PNRACCOUNTINGS SET K9_BREAK3="+ qry_bk3.FieldByName("K9_BREAK3").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

					upt_bk3.ExecSQL

					If InTransaction Then Commit
					Set upt_bk3 = Nothing

					ER0_151:
					  If InTransaction Then Rollback
					  Set upt_bk3 = Nothing
				End If
			End If
		End If

		'Ajusta UDID1
		If qry_vendas.FieldByName("K9_UDID1").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_UDID1").AsInteger = 93 Then
                Dim qry_ud1 As Object
                Set qry_ud1 = NewQuery
                qry_ud1.Active = False
                qry_ud1.Clear

                qry_ud1.Add "SELECT TOP 1 ACC.K9_UDID1 FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_UDID1 IS NOT NULL " & _
                             "GROUP BY ACC.K9_UDID1"

                qry_ud1.Active = True

                If Not qry_ud1.EOF Then
                    Dim upt_UD1 As Object
                    Set upt_UD1 = NewQuery
                    upt_UD1.Active = False
                    upt_UD1.Clear

                    upt_UD1.Add "UPDATE VM_PNRACCOUNTINGS SET K9_UDID1='" & qry_ud1.FieldByName("K9_UDID1").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_UD1.ExecSQL
                    Set upt_UD1 = Nothing
                End If

                Set upt_UD1 = Nothing
            End If
        End If

		'UDID2
		If qry_vendas.FieldByName("K9_UDID2").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_UDID2").AsInteger = 93 Then
                Dim qry_ud2 As Object
                Set qry_ud2 = NewQuery
                qry_ud2.Active = False
                qry_ud2.Clear

                qry_ud2.Add "SELECT TOP 1 ACC.K9_UDID2 FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_UDID2 IS NOT NULL " & _
                             "GROUP BY ACC.K9_UDID2"

                qry_ud2.Active = True

                If Not qry_ud2.EOF Then
                    Dim upt_ud2 As Object
                    Set upt_ud2 = NewQuery
                    upt_ud2.Active = False
                    upt_ud2.Clear

                    upt_ud2.Add "UPDATE VM_PNRACCOUNTINGS SET K9_UDID2='" & qry_ud2.FieldByName("K9_UDID2").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_ud2.ExecSQL
                    Set upt_ud2 = Nothing
                End If

                Set qry_ud2 = Nothing
            End If
        End If

		'Ajusta UDID3
		If qry_vendas.FieldByName("K9_UDID3").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_UDID3").AsInteger = 93 Then
                Dim qry_ud3 As Object
                Set qry_ud3 = NewQuery
                qry_ud3.Active = False
                qry_ud3.Clear

                qry_ud3.Add "SELECT TOP 1 ACC.K9_UDID3 FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_UDID3 IS NOT NULL " & _
                             "GROUP BY ACC.K9_UDID3"

                qry_ud3.Active = True

                If Not qry_ud3.EOF Then
                    Dim upt_ud3 As Object
                    Set upt_ud3 = NewQuery
                    upt_ud3.Active = False
                    upt_ud3.Clear

                    upt_ud3.Add "UPDATE VM_PNRACCOUNTINGS SET K9_UDID3='" & qry_ud3.FieldByName("K9_UDID3").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_ud3.ExecSQL
                    Set upt_ud3 = Nothing
                End If

                Set upt_ud3 = Nothing
            End If
        End If

		'Ajusta CARGO
		If qry_vendas.FieldByName("K9_CARGO").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_CARGO").AsInteger = 93 Then
                Dim qry_cargo As Object
                Set qry_cargo = NewQuery
                qry_cargo.Active = False
                qry_cargo.Clear

                qry_cargo.Add "SELECT TOP 1 ACC.K9_CARGO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_CARGO IS NOT NULL " & _
                             "GROUP BY ACC.K9_CARGO"

                qry_cargo.Active = True

                If Not qry_cargo.EOF Then
                    Dim upt_carg As Object
                    Set upt_carg = NewQuery
                    upt_carg.Active = False
                    upt_carg.Clear

                    upt_carg.Add "UPDATE VM_PNRACCOUNTINGS SET K9_CARGO='" & qry_cargo.FieldByName("K9_CARGO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_carg.ExecSQL
                    Set upt_carg = Nothing
                End If

                Set qry_cargo = Nothing
            End If
        End If

		'Ajusta DIVISAO
		If qry_vendas.FieldByName("INFDIVISAO").AsString = "" Then
            If qry_campos_obg.FieldByName("DIVISAO").AsInteger = 3 Then
                Dim qry_div As Object
                Set qry_div = NewQuery
                qry_div.Active = False
                qry_div.Clear

                qry_div.Add "SELECT TOP 1 ACC.INFDIVISAO FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFDIVISAO IS NOT NULL " & _
                             "GROUP BY ACC.INFDIVISAO"

                qry_div.Active = True

                If Not qry_div.EOF Then
                    Dim upt_divpax As Object
                    Set upt_divpax = NewQuery
                    upt_divpax.Active = False
                    upt_divpax.Clear

                    upt_divpax.Add "UPDATE VM_PNRACCOUNTINGS SET INFDIVISAO='" & qry_div.FieldByName("INFDIVISAO").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_divpax.ExecSQL
                    Set upt_divpax = Nothing
                End If

                Set qry_div = Nothing
            End If
        End If

        If qry_vendas.FieldByName("INFCFI").AsString = "" Then
            If qry_campos_obg.FieldByName("CFI").AsInteger = 3 Then
                Dim QRY_INFCFI As Object
                Set QRY_INFCFI = NewQuery
                QRY_INFCFI.Active = False
                QRY_INFCFI.Clear

                QRY_INFCFI.Add "SELECT TOP 1 ACC.INFCFI FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFCFI IS NOT NULL " & _
                             "GROUP BY ACC.INFCFI"

                QRY_INFCFI.Active = True

                If Not QRY_INFCFI.EOF Then
                    Dim UPT_INFCFI As Object
                    Set UPT_INFCFI = NewQuery
                    UPT_INFCFI.Active = False
                    UPT_INFCFI.Clear

                    UPT_INFCFI.Add "UPDATE VM_PNRACCOUNTINGS SET INFCFI='" & QRY_INFCFI.FieldByName("INFCFI").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    UPT_INFCFI.ExecSQL
                    Set UPT_INFCFI = Nothing
                End If

                Set QRY_INFCFI = Nothing
            End If
        End If

		'Ajusta DIVISAO PAX
		If qry_vendas.FieldByName("K9_DIVISAOPAX").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_DIVISAOPAX").AsInteger = 93 Then
                Dim qry_divpax As Object
                Set qry_divpax = NewQuery
                qry_divpax.Active = False
                qry_divpax.Clear

                qry_divpax.Add "SELECT TOP 1 ACC.K9_DIVISAOPAX FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_DIVISAOPAX IS NOT NULL " & _
                             "GROUP BY ACC.K9_DIVISAOPAX"

                qry_divpax.Active = True

                If Not qry_divpax.EOF Then
                    Dim upt_div As Object
                    Set upt_div = NewQuery
                    upt_div.Active = False
                    upt_div.Clear

                    upt_div.Add "UPDATE VM_PNRACCOUNTINGS SET K9_DIVISAOPAX='" & qry_divpax.FieldByName("K9_DIVISAOPAX").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_div.ExecSQL
                    Set upt_div = Nothing
                End If

                Set qry_divpax = Nothing
            End If
        End If


		'Ajusta MATRICULA
		If qry_vendas.FieldByName("MATRICULA").AsString = "" Then
			If qry_campos_obg.FieldByName("MATRICULA").AsInteger = 3 Then
				Contador=Contador+1

				Dim qry_mat As Object
				Set qry_mat = NewQuery
				qry_mat.Active = False
				qry_mat.Clear

				qry_mat.Add ("SELECT TOP 1 ACC.MATRICULA, COUNT(ACC.HANDLE) AS QTDE "+ _
							  "FROM VM_PNRACCOUNTINGS ACC "+ _
							  "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) "+ _
							  "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 AND ACC.PASSAGEIRONAOCAD='"+qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString+"' "+ _
							  "AND PNR.CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							  "AND ACC.MATRICULA IS NOT NULL "+ _
							  "GROUP BY ACC.MATRICULA "+ _
							  "ORDER BY QTDE DESC")

				qry_mat.Active = True

				If Not qry_mat.EOF () Then

					Contador=Contador+1
					Dim upt_mat As Object
					Set upt_mat = NewQuery

					On Error GoTo ER0_156
					If Not InTransaction Then StartTransaction

					upt_mat.Active = False
					upt_mat.Clear

					upt_mat.Add("UPDATE VM_PNRACCOUNTINGS SET MATRICULA="+ qry_mat.FieldByName("MATRICULA").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

					upt_mat.ExecSQL

					If InTransaction Then Commit
					Set upt_mat = Nothing

					ER0_156:
					  If InTransaction Then Rollback
					  Set upt_mat = Nothing
				End If
			End If
		End If

		'Ajusta APROVADOR
		If qry_vendas.FieldByName("INFAPROVADOR").AsString = "" Then
            If qry_campos_obg.FieldByName("APROVADOR").AsInteger = 3 Then
                Dim qry_apv As Object
                Set qry_apv = NewQuery
                qry_apv.Active = False
                qry_apv.Clear

                qry_apv.Add "SELECT TOP 1 ACC.INFAPROVADOR FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.INFAPROVADOR IS NOT NULL " & _
                             "GROUP BY ACC.INFAPROVADOR"

                qry_apv.Active = True

                If Not qry_apv.EOF Then
                    Dim upt_apv As Object
                    Set upt_apv = NewQuery
                    upt_apv.Active = False
                    upt_apv.Clear

                    upt_apv.Add "UPDATE VM_PNRACCOUNTINGS SET INFAPROVADOR='" & qry_apv.FieldByName("INFAPROVADOR").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_apv.ExecSQL
                    Set upt_apv = Nothing
                End If

                Set qry_apv = Nothing
            End If
        End If
		'Reason Code
		If qry_vendas.FieldByName("K9_REASONCODE").AsInteger = 0 Then
			If qry_campos_obg.FieldByName("K9_REASONCODE").AsInteger = 93 Then

				Dim qry_reason As Object
				Set qry_reason = NewQuery
				qry_reason.Active = False
				qry_reason.Clear

				qry_reason.Add ("SELECT TOP 1 ACC.K9_REASONCODE, COUNT(ACC.HANDLE) AS QTDE "+ _
							  "FROM VM_PNRACCOUNTINGS ACC "+ _
							  "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) "+ _
							  "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 AND ACC.PASSAGEIRONAOCAD='"+qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString+"' "+ _
							  "AND PNR.CLIENTE="+qry_vendas.FieldByName("HANDLE_CLI").AsInteger+" "+ _
							  "AND ACC.K9_REASONCODE IS NOT NULL "+ _
							  "GROUP BY ACC.K9_REASONCODE "+ _
							  "ORDER BY QTDE DESC")

				qry_reason.Active = True

				If Not qry_reason.EOF () Then

					Contador=Contador+1
					Dim upt_reason As Object
					Set upt_reason = NewQuery

					On Error GoTo ER0_158
					If Not InTransaction Then StartTransaction

					upt_reason.Active = False
					upt_reason.Clear

					upt_reason.Add("UPDATE VM_PNRACCOUNTINGS SET K9_REASONCODE="+ qry_reason.FieldByName("K9_REASONCODE").AsInteger +" WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_ACC").AsInteger+"")

					upt_reason.ExecSQL

					If InTransaction Then Commit
					Set upt_reason = Nothing

					ER0_158:
					  If InTransaction Then Rollback
					  Set upt_reason = Nothing
				End If
			End If
		End If

		'Finalidade pré definida
		If qry_vendas.FieldByName("K9_FINALIDADEPREDEFINIDA").AsString = "" Then
            If qry_campos_obg.FieldByName("K9_FINALIDADEPREDEFINIDA").AsInteger = 93 Then
                Dim QRY_FINP As Object
                Set QRY_FINP = NewQuery
                QRY_FINP.Active = False
                QRY_FINP.Clear

                QRY_FINP.Add "SELECT TOP 1 ACC.K9_FINALIDADEPREDEFINIDA FROM VM_PNRACCOUNTINGS ACC " & _
                             "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " & _
                             "WHERE PNR.DATAINCLUSAO>=GETDATE()-45 " & _
                             "AND ACC.PASSAGEIRONAOCAD='" & qry_vendas.FieldByName("PASSAGEIRONAOCAD").AsString & "' " & _
                             "AND PNR.CLIENTE=" & qry_vendas.FieldByName("HANDLE_CLI").AsInteger & " " & _
                             "AND ACC.K9_FINALIDADEPREDEFINIDA IS NOT NULL " & _
                             "GROUP BY ACC.K9_FINALIDADEPREDEFINIDA"

                QRY_FINP.Active = True

                If Not QRY_FINP.EOF Then
                    Dim upt_finp As Object
                    Set upt_finp = NewQuery
                    upt_finp.Active = False
                    upt_finp.Clear

                    upt_finp.Add "UPDATE VM_PNRACCOUNTINGS SET K9_FINALIDADEPREDEFINIDA='" & QRY_FINP.FieldByName("K9_FINALIDADEPREDEFINIDA").AsString & "' " & _
                                "WHERE HANDLE = " & qry_vendas.FieldByName("HANDLE_ACC").AsInteger

                    upt_finp.ExecSQL
                    Set upt_finp = Nothing
                End If

                Set QRY_FINP = Nothing
            End If
        End If
		'Coloca a venda no agendamento novamente
		If Contador>0 Then
			Dim qryUPDT15 As Object
			Set qryUPDT15 = NewQuery

			On Error GoTo ER15
			If Not InTransaction Then StartTransaction

			qryUPDT15.Active = False
			qryUPDT15.Clear

				qryUPDT15.Add("UPDATE VM_PNRS SET SITUACAO=3, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ qry_vendas.FieldByName("HANDLE_PNR").AsInteger+"")

			qryUPDT15.ExecSQL

			If InTransaction Then Commit
			Set qryUPDT15 = Nothing

			ER15:
			  If InTransaction Then Rollback
			  Set qryUPDT15 = Nothing
		End If

		Contador=0
		NovaTaxa=Empty
		Titular=Empty

		qry_vendas.Next
	Wend

End Sub

