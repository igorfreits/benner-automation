Public Sub Main()

    Dim QRY_bk3_cc As Object
    Dim SQL_bk3_cc As Object
	Dim UPDT_bk3_cc As Object

	Dim QRY_bk2_cc As Object
	Dim SQL_bk2_cc As Object
    Dim UPDT_bk2_cc As Object

	'UPDATE Centro de custo - BK3
    Set QRY_bk3_cc = NewQuery
    QRY_bk3_cc.Active = False
    QRY_bk3_cc.Clear

    QRY_bk3_cc.Add("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, ACC.K9_BREAK3 BATATA, LOCALIZADORA, PNR.DATAINCLUSAO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (68436,71724)) " + _
               "AND ACC.K9_BREAK3 IS NOT NULL "+ _
               "AND (ACC.CENTRODECUSTO IS NULL OR ACC.CENTRODECUSTO IN (0))")

    QRY_bk3_cc.Active = True

    While Not QRY_bk3_cc.EOF
        Set SQL_bk3_cc = NewQuery
            SQL_bk3_cc.Active = False
            SQL_bk3_cc.Clear

            SQL_bk3_cc.Add("SELECT CENTROCUSTO, DESCRICAO, HANDLE, CODIGOINTERNO, VALIDO " + _
             "FROM BB_CLIENTECC " + _
             "WHERE CENTROCUSTO = '"+ QRY_bk3_cc.FieldByName("BATATA").AsString +"' " + _
             "AND DESCRICAO = '"+ QRY_bk3_cc.FieldByName("BATATA").AsString +"' " + _
             "AND VALIDO = 'S' " + _
             "AND BB_CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (68436,71724))")

            SQL_bk3_cc.Active = True

                    Set UPDT_bk3_cc = NewQuery
                    UPDT_bk3_cc.Active = False
                    UPDT_bk3_cc.Clear

                    UPDT_bk3_cc.Add("UPDATE VM_PNRACCOUNTINGS SET CENTRODECUSTO = '" + SQL_bk3_cc.FieldByName("CODIGOINTERNO").AsString + "' WHERE HANDLE = " + QRY_bk3_cc.FieldByName("HANDLE_ACC").AsString)

                    UPDT_bk3_cc.ExecSQL


                    Dim qryUPDT_1 As Object
                    Set qryUPDT_1 = NewQuery
                    qryUPDT_1.Active = False
                    qryUPDT_1.Clear

                    qryUPDT_1.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_bk3_cc.FieldByName("HANDLE_PNR").AsInteger+"")

                    qryUPDT_1.ExecSQL

                    QRY_bk3_cc.Next

    Wend

	'UPDATE Centro de custo - BK2
	Set QRY_bk2_cc = NewQuery
    QRY_bk2_cc.Active = False
    QRY_bk2_cc.Clear

    QRY_bk2_cc.Add("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, ACC.K9_BREAK2 BATATA, LOCALIZADORA, PNR.DATAINCLUSAO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (68873,71847,70215,73869,74512,68873)) " + _
               "AND ACC.K9_BREAK2 IS NOT NULL "+ _
               "AND (ACC.CENTRODECUSTO IS NULL OR ACC.CENTRODECUSTO IN (0))")

    QRY_bk2_cc.Active = True

    While Not QRY_bk2_cc.EOF
        Set SQL_bk2_cc = NewQuery
            SQL_bk2_cc.Active = False
            SQL_bk2_cc.Clear

            SQL_bk2_cc.Add("SELECT CENTROCUSTO, DESCRICAO, HANDLE, CODIGOINTERNO, VALIDO " + _
             "FROM BB_CLIENTECC " + _
             "WHERE CENTROCUSTO = '"+ QRY_bk2_cc.FieldByName("BATATA").AsString +"' " + _
             "AND DESCRICAO = '"+ QRY_bk2_cc.FieldByName("BATATA").AsString +"' " + _
             "AND VALIDO = 'S' " + _
             "AND BB_CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (68873,71847,70215,73869,74512,68873))")

            SQL_bk2_cc.Active = True

            Dim qryUPDT_2 As Object
            Set qryUPDT_2 = NewQuery
            qryUPDT_2.Active = False
            qryUPDT_2.Clear

            qryUPDT_2.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_bk2_cc.FieldByName("HANDLE_PNR").AsInteger+"")

            qryUPDT_2.ExecSQL

            QRY_bk2_cc.Next

	Wend

	'UPDATE Departamento/Empenho - BK3
	Dim QRY_bk3_dept As Object
	Set QRY_bk3_dept = NewQuery
	QRY_bk3_dept.Active = False
	QRY_bk3_dept.Clear

	QRY_bk3_dept.Add ("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, ACC.K9_BREAK3, LOCALIZADORA, PNR.DATAINCLUSAO,ACC.MATRICULA, ACC.EMPENHO,ACC.INFNIVELFUNCIONARIO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (49972,49972)) " + _
               "AND ACC.K9_BREAK3 IS NOT NULL " + _
               "AND ACC.EMPENHO IS NOT NULL")

	QRY_bk3_dept.Active = True

	While Not QRY_bk3_dept.EOF()

		Dim UPT_bk3_dept As Object
		Set UPT_bk3_dept = NewQuery

		UPT_bk3_dept.Active = False
		UPT_bk3_dept.Clear

		UPT_bk3_dept.Add("UPDATE VM_PNRACCOUNTINGS SET EMPENHO ='"+ QRY_bk3_dept.FieldByName("K9_BREAK3").AsString +"' WHERE HANDLE ="+ QRY_bk3_dept.FieldByName("HANDLE_ACC").AsInteger+"")

		UPT_bk3_dept.ExecSQL

                    Dim qryUPDT_3 As Object
                    Set qryUPDT_3 = NewQuery
                    qryUPDT_3.Active = False
                    qryUPDT_3.Clear

                    qryUPDT_3.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_bk3_dept.FieldByName("HANDLE_PNR").AsInteger+"")

                    qryUPDT_3.ExecSQL

                    QRY_bk3_dept.Next


	Wend

	'UPDATE Departamento/Empenho - BK2
	Dim QRY_bk2_dept As Object
	Set QRY_bk2_dept = NewQuery
	QRY_bk2_dept.Active = False
	QRY_bk2_dept.Clear

	QRY_bk2_dept.Add ("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, ACC.K9_BREAK2, LOCALIZADORA, PNR.DATAINCLUSAO,ACC.MATRICULA, ACC.EMPENHO,ACC.INFNIVELFUNCIONARIO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (68436,71724)) " + _
               "AND ACC.K9_BREAK2 IS NOT NULL " + _
               "AND ACC.EMPENHO IS NOT NULL")

	QRY_bk2_dept.Active = True

	While Not QRY_bk2_dept.EOF()

		Dim UPT_bk2_dept As Object
		Set UPT_bk2_dept = NewQuery

		UPT_bk2_dept.Active = False
		UPT_bk2_dept.Clear

		UPT_bk2_dept.Add("UPDATE VM_PNRACCOUNTINGS SET EMPENHO ='"+ QRY_bk2_dept.FieldByName("K9_BREAK2").AsString +"' WHERE HANDLE ="+ QRY_bk2_dept.FieldByName("HANDLE_ACC").AsInteger+"")

		UPT_bk2_dept.ExecSQL

                    Dim qryUPDT_4 As Object
                    Set qryUPDT_4 = NewQuery
                    qryUPDT_4.Active = False
                    qryUPDT_4.Clear

                    qryUPDT_4.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_bk2_dept.FieldByName("HANDLE_PNR").AsInteger+"")

                    qryUPDT_4.ExecSQL

                    QRY_bk2_dept.Next


	Wend

	'UPDATE Matricula - BK1
	Dim QRY_bk1_matr As Object
	Set QRY_bk1_matr = NewQuery
	QRY_bk1_matr.Active = False
	QRY_bk1_matr.Clear

	QRY_bk1_matr.Add ("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, ACC.K9_BREAK1, LOCALIZADORA, PNR.DATAINCLUSAO,ACC.MATRICULA, ACC.EMPENHO,ACC.INFNIVELFUNCIONARIO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (71847,70215,73869,71724,74512)) " + _
               "AND ACC.K9_BREAK1 IS NOT NULL " + _
               "AND ACC.MATRICULA IS NOT NULL")

	QRY_bk1_matr.Active = True

	While Not QRY_bk1_matr.EOF()

		Dim UPT_bk1_matr As Object
		Set UPT_bk1_matr = NewQuery

		UPT_bk1_matr.Active = False
		UPT_bk1_matr.Clear

		UPT_bk1_matr.Add("UPDATE VM_PNRACCOUNTINGS SET MATRICULA ='"+ QRY_bk1_matr.FieldByName("K9_BREAK1").AsString +"' WHERE HANDLE ="+ QRY_bk1_matr.FieldByName("HANDLE_ACC").AsInteger+"")

		UPT_bk1_matr.ExecSQL

                    Dim qryUPDT_5 As Object
                    Set qryUPDT_5 = NewQuery
                    qryUPDT_5.Active = False
                    qryUPDT_5.Clear

                    qryUPDT_5.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_bk1_matr.FieldByName("HANDLE_PNR").AsInteger+"")

                    qryUPDT_5.ExecSQL

                    QRY_bk1_matr.Next

	Wend

	'UPDATE Matricula - BK3
	Dim QRY_bk3_matr As Object
	Set QRY_bk3_matr = NewQuery
	QRY_bk3_matr.Active = False
	QRY_bk3_matr.Clear

	QRY_bk3_matr.Add ("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, ACC.K9_BREAK3, LOCALIZADORA, PNR.DATAINCLUSAO,ACC.MATRICULA, ACC.EMPENHO, ACC.INFNIVELFUNCIONARIO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (49972,49972)) " + _
               "AND ACC.K9_BREAK3 IS NOT NULL " + _
               "AND ACC.MATRICULA IS NOT NULL")

	QRY_bk3_matr.Active = True

	While Not QRY_bk3_matr.EOF()

		Dim UPT_bk3_matr As Object
		Set UPT_bk3_matr = NewQuery

		UPT_bk3_matr.Active = False
		UPT_bk3_matr.Clear

		UPT_bk3_matr.Add("UPDATE VM_PNRACCOUNTINGS SET MATRICULA ='"+ QRY_bk3_matr.FieldByName("K9_BREAK3").AsString +"' WHERE HANDLE ="+ QRY_bk3_matr.FieldByName("HANDLE_ACC").AsInteger+"")

		UPT_bk3_matr.ExecSQL

                    Dim qryUPDT_6 As Object
                    Set qryUPDT_6 = NewQuery
                    qryUPDT_6.Active = False
                    qryUPDT_6.Clear

                    qryUPDT_6.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_bk3_matr.FieldByName("HANDLE_PNR").AsInteger+"")

                    qryUPDT_6.ExecSQL

                    QRY_bk3_matr.Next


	Wend

	'UPDATE Nivel Funcionario - BK3
	Dim QRY_bk3_nv_func As Object
	Set QRY_bk3_nv_func = NewQuery
	QRY_bk3_nv_func.Active = False
	QRY_bk3_nv_func.Clear

	QRY_bk3_nv_func.Add ("SELECT PNR.HANDLE AS HANDLE_PNR, ACC.HANDLE AS HANDLE_ACC, ACC.K9_BREAK3, LOCALIZADORA, PNR.DATAINCLUSAO,ACC.MATRICULA, ACC.EMPENHO, ACC.INFNIVELFUNCIONARIO " + _
               "FROM VM_PNRACCOUNTINGS ACC " + _
               "LEFT JOIN VM_PNRS PNR ON (PNR.HANDLE=ACC.PNR) " + _
               "WHERE SITUACAO IN (3) " + _
               "AND CLIENTE IN (SELECT HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL IN (73869)) " + _
               "AND ACC.K9_BREAK3 IS NOT NULL " + _
               "AND ACC.INFNIVELFUNCIONARIO IS NOT NULL")

	QRY_bk3_nv_func.Active = True

	While Not QRY_bk3_nv_func.EOF()

		Dim UPT_bk3_nv_func As Object
		Set UPT_bk3_nv_func = NewQuery

		UPT_bk3_nv_func.Active = False
		UPT_bk3_nv_func.Clear

		UPT_bk3_nv_func.Add("UPDATE VM_PNRACCOUNTINGS SET INFNIVELFUNCIONARIO ='"+ QRY_bk3_nv_func.FieldByName("K9_BREAK3").AsString +"' WHERE HANDLE ="+ QRY_bk3_nv_func.FieldByName("HANDLE_ACC").AsInteger+"")

		UPT_bk3_nv_func.ExecSQL

                    Dim qryUPDT_7 As Object
                    Set qryUPDT_7 = NewQuery
                    qryUPDT_7.Active = False
                    qryUPDT_7.Clear

                    qryUPDT_7.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_bk3_nv_func.FieldByName("HANDLE_PNR").AsInteger+"")

                    qryUPDT_7.ExecSQL

                    QRY_bk3_nv_func.Next


	Wend


End Sub
