
Public Sub MAIN

	Dim QRY_GTNE As Object
	Set QRY_GTNE = NewQuery
	QRY_GTNE.Active = False
	QRY_GTNE.Clear

	QRY_GTNE.Add("Select PNR.HANDLE HANDLE_PNR, ACC.HANDLE HANDLE_ACC, PNR.DATAINCLUSAO,PNR.LOCALIZADORA,ACC.REQUISICAO,PNR.TIPORESERVA, " + _
                "Case WHEN CHARINDEX('<CodigoGTNE>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) > 0 " + _
                "Then LTRIM(RTRIM( " + _
                "SUBSTRING(CONVERT(VARCHAR(MAX), Log.XMLRESERVA), " + _
                "CHARINDEX('<CodigoGTNE>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CodigoGTNE>'), " + _
                "CHARINDEX('</CodigoGTNE>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) - (CHARINDEX('<CodigoGTNE>', CONVERT(VARCHAR(MAX), LOG.XMLRESERVA)) + LEN('<CodigoGTNE>')) " + _
                "))) " + _
                "Else '' " + _
                "End As CODE_GTNE " + _
                "FROM VM_PNRS PNR " + _
                "Left Join VM_PNRACCOUNTINGS ACC On ACC.PNR = PNR.Handle " + _
                "Left Join BB_LOGINTEGRACOES Log On Log.Handle = PNR.LOGINTEGRACAO " + _
                "WHERE CAST(PNR.DATAINCLUSAO As Date) = CAST(GETDATE() -1 As Date) " + _
                "And PNR.TIPORESERVA In (27) And ACC.INFOS Is Null And PNR.CLIENTE In (Select HANDLE FROM GN_PESSOAS WHERE GRUPOEMPRESARIAL = 70610)")

	QRY_GTNE.Active = True

	While Not QRY_GTNE.EOF()

		Dim UPT_GTNE As Object
		Set UPT_GTNE = NewQuery

		UPT_GTNE.Active = False
		UPT_GTNE.Clear

		UPT_GTNE.Add("UPDATE VM_PNRACCOUNTINGS SET INFOS ='"+ QRY_GTNE.FieldByName("CODE_GTNE").AsString +"' WHERE HANDLE ="+ QRY_GTNE.FieldByName("HANDLE_ACC").AsInteger+"")

		UPT_GTNE.ExecSQL

                    Dim qryUPDT_3 As Object
                    Set qryUPDT_3 = NewQuery
                    qryUPDT_3.Active = False
                    qryUPDT_3.Clear

                    qryUPDT_3.Add("UPDATE VM_PNRS SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' WHERE HANDLE ="+ QRY_GTNE.FieldByName("HANDLE_PNR").AsInteger+"")

                    qryUPDT_3.ExecSQL

                    QRY_GTNE.Next


	Wend
End Sub
