<%
Class Authorize

	Private cmd, rs

	Private pApproved
	Public Property Get Approved
		Approved = pApproved
	End Property

	Private pTransactionID
	Private pOrderNumber
	Private pAuthorizeID
	Private pResponseCode
	Private pReceiptType
	Private pReceiptDestination
	Private pCustomerEmail
	Private pOriginalAmount
	Private pOriginalSalesTax
	Private pOriginalType
	Private pOriginalID
	Private pCustomerID
	Private pPONumber
	Private pCardNumber
	Private pFirstName
	Private pLastName
	Private pCompanyName
	Private pAddress
	Private pCity
	Private pState
	Private pZip
	Private pExpireDate
	Private pReferenceCode
	Private pPostToCode
	Private pCompanyDivision
	Private pSalespersonCode
	Private pCustomerType
	Private pCustomerSubType
	Private pPurgeClass
	Private pReleaseNumber
	Private pRecordNumber
	Public Property Get TransactionID
		TransactionID = pTransactionID
	End Property
	Public Property Let TransactionID(value)
		Call LoadTransaction(value)
	End Property

	Private pCardID
	Public Property Get CardID
		CardID = pCardID
	End Property
	Public Property Let CardID(value)
		If value <> "" And IsNumeric(value) Then
			cmd.CommandText = _
				"SELECT CreditCardID " & _
				"FROM CreditCards " & _
				"WHERE CreditCardID = ?; "
			Call DeleteAllParameters(cmd)
			cmd.Parameters.Append cmd.CreateParameter("CardID", adInteger, adParamInput, , value)
			Set rs = cmd.Execute()
			If Not rs.EOF Then
				pCardID = rs("CreditCardID")
			End If
			rs.Close
			Set rs = Nothing
		End If
	End Property

	Private pAmount
	Public Property Get Amount
		Amount = pAmount
	End Property
	Public Property Let Amount(value)
		If value <> "" And IsNumeric(value) Then
			If value <> 0 Then
				pAmount = Abs(value)
			End If
		End If
	End Property

	Private pSalesTax
	Public Property Get SalesTax
		SalesTax = pSalesTax
	End Property
	Public Property Let SalesTax(value)
		If value <> "" And IsNumeric(value) Then
			If value <> 0 Then
				pSalesTax = Abs(value)
			End If
		End If
	End Property

	Private pTransactionType
	Public Property Get TransactionType
		TransactionType = pTransactionType
	End Property
	Public Property Let TransactionType(value)
		Select Case value
			Case "VALIDATE" ,"AUTH_CAPTURE", "AUTH_ONLY", "PRIOR_AUTH_CAPTURE", "CREDIT", "VOID"
				pTransactionType = value
		End Select
	End Property

	Private Sub Class_Initialize()
		Set cmd = Server.CreateObject("ADODB.Command")
		ConnectToDb("Refron")
		cmd.ActiveConnection = conRefron
		pApproved = False
		pTransactionID = Null
		pOrderNumber = Null
		pAuthorizeID = Null
		pResponseCode = Null
		pReceiptType = Null
		pReceiptDestination = Null
		pCustomerEmail = Null
		pOriginalAmount = 0
		pOriginalSalesTax = Null
		pOriginalType = Null
		pOriginalID = Null
		pCustomerID = Null
		pPONumber = Null
		pCardNumber = Null
		pFirstName = Null
		pLastName = Null
		pCompanyName = Null
		pAddress = Null
		pCity = Null
		pState = Null
		pZip = Null
		pExpireDate = Null
		pCardID = Null
		pAmount = Null
		pSalesTax = Null
		pTransactionType = Null
		pReferenceCode = Null
		pPostToCode = Null
		pCompanyDivision = Null
		pSalespersonCode = Null
		pCustomerType = Null
		pCustomerSubType = Null
		pPurgeClass = Null
		pReleaseNumber = Null
		pRecordNumber = Null
	End Sub

	Private Sub Class_Terminate()
		Set cmd = Nothing
	End Sub

	Private Sub LoadTransaction(intID)
		If intID <> "" And IsNumeric(intID) Then
			cmd.CommandText = _
				"SELECT " & _
					"CreditCardTransactionID, " & _
					"CreditCardTransactionOrderNumber, " & _
					"CreditCardTransactionAuthorizeID, " & _
					"CreditCardTransactionResponseCode, " & _
					"CreditCardTransactionReceiptType, " & _
					"CreditCardTransactionReceiptDestination, " & _
					"CreditCardTransactionAmount AS OriginalAmount, " & _
					"CreditCardTransactionSalesTax AS OriginalSalesTax, " & _
					"CreditCardTransactionOriginalID, " & _
					"COALESCE(OriginalType, CreditCardTransactionType) AS OriginalType, " & _
					"CASE " & _
						"WHEN LTRIM(RTRIM(OHPO#)) = '' THEN NULL " & _
						"ELSE LTRIM(RTRIM(OHPO#)) " & _
					"END AS PONumber, " & _
					"ISNULL(OHCSCD, CreditCardAccountCode) AS AccountCode, " & _
					"CASE " & _
						"WHEN LTRIM(RTRIM(OHPOST)) = '' THEN NULL " & _
						"ELSE LTRIM(RTRIM(OHPOST)) " & _
					"END AS PostToCode, " & _
					"OHRLS AS ReleaseNumber, " & _
					"PostCompany.CMDIVN AS CompanyDivision, " & _
					"PostCompany.CMPURG AS PurgeClass, " & _
					"ISNULL(CustCompany.CMRFCD, PostCompany.CMRFCD) AS ReferenceCode, " & _
					"ISNULL(CustCompany.CMSLMN, PostCompany.CMSLMN) AS SalespersonCode, " & _
					"ISNULL(CustCompany.CMCSTY, PostCompany.CMCSTY) AS CustomerType, " & _
					"ISNULL(CustCompany.CMCSSB, PostCompany.CMCSSB) AS CustomerSubType, " & _
					"CreditCardID, " & _
					"CreditCardNumberEnc, " & _
					"CreditCardFirstName, " & _
					"CreditCardLastName, " & _
					"CreditCardCompanyName, " & _
					"CreditCardAddress, " & _
					"CreditCardCity, " & _
					"CreditCardState, " & _
					"CreditCardZip, " & _
					"CreditCardZip4, " & _
					"CreditCardExpireDate " & _
				"FROM CreditCardTransactions " & _
					"INNER JOIN CreditCards ON CreditCardTransactionCardID = CreditCardID " & _
					"LEFT JOIN ORDRHDR ON CreditCardTransactionOrderNumber = OHORNU " & _
					"LEFT JOIN CUSTMST AS CustCompany ON OHCSCD = CustCompany.CMCSCD AND OHCSCD <> '' " & _
					"LEFT JOIN CUSTMST AS PostCompany ON OHPOST = PostCompany.CMCSCD AND OHPOST <> '' " & _
					"LEFT JOIN " & _
						"(" & _
							"SELECT " & _
								"CreditCardTransactionAuthorizeID AS OriginalID, " & _
								"CreditCardTransactionType AS OriginalType " & _
							"FROM CreditCardTransactions " & _
							"WHERE CreditCardTransactionOriginalID IS NULL " & _
								"AND CreditCardTransactionResponseCode = 1 " & _
						") AS OriginalTransaction ON CreditCardTransactionOriginalID = OriginalID " & _
				"WHERE CreditCardTransactionID = ?; "
			Call DeleteAllParameters(cmd)
			cmd.Parameters.Append cmd.CreateParameter("TransactionID", adInteger, adParamInput, , intID)
			Set rs = cmd.Execute()
			If Not rs.EOF Then
				pTransactionID = rs("CreditCardTransactionID")
				pOrderNumber = rs("CreditCardTransactionOrderNumber")
				pAuthorizeID = rs("CreditCardTransactionAuthorizeID")
				pResponseCode = rs("CreditCardTransactionResponseCode")
				pReceiptType = rs("CreditCardTransactionReceiptType")
				pReceiptDestination = rs("CreditCardTransactionReceiptDestination")
				If Not IsNull(pReceiptDestination) Then
					Select Case pReceiptType
						Case "F"
'							pCustomerEmail = "1" & pReceiptDestination & "@fax.refron.com"
							pCustomerEmail = GetToAddress("", FixFaxNumber(Trim(pReceiptDestination)))
						Case "E"
							pCustomerEmail = pReceiptDestination
					End Select
				End If
				pOriginalAmount = rs("OriginalAmount")
				pOriginalSalesTax = rs("OriginalsalesTax")
				pOriginalType = rs("OriginalType")
				pOriginalID = rs("CreditCardTransactionOriginalID")
				pCustomerID = rs("AccountCode")
				pPONumber = rs("PONumber")
				pCardID = rs("CreditCardID")
				pCardNumber = rs("CreditCardNumberEnc")
				pFirstName = rs("CreditCardFirstName")
				pLastName = rs("CreditCardLastName")
				pCompanyName = rs("CreditCardCompanyName")
				pAddress = rs("CreditCardAddress")
				pCity = rs("CreditCardCity")
				pState = rs("CreditCardState")
				If Not IsNull(rs("CreditCardZip4")) Then
					pZip = PadZeros(rs("CreditCardZip"), 5) & "-" & PadZeros(rs("CreditCardZip4"), 4)
				Else
					pZip = PadZeros(rs("CreditCardZip"), 5)
				End If
				pExpireDate = PadZeros(Month(rs("CreditCardExpireDate")), 2) & Right(Year(rs("CreditCardExpireDate")), 2)
				pReferenceCode = rs("ReferenceCode")
				pPostToCode = rs("PostToCode")
				pCompanyDivision = rs("CompanyDivision")
				pSalespersonCode = rs("SalespersonCode")
				pCustomerType = rs("CustomerType")
				pCustomerSubType = rs("CustomerSubType")
				pPurgeClass = rs("PurgeClass")
				pReleaseNumber = rs("ReleaseNumber")
				pRecordNumber = Null
			End If
			rs.Close
			Set rs = Nothing
		End If
	End Sub

	Private Sub ProcessTransaction(strType, dblAmount, dblSalesTax)
		If Not IsNull(pTransactionID) _
			And Not IsNull(strType) _
			And Not IsNull(dblAmount) Then

			Dim strRequestText, strPostURL
			
			'--- Default values ---
			strRequestText = _
				"x_version=3.1" & _
				"&x_delim_data=True" & _
				"&x_delim_char=|" & _
				"&x_encap_char=" & _
				"&x_relay_response=False" & _
				"&x_country=US" & _
				"&x_currency_code=USD" & _
				"&x_method=CC" & _
                "&x_login=36Es2Gwk" & _
                "&x_tran_key=3b72jAn47EV6L3Kp" & _
				"&x_duplicate_window=30"

				'Old Refron Values - Before switch to new merchant account 3/6/07
				'"&x_login=Refron79" & _
				'"&x_tran_key=m9GHR8urV1VrjBx6" & _

                'New Refron Values - After switch to new merchant account 3/6/07
    			'"&x_login=5rz7PGCSV87" & _
				'"&x_tran_key=8Uq526B5gp5RWFwe" & _

                'New Airgas Values - After switch to Airgas on 8/1/08
                '"&x_login=36Es2Gwk" & _
                '"&x_tran_key=3b72jAn47EV6L3Kp" & _

			'--- Parameters ---
			strRequestText = strRequestText & _
				"&x_type=" & Server.URLEncode(strType) & _
				"&x_amount=" & Server.URLEncode(FormatCurrency(dblAmount, 2, True, False, True))
			If Not IsNull(dblSalesTax) Then
				strRequestText = strRequestText & _
					"&x_tax=" & Server.URLEncode(FormatCurrency(dblSalesTax, 2, True, False, True)) & _
					"&x_tax_exempt=False" '& _
					'"&x_duty=0.00" & _
					'"&x_freight=0.00"
			Else
				strRequestText = strRequestText & _
					"&x_tax=" & Server.URLEncode("$0.00") & _
					"&x_tax_exempt=True" '& _
					'"&x_duty=0.00" & _
					'"&x_freight=0.00"
			End If

			'--- Live Variables ---
			strPostURL = "https://secure.authorize.net/gateway/transact.dll"

			'--- Test Variables ---
			'strPostURL = "https://certification.authorize.net/gateway/transact.dll"
			'strRequestText = strRequestText & "&x_test_request=True"
			'strRequestText = strRequestText & "&x_merchant_email=" & Server.URLEncode("ben@refron.com")
				
			'--- Type Specific Parameters ---
			Select Case strType
				Case "AUTH_CAPTURE", "AUTH_ONLY"
					'Do Nothing
				Case "CREDIT", "VOID", "PRIOR_AUTH_CAPTURE"
					strRequestText = strRequestText & "&x_trans_id=" & Server.URLEncode(pOriginalID)
			End Select

			'--- Credit Card & Billing ---

            dim objEncrypt
            set objEncrypt = server.CreateObject("CUUtils.Encrypt")
            
            
            

			strRequestText = strRequestText & _
				"&x_card_num=" & Server.URLEncode(objEncrypt.EncryptStringAscii(cstr(pCardNumber))) & _
				"&x_exp_date=" & Server.URLEncode(pExpireDate) & _
				"&x_address=" & Server.URLEncode(pAddress) & _
				"&x_city=" & Server.URLEncode(pCity) & _
				"&x_state=" & Server.URLEncode(pState) & _
				"&x_zip=" & Server.URLEncode(pZip)
				
			set objEncrypt = Nothing

			If Not IsNull(pFirstName) Then
				strRequestText = strRequestText & "&x_first_name=" & Server.URLEncode(pFirstName)
			End If
			If Not IsNull(pLastName) Then
				strRequestText = strRequestText & "&x_last_name=" & Server.URLEncode(pLastName)
			End If
			If Not IsNull(pCompanyName) Then
				strRequestText = strRequestText & "&x_company=" & Server.URLEncode(pCompanyName)
			End If
			If Not IsNull(pCustomerID) Then
				strRequestText = strRequestText & "&x_cust_id=" & Server.URLEncode(pCustomerID)
			End If
			If Not IsNull(pOrderNumber) Then
				strRequestText = strRequestText & "&x_invoice_num=" & Left(Server.URLEncode(pOrderNumber), 6) & "-" & Right(Server.URLEncode(pOrderNumber), 1)
			End If
			If Not IsNull(pPONumber) Then
				strRequestText = strRequestText & "&x_po_num=" & Server.URLEncode(pPONumber)
			End If
			If Not IsNull(pCustomerEmail) And strType <> "AUTH_ONLY" And strType <> "VOID" Then
				strRequestText = strRequestText & _
					"&x_email_customer=True" & _
					"&x_email=" & Server.URLEncode(pCustomerEmail)
			End If
			
			Dim HttpRequest, aryResponse, i
			Set HttpRequest = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
			HttpRequest.Open "POST", strPostURL, False
			HttpRequest.SetRequestHeader "Content-Type", "application/x-www-form-URLEncoded"
'Response.Write "<p>Sending Request:<ul><li>" & Replace(strRequestText, "&", "<li>") & "</ul></p>"
			HttpRequest.Send strRequestText
'Response.Write "<p>Received Response:<ul><li>" & Replace(HttpRequest.ResponseText, "|", "<li>") & "</ul></p>"
			aryResponse = Split("Start At 1, Not 0|" & HttpRequest.ResponseText, "|")
			Set HttpRequest = Nothing

			If UBound(aryResponse) > 2 Then

				For i = 0 To UBound(aryResponse)
					If aryResponse(i) = "" Then
						aryResponse(i) = Null
					End If
				Next

				cmd.CommandText = _
					"UPDATE CreditCardTransactions " & _
					"SET " & _
						"CreditCardTransactionResponseCode = ?, " & _
						"CreditCardTransactionAmount = ?, " & _
						"CreditCardTransactionSalesTax = ?, " & _
						"CreditCardTransactionType = ?, " & _
						"CreditCardTransactionDateProcessed = ?, " & _
						"CreditCardTransactionUserCode = ?, " & _
						"CreditCardTransactionResponseSubCode = ?, " & _
						"CreditCardTransactionReasonCode = ?, " & _
						"CreditCardTransactionApprovalCode = ?, " & _
						"CreditCardTransactionAvsResultCode = ?, " & _
						"CreditCardTransactionAuthorizeID = ?, " & _
						"CreditCardTransactionCvc2ResponseCode = ?, " & _
						"CreditCardTransactionCavvResponseCode = ? " & _
					"WHERE CreditCardTransactionID = ?; "
				Call DeleteAllParameters(cmd)
				cmd.Parameters.Append cmd.CreateParameter("ResponseCode", adTinyInt, adParamInput, , aryResponse(1))
				If strType = "CREDIT" Or strType = "VOID" Then
					cmd.Parameters.Append cmd.CreateParameter("Amount", adCurrency, adParamInput, , (-1 * dblAmount))
					cmd.Parameters.Append cmd.CreateParameter("SalesTax", adCurrency, adParamInput, , (-1 * dblSalesTax))
				Else
					cmd.Parameters.Append cmd.CreateParameter("Amount", adCurrency, adParamInput, , dblAmount)
					cmd.Parameters.Append cmd.CreateParameter("SalesTax", adCurrency, adParamInput, , dblSalesTax)
				End If
				cmd.Parameters.Append cmd.CreateParameter("Type", adVarChar, adParamInput, 18, strType)
				cmd.Parameters.Append cmd.CreateParameter("DateProcessed", adDate, adParamInput, , Now())
				cmd.Parameters.Append cmd.CreateParameter("UserCode", adChar, adParamInput, 2, Session("UserCode"))
				cmd.Parameters.Append cmd.CreateParameter("ResponseSubCode", adTinyInt, adParamInput, , aryResponse(2))
				cmd.Parameters.Append cmd.CreateParameter("ReasonCode", adSmallInt, adParamInput, , aryResponse(3))
				cmd.Parameters.Append cmd.CreateParameter("ApprovalCode", adChar, adParamInput, 6, aryResponse(5))
				cmd.Parameters.Append cmd.CreateParameter("AvsResultCode", adChar, adParamInput, 1, aryResponse(6))
				cmd.Parameters.Append cmd.CreateParameter("AuthorizeID", adBigInt, adParamInput, , aryResponse(7))
				cmd.Parameters.Append cmd.CreateParameter("Cvc2ResponseCode", adChar, adParamInput, 1, aryResponse(39))
				cmd.Parameters.Append cmd.CreateParameter("CavvResponseCode", adChar, adParamInput, 1, aryResponse(40))
				cmd.Parameters.Append cmd.CreateParameter("TransactionID", adInteger, adParamInput, , pTransactionID)
				cmd.Execute , , adExecuteNoRecords
			
				cmd.CommandText = _
					"UPDATE CreditCardReasonCodes " & _
					"SET " & _
						"CreditCardReasonCodeName = ? " & _
					"WHERE CreditCardReasonCodeID = ?; "
				Call DeleteAllParameters(cmd)
				cmd.Parameters.Append cmd.CreateParameter("ReasonCodeName", adVarChar, adParamInput, 255, Left(Replace(Trim(aryResponse(4)), "(TESTMODE) ", ""), 255))
				cmd.Parameters.Append cmd.CreateParameter("ReasonCodeID", adSmallInt, adParamInput, , aryResponse(3))
				cmd.Execute , , adExecuteNoRecords
				
				If aryResponse(1) = 1 And aryResponse(3) = 1 Then
					pApproved = True
					
					If Not IsNull(pPostToCode) Then
						Select Case strType
							Case "AUTH_CAPTURE", "PRIOR_AUTH_CAPTURE", "CREDIT", "VOID"
								If strType = "VOID" And pOriginalType = "AUTH_ONLY" Then
									'Do Nothing - Charge was not added to the open item table and does not need to be removed.
								Else
								    ConnectToDb("Refron")
								    'Removed 7/31/08 by Brian - Turn off OPNITEM updates
								    'Added back 8/8/08 by Ben - Turn on OPNITEM updates
									conRefron.Execute("UPDATE OPNITEM SET OIINV# = CAST(CAST(OIINV# AS INTEGER) + 1 AS CHAR(7)) WHERE OIREC# = 1")
									pRecordNumber = conRefron.Execute("SELECT CAST(OIINV# AS INTEGER) AS RecordNumber FROM OPNITEM WHERE OIREC# = 1")("RecordNumber")
									If IsNumeric(pRecordNumber) And pRecordNumber <> "" Then
										cmd.CommandText = _
											"INSERT INTO OPNITEM (" & _
												"OITRCD, " & _
												"OIPOST, " & _
												"OIINV#, " & _
												"OIIVMM, " & _
												"OIIVDD, " & _
												"OIIVYY, " & _
												"OICO, " & _
												"OISLMN, " & _
												"OICSTP, " & _
												"OISBTP, " & _
												"OICLAS, " & _
												"OIAR$, " & _
												"OIDTEN, " & _
												"OISTC, " & _
												"OIBTCH, " & _
												"OIB294, " & _
												"OIPRCD, " & _
												"OIDTLP, " & _
												"OICUPO, " & _
												"OICSCD, " & _
												"OIREC#, " & _
												"OIIVCC, " & _
												"OIRLS, " & _
												"OIIVD8, " & _
												"OIDLP8, " & _
												"OIDTE8, " & _
												"OIRFCD, " & _
												"OIFRT$, " & _
												"OITAX$, " & _
												"OIPAID, " & _
												"OIAPPD, " & _
												"OIWRTO, " & _
												"OIPATD, " & _
												"OIMRCH, " & _
												"OIDTPP, " & _
												"OIHIDE, " & _
												"OITRM$, " & _
												"OIORD#, " & _
												"OIDUDT, " & _
												"OIFUD8, " & _
												"OIDIVN, " & _
												"OITPOR, " & _
												"OIOTRM, " & _
												"OIWRTC, " & _
												"OIBNCD, " & _
												"OITERM, " & _
												"OITMPS, " & _
												"OITXST, " & _
												"OITMPC, " & _
												"OIPPDF, " & _
												"OIWSID, " & _
												"OIFUBY, " & _
												"OIB158, " & _
												"OISTAT, " & _
												"OINTXC, " & _
												"OINTMC, " & _
												"OIFUNT, " & _
												"OICMNT) " & _
											"VALUES (" & _
												"?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " & _
												"0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " & _
												"'', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '' " & _
												") "
										Call DeleteAllParameters(cmd)
										Dim DateProcessed
										If Time() >= #18:00:00# Then
											DateProcessed = DateAdd("d", 1, Date())
										Else
											DateProcessed = Date()
										End If
										If strType = "CREDIT" Or strType = "VOID" Then
											cmd.Parameters.Append cmd.CreateParameter("TransactionCode", adChar, adParamInput, 2, "CD")
										Else
											cmd.Parameters.Append cmd.CreateParameter("TransactionCode", adChar, adParamInput, 2, "CR")
										End If
										cmd.Parameters.Append cmd.CreateParameter("PostToCode", adChar, adParamInput, 7, pPostToCode)
										cmd.Parameters.Append cmd.CreateParameter("InvoiceNumber", adChar, adParamInput, 7, pOrderNumber)
										cmd.Parameters.Append cmd.CreateParameter("MonthProcessed", adInteger, adParamInput, , PadZeros(Month(DateProcessed), 2))
										cmd.Parameters.Append cmd.CreateParameter("DayProcessed", adInteger, adParamInput, , PadZeros(Day(DateProcessed), 2))
										cmd.Parameters.Append cmd.CreateParameter("YearProcessed", adInteger, adParamInput, , Right(Year(DateProcessed), 2))
										cmd.Parameters.Append cmd.CreateParameter("CompanyDivision", adChar, adParamInput, 1, pCompanyDivision)
										cmd.Parameters.Append cmd.CreateParameter("SalespersonCode", adChar, adParamInput, 3, pSalespersonCode)
										cmd.Parameters.Append cmd.CreateParameter("CustomerType", adChar, adParamInput, 1, pCustomerType)
										cmd.Parameters.Append cmd.CreateParameter("CustomerSubType", adChar, adParamInput, 1, pCustomerSubType)
										cmd.Parameters.Append cmd.CreateParameter("PurgeClass", adChar, adParamInput, 1, pPurgeClass)
										If strType = "CREDIT" Or strType = "VOID" Then
											cmd.Parameters.Append cmd.CreateParameter("Amount", adCurrency, adParamInput, , dblAmount)
										Else
											cmd.Parameters.Append cmd.CreateParameter("Amount", adCurrency, adParamInput, , (-1 * dblAmount))
										End If
										cmd.Parameters.Append cmd.CreateParameter("DateEntered6", adInteger, adParamInput, , PadZeros(Month(Date()), 2) & PadZeros(Day(Date()), 2) & Right(Year(Date()), 2))
										cmd.Parameters.Append cmd.CreateParameter("SubTransactionCode", adChar, adParamInput, 2, "CC")
										cmd.Parameters.Append cmd.CreateParameter("Batch", adChar, adParamInput, 2, PadZeros(Day(DateProcessed), 2))
										cmd.Parameters.Append cmd.CreateParameter("BlankField", adChar, adParamInput, 37, "Auto CC Charge")
										cmd.Parameters.Append cmd.CreateParameter("PrintCode", adChar, adParamInput, 1, "Y")
										cmd.Parameters.Append cmd.CreateParameter("DateLastPaid", adInteger, adParamInput, , PadZeros(Month(Date()), 2) & PadZeros(Day(Date()), 2) & Right(Year(Date()), 2))
										If IsNull(pPONumber) Then
											cmd.Parameters.Append cmd.CreateParameter("PurchaseOrderNumber", adChar, adParamInput, 13, "")
										Else
											cmd.Parameters.Append cmd.CreateParameter("PurchaseOrderNumber", adChar, adParamInput, 13, pPONumber)
										End If
										If IsNull(pCustomerID) Then
											cmd.Parameters.Append cmd.CreateParameter("AccountCode", adChar, adParamInput, 7, "")
										Else
											cmd.Parameters.Append cmd.CreateParameter("AccountCode", adChar, adParamInput, 7, pCustomerID)
										End If
										cmd.Parameters.Append cmd.CreateParameter("RecordNumber", adInteger, adParamInput, , pRecordNumber)
										cmd.Parameters.Append cmd.CreateParameter("CenturyProcessed", adInteger, adParamInput, , Left(Year(DateProcessed), 2))
										cmd.Parameters.Append cmd.CreateParameter("ReleaseNumber", adChar, adParamInput, 14, pReleaseNumber)
										cmd.Parameters.Append cmd.CreateParameter("DateProcessed", adInteger, adParamInput, , Year(DateProcessed) & PadZeros(Month(DateProcessed), 2) & PadZeros(Day(DateProcessed), 2))
										cmd.Parameters.Append cmd.CreateParameter("DateLastPaid", adInteger, adParamInput, , Year(Date()) & PadZeros(Month(Date()), 2) & PadZeros(Day(Date()), 2))
										cmd.Parameters.Append cmd.CreateParameter("DateEntered8", adInteger, adParamInput, , Year(Date()) & PadZeros(Month(Date()), 2) & PadZeros(Day(Date()), 2))
										cmd.Parameters.Append cmd.CreateParameter("ReferenceCode", adChar, adParamInput, 7, pReferenceCode)
										ConnectToDb("Refron")
										cmd.ActiveConnection = conRefron
								        'Removed 7/31/08 by Brian - Turn off OPNITEM updates
								        'Added back 8/8/08 by Ben - Turn on OPNITEM updates
										cmd.Execute , , adExecuteNoRecords

										cmd.CommandText = _
											"UPDATE CreditCardTransactions " & _
											"SET " & _
												"CreditCardTransactionOpnitemRecordNumber = ? " & _
											"WHERE CreditCardTransactionID = ?; "
										Call DeleteAllParameters(cmd)
										cmd.Parameters.Append cmd.CreateParameter("OpnitemRecordNumber", adInteger, adParamInput, , pRecordNumber)
										cmd.Parameters.Append cmd.CreateParameter("TransactionID", adInteger, adParamInput, , pTransactionID)
										ConnectToDb("Refron")
										cmd.ActiveConnection = conRefron
								        'Removed 7/31/08 by Brian - Turn off OPNITEM updates
								        'Added back 8/8/08 by Ben - Turn on OPNITEM updates
										cmd.Execute , , adExecuteNoRecords
									End If
								End If
							Case Else
								'Do Nothing
						End Select
					End If
				
				End If

			End If
		End If
	End Sub

	Private Function PadZeros(strNumber, intLen)
		Dim i, strTemp
		For i = 1 To intLen - Len(strNumber)
			strTemp = strTemp & "0"
		Next
		PadZeros = strTemp & CStr(strNumber)
	End Function

	Public Sub Process
		Select Case pTransactionType
			Case "VALIDATE" 'Custom type to validate address & cc by authorizing $0.01 then voiding it
				If Not IsNull(pCardID) And IsNull(pTransactionID) Then
					cmd.CommandText = _
						"SET NOCOUNT ON; " & _
						"INSERT INTO CreditCardTransactions (" & _
							"CreditCardTransactionCardID, " & _
							"CreditCardTransactionDatePosted, " & _
							"CreditCardTransactionOriginalID, " & _
							"CreditCardTransactionNote) " & _
						"VALUES (?, ?, ?, ?); " & _
						"SELECT CAST(@@IDENTITY AS INTEGER) AS CreditCardTransactionID; "
					Call DeleteAllParameters(cmd)
					cmd.Parameters.Append cmd.CreateParameter("CardID", adInteger, adParamInput, , pCardID)
					cmd.Parameters.Append cmd.CreateParameter("DatePosted", adDate, adParamInput, , Now())
					cmd.Parameters.Append cmd.CreateParameter("OriginalID", adInteger, adParamInput, , pAuthorizeID) 'Null value must come from variable
					cmd.Parameters.Append cmd.CreateParameter("Note", adVarChar, adParamInput, 250, "Transaction to validate credit card.")
					Call LoadTransaction(cmd.Execute()("CreditCardTransactionID"))
					Call ProcessTransaction("AUTH_ONLY", 0.01, Null)
					If pResponseCode = 1 Then
						cmd.Parameters("CardID") = pCardID
						cmd.Parameters("DatePosted") = Now()
						cmd.Parameters("OriginalID") = pAuthorizeID
						cmd.Parameters("Note") = "Transaction to void validation."
						Call LoadTransaction(cmd.Execute()("CreditCardTransactionID"))
						Call ProcessTransaction("VOID", 0.01, Null)
					End If
				End If
			Case "AUTH_CAPTURE", "AUTH_ONLY"
				If Not IsNull(pTransactionID) _
					And Not IsNull(pAmount) _
					And pResponseCode = 0 Then
					Call ProcessTransaction(pTransactionType, pAmount, pSalesTax)
				End If
			Case "CREDIT"
				If Not IsNull(pAuthorizeID) _
					And Not IsNull(pAmount) _
					And pAmount <= Abs(pOriginalAmount) _
					And (pOriginalType = "AUTH_CAPTURE" Or pOriginalType = "PRIOR_AUTH_CAPTURE") _
					And pResponseCode = 1 Then
					cmd.CommandText = _
						"SET NOCOUNT ON; " & _
						"INSERT INTO CreditCardTransactions (" & _
							"CreditCardTransactionCardID, " & _
							"CreditCardTransactionDatePosted, " & _
							"CreditCardTransactionOrderNumber, " & _
							"CreditCardTransactionOriginalID, " & _
							"CreditCardTransactionReceiptType, " & _
							"CreditCardTransactionReceiptDestination, " & _
							"CreditCardTransactionNote) " & _
						"VALUES (?, ?, ?, ?, ?, ?, ?); " & _
						"SELECT CAST(@@IDENTITY AS INTEGER) AS CreditCardTransactionID; "
					Call DeleteAllParameters(cmd)
					cmd.Parameters.Append cmd.CreateParameter("CardID", adInteger, adParamInput, , pCardID)
					cmd.Parameters.Append cmd.CreateParameter("DatePosted", adDate, adParamInput, , Now())
					response.write "<br/>pOrderNumber: " & pOrderNumber & "<br/>"
					cmd.Parameters.Append cmd.CreateParameter("OrderNumber", adInteger, adParamInput, , pOrderNumber)
					response.write "<br/>pAuthorizeID: " & pAuthorizeID & "<br/>"
					cmd.Parameters.Append cmd.CreateParameter("OriginalID", adInteger, adParamInput, , pAuthorizeID)
					response.write "<br/>pReceiptType: " & pReceiptType & "<br/>"
					cmd.Parameters.Append cmd.CreateParameter("ReceiptType", adVarChar, adParamInput, 75, pReceiptType)
					cmd.Parameters.Append cmd.CreateParameter("ReceiptDestination", adVarChar, adParamInput, 250, pReceiptDestination)
					cmd.Parameters.Append cmd.CreateParameter("Note", adVarChar, adParamInput, 250, "Credit for previous transaction.")
					Call LoadTransaction(cmd.Execute()("CreditCardTransactionID"))
					Call ProcessTransaction(pTransactionType, pAmount, pSalesTax)
				End If
			Case "VOID"
				If Not IsNull(pAuthorizeID) _
					And pResponseCode = 1 Then
					pAmount = pOriginalAmount
					pSalesTax = pOriginalSalesTax
					cmd.CommandText = _
						"SET NOCOUNT ON; " & _
						"INSERT INTO CreditCardTransactions (" & _
							"CreditCardTransactionCardID, " & _
							"CreditCardTransactionDatePosted, " & _
							"CreditCardTransactionOrderNumber, " & _
							"CreditCardTransactionOriginalID, " & _
							"CreditCardTransactionReceiptType, " & _
							"CreditCardTransactionReceiptDestination, " & _
							"CreditCardTransactionNote) " & _
						"VALUES (?, ?, ?, ?, ?, ?, ?); " & _
						"SELECT CAST(@@IDENTITY AS INTEGER) AS CreditCardTransactionID; "
					Call DeleteAllParameters(cmd)
					cmd.Parameters.Append cmd.CreateParameter("CardID", adInteger, adParamInput, , pCardID)
					cmd.Parameters.Append cmd.CreateParameter("DatePosted", adDate, adParamInput, , Now())
					cmd.Parameters.Append cmd.CreateParameter("OrderNumber", adInteger, adParamInput, , pOrderNumber)
					cmd.Parameters.Append cmd.CreateParameter("OriginalID", adInteger, adParamInput, , pAuthorizeID)
					cmd.Parameters.Append cmd.CreateParameter("ReceiptType", adVarChar, adParamInput, 75, pReceiptType)
					cmd.Parameters.Append cmd.CreateParameter("ReceiptDestination", adVarChar, adParamInput, 250, pReceiptDestination)
					cmd.Parameters.Append cmd.CreateParameter("Note", adVarChar, adParamInput, 250, "Void previous transaction.")
					Call LoadTransaction(cmd.Execute()("CreditCardTransactionID"))
					Call ProcessTransaction(pTransactionType, pAmount, pSalesTax)
				End If
			Case "PRIOR_AUTH_CAPTURE"
				If Not IsNull(pAuthorizeID) _
					And pAmount <= Abs(pOriginalAmount) _
					And pOriginalType = "AUTH_ONLY" _
					And pResponseCode = 1 Then
					cmd.CommandText = _
						"SET NOCOUNT ON; " & _
						"INSERT INTO CreditCardTransactions (" & _
							"CreditCardTransactionCardID, " & _
							"CreditCardTransactionDatePosted, " & _
							"CreditCardTransactionOrderNumber, " & _
							"CreditCardTransactionOriginalID, " & _
							"CreditCardTransactionReceiptType, " & _
							"CreditCardTransactionReceiptDestination, " & _
							"CreditCardTransactionNote) " & _
						"VALUES (?, ?, ?, ?, ?, ?, ?); " & _
						"SELECT CAST(@@IDENTITY AS INTEGER) AS CreditCardTransactionID; "
					Call DeleteAllParameters(cmd)
					cmd.Parameters.Append cmd.CreateParameter("CardID", adInteger, adParamInput, , pCardID)
					cmd.Parameters.Append cmd.CreateParameter("DatePosted", adDate, adParamInput, , Now())
					cmd.Parameters.Append cmd.CreateParameter("OrderNumber", adInteger, adParamInput, , pOrderNumber)
					cmd.Parameters.Append cmd.CreateParameter("OriginalID", adInteger, adParamInput, , pAuthorizeID)
					cmd.Parameters.Append cmd.CreateParameter("ReceiptType", adVarChar, adParamInput, 75, pReceiptType)
					cmd.Parameters.Append cmd.CreateParameter("ReceiptDestination", adVarChar, adParamInput, 250, pReceiptDestination)
					cmd.Parameters.Append cmd.CreateParameter("Note", adVarChar, adParamInput, 250, "Process previous authorization.")
					Call LoadTransaction(cmd.Execute()("CreditCardTransactionID"))
					Call ProcessTransaction(pTransactionType, pAmount, pSalesTax)
				End If
		End Select
	End Sub

End Class
%>