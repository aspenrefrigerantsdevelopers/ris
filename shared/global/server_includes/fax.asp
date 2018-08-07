<%
Function FaxEnabled()
	strSql = _
		"SELECT " & _
			"CASE IntranetSettingFaxTestUserCode " & _
				"WHEN '" & Session("UserCode") & "' THEN IntranetSettingFaxTestStatus " & _
				"ELSE IntranetSettingFaxStatus " & _
				"END AS FaxStatus " & _
		"FROM IntranetSettings " & _
		"WHERE IntranetSettingID = 1 "
	ConnectToDb("Refron")
	FaxEnabled = CBool(conRefron.Execute(strSql)("FaxStatus"))
End Function

Function GetToAddress(Prefix, Number)
	If Not FaxEnabled Then
		If Trim(Session("WorkEmail")) <> "" Then
			GetToAddress = Trim(Session("WorkEmail"))
		Else
			GetToAddress = "faxbackup@refron.com"
		End If
	Else
		strSql = _
			"SELECT " & _
				"CASE IntranetSettingFaxTestUserCode " & _
					"WHEN '" & Session("UserCode") & "' THEN " & _
						"CASE TestVendor.IntranetFaxVendorUsePrefix " & _
							"WHEN 1 THEN '" & Prefix & Number & "' + TestVendor.IntranetFaxVendorEmailSuffix " & _
							"ELSE '" & Number & "' + TestVendor.IntranetFaxVendorEmailSuffix " & _
							"END " & _
					"ELSE " & _
						"CASE PrimaryVendor.IntranetFaxVendorUsePrefix " & _
							"WHEN 1 THEN '" & Prefix & Number & "' + PrimaryVendor.IntranetFaxVendorEmailSuffix " & _
							"ELSE '" & Number & "' + PrimaryVendor.IntranetFaxVendorEmailSuffix " & _
							"END " & _
					"END AS ToAddress " & _
			"FROM IntranetSettings " & _
				"LEFT JOIN IntranetFaxVendors AS PrimaryVendor ON IntranetSettingFaxVendorID = PrimaryVendor.IntranetFaxVendorID " & _
				"LEFT JOIN IntranetFaxVendors AS TestVendor ON IntranetSettingFaxTestVendorID = TestVendor.IntranetFaxVendorID " & _
			"WHERE IntranetSettingID = 1 "
		ConnectToDb("Refron")
		GetToAddress = conRefron.Execute(strSql)("ToAddress")
	End If
End Function

Function GetFromAddress(CurrentFromAddress)
	strSql = _
		"SELECT " & _
			"CASE IntranetSettingFaxTestUserCode " & _
				"WHEN '" & Session("UserCode") & "' THEN TestVendor.IntranetFaxVendorFromAddress " & _
				"ELSE PrimaryVendor.IntranetFaxVendorFromAddress " & _
				"END AS FromAddress " & _
		"FROM IntranetSettings " & _
			"LEFT JOIN IntranetFaxVendors AS PrimaryVendor ON IntranetSettingFaxVendorID = PrimaryVendor.IntranetFaxVendorID " & _
			"LEFT JOIN IntranetFaxVendors AS TestVendor ON IntranetSettingFaxTestVendorID = TestVendor.IntranetFaxVendorID " & _
		"WHERE IntranetSettingID = 1 "
	ConnectToDb("Refron")
	Dim strFromAddress
	strFromAddress = conRefron.Execute(strSql)("FromAddress")
	If strFromAddress = "" Then
		GetFromAddress = CurrentFromAddress
	Else
		GetFromAddress = strFromAddress
	End If
End Function

Function GetSubject(Subject)
	If Not FaxEnabled Then
		GetSubject = "Message NOT faxed! Print and fax manually..."
	Else
		GetSubject = Subject
	End If
End Function
%>