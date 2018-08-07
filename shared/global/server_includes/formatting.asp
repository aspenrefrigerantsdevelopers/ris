<%
Function PadZeros(strNumber, intLen)
	Dim i, strTemp
	For i = 1 To intLen - Len(strNumber)
		strTemp = strTemp & "0"
	Next
	PadZeros = strTemp & CStr(strNumber)
End Function

Function PadSpacesRight(strText, intLen)
	Dim i, strTemp
	For i = 1 To intLen - Len(Trim(strText))
		strTemp = strTemp & " "
	Next
	PadSpacesRight = strTemp & Trim(strText)
End Function

Function PadLeft(strText, strValue, intLen)
	Dim i, strTemp
	For i = 1 To intLen - Len(Trim(strText))
		strTemp = strTemp & strValue
	Next
	PadLeft = strTemp & Trim(strText)
End Function

Function PadRight(strText, strValue, intLen)
	Dim i, strTemp
	For i = 1 To intLen - Len(Trim(strText))
		strTemp = strTemp & strValue
	Next
	PadRight = Trim(strText) & strTemp
End Function

Function TrimNumber(Number, DigitsAfterDecPoint)
	If InStr(1, CStr(Number), ".") > 0 Then
		TrimNumber = CDbl(Left(CStr(Number), InStr(1, CStr(Number), ".") + DigitsAfterDecPoint))
	Else
		TrimNumber = Number
	End If
End Function

'Changes SQL Text from AS400 Syntax to SQL Server Syntax
Function FixSQL(sql)
	Dim tmpString
	tmpString = sql
	tmpString = Replace(tmpString, "WGWRK.", "")
	tmpString = Replace(tmpString, "WEBLIB.", "")
	tmpString = Replace(tmpString, "QS36F.", "")
	tmpString = Replace(tmpString, "||", "+")
	tmpString = Replace(tmpString, "IFNULL", "ISNULL")
	tmpString = Replace(tmpString, "FLOAT(", "CONVERT(FLOAT, ")
	FixSQL = tmpString
End Function

'Adds a zero to short zip codes and returns blank for zero
Function FixZip(ZipCode)
	Select Case Len(ZipCode)
		Case 5:
			FixZip = CStr(ZipCode)
		Case 4:
			FixZip = "0" & CStr(ZipCode)
		Case Else
			FixZip = ""
	End Select
End Function

'Adds a zero to short zip codes and returns blank for zero
Function FixZip4(ZipCode)
	Select Case Len(ZipCode)
		Case 4:
			FixZip4 = CStr(ZipCode)
		Case 3:
			FixZip4 = "0" & CStr(ZipCode)
		Case Else
			FixZip4 = ""
	End Select
End Function

'Returns multipart zip code formatted with hyphen
Function FullZip(Zip, Zip4)
	If FixZip(Zip) <> "" Then
		If FixZip4(Zip4) <> "" Then
			FullZip = FixZip(Zip) & "-" & FixZip4(Zip4)
		Else
			FullZip = FixZip(Zip)
		End If
	End If
End Function

'Used to format a date field as "long date"
'Dim LongDate
'LongDate = MonthName(Month(Now())) & " " & Day(Now())& ", " & Year(Now())

'Used to convert phone numbers into display format
Function FormatPhone(strTxt, strFormat)
	If IsNull(strTxt) Or strTxt = "" Or strTxt = "0" Then FormatPhone = "" : Exit Function
	Dim strExt
	strTxt = StripPhone(CStr(strTxt))
	If strTxt = "" Or strTxt = "0" Then FormatPhone = "" : Exit Function
	Select Case strFormat
		Case "(PPP) PPP-PPPP"
			Select Case Len(strTxt)
				Case 1,2,3,4:
					FormatPhone = strTxt
				Case 5,6,7:
					FormatPhone = Left(strTxt,Len(strTxt)-4) & "-" & Right(strTxt,4)
				Case 8,9,10:
					FormatPhone = "(" & Left(strTxt,Len(strTxt)-7) & ") " & Mid(strTxt,Len(strTxt)-6,3) & "-" & Right(strTxt,4)
				Case Else
					strExt = Mid(strTxt,11)
					strTxt = Left(strTxt,10)
					FormatPhone = "(" & Left(strTxt,Len(strTxt)-7) & ") " & Mid(strTxt,Len(strTxt)-6,3) & "-" & Right(strTxt,4) & " x" & strExt
			End Select
		Case "PPP-PPP-PPPP"
			Select Case Len(strTxt)
				Case 1,2,3,4:
					FormatPhone = strTxt
				Case 5,6,7:
					FormatPhone = Left(strTxt,Len(strTxt)-4) & "-" & Right(strTxt,4)
				Case 8,9,10:
					FormatPhone = Left(strTxt,Len(strTxt)-7) & "-" & Mid(strTxt,Len(strTxt)-6,3) & "-" & Right(strTxt,4)
				Case Else
					strExt = Mid(strTxt, 11)
					strTxt = Left(strTxt, 10)
					FormatPhone = Left(strTxt,Len(strTxt)-7) & "-" & Mid(strTxt,Len(strTxt)-6,3) & "-" & Right(strTxt,4) & " x" & strExt
			End Select
	End Select
End Function

Function FixNumber(strTxt)
	If IsNull(strTxt) Or strTxt = "" Or strTxt = "0" Then FixNumber = "" : Exit Function
	strTxt = StripPhone(CStr(strTxt))
	If strTxt = "" Or strTxt = "0" Then FixNumber = "" : Exit Function
	FixNumber = strTxt
End Function

Function FixDate(strDate)
	If IsNull(strDate) Or strDate = "" Or Not IsDate(strDate) Or strDate = "1/1/1900" Then FixDate = "" : Exit Function
	FixDate = strDate
End Function

Function FixTime(strTime)
	If IsNull(strTime) Or strTime = "" Or Not IsDate(strTime) Or strTime = "0:0:0" Then FixTime = "" : Exit Function
	FixTime = strTime
End Function

'Used to convert zip codes into display format
Function FormatZip(strTxt,numChars)
	strTxt = CStr(strTxt)
	Select Case numChars-Len(strTxt)
		Case 1:
			FormatZip = "0" & strTxt
		Case 2:
			FormatZip = "00" & strTxt
		Case 3:
			FormatZip = "000" & strTxt
		Case Else
			FormatZip = strTxt
	End Select
End Function

Function FixSqlVariable(strIn)
	Dim strTxt : strTxt = strIn
	strTxt = Replace(strTxt, "'", "''")
	strTxt = Replace(strTxt, "\", "\\")
	strTxt = Replace(strTxt, "%", "\%")
	strTxt = Replace(strTxt, "_", "\_")
	FixSqlVariable = strTxt
End Function

'Strips all extraneous characters from phone numbers
Function StripPhone(strIn)
	Dim strTxt : strTxt = strIn
	'strTxt = Replace(strTxt,"-","")
	'strTxt = Replace(strTxt,"(","")
	'strTxt = Replace(strTxt,")","")
	'strTxt = Replace(strTxt," ","")
	'strTxt = Replace(strTxt,"/","")
	'strTxt = Replace(strTxt,"\","")
	'strTxt = Replace(strTxt, ".", "")
	strTxt = StripNonAlphaNumericCharacters(strTxt)
	strTxt = Replace(strTxt, "x", "")
	strTxt = Replace(strTxt, "X", "")
	strTxt = Replace(strTxt, "ext", "")
	strTxt = Replace(strTxt, "Ext", "")
	'strTxt = Left(strTxt,10) - commented out on 6/25/04 because it removed the extension passed to the format phone function
	StripPhone = strTxt
End Function

Function StripNonAlphaNumericCharacters(strIn)
	Dim strValue : strValue = strIn
	strValue = Replace(strValue, " ", "")
	strValue = Replace(strValue, "`", "")
	strValue = Replace(strValue, "~", "")
	strValue = Replace(strValue, "!", "")
	strValue = Replace(strValue, "@", "")
	strValue = Replace(strValue, "#", "")
	strValue = Replace(strValue, "$", "")
	strValue = Replace(strValue, "%", "")
	strValue = Replace(strValue, "^", "")
	strValue = Replace(strValue, "&", "")
	strValue = Replace(strValue, "*", "")
	strValue = Replace(strValue, "(", "")
	strValue = Replace(strValue, ")", "")
	strValue = Replace(strValue, "-", "")
	strValue = Replace(strValue, "_", "")
	strValue = Replace(strValue, "=", "")
	strValue = Replace(strValue, "+", "")
	strValue = Replace(strValue, "[", "")
	strValue = Replace(strValue, "{", "")
	strValue = Replace(strValue, "]", "")
	strValue = Replace(strValue, "}", "")
	strValue = Replace(strValue, "\", "")
	strValue = Replace(strValue, "|", "")
	strValue = Replace(strValue, ";", "")
	strValue = Replace(strValue, ":", "")
	strValue = Replace(strValue, "'", "")
	strValue = Replace(strValue, """", "")
	strValue = Replace(strValue, ",", "")
	strValue = Replace(strValue, "<", "")
	strValue = Replace(strValue, ".", "")
	strValue = Replace(strValue, ">", "")
	strValue = Replace(strValue, "/", "")
	strValue = Replace(strValue, "?", "")
	StripNonAlphaNumericCharacters = strValue
End Function

Function StripInvalidFilenameCharacters(strIn)
	Dim strValue : strValue = strIn
	strValue = Replace(strValue, "\", "")
	strValue = Replace(strValue, "/", "")
	strValue = Replace(strValue, ":", "")
	strValue = Replace(strValue, "*", "")
	strValue = Replace(strValue, "?", "")
	strValue = Replace(strValue, """", "")
	strValue = Replace(strValue, "<", "")
	strValue = Replace(strValue, ">", "")
	strValue = Replace(strValue, "|", "")
	StripInvalidFilenameCharacters = strValue
End Function

Function StripNumber(strIn)
	StripNumber = StripNonAlphaNumericCharactersExceptNumericCharacters(strIn)
End Function

Function StripNonAlphaNumericCharactersExceptNumericCharacters(strIn)
	'--- Does not remove / . - from string ---
	Dim strValue : strValue = strIn
	strValue = Replace(strValue, " ", "")
	strValue = Replace(strValue, "`", "")
	strValue = Replace(strValue, "~", "")
	strValue = Replace(strValue, "!", "")
	strValue = Replace(strValue, "@", "")
	strValue = Replace(strValue, "#", "")
	strValue = Replace(strValue, "$", "")
	strValue = Replace(strValue, "%", "")
	strValue = Replace(strValue, "^", "")
	strValue = Replace(strValue, "&", "")
	strValue = Replace(strValue, "*", "")
	strValue = Replace(strValue, "(", "")
	strValue = Replace(strValue, ")", "")
	strValue = Replace(strValue, "_", "")
	strValue = Replace(strValue, "=", "")
	strValue = Replace(strValue, "+", "")
	strValue = Replace(strValue, "[", "")
	strValue = Replace(strValue, "{", "")
	strValue = Replace(strValue, "]", "")
	strValue = Replace(strValue, "}", "")
	strValue = Replace(strValue, "\", "")
	strValue = Replace(strValue, "|", "")
	strValue = Replace(strValue, ";", "")
	strValue = Replace(strValue, ":", "")
	strValue = Replace(strValue, "'", "")
	strValue = Replace(strValue, """", "")
	strValue = Replace(strValue, ",", "")
	strValue = Replace(strValue, "<", "")
	strValue = Replace(strValue, ">", "")
	strValue = Replace(strValue, "?", "")
	StripNonAlphaNumericCharactersExceptNumericCharacters = strValue
End Function

Function FixFaxNumber(strValue)
	strValue = StripNonAlphaNumericCharacters(strValue)
	If strValue <> "" And IsNumeric(strValue) Then
		Select Case Len(strValue)
			Case 11	'12035551212
				If Left(strValue, 1) = "1" And Mid(strValue, 2, 1) <> "1" Then
					FixFaxNumber = "1-" & Mid(strValue, 2, 3) & "-" & Mid(strValue, 5, 3) & "-" & Mid(strValue, 8, 4)
				Else
					AddError("Invalid 11 digit number")
				End If
			Case 10 '2035551212
				If Left(strValue, 1) <> "1" Then
					FixFaxNumber = "1-" & Mid(strValue, 1, 3) & "-" & Mid(strValue, 4, 3) & "-" & Mid(strValue, 7, 4)
				Else
					AddError("Invalid 10 digit number " & strValue)
				End If
			Case 7 '5551212
				If Left(strValue, 1) <> "1" Then
					FixFaxNumber = "1-718-" & Mid(strValue, 1, 3) & "-" & Mid(strValue, 4, 4)
				Else
					AddError("Invalid 7 digit number")
				End If
			Case Else
				AddError("Invalid " & Len(strValue) & " digit number")
		End Select
	Else
		AddError("Blank or non-numeric number")
	End If
End Function

'Returns a date in the AS/400 format YYYYMMDD as a number
Function MakeDate()
	Dim strDate
	strDate = Year(date())
	If Month(date()) < 10 Then strDate = strDate & "0" & Month(date()) Else strDate = strDate & Month(date())
	If Day(date()) < 10 Then strDate = strDate & "0" & Day(date()) Else strDate = strDate & Day(date())
	MakeDate = strDate
End Function

Function URLDecode(str)
	Dim re
	Set re = New RegExp
	str = Replace(str, "+", " ")
	re.Pattern = "%([0-9a-fA-F]{2})"
	re.Global = True
	URLDecode = re.Replace(str, GetRef("UrldecodeHex"))
End Function

Function URLDecodeHex(match, hex_digits, pos, source)
	URLDecodeHex = Chr("&H" & hex_digits)
End Function

'Creates a comma delimited string out of a recordset for a specific column
Function GetString(rsTemp, colNum)
	Dim aryTemp, strTemp, i
	If VarType(rsTemp) = vbObject Then
		If Not rs.EOF Then
			rsTemp.MoveFirst
			aryTemp = rsTemp.GetRows(, , colNum)
			rsTemp.MoveFirst
			For i = 0 To UBound(aryTemp, 2)
				strTemp = strTemp & ", """ & aryTemp(0, i) & """"
			Next
			GetString = Mid(strTemp, 3)
		Else
			GetString = ""
		End If
	Else
		GetString = ""
	End If
End Function

Function MakeList(ListType, ListName, ListDefaultValue, ListCurrentValue, ListValues, ListLabels, ListStatus)
    MakeList = MakeList2(ListType, ListName, ListDefaultValue, ListCurrentValue, ListValues, ListLabels, ListStatus, "")
End Function

'Creates a SELECT, RADIO or CHECKBOX list out of the values provided
Function MakeList2(ListType, ListName, ListDefaultValue, ListCurrentValue, ListValues, ListLabels, ListStatus, ExtraHtml)
	Dim strList, aryTemp, aryListValues, aryListLabels, i, Selected, strExtraHtml, strStyle
	strExtraHtml = ExtraHtml
	If Left(ListValues, 1) = """" Then
		If Right(ListValues, 2) = ", " Then
			ListValues = Left(ListValues, Len(ListValues) - 2)
		End If
		If ListValues = """""" Then
			ListValues = """ """
		End If
		aryListValues = Split(Mid(ListValues, 2, Len(ListValues) - 2), """, """)
		If Right(ListLabels, 2) = ", " Then
			ListLabels = Left(ListLabels, Len(ListLabels) - 2)
		End If
		If ListLabels = """""" Then
			ListLabels = """ """
		End If
		aryListLabels = Split(Mid(ListLabels, 2, Len(ListLabels) - 2), """, """)
	Else
		If ListValues = "" Then
			ListValues = " "
		End If
		aryListValues = Split(ListValues, ", ")
		If ListLabels = "" Then
			ListLabels = " "
		End If
		aryListLabels = Split(ListLabels, ", ")
	End If
	If UBound(aryListValues) <> UBound(aryListLabels) Then
		MakeList2 = "# of values does not match # of labels"
		Exit Function
	End If
	Select Case UCase(ListStatus)
		Case "DISABLED"
			strExtraHtml = strExtraHtml & " disabled"
		Case "READONLY"
			strList = "<input type=""hidden"" name=""" & ListName & """ value=""" & ListCurrentValue & """>"
			strExtraHtml = strExtraHtml & " disabled"
	End Select
	If aryListLabels(UBound(aryListLabels)) = UCase(aryListLabels(UBound(aryListLabels))) Then
		For i = 0 To UBound(aryListLabels)
			aryListLabels(i) = LCase(aryListLabels(i))
		Next
		strExtraHtml = strExtraHtml & " style=""text-transform:capitalize;"""
	End If
	If UCase(ListType) = "SELECT" Then
		strList = strList & "<select name=""" & ListName & """" & strExtraHtml & ">"
	End If
	For i = 0 To UBound(aryListValues)
		If i > 0 And UCase(ListType) <> "SELECT" Then
			strList = strList & " &nbsp; "
		End If
		If (Trim(aryListValues(i)) = Trim(ListCurrentValue) And VarType(ListCurrentValue) <> vbEmpty) _
			Or (UCase(ListType) = "CHECKBOX" And InStr(1, ListCurrentValue, aryListValues(i)) > 0 And VarType(ListCurrentValue) <> vbEmpty) _
			Or (VarType(ListCurrentValue) = vbEmpty And Trim(aryListValues(i)) = Trim(ListDefaultValue)) _
			Or (IsNumeric(ListDefaultValue) And Trim(ListCurrentValue) = "0" And Trim(aryListValues(i)) = Trim(ListDefaultValue)) Then
			If UCase(ListType) = "SELECT" Then
				Selected = " selected"
			Else
				Selected = " checked"
			End If
		Else
			Selected = ""
		End If
		If Trim(aryListValues(i)) = Trim(ListDefaultValue) Then
			Select Case UCase(ListType)
				Case "SELECT"
					strStyle = " style=""background-color:yellow;"""
				Case Else
					strStyle = " style=""font-weight:bold;"""
			End Select
			'strStyle = " style=""color:blue;"""
		Else
			strStyle = ""
		End If
		Select Case UCase(ListType)
			Case "SELECT"
				strList = strList & "<option id=""" & ListName & i & """ value=""" & aryListValues(i) & """" & Selected & strStyle & ">" & Trim(aryListLabels(i)) & "</option>"
			Case Else
				strList = strList & "<input type=""" & LCase(ListType) & """ name=""" & ListName & """ id=""" & ListName & i & """ value=""" & aryListValues(i) & """" & Selected & " style=""vertical-align:middle; margin-left:-5px;""" & strExtraHtml & ">&nbsp;<span" & strStyle & ">" & Trim(aryListLabels(i)) & "</span>"
		End Select
	Next
	If UCase(ListType) = "SELECT" Then
		strList = strList & "</select>"
	End If
	MakeList2 = strList
End Function

'Changes a checkbox value to "Y" or "N"
Function ChkToYN(str)
	If str = "on" Then
		ChkToYN = "Y"
	Else
		ChkToYN = "N"
	End If
End Function

'Changes "Y" or "N" to a checkbox value
Function YNToChk(str)
	If str = "Y" Then
		YNToChk = "on"
	Else
		YNToChk = ""
	End If
End Function

Function ReturnPagingNavigation(rs, PagesToDisplay, URL, ViewAllURL)

	If rs.EOF Then
		ReturnPagingNavigation = ""
	Else
		Dim NavText, StartPage, EndPage, Page, RecordsOnThisPage

	'	URL = Trim(URL)
	'	URL = Replace(URL, "&Page=" & Request.QueryString("Page"), "")
	'	URL = Replace(URL, "?Page=" & Request.QueryString("Page"), "?")
	'	URL = Replace(URL, "Page=" & Request.QueryString("Page"), "")
	'	Select Case Right(URL, 1)
	'		Case "?", "&"
	'			URL = URL
	'		Case Else
	'			URL = URL & "&"
	'	End Select

		URL = RemoveQueryStringVariable("Page", URL)

		'Set the current page (default = 1 if no page specified)
		If Request.QueryString("Page") <> "" Then
			If IsNumeric(Request.QueryString("Page")) Then
				rs.AbsolutePage = Request.QueryString("Page")
			Else
				rs.AbsolutePage = 1
			End If
		Else
			rs.AbsolutePage = 1
		End If

		'Determine # of records on current page since it could be less than PageSize if on last page
		If (rs.AbsolutePage * rs.PageSize) > rs.RecordCount Then
			RecordsOnThisPage = rs.RecordCount
		Else
			RecordsOnThisPage = (rs.AbsolutePage * rs.PageSize)
		End If

		If rs.PageCount <= PagesToDisplay Then
			'Adjust the # of pages to display if less than defined value
			PagesToDisplay = rs.PageCount
			'Set start & end pages when total pages is less than # of pages to display defined value
			StartPage = 1
			EndPage = rs.PageCount
		Else
			'Catch first 5 pages
			If rs.AbsolutePage <= Int(PagesToDisplay / 2) + 1 Then
				StartPage = 1
			'Catch last 5 pages
			ElseIf rs.AbsolutePage > rs.PageCount - Int(PagesToDisplay / 2) Then
				StartPage = rs.PageCount - PagesToDisplay + 2
			'Catch all pages in between
			Else
				StartPage = rs.AbsolutePage - Int(PagesToDisplay / 2) + 1
			End If
			'Catch first 5 pages
			If rs.AbsolutePage <= Int(PagesToDisplay / 2) + 1 Then
				EndPage = PagesToDisplay - 1
			'Catch last 5 pages
			ElseIf rs.AbsolutePage >= rs.PageCount - Int(PagesToDisplay / 2) Then
				EndPage = rs.PageCount
			'Catch all pages in between
			Else
				EndPage = rs.AbsolutePage + Int(PagesToDisplay / 2) - 1
			End If
		End If

		NavText = "Page "
		
		'Only run when 2nd number is not 2
		If StartPage <> 1 Then
			If rs.AbsolutePage = 1 Then
				NavText = "1&nbsp;...&nbsp;"
			Else
				NavText = NavText & "<a href=""" & URL & "Page=1"" style=""text-decoration:underline; color:blue;"">1</a>&nbsp;...&nbsp;"
			End If
		End If

		For Page = StartPage To EndPage
			If Page <> StartPage Then
				NavText = NavText & ",&nbsp;"
			End If
			If rs.AbsolutePage = Page Then
				NavText = NavText & Page
			Else
				NavText = NavText & "<a href=""" & URL & "Page=" & Page & """ style=""text-decoration:underline; color:blue;"">" & Page & "</a>"
			End If
		Next

		'Only run when 2nd to last number is not PageCount - 1
		If EndPage <> rs.PageCount Then
			If rs.AbsolutePage = rs.PageCount Then
				NavText = "&nbsp;...&nbsp;" & rs.PageCount
			Else
				NavText = NavText & "&nbsp;...&nbsp;<a href=""" & URL & "Page=" & rs.PageCount & """ style=""text-decoration:underline; color:blue;"">" & rs.PageCount & "</a>"
			End If
		End If

		'Display view all link if URL was passed
		If ViewAllURL <> "" And StartPage <> EndPage Then
			NavText = NavText & "&nbsp;&nbsp;|&nbsp;&nbsp;<a target=""_blank"" href=""" & ViewAllURL & """ style=""color:blue; text-decoration:underline;"">View All</a>"
		End If

		'Display record summary
		NavText = NavText & "&nbsp;&nbsp;(Records " & FormatNumber(((rs.PageSize * rs.AbsolutePage) - rs.PageSize + 1), 0, True, True, True) & " - " & FormatNumber(RecordsOnThisPage, 0, True, True, True) & " of " & FormatNumber(rs.RecordCount, 0, True, True, True) & ")"

		'Return navigation text
		ReturnPagingNavigation = NavText
	End If

End Function

Function RemoveQueryStringVariable(strVar, strURL)
	Dim strQueryString
	strQueryString = Mid(strURL, InStr(1, strURL, "?") + 1)
	strURL = Left(strURL, InStr(1, strURL, "?") - 1)
	If strQueryString = "" Then
		strURL = strURL & "?"
	Else
		'Response.Write "&" & strVar & "=" & Server.URLPathEncode(Request.QueryString(strVar))
		'Response.End
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & Request.QueryString(strVar), "")
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & Server.URLEncode(Request.QueryString(strVar)), "")
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & Server.URLPathEncode(Request.QueryString(strVar)), "")
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & Server.HTMLEncode(Request.QueryString(strVar)), "")
		strQueryString = Replace(strQueryString, strVar & "=" & Request.QueryString(strVar), "")
		strQueryString = Replace(strQueryString, strVar & "=" & Server.URLEncode(Request.QueryString(strVar)), "")
		strQueryString = Replace(strQueryString, strVar & "=" & Server.URLPathEncode(Request.QueryString(strVar)), "")
		strQueryString = Replace(strQueryString, strVar & "=" & Server.HTMLEncode(Request.QueryString(strVar)), "")
		If Left(strQueryString, 1) = "&" Then
			strURL = strURL & "?" & Mid(strQueryString, 2)
		Else
			strURL = strURL & "?" & strQueryString
		End If
		If Right(strURL, 1) <> "?" And Right(strURL, 1) <> "&" Then
			strURL = strURL & "&"
		End If
	End If
	RemoveQueryStringVariable = strURL
End Function

Function ReturnStringVariables(strVars, strURL)
	Dim aryVars, i, strValue, strValues
	aryVars = Split(strVars, ",")
	For i = 0 To UBound(aryVars)
		strValue = ReturnStringVariable(aryVars(i), strURL)
		If Len(strValue) > 0 Then 
			If Len(strValues) > 0 Then
				strValues = strValues & ", "
			End If
			strValues = strValues & strValue
		End If
	Next
	ReturnStringVariables = strValues
End Function

Function ReturnStringVariable(strVar, strURL)
	Dim p
	p = InStr(1, LCase(strURL), "&" & LCase(strVar) & "=")
	If p > 0 Then
		p = p + 1
	Else
		If Left(LCase(strURL), Len(strVar) + 1) = LCase(strVar) & "=" Then
			p = 1
		End If
	End If
	If p > 0 Then
		If InStr(p + Len(strVar) + 1, strURL, "&") > 0 Then
			ReturnStringVariable = URLDecode(Mid(strURL, p + Len(strVar) + 1, InStr(p + Len(strVar) + 1, strURL, "&") - (p + Len(strVar) + 1)))
		Else
			ReturnStringVariable = URLDecode(Mid(strURL, p + Len(strVar) + 1))
		End If
	End If
End Function

Function RemoveStringVariable(strVar, strURL)
	Dim strQueryString, aryVariables, i, strValue ', dicVariables
	strQueryString = Mid(strURL, InStr(1, strURL, "?") + 1)
	strURL = Left(strURL, InStr(1, strURL, "?") - 1)
	aryVariables = Split(strQueryString, "&")
'	Set dicVariables = Server.CreateObject("Scripting.Dictionary")
'	For i = 0 To UBound(aryVariables)
'		If InStr(1, aryVariables(i), "=") = 0 Then
'			dicVariables.Add aryVariables(i), ""
'		Else
'			dicVariables.Add Left(aryVariables(i), InStr(1, aryVariables(i), "=") - 1), Mid(aryVariables(i), InStr(1, aryVariables(i), "=") + 1)
'		End If
'	Next
	For i = 0 To UBound(aryVariables)
		If InStr(1, aryVariables(i), "=") > 0 Then
			If Left(aryVariables(i), InStr(1, aryVariables(i), "=") - 1) = strVar Then
				strValue = Mid(aryVariables(i), InStr(1, aryVariables(i), "=") + 1)
			End If
		End If
	Next
	If strQueryString = "" Then
		strURL = strURL & "?"
	Else
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & strValue, "")
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & Server.URLEncode(strValue), "")
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & Server.URLPathEncode(strValue), "")
		strQueryString = Replace(strQueryString, "&" & strVar & "=" & Server.HTMLEncode(strValue), "")
		strQueryString = Replace(strQueryString, strVar & "=" & strValue, "")
		strQueryString = Replace(strQueryString, strVar & "=" & Server.URLEncode(strValue), "")
		strQueryString = Replace(strQueryString, strVar & "=" & Server.URLPathEncode(strValue), "")
		strQueryString = Replace(strQueryString, strVar & "=" & Server.HTMLEncode(strValue), "")
		If Left(strQueryString, 1) = "&" Then
			strURL = strURL & "?" & Mid(strQueryString, 2)
		Else
			strURL = strURL & "?" & strQueryString
		End If
		If Right(strURL, 1) <> "?" And Right(strURL, 1) <> "&" Then
			strURL = strURL & "&"
		End If
	End If
	RemoveStringVariable = strURL
End Function

Function ReturnSortColumnTitle(strSortByField, strTitle)
	Dim SortURL
	SortURL = Request.ServerVariables("URL") & "?" & Request.QueryString()
	SortURL = RemoveQueryStringVariable("SortBy", SortURL)
	SortURL = RemoveQueryStringVariable("SortDir", SortURL)

	If strSortByField = strSortBy Then
		If strSortDir = "ASC" Then
			ReturnSortColumnTitle = "<a href=""" & SortURL & "SortBy=" & Server.URLEncode(strSortByField) & "&SortDir=DESC"">" & strTitle & "</a>&nbsp<img align=""absmiddle"" src=""/shared/images/icons/resort_az.gif"">"
		Else
			ReturnSortColumnTitle = "<a href=""" & SortURL & "SortBy=" & Server.URLEncode(strSortByField) & "&SortDir=ASC"">" & strTitle & "</a>&nbsp<img align=""absmiddle"" src=""/shared/images/icons/resort_za.gif"">"
		End If
	Else
		ReturnSortColumnTitle = "<a href=""" & SortURL & "SortBy=" & Server.URLEncode(strSortByField) & "&SortDir=ASC"">" & strTitle & "</a>"
	End If
End Function
%>