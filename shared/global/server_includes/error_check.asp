<%
'---- Adds an error message to a session variable and increments the number of errors ----
Sub AddError(Msg)
	Session("NumErrors") = Session("NumErrors") + 1
	Session("ErrorMsg" & Session("NumErrors")) = Msg
End Sub

'---- Clears all errors and sets the number of errors to zero
Sub ClearErrors()
	Dim ErrNum
	For ErrNum = 1 To Session("NumErrors")
		Set Session("ErrorMsg" & ErrNum) = Nothing
	Next
	Session("NumErrors") = 0
End Sub

'---- Returns a red bullet list of the current error messages
Function ReturnErrorMsg()
	Dim ErrNum, ErrMsg
	If Session("NumErrors") = 1 Then
		ErrMsg = _
			"<div style=""color:red; padding-top:4px; padding-bottom:4px;"">" &_
			Session("ErrorMsg1") &_
			"</div>"
	ElseIf Session("NumErrors") > 1 Then
		ErrMsg = _
			"<div style=""color:red; padding-top:4px; padding-bottom:4px;"">" &_
			"Encountered " & Session("NumErrors") & " error"
		If Session("NumErrors") > 1 Then
			ErrMsg = ErrMsg & "s"
		End If
		ErrMsg = ErrMsg & ":</DIV>"
		For ErrNum = 1 To Session("NumErrors")
			ErrMsg = ErrMsg &_
				"<div style=""color:red; padding-left:15px; padding-bottom:4px;"">" &_
				"<li>" & Session("ErrorMsg" & ErrNum) &_
				"</div>"
		Next
	End If
	Call ClearErrors()
	ReturnErrorMsg = ErrMsg
End Function

'---- Prints a standard error table using the title passed for the heading ----
Sub PrintErrorTable(varTitleText)
	%>
	<table class="Standard">
		<thead>
			<tr class="Standard">
				<td class="Standard">
					<%=varTitleText%>
				</td>
			</tr>
		</thead>
		<tbody>
			<tr class="Standard">
				<td class="Standard">
					<%=ReturnErrorMsg()%>
					<input type="button" class="Standard" onclick="JavaScript:history.back(1)" value="Back">
				</td>
			</tr>
		</tbody>
	</table>
	<%
End Sub

'---- Will return true if data has been posted from a different website ----
Function CheckForeignSitePost()
	Dim lstrReferer, lstrHost
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		lstrReferer = Request.ServerVariables("HTTP_REFERER")
		lstrHost = Request.ServerVariables("HTTP_HOST")
		If InStr(1, lstrReferer, "//" & lstrHost & "/", vbTextCompare) = 0 Then
			CheckForeignSitePost = True
		Else
			CheckForeignSitePost = False
		End if
	Else
		CheckForeignSitePost = False
	End If
End Function

'----  ----
Function CheckDuplicateAS400User(Var, UserID)
	If Var <> "" Then
		Dim rsUser
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		If UserID <> "" Then
			strSql = _
				"SELECT UMUSCD " &_
				"FROM USERMST " &_
				"WHERE UMRCNU <> " & UserID & " " &_
				"AND UMUSCD LIKE '" & Var & "'"
		Else
			strSql = _
				"SELECT UMUSCD " &_
				"FROM USERMST " &_
				"WHERE UMUSCD LIKE '" & Var & "'"
		End If
		rsUser.Open strSql, conRefron, adOpenForwardOnly, adLockReadOnly, adCmdText
		If Not rsUser.EOF Then
			rsUser.Close
			Set rsUser = Nothing
			CheckDuplicateAS400User = True
		Else
			rsUser.Close
			Set rsUser = Nothing
			CheckDuplicateAS400User = False
		End If
	Else
		CheckDuplicateAS400User = True
	End If
End Function

'----  ----
Function CheckAS400User(Var)
	If Var <> "" Then
		Dim rsUser
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		strSql = _
			"SELECT UMUSCD " &_
			"FROM USERMST " &_
			"WHERE UMUSCD LIKE '" & Var & "'"
		ConnectToDb("Refron")
		rsUser.Open strSql, conRefron, adOpenForwardOnly, adLockReadOnly, adCmdText
		If rsUser.EOF Then
			rsUser.Close
			Set rsUser = Nothing
			CheckAS400User = True
		Else
			rsUser.Close
			Set rsUser = Nothing
			CheckAS400User = False
		End If
	Else
		CheckAS400User = True
	End If
End Function

'---- Returns true if the manufacturer code is not found in the manufacturer master ----
Function CheckManufacturer(Var)
	If Var <> "" Then
		Dim rsTemp
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSql = _
			"SELECT MNMFCD " &_
			"FROM MFRMST " &_
			"WHERE MNMFCD LIKE '" & Var & "'"
		rsTemp.Open strSql, conRefron, adOpenForwardOnly, adLockReadOnly, adCmdText
		If rsTemp.EOF Then
			rsTemp.Close
			Set rsTemp = Nothing
			CheckManufacturer = True
		Else
			rsTemp.Close
			Set rsTemp = Nothing
			CheckManufacturer = False
		End If
	Else
		CheckManufacturer = True
	End If
End Function

'---- Returns true if the class code is not found in the class master ----
Function CheckClass(Var)
	If Var <> "" Then
		Dim rsTemp
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSql = _
			"SELECT CSCLAS " &_
			"FROM CLASMST " &_
			"WHERE CSCLAS LIKE '" & Var & "'"
		rsTemp.Open strSql, conRefron, adOpenForwardOnly, adLockReadOnly, adCmdText
		If rsTemp.EOF Then
			rsTemp.Close
			Set rsTemp = Nothing
			CheckClass = True
		Else
			rsTemp.Close
			Set rsTemp = Nothing
			CheckClass = False
		End If
	Else
		CheckClass = True
	End If
End Function

'---- Returns true if the class code is not found in the class master ----
Function CheckOffice(Var)
	If Var <> "" Then
		Dim rsTemp
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSql = _
			"SELECT WFID " &_
			"FROM WOFFICE " &_
			"WHERE WFID = " & Var
		ConnectToDb("Refron")
		rsTemp.Open strSql, conRefron, adOpenForwardOnly, adLockReadOnly, adCmdText
		If rsTemp.EOF Then
			rsTemp.Close
			Set rsTemp = Nothing
			CheckOffice = True
		Else
			rsTemp.Close
			Set rsTemp = Nothing
			CheckOffice = False
		End If
	Else
		CheckOffice = True
	End If
End Function

'---- Returns true if the item code is not found in the item master ----
Function CheckItem(Var)
	If Var <> "" Then
		Dim rsTemp
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSql = _
			"SELECT ITITEM " &_
			"FROM ITEMMST " &_
			"WHERE ITITEM LIKE '" & Var & "'"
		ConnectToDb("Refron")
		rsTemp.Open strSql, conRefron, adOpenForwardOnly, adLockReadOnly, adCmdText
		If rsTemp.EOF Then
			rsTemp.Close
			Set rsTemp = Nothing
			CheckItem = True
		Else
			rsTemp.Close
			Set rsTemp = Nothing
			CheckItem = False
		End If
	Else
		CheckItem = True
	End If
End Function

'---- Returns true if the variable is blank ----
Function CheckBlank(Var)
	If Var = "" Or IsEmpty(Var) Or IsNull(Var) Then
		CheckBlank = True
	Else
		CheckBlank = False
	End If
End Function

'---- Returns true if the variable is not a number ----
Function CheckNum(Var)
	If IsNumeric(Var) Then
		CheckNum = False
	Else
		CheckNum = True
	End If
End Function			
			
'---- Returns true if the variable is greater than the maximum ----
Function CheckMaxNum(Var,Max)
	If Var <= Max Then
		CheckMaxNum = False
	Else
		CheckMaxNum = True
	End If
End Function

'---- Returns true if the variable is less than the minimum----
Function CheckMinNum(Var,Min)
	If Var >= Min Then
		CheckMinNum = False
	Else
		CheckMinNum = True
	End If
End Function

'---- Returns true if the variable is not a date ----
Function CheckDate(Var)
	If IsDate(Var) Then
		If CDate(Var) >= CDate("01/01/1753") And CDate(Var) <= CDate("12/31/9999") Then
			CheckDate = False
		Else
			CheckDate = True
		End If
	Else
		CheckDate = True
	End If
End Function

'---- Returns ture if the variable has more characters than the maximum ----
Function CheckMaxLength(Var,Max)
	If Len(Var) <= Max Then
		CheckMaxLength = False
	Else
		CheckMaxLength = True
	End If
End Function

'---- Returns true if the variable has fewer characters than the minimum ----
Function CheckMinLength(Var,Min)
	If Len(Var) >= Min Then
		CheckMinLength = False
	Else
		CheckMinLength = True
	End If
End Function
%>