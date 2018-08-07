<%
Function ReturnUsername()
	Dim regEx, Matches
	Set regEx = New RegExp
	regEx.Pattern = "\w+"
	regEx.IgnoreCase = True
	regEx.Global = True
	Set Matches = regEx.Execute(Request.ServerVariables("REMOTE_USER"))
	If Matches.Count = 0 Then
		ReturnUserName = ""
	Else
		ReturnUserName = UCase(Left(Matches(Matches.Count-1).Value,1)) & LCase(Mid(Matches(Matches.Count-1).Value,2))
	End If
End Function

Function ReturnUserRecord(conDB)
	Dim rsUser
	Set rsUser = Server.CreateObject("ADODB.Recordset")
	strSql = "SELECT * FROM USERMST WHERE UPPER(UMNNAM) = '" & UCase(ReturnUserName()) & "'"
	rsUser.Open strSql, conDB, adOpenForwardOnly, adLockReadOnly
	Set ReturnUserRecord = rsUser
End Function

Function ReturnUserField(conDB, strField)
	Dim rsUser
	ConnectToDb("Refron")
	Set rsUser = ReturnUserRecord(conRefron)
	If rsUser.EOF Then
		ReturnUserField = ""
	Else
		ReturnUserField = Trim(rsUser(strField))
	End If
	rsUser.Close
	Set rsUser = Nothing
End Function
%>