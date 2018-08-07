<%
Sub UpdateField(rsTemp, DBField, FieldType, NewValue)
	Select Case FieldType
		Case "N":
			If IsNull(NewValue) Then
				'If it is null, set the value to 0.
				'Response.Write "<br>Setting null " & DBField & " field to zero."
				rsTemp(DBField) = 0
			ElseIf Trim(NewValue) = "" Then
				'If it is a blank numeric field, set the value to 0.
				'Response.Write "<br>Setting " & DBField & " field to zero."
				rsTemp(DBField) = 0
			ElseIf Not IsNumeric(NewValue) Then
				'Do nothing since numeric fields require numeric values.
				'Response.Write "<br>Not numeric, not updating."
			ElseIf CDbl(Trim(NewValue)) = CDbl(rsTemp(DBField)) Then
				rsTemp(DBField) = CDbl(Trim(NewValue))
				'They match so don't change.
				'Response.Write "<br>Field " & DBField & " = " & NewValue & ", not updating."
			Else
			'Removed last check since it is always true if 2nd If is false.
			'ElseIf CDbl(Trim(NewValue)) <> CDbl(rsTemp(DBField)) Then
				'Set field equal to new value.
				'Response.Write "<br>Updating the field " & DBField & " = " & NewValue & "."
				rsTemp(DBField) = CDbl(Trim(NewValue))
			End If
		Case "A":
		    IF rsTemp(DBField) = "" Then
		        rsTemp(DBField) = ""
		    End IF
			If IsNull(NewValue) Then
				'If it is null, set the value to blank.
				'Response.Write "<br>Setting null " & DBField & " field to blank."
				rsTemp(DBField) = ""
			ElseIf Trim(NewValue) = Trim(rsTemp(DBField)) Then
				'They match so don't change.
				'Response.Write "<br>Field " & DBField & " = " & NewValue & ", not updating."
			Else
			'Removed last check since it is always true if 2nd If is false.
			'ElseIf CDbl(Trim(NewValue)) <> CDbl(rsTemp(DBField)) Then
				'Set field equal to new value.
				'Response.Write "<br>Updating the field " & DBField & " = " & NewValue & "."
				rsTemp(DBField) = Trim(NewValue)
			End If
	End Select
End Sub

'Increments the counter and creates a unique number for new contact records
Function GetNewRecNum(strKey, strCounter)
	Dim rsCounter, newNum, newPrefix
	Set rsCounter = Server.CreateObject("ADODB.Recordset")
	'Get the current number from the counter file
	strSql = "SELECT " & Trim(strCounter) & " FROM COUNTER WHERE CZKEY = '" & Trim(strKey) & "'"
	ConnectToDb("Refron")
	rsCounter.Open strSql, conRefron, adOpenKeySet, adLockPessimistic
	If Not rsCounter.EOF Then
		Select Case UCase(Trim(strKey))
			Case "CONTACT#"
				'Increment the current number to create the new number
				newNum = CDbl(rsCounter(strCounter)) + 1 
				'Save the new number to the counter file
				rsCounter(strCounter) = newNum
				'Update, close and release the counter file for use by other programs
				rsCounter.Update
				rsCounter.Close
				Set rsCounter = Nothing
				'Check if the new number is less than six digits long and add a leading zero
				If Len(CStr(newNum)) < 6 Then
					newNum = "0" & CStr(newNum)
				Else
					newNum = CStr(newNum)
				End If
				'Create a new prefix (check digit) for the record number
				newPrefix = CStr( 10 + CInt(Mid(newNum,1,1)) + CInt(Mid(newNum,2,1)) + CInt(Mid(newNum,3,1)) + CInt(Mid(newNum,4,1)) + CInt(Mid(newNum,5,1)) + CInt(Mid(newNum,6,1)) )
				'Return the new prefix and number
				GetNewRecNum = CDbl( newPrefix & newNum )
			Case Else
				'Increment the current number to create the new number
				newNum = CDbl(rsCounter(strCounter)) + 1 
				'Save the new number to the counter file
				rsCounter(strCounter) = newNum
				'Update, close and release the counter file for use by other programs
				rsCounter.Update
				rsCounter.Close
				Set rsCounter = Nothing
				'Return the new prefix and number
				GetNewRecNum = newNum
		End Select
	End If
End Function
%>