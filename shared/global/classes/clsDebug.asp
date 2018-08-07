<%
Class clsDebug

	'Put in Header
	'Dim Debug
	'Set Debug = New clsDebug

	'Use in Code
	'Debug.Print("Label", OutputVariable)

	'Put in Footer
	'Debug.End
	'Set Debug = Nothing
	
    Private dteRequestTime
    Private dteFinishTime
    Private errLastError
    Private objVariableStorage
    Private objSummaryStorage
    Private objLastErrorStorage
	
    Private Sub Class_Initialize()
		Set errLastError = Server.GetLastError()
		blnEnabled = False
        dteRequestTime = Now()
        Set objVariableStorage =  Server.CreateObject("Scripting.Dictionary")
        Set objSummaryStorage = Server.CreateObject("Scripting.Dictionary")
        Set objLastErrorStorage = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
		Set errLastError = Nothing
        objVariableStorage.RemoveAll
        Set objVariableStorage = Nothing
        objSummaryStorage.RemoveAll
        Set objSummaryStorage = Nothing
        objLastErrorStorage.RemoveAll
        Set objLastErrorStorage = Nothing
    End Sub

    Private Function ReturnCollection(Byval Name, Byval Collection)
        Dim strCollection, varItem
        If Collection.Count = 0 Then
			strCollection = ""
		Else
			strCollection = vbNewLine & "<tr><td colspan=""3"" style=""background-color:white; border-bottom:1px solid black;"">&nbsp;</td></tr>"
			strCollection = strCollection & vbNewLine & "<tr><th colspan=""3"" class=""ErrorHead"">" & Name & "</th></tr>"
	        For Each varItem In Collection
				strCollection = strCollection & vbNewLine & "<tr>"
				strCollection = strCollection & "<td class=""ErrorCell"">" & varItem & "</td>"
				strCollection = strCollection & "<td class=""ErrorCell"">" & ReturnType(Collection(varItem)) & "</td>"
				strCollection = strCollection & "<td class=""ErrorCell"">" & ReturnValue(Collection(varItem)) & "</td>"
		        strCollection = strCollection & "</tr>"
	        Next
		End If
		ReturnCollection = strCollection
    End Function

    Private Function ReturnType(Byval Item)
		Select Case VarType(Item)
			Case vbEmpty
				ReturnType = "Empty"
			Case vbNull
				ReturnType = "Null"
			Case vbInteger
				ReturnType = "Integer"
			Case vbLong
				ReturnType = "Long"
			Case vbSingle
				ReturnType = "Single"
			Case vbDouble
				ReturnType = "Double"
			Case vbCurrency
				ReturnType = "Currency"
			Case vbDate
				ReturnType = "Date"
			Case vbString
				ReturnType = "String"
			Case vbObject
				ReturnType = "Object"
			Case vbError
				ReturnType = "Error"
			Case vbBoolean
				ReturnType = "Boolean"
			Case vbVariant
				ReturnType = "Variant"
			Case vbDataObject
				ReturnType = "Data Object"
			Case vbDecimal
				ReturnType = "Decimal"
			Case vbByte
				ReturnType = "Byte"
			Case vbArray
				ReturnType = "Array"
			Case Else
				ReturnType = "Unknown - " & VarType(Item)
		End Select
	End Function

    Private Function ReturnValue(Byval Item)
		Select Case VarType(Item)
			Case vbEmpty
				ReturnValue = ""
			Case vbNull
				ReturnValue = ""
			Case vbInteger
				ReturnValue = Item
			Case vbLong
				ReturnValue = Item
			Case vbSingle
				ReturnValue = Item
			Case vbDouble
				ReturnValue = Item
			Case vbCurrency
				ReturnValue = Item
			Case vbDate
				ReturnValue = Item
			Case vbString
				If Item = "" Then
					ReturnValue = "&nbsp;"
				ElseIf Left(Item, 1) = "!" Then
					ReturnValue = Mid(Item, 2)
				Else
					'Removed by Ben on 5/29/07 - Not sure why we were replacing "/" with "/<span style=""width:0px;""> </span>"
					'ReturnValue = Replace(Replace(Replace(Server.HTMLEncode(Item), vbNewLine, "<br>"), Chr(10), "<br>"), "/", "/<span style=""width:0px;""> </span>")
					ReturnValue = Replace(Replace(Server.HTMLEncode(Item), vbNewLine, "<br>"), Chr(10), "<br>")
				End If
			Case vbObject
				ReturnValue = ""
			Case vbError
				ReturnValue = ""
			Case vbBoolean
				ReturnValue = Item
			Case vbVariant
				ReturnValue = Item
			Case vbDataObject
				ReturnValue = Item
			Case vbDecimal
				ReturnValue = Item
			Case vbByte
				ReturnValue = Item
			Case vbArray
				ReturnValue = ""
			Case Else
				ReturnValue = Item
		End Select
	End Function

    Private blnEnabled
    Public Property Get Enabled()
        Enabled = blnEnabled
    End Property
    Public Property Let Enabled(newValue)
        blnEnabled = newValue
    End Property

    Private strDebugInfo
    Public Property Get DebugInfo()
        DebugInfo = strDebugInfo
    End Property
    Public Property Let DebugInfo(newValue)
        strDebugInfo = newValue
    End Property

    Public Sub [Print](label, output)
		If Not IsNull(label) Then
			label = CStr(label)
			If Enabled And Len(label) > 0 Then
				If objVariableStorage.Exists(label) Then
					If Len(label) >= 5 Then
						If Mid(label, (Len(label) - 1), 1) = "_" And IsNumeric(Right(label, 1)) Then
							Call [Print](Left(label, (Len(label) - 2)) & "_" & (Right(label, 1) + 1), output)
						ElseIf Mid(label, (Len(label) - 2), 1) = "_" And IsNumeric(Right(label, 2)) Then
							Call [Print](Left(label, (Len(label) - 3)) & "_" & (Right(label, 2) + 1), output)
						ElseIf Mid(label, (Len(label) - 3), 1) = "_" And IsNumeric(Right(label, 3)) Then
							Call [Print](Left(label, (Len(label) - 4)) & "_" & (Right(label, 3) + 1), output)
						ElseIf Mid(label, (Len(label) - 4), 1) = "_" And IsNumeric(Right(label, 4)) Then
							Call [Print](Left(label, (Len(label) - 5)) & "_" & (Right(label, 4) + 1), output)
						Else
							Call [Print](label & "_1", output)
						End If
					ElseIf Len(label) >= 4 Then
						If Mid(label, (Len(label) - 1), 1) = "_" And IsNumeric(Right(label, 1)) Then
							Call [Print](Left(label, (Len(label) - 2)) & "_" & (Right(label, 1) + 1), output)
						ElseIf Mid(label, (Len(label) - 2), 1) = "_" And IsNumeric(Right(label, 2)) Then
							Call [Print](Left(label, (Len(label) - 3)) & "_" & (Right(label, 2) + 1), output)
						ElseIf Mid(label, (Len(label) - 3), 1) = "_" And IsNumeric(Right(label, 3)) Then
							Call [Print](Left(label, (Len(label) - 4)) & "_" & (Right(label, 3) + 1), output)
						Else
							Call [Print](label & "_1", output)
						End If
					ElseIf Len(label) = 3 Then
						If Mid(label, (Len(label) - 1), 1) = "_" And IsNumeric(Right(label, 1)) Then
							Call [Print](Left(label, (Len(label) - 2)) & "_" & (Right(label, 1) + 1), output)
						Else
							Call [Print](label & "_1", output)
						End If
					Else
						Call [Print](label & "_1", output)
					End If
			        'objVariableStorage.Item(label) = output
			    Else
			        objVariableStorage.Add label, output
			    End If
			End If
		End If
    End Sub

    Public Sub AddSummary(label, output)
		If Not IsNull(label) Then
			label = CStr(label)
			If Enabled And Len(label) > 0 Then
				If objSummaryStorage.Exists(label) Then
			        objSummaryStorage.Item(label) = output
			    Else
			        objSummaryStorage.Add label, output
			    End If
			End If
		End If
    End Sub

    Public Sub AddLastError(label, output)
		If Not IsNull(label) Then
			label = CStr(label)
			If Enabled And Len(label) > 0 Then
				If objLastErrorStorage.Exists(label) Then
			        objLastErrorStorage.Item(label) = output
			    Else
			        objLastErrorStorage.Add label, output
			    End If
			End If
		End If
    End Sub

    Public Sub [End]()
        dteFinishTime = Now()
'		If Request.ServerVariables("QUERY_STRING") = "" Then
			Call AddSummary("Current Page", "!<a target=""_blank"" href=""" & Request.ServerVariables("PATH_INFO") & """ style=""color:blue; text-decoration:underline;"">http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO") & "</a>")
'		Else
'			Call AddSummary("Current Page", "<a target=""_blank"" href=""" & Request.ServerVariables("PATH_INFO") & "?" &  Request.ServerVariables("QUERY_STRING") & """ style=""color:blue; text-decoration:underline;"">http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO") & "?" &  Request.ServerVariables("QUERY_STRING") & "</a>")
'		End If

		Call AddSummary("Time Started", dteRequestTime)
		Call AddSummary("Time Finished", dteFinishTime)
		Call AddSummary("Elapsed Time", DateDiff("s", dteRequestTime, dteFinishTime) & " sec")
		Call AddSummary("Request Type", Request.ServerVariables("REQUEST_METHOD"))
		Call AddSummary("Status Code", Response.Status)
		If errLastError.Number <> 0 Then
			Call AddLastError("Catgegory", errLastError.Category)
			Call AddLastError("Description", errLastError.Description)
			Call AddLastError("File", errLastError.File)
			Call AddLastError("Line", errLastError.Line)
			Call AddLastError("Column", errLastError.Column)
			If errLastError.Column > 0 Then
				Call AddLastError("Source", "!" & Left(errLastError.Source, errLastError.Column) & "<span style=""font-weight:bold; color:red; background-color:yellow;"">" & Mid(errLastError.Source, (errLastError.Column + 1), 1) & "</span>" & Mid(errLastError.Source, (errLastError.Column + 2)))
			Else
				Call AddLastError("Source", errLastError.Source)
			End If
			Call AddLastError("ASP Code", errLastError.ASPCode)
			Call AddLastError("ASP Description", errLastError.ASPDescription)
			Call AddLastError("Error #", errLastError.Number)
        End If
        strDebugInfo = _
			"<style>.ErrorHead {text-align:left; background-color:9cb19e; border:1px black solid; padding-right:4px; padding-left:4px; padding-bottom:1px; padding-top:1px;} .ErrorCell {border:1px black solid; background-color:#dadada; padding-right:4px; padding-left:4px; padding-bottom:1px; padding-top:1px;}</style>" & _
			vbNewLine & "<table style=""border:none; FONT-SIZE: 9pt; MARGIN-LEFT: 2%; WIDTH: 96%; MARGIN-RIGHT: 1.9%; FONT-FAMILY: Arial; BORDER-COLLAPSE: collapse"">" & _
			ReturnCollection("SUMMARY INFO:", objSummaryStorage) & _
			ReturnCollection("LAST ASP ERROR:", objLastErrorStorage) & _
			ReturnCollection("VARIABLE STORAGE:", objVariableStorage) & _
			ReturnCollection("QUERYSTRING COLLECTION:", Request.QueryString())
		On Error Resume Next
			strDebugInfo = strDebugInfo & ReturnCollection("FORM COLLECTION:", Request.Form())
		On Error GoTo 0
		strDebugInfo = strDebugInfo & _
			ReturnCollection("SESSION CONTENTS COLLECTION:", Session.Contents()) & _
			ReturnCollection("COOKIES COLLECTION:", Request.Cookies()) & _
			ReturnCollection("SERVER VARIABLES COLLECTION:", Request.ServerVariables()) & _
			ReturnCollection("CLIENT CERTIFICATE COLLECTION:", Request.ClientCertificate()) & _
			ReturnCollection("SESSION STATIC OBJECTS COLLECTION:", Session.StaticObjects()) & _
			ReturnCollection("APPLICATION STATIC OBJECTS COLLECTION:", Application.StaticObjects()) & _
		    vbNewLine & "</table>" & _
	        "<br>"
'			ReturnCollection("APPLICATION CONTENTS COLLECTION:", Application.Contents()) & _
        If blnEnabled _
			And (Session("UserCode") = "BM" Or Session("UserCode") = "BR" Or Session("UserCode") = "WG" Or LCase(Left(Request.ServerVariables("SERVER_NAME"), 3)) = "dev") Then
			'And (LCase(Left(Request.ServerVariables("SERVER_NAME"), 4)) = "temp" _
			'Or LCase(Left(Request.ServerVariables("SERVER_NAME"), 3)) = "dev" _
			'Or LCase(Left(Request.ServerVariables("SERVER_NAME"), 5)) = "wolfe") Then
			Response.Write strDebugInfo
		ElseIf errLastError.Number <> 0 Then
			Response.Clear
			Response.Write "You have encountered an error. The IT department has already been notified about this particular error and they are working to correct it ASAP. Please read the error description below and if you think you understand why it happened you may go <a href=""JavaScript: history.back(1)"" style=""color:blue;"">back</a> and try again."
			Response.Write "<br><br><b>Error Description:</b> " & errLastError.Description
			Response.Write "<br><br>Thank you for your patience with regards to this problem!"
        End if
    End Sub

End Class
%>