<%
'UploadFile DestURL, FileName

Sub UploadFile(DestURL, FileName)
	Const Boundary = "---------------------------0123456789012"
	Const FieldName = "LocalFile"

	Dim FileContents, FormData
	FileContents = GetFile(FileName)
	FormData = BuildFormData(FileContents, Boundary, FileName, FieldName)

	WinHTTPPostRequest DestURL, FormData, Boundary
End Sub

Function BuildFormData(FileContents, Boundary, FileName, FieldName)
	Dim FormData, Pre, Po
	Const ContentType = "application/upload"

	Pre = "--" + Boundary + vbCrLf + mpFields(FieldName, FileName, ContentType)
	Po = vbCrLf + "--" + Boundary + "--" + vbCrLf

	Const adLongVarBinary = 205
	Dim RS: Set RS = CreateObject("ADODB.Recordset")
	RS.Fields.Append "b", adLongVarBinary, Len(Pre) + LenB(FileContents) + Len(Po)
	RS.Open
	RS.AddNew
	Dim LenData
	LenData = Len(Pre)
	RS("b").AppendChunk (StringToMB(Pre) & ChrB(0))
	Pre = RS("b").GetChunk(LenData)
	RS("b") = ""
	    
	LenData = Len(Po)
	RS("b").AppendChunk (StringToMB(Po) & ChrB(0))
	Po = RS("b").GetChunk(LenData)
	RS("b") = ""
	    
	RS("b").AppendChunk (Pre)
	RS("b").AppendChunk (FileContents)
	RS("b").AppendChunk (Po)
	RS.Update
	FormData = RS("b")
	RS.Close
	BuildFormData = FormData
End Function

Function WinHTTPPostRequest(URL, FormData, Boundary)
	Dim Http
	Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
	'Set Http = CreateObject("MSXML2.ServerXMLHTTP")

	Http.Open "POST", URL, False
	Http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" + Boundary
	Http.SetTimeouts 60000, 60000, 60000, 60000
	Http.Send FormData
	WinHTTPPostRequest = Http.ResponseText
End Function

Function mpFields(FieldName, FileName, ContentType)
	Dim MPTemplate
	MPTemplate = "Content-Disposition: form-data; name=""{field}"";" + _
		" filename=""{file}""" + vbCrLf + _
		"Content-Type: {ct}" + vbCrLf + vbCrLf
	Dim Out
	Out = Replace(MPTemplate, "{field}", FieldName)
	Out = Replace(Out, "{file}", FileName)
	mpFields = Replace(Out, "{ct}", ContentType)
End Function

Function GetFile(FileName)
	Dim Stream: Set Stream = CreateObject("ADODB.Stream")
	Stream.Type = 1 'Binary
	Stream.Open
	Stream.LoadFromFile FileName
	GetFile = Stream.Read
	Stream.Close
End Function

Function GetURL(URL)
	Dim Http
	'Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set Http = CreateObject("MSXML2.ServerXMLHTTP")
	  
	Http.Open "GET", URL, False
	Http.SetTimeouts 60000, 60000, 60000, 60000
	Http.Send
	GetURL = Http.ResponseBody
End Function

Function StringToMB(S)
	Dim I, B
	For I = 1 To Len(S)
		B = B & ChrB(Asc(Mid(S, I, 1)))
	Next
	StringToMB = B
End Function
%>