<%
Sub LogError(source,keyName,keyValue,message)
	On Error Resume Next
		Set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = conRefron
		cmd.CommandType = 4 'adCmdStoredProc            
		cmd.CommandText = "dbo.z_Debug"
		cmd.Parameters.Append cmd.CreateParameter("@Source", adVarChar, adParamInput, 100, source)
		cmd.Parameters.Append cmd.CreateParameter("@KeyName", adVarChar, adParamInput, 50, keyName)
		cmd.Parameters.Append cmd.CreateParameter("@KeyValue", adVarChar, adParamInput, 250, keyValue)
		cmd.Parameters.Append cmd.CreateParameter("@Message", adVarChar, adParamInput, 5000, message)
		cmd.NamedParameters = true
		cmd.Execute
	On Error GoTo 0
	Set cmd = nothing
End Sub
%>
