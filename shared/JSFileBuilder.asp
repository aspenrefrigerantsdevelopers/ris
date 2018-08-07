<!--#include virtual="/shared/ErrorLogging.asp"-->
<!--Sub BuildNewItemLookupFileJSON-->
<%
Sub BuildNewItemLookupFileJSON()
	LogError "JSFileBuider.asp","Info","Line Number: 5", Session.Contents("LogonUser")
	Dim cmd,rs,fs,jsFile
	Dim bulkItems,itemList,fileVersion
	Dim addComma,bulkComplete,firstTable
	fileVersion = 0
	firstTable = True
	addComma = False
	bulkComplete = False

	On Error Resume Next
		Set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = conRefron
		cmd.CommandType = 4 'adCmdStoredProc            
		cmd.CommandText = "dbo.JavascriptLookupFile_OrderItems_JSON"
		cmd.Parameters.Append cmd.CreateParameter("@outFileVersion",adInteger,adParamOutput, , fileVersion)
		Set rs = cmd.Execute

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","Line Number: 23", Err.description
			Err.Clear()
		End If

		Do While Not rs.EOF
			'first resultset is bulkItems
			If addComma Then
				bulkItems = bulkItems + "," + rs(0)
			Else
				bulkItems = bulkItems + rs(0)
				addComma = True
			End If
			rs.MoveNext
		Loop

		Set rs = rs.NextRecordset()
		addComma = False

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","Line Number: 42", Err.description
			Err.Clear()
		End If

		Do While Not rs.EOF
			'second resultset
			If addComma Then
				itemList = itemList + "," + rs(0)
			Else
				itemList = itemList + rs(0)
				addComma = True
			End If
			rs.MoveNext
		Loop

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","Line Number: 58", Err.description
			Err.Clear()
		End If

		'NOTE: the output prarameter is not available until all result sets have been interated over
		fileVersion = cmd.Parameters("@outFileVersion")
		LogError "JSFileBuider.asp","Info","Line Number: 64", fileVersion

		jsFile =  Server.MapPath("\shared\jsLookup\OrderItemsJSON_" & fileVersion & ".js") 
		LogError "JSFileBuider.asp","Info","Line Number: 67", jsFile

		Set fs = CreateObject("Scripting.FileSystemObject")
		Set jsFile = fs.CreateTextFile(jsFile, True)
		Dim BulkFunction : BulkFunction = "function BulkItemsJSON(){return [" + bulkItems + "];}"
		Dim ItemsFunction : ItemsFunction = "function OrderItemsJSON(){return [" + itemList + "];}"
		jsFile.WriteLine(BulkFunction)
		jsFile.WriteLine("")
		jsFile.WriteLine(ItemsFunction)
		jsFile.Close

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","Line Number: 79", Err.description
			Err.Clear()
		End If
		rs.Close
	On Error GoTo 0
	Set jsFile = nothing
	Set fs = nothing
	Set cmd = nothing
	Set rs = nothing
End Sub
%>
<!--Sub BuildNewBankDepositItemLookupFileJSON-->
<%
Sub BuildNewBankDepositItemLookupFileJSON()
	Dim cmd,rs,fs,jsFile
	Dim items,fileVersion,addComma
	fileVersion = 0
	addComma = False

	On Error Resume Next
		Set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = conRefron
		cmd.CommandTimeout = 10
		cmd.CommandType = 4 'adCmdStoredProc            
		cmd.CommandText = "dbo.JavascriptLookupFile_BankingDepositItems_JSON"
		cmd.Parameters.Append cmd.CreateParameter("@outFileVersion",adInteger,adParamOutput, , fileVersion)
		Set rs = cmd.Execute

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","1-BuildNewBankDepositItemLookupFileJSON", Err.description
			Err.Clear()
		End If

		Do While Not rs.EOF
			'first resultset is BankingDepositItems
			If addComma Then
				items = items + "," + rs(0)
			Else
				items = items + rs(0)
				addComma = True
			End If
			rs.MoveNext
		Loop
	
		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","2-BuildNewBankDepositItemLookupFileJSON", Err.description
			Err.Clear()
		End If

		'NOTE: the output prarameter is not available until all result sets have been interated over
		fileVersion = cmd.Parameters("@outFileVersion")
		jsFile =  Server.MapPath("\shared\jsLookup\BankDepositItemsJSON_" & fileVersion & ".js") 

		Set fs = CreateObject("Scripting.FileSystemObject")
		Set jsFile = fs.CreateTextFile(jsFile, True)
		Dim BankDepositItems : BankDepositItems = "function BankDepositItemsJSON(){return [" + items + "];}"
		jsFile.WriteLine(BankDepositItems)
		jsFile.Close

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","3-BuildNewBankDepositItemLookupFileJSON", Err.description
			Err.Clear()
		End If
		rs.Close
	On Error GoTo 0
	Set jsFile = nothing
	Set fs = nothing
	Set cmd = nothing
	Set rs = nothing
End Sub
%>
<!--Sub BuildNewSAPHouseAccountsFile-->
<%
Sub BuildNewSAPHouseAccountsFile()
	Dim cmd,rs,fs,jsFile
	Dim houseAccounts,fileVersion,addComma
	fileVersion = 0
	addComma = False

	On Error Resume Next
		Set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = conRefron
		cmd.CommandType = 4 'adCmdStoredProc            
		cmd.CommandText = "dbo.JavascriptLookupFile_SAP_HouseAccounts"
		cmd.Parameters.Append cmd.CreateParameter("@outFileVersion",adInteger,adParamOutput, , fileVersion)
		Set rs = cmd.Execute

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","1-BuildNewSAPHouseAccountsFile", Err.description
			Err.Clear()
		End If

		Do While Not rs.EOF
			'first resultset is houseAccounts
			If addComma Then
				houseAccounts = houseAccounts + "," + rs(0)
			Else
				houseAccounts = houseAccounts + rs(0)
				addComma = True
			End If
			rs.MoveNext
		Loop
	
		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","2-BuildNewSAPHouseAccountsFile", Err.description
			Err.Clear()
		End If

		'NOTE: the output prarameter is not available until all result sets have been interated over
		fileVersion = cmd.Parameters("@outFileVersion")
		jsFile =  Server.MapPath("\shared\jsLookup\SAP_HouseAccts_" & fileVersion & ".js") 

		Set fs = CreateObject("Scripting.FileSystemObject")
		Set jsFile = fs.CreateTextFile(jsFile, True)
		Dim HouseAccts : HouseAccts = "function HouseAccts(){return [" + houseAccounts + "];}"
		jsFile.WriteLine(HouseAccts)
		jsFile.Close

		If Err.number <> 0 then
			LogError "JSFileBuider.asp","Error","3-BuildNewSAPHouseAccountsFile", Err.description
			Err.Clear()
		End If
		rs.Close
	On Error GoTo 0
	Set jsFile = nothing
	Set fs = nothing
	Set cmd = nothing
	Set rs = nothing
End Sub
%>