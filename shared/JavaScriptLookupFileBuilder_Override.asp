<!--#include virtual="/shared/global/server_includes/global.asp"-->
<!--#include virtual="/shared/JSFileBuilder.asp"-->
<%
'####################  ENVIRONMENT  #############################################################
' the shared\global\server_includes\global.asp has a Const variable "gIsLive"
' in the Dev environment this is set to False
' in the Live environment this is set to True
' this provides a way to deploy pages without commenting code due to the environment.
' use this section to manage the variables required for the page to respond to the correct environment

Dim RISserverName
Dim RRSserverName

If gIsLive Then
	RISserverName = "intranet.refron.com"
	RRSserverName = "rrs"
Else
	RISserverName = "gosrvwebdev-web"
	RRSserverName = "gosrvwebdev-web/rrs"
End If

Dim cmd,rs
'lookup file versions
Dim BankDepositItemLookupVersion
Dim OrderItemLookupVersion
Dim SAPHouseAccountsLookupVersion
Dim selFileTypeSupported : selFileTypeSupported = ""

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		'form was submitted
		Dim selFileType
		Dim errorMsg : errorMsg = ""
		selFileType = Request.Form("selFileType")
		If selFileType = "I" Then
			Call BuildNewItemLookupFileJSON
			selFileTypeSupported = True
		End If
		If selFileType = "D" Then
			Call BuildNewBankDepositItemLookupFileJSON
			selFileTypeSupported = True
		End If
		If selFileType = "H" Then
			Call BuildNewSAPHouseAccountsFile
			selFileTypeSupported = True
		End If

		If selFileTypeSupported = "" Then
			selFileTypeSupported = False
		End If
	Else
		'form GET request
	End If

'#####################################################################################################
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = conRefron
	cmd.Prepared = True
	cmd.CommandType = adCmdText
	
	'OrderItems_<version>.js version call required to use the current list of order items
	cmd.CommandText = "SELECT [OrderItemLookupJsFileVersion],[SAPHouseAccountsLookupJsFileVersion],[BankDepositItemLookupJsFileVersion]FROM [Refron].[dbo].[_Config];"
	cmd.ActiveConnection = conRefron
	cmd.CommandTimeout = 10
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.CursorLocation = adUseClient	
	
	'cmd.ActiveConnection.Open
	cmd.CommandTimeout = 10
	rs = cmd.Execute(cmd.CommandText,,adCmdText)
	
	'rs.Open cmd, , adOpenStatic, adLockReadOnly, adCmdText
	BankDepositItemLookupVersion = rs("BankDepositItemLookupJsFileVersion")
	OrderItemLookupVersion = rs("OrderItemLookupJsFileVersion")
	SAPHouseAccountsLookupVersion = rs("SAPHouseAccountsLookupJsFileVersion")
	'rs.Close
	Set rs = Nothing
	Response.Write("BankDepositItemLookupVersion=" & BankDepositItemLookupVersion & "<br>")
	Response.Write("OrderItemLookupVersion=" & OrderItemLookupVersion & "<br>")
	Response.Write("SAPHouseAccountsLookupVersion=" & SAPHouseAccountsLookupVersion & "<br>")
'#####################################################################################################

	'Response.Redirect(Request.QueryString("returnUrl"))
%>
<html>
<head>
	<title>Lookup File Builder</title>
	<link rel='Stylesheet' type='text/css' href='http://<%=RISserverName%>/shared/global/styles/jquery-ui.min.css' />
	<link rel='Stylesheet' type='text/css' href='http://<%=RISserverName%>/shared/global/styles/jquery-autocomplete.css' />
	<script src='http://<%=RISserverName%>/shared/global/client_scripts/jquery-1.8.3.min.js' type="text/javascript"></script>
	<script src='http://<%=RISserverName%>/shared/global/client_scripts/jquery-ui.js' type="text/javascript"></script>
	<script src="http://<%=RISserverName%>/shared/jsLookup/BankDepositItems_<%=BankDepositItemLookupVersion%>.js" type="text/javascript"></script>
	<script src="http://<%=RISserverName%>/shared/jsLookup/OrderItems_<%=OrderItemLookupVersion%>.js" type="text/javascript"></script>
	<script src="http://<%=RISserverName%>/shared/jsLookup/SAP_HouseAccts_<%=SAPHouseAccountsLookupVersion%>.js" type="text/javascript"></script>
<script type="text/javascript">
	$(function ()
	{
		var depositItems;
		depositItems = BankDepositItems();
		$("#tbDepositItems").autocomplete({
			source: function (request, response)
			{
				var results = $.ui.autocomplete.filter(depositItems, request.term);
				response(results.slice(0, 15));
				$("ul.ui-menu").width("370px");
				$("ul.ui-menu").children().addClass('autocomplete-bgcolor');
			}
		});
		var orderItems;
		orderItems = OrderItems();
		$("#tbOrderItems").autocomplete({
			source: function (request, response)
			{
				var results = $.ui.autocomplete.filter(orderItems, request.term);
				response(results.slice(0, 15));
				$("ul.ui-menu").width("370px");
				$("ul.ui-menu").children().addClass('autocomplete-bgcolor');
			}
		});
		var sapHouseAccts;
		sapHouseAccts = HouseAccts();
		$("#tbSapHouseAccts").autocomplete({
			source: function (request, response)
			{
				var results = $.ui.autocomplete.filter(sapHouseAccts, request.term);
				response(results.slice(0, 15));
				$("ul.ui-menu").width("370px");
				$("ul.ui-menu").children().addClass('autocomplete-bgcolor');
			}
		});
	});
</script>
</head>
<body>
	<form name="Override" id="Override" action="JavaScriptLookupFileBuilder_Override.asp" method="post">
		<table>
			<tr>
				<th colspan="3">JavaScript Lookup File Builder</th>
			</tr>
			<tr>
				<td>
					File Type:
				</td>
				<td>
					<select id="selFileType" name="selFileType">
						<option value="D">Bank Deposit Items</option>
						<option value="B">Bulk Items</option>
						<option value="I">Order Items</option>
					</select>
				</td>
				<td>
					<input id="btnBuildFile" type="submit" value="Build File" />
				</td>
			</tr>
			<tr>
				<td colspan="3">
					<hr />
				</td>
			</tr>
			<tr>
				<td>
					Bank Deposit Items:
				</td>
				<td colspan="2">
					<input id="tbDepositItems" name="tbDepositItems" type="text" size="30" />
				</td>
			</tr>
			<tr>
				<td>
					Order Items:
				</td>
				<td colspan="2">
					<input id="tbOrderItems" name="tbOrderItems" type="text" size="30" />
				</td>
			</tr>
			<tr>
				<td>
					SAP House Accts:
				</td>
				<td colspan="2">
					<input id="tbSapHouseAccts" name="tbSapHouseAccts" type="text" size="30" />
				</td>
			</tr>
		</table>
		<%
		If selFileTypeSupported = False Then
			Response.Write("File Type Not Supported Yet<br>")
		End If
		%>
	</form>
</body>
</html>