<iframe style="display:none;position:absolute;width:148;height:194;z-index=100" id="CalFrame" marginheight="0" marginwidth="0" noresize frameborder="0" scrolling="NO" src="/shared/global/client_scripts/calendar/calendar.htm" onMouseOut="CloseCal()"></iframe>
<script language="JavaScript" src="/shared/global/client_scripts/calendar/calendar.js"></script>

<%
Sub CreateCal(varForm, varTextBox, varDefault, varNum, varFrom, varTo)
	If InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") Then
		If varDefault = "" Then varDefault = Date()
		If varNum = "" Then varNum = 1
		If varFrom = "" Then varFrom = "1/1/1753"
		If varTo = "" Then varTo = "1/1/2999"
		Response.Write "<input type=hidden name=""HiddenCalDate" & varNum & """ id=""HiddenCalDate" & varNum & """ value='" & varDefault & "'>"
		Response.Write "<a target=_self href=""JavaScript:ShowCalendar(document." & varForm & ".calimg" & varNum & ", document." & varForm & "." & varTextBox & ", document." & varForm & ".HiddenCalDate" & varNum & ", '" & varFrom & "', '" & varTo & "')"" onClick=""event.cancelBubble=true;""><img align=absmiddle border=0 id=calimg" & varNum & " src=""/shared/global/client_scripts/calendar/calendar.gif"" style=""POSITION: relative""></a>"
	End If
End Sub
%>
