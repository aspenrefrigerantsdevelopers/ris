<script runat="server" language="jscript">
	function getTimer()
		{
		var T = new Date();
		return (T.getHours() * 3600000) + (T.getMinutes() * 60000) + (T.getSeconds() * 1000) + (T.getMilliseconds());
		}
</script>

<%
'Put in code...
'First event
'SetTimer("Description #1")
'Second event
'SetTimer("Description #2")
'Before footer
'Call WriteTimers()

Dim Timers(30,1), TimerCount, TotalTime

TimerCount = -1
SetTimer("Start Timer")

Sub SetTimer(Description)
	TimerCount = TimerCount + 1
	Timers(TimerCount, 0) = Description
	Timers(TimerCount, 1) = getTimer()
End Sub

Sub WriteTimers()
	SetTimer("End Timer")
	Dim Timer, TotalTime, CurrentTime, TimerRow
	TotalTime = (Timers(TimerCount, 1) - Timers(0, 1))
	TimerRow = 1
	With Response
		.Write "<br>"
		.Write "<table>"
			.Write "<thead>"
				.Write "<tr>"
					.Write "<th nowrap>Description</th>"
					.Write "<th nowrap align=""right"">&nbsp;&nbsp;&nbsp;Milliseconds</th>"
					.Write "<th nowrap align=""right"">&nbsp;&nbsp;&nbsp;% of Total</th>"
					.Write "<th width=""100%"">&nbsp;</th>"
				.Write "</tr>"
			.Write "</thead>"
			.Write "<tbody>"
	End With
	For Timer = 1 To TimerCount
		CurrentTime = (Timers(Timer, 1) - Timers((Timer - 1), 1))
		With Response
				.Write "<tr class=""Row" & TimerRow & """>"
					.Write "<td nowrap>" & Timers(Timer, 0) & "</td>"
					.Write "<td nowrap align=""right"">" & CurrentTime & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
					.Write "<td nowrap align=""right"">" & FormatPercent(((CurrentTime) / TotalTime), 0, True, False, True) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
					.Write "<td width=""100%"">&nbsp;</td>"
				.Write "</tr>"
		End With
		If TimerRow = 1 Then TimerRow = 2 Else TimerRow = 1
	Next
	With Response
			.Write "</tbody>"
			.Write "<tfoot>"
				.Write "<tr>"
					.Write "<td nowrap>Total Time</td>"
					.Write "<td nowrap align=""right"">" & TotalTime & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
					.Write "<td nowrap align=""right"">100%&nbsp;&nbsp;&nbsp;&nbsp;</td>"
					.Write "<td width=""100%"">&nbsp;</td>"
				.Write "</tr>"
			.Write "</tfoot>"
		.Write "</table>"
	End With
End Sub
%>