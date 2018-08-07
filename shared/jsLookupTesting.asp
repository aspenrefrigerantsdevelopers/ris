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

If gIsLive Then
	RISserverName = "intranet.refron.com"
Else
	RISserverName = "gosrvwebdev-web"
End If

'################################################################################################
	
'###########################################################################################################
'	this call will auto update the OrderItems_xxx.js file to a new version
'Call BuildNewItemLookupFile
'###########################################################################################################

'Response.Write("New OrderItems js lookup file created")
'Response.Write(Date())
'	Dim Page
'	Set Page = Server.CreateObject("RefronIntranetPage.WSC")
'	Page.Title = "Airgas Intranet: Banner"
'	Page.AddHeader("<script language=""JavaScript"" type=""text/javascript"" src=""/shared/global/client_scripts/get_in_frames.js""></script>")
'	Page.PrintHeader()

%>
<script language="JavaScript">
	<!--
	var timerID ;

	function tzone(tz, os, ds, cl)
		{
		this.ct = new Date(0) ;		// datetime
		this.tz = tz ;				// code
		this.os = os ;				// GMT offset
		this.ds = ds ;				// has daylight savings
		this.cl = cl ;				// font color
		}

	function UpdateClocks()
		{
		var ct = new Array(
			new tzone('PST: ', -8, 0, 'pink'),
			new tzone('MST: ', -7, 0, 'cyan'),
			new tzone('CST: ', -6, 0, 'lime'),
			new tzone('EST: ', -5, 0, 'yellow')
			);

		var dt = new Date(); // [GMT] time according to machine clock
//	alert(dt);
//	alert(dt.getHours() + ' - ' + dt.getUTCHours());

		var startDST = new Date(dt.getFullYear(), 3, 1);
//	alert(startDST);
		while (startDST.getDay() != 0)
			startDST.setDate(startDST.getDate() + 1) ;

		var endDST = new Date(dt.getFullYear(), 9, 31) ;
//	alert(endDST);
		while (endDST.getDay() != 0)
			endDST.setDate(endDST.getDate() - 1) ;

		var ds_active ;		// DS currently active

//	alert(startDST + '~' + endDST);
		if (startDST < dt && dt < endDST)
			ds_active = 1 ;
		else
			ds_active = 0 ;
//	alert(ds_active);

		// Adjust each clock offset if that clock has DS and in DS.
		for(n=0 ; n<ct.length ; n++)
			if (ct[n].ds == 1 && ds_active == 1) ct[n].os++ ;

		// compensate time zones
		gmdt = new Date() ;

		for (n=0 ; n<ct.length ; n++)
			ct[n].ct = new Date(gmdt.getTime() + ct[n].os * 3600 * 1000) ;

		document.all.Clock0.innerHTML =	'<font color="' + ct[0].cl + '">' + ct[0].tz + ClockString(ct[0].ct, 'ShortTime') + '</font>';
		document.all.Clock1.innerHTML =	'<font color="' + ct[1].cl + '">' + ct[1].tz + ClockString(ct[1].ct, 'ShortTime') + '</font>';
		document.all.Clock2.innerHTML =	'<font color="' + ct[2].cl + '">' + ct[2].tz + ClockString(ct[2].ct, 'ShortTime') + '</font>';
		document.all.Clock3.innerHTML =	'<font color="' + ct[3].cl + '">' + ct[3].tz + ClockString(ct[3].ct, 'ShortTime') + '</font>';

		timerID = window.setTimeout("UpdateClocks()", 1001);
		}

	function ClockString(dt, mode)
		{
		var stemp, ampm;
		var dt_year = dt.getUTCFullYear();
		var dt_month = dt.getUTCMonth() + 1;
		var dt_day = dt.getUTCDate();
		var dt_hour = dt.getUTCHours();
		var dt_minute = dt.getUTCMinutes();
		var dt_second = dt.getUTCSeconds();
	//	alert(dt);
	//	alert(dt.getHours());
		dt_year = dt_year.toString();

		if (0 <= dt_hour && dt_hour < 12)
			{
			ampm = 'AM';
			if (dt_hour == 0) dt_hour = 12;
			}
		else
			{
			ampm = 'PM';
			dt_hour = dt_hour - 12;
			if (dt_hour == 0) dt_hour = 12;
			}

		if (dt_minute < 10)
			{
			dt_minute = '0' + dt_minute;
			}

		if (dt_second < 10)
			{
			dt_second = '0' + dt_second;
			}

		switch (mode)
			{
			case 'DateTime':
				stemp = dt_month + '/' + dt_day + '/' + dt_year.substr(2,2) + ' ' + dt_hour + ":" + dt_minute + ":" + dt_second + ' ' + ampm;
				break;
			case 'LongTime':
				stemp = dt_hour + ":" + dt_minute + ":" + dt_second + ' ' + ampm;
				break;
			case 'ShortTime':
				stemp = dt_hour + ":" + dt_minute + ' ' + ampm;
				break;
			default:
				stemp = dt_month + '/' + dt_day + '/' + dt_year.substr(2,2) + ' ' + dt_hour + ":" + dt_minute + ":" + dt_second + ' ' + ampm;
				break;
			}
		return stemp;
		}
	//-->
</script>
<body onLoad="UpdateClocks()">
<table>
	<tr style="background-color:Black;">
		<td nowrap class="BannerClockCell"><%=FormatDateTime(Date(),vbLongDate)%></td>
		<td nowrap id="Clock0" class="BannerClockCell">&nbsp;</td>
		<td nowrap id="Clock1" class="BannerClockCell">&nbsp;</td>
		<td nowrap id="Clock2" class="BannerClockCell">&nbsp;</td>
		<td nowrap id="Clock3" class="BannerClockCell">&nbsp;</td>
	</tr>
</table>
</body>
