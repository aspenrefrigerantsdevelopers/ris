if(top == self) 
	{ 
	var sContainer = "/default.asp";
	var sURL = "";
	var sQuery = "";
	switch (unescape(window.location.pathname).toLowerCase())
		{
		case '/menu.asp':
			break;
		case '/banner.asp':
			break;
		case '/copyright.asp':
			break;
		default:
			sURL = "?URL=" + unescape(window.location.pathname);
			if (unescape(window.location.search) != '')
				{
				sQuery = "&" + unescape(window.location.search);
				}
			break;
		}
	sQuery = sQuery.substring(1,sQuery.length);
	var sGotoURL = sContainer + sURL + sQuery;
	var oAppVer = navigator.appVersion;
	var bIsNetscape = (navigator.appName == 'Netscape') && ((oAppVer.indexOf('3') != -1) || (oAppVer.indexOf('4') != -1));
	var bIsMicrosoftIE = (oAppVer.indexOf('MSIE 4') != -1);
	if (bIsNetscape || bIsMicrosoftIE)
		{
		location.replace(sGotoURL);
		}
	else
		{
		location.href = sGotoURL;
		}
	}