if(top != self)
	{
	var sThisURL = unescape(location.pathname);
	var sThisQuery = unescape(location.search);
	var sGotoURL = sThisURL + sThisQuery;
	var oAppVer = navigator.appVersion;
	var bIsNetscape = (navigator.appName == 'Netscape') && ((oAppVer.indexOf('3') != -1) || (oAppVer.indexOf('4') != -1));
	var bIsMicrosoftIE = (oAppVer.indexOf('MSIE 4') != -1);
	if (bIsNetscape || bIsMicrosoftIE)
		{
		top.location.replace(sGotoURL);
		}
	else
		{
		top.location.href = sGotoURL;
		}
	}