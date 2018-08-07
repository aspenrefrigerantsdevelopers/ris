function setMenus(menuList)
	{
	var menuArray = menuList.split(',');
	var cookieName = null;
	var cookieValue = null;
	for (var i=0; i<menuArray.length; i++)
		{
		cookieName = document.location.pathname + menuArray[i];
		cookieValue = getSubCookie('MenuSections', cookieName);
		if (cookieValue != null && cookieValue != '')
			{
			expandRetract(menuArray[i], document.all(menuArray[i] + '_image'), cookieValue, 2);
			}
		else
			{
			expandRetract(menuArray[i], document.all(menuArray[i] + '_image'), 'hidden', 2);
			}
		}
	}

function expandRetract(obj, image, state, expDays)
	{
	if (expDays == null)
		{
		expDays = 2;
		}
	if (state == 'hidden' || (state == null && document.all(obj).style.visibility == 'visible'))
		{
		document.all(obj).style.visibility = 'hidden';
		document.all(obj).style.display = 'none';
		image.src = '/shared/images/icons/expand.gif';
		setSubCookie('MenuSections', document.location.pathname + obj, 'hidden', expDays);
		}
	else
		{
		document.all(obj).style.visibility = 'visible';
		document.all(obj).style.display = 'inline';
		image.src = '/shared/images/icons/retract.gif';
		setSubCookie('MenuSections', document.location.pathname + obj, 'visible', expDays);
		}
	}

function expandAll(menuList, expDays)
	{
	if (expDays == null)
		{
		expDays = 30;
		}
	var menuArray = menuList.split(',');
	for (var i=0; i<menuArray.length; i++)
		{
		expandRetract(menuArray[i], document.all(menuArray[i] + '_image'), 'visible', expDays);
		}
	}

function retractAll(menuList, expDays)
	{
	if (expDays == null)
		{
		expDays = 30;
		}
	var menuArray = menuList.split(',');
	for (var i=0; i<menuArray.length; i++)
		{
		expandRetract(menuArray[i], document.all(menuArray[i] + '_image'), 'hidden', expDays);
		}
	}

function setCookie(name, value, expDays)
	{
	if (expDays == null)
		{
		expDays = 30;
		}
	var exp = new Date();
	var cookieTimeToLive = exp.getTime() + (expDays * 24 * 60 * 60 * 1000);
	exp.setTime(cookieTimeToLive);
	document.cookie = name + "=" + escape(value) + "; expires=" + exp.toGMTString();
	}

function getCookie(sName)
	{
	var aCookie = document.cookie.split("; ");
	for (var i=0; i < aCookie.length; i++)
		{
		var aCrumb = aCookie[i].split("=");
		if (sName == aCrumb[0])
			{
			return unescape(aCrumb[1]);
			}
		}
	return null;
	}

function deleteCookie(name)
	{
	var exp = new Date();
	exp.setTime (exp.getTime() - 1000000000);  // This cookie is history
	var cval = getCookie (name);
	document.cookie = name + "=" + cval + "; expires=" + exp.toGMTString();    
	}

function setSubCookie(name, subName, value, expDays)
	{
	if (expDays == null)
		{
		expDays = 30;
		}
	var exp = new Date();
	var cookieTimeToLive = exp.getTime() + (expDays * 24 * 60 * 60 * 1000);
	exp.setTime(cookieTimeToLive);

	if (value == getSubCookie(name, subName))
		{
		document.cookie = name + "=" + escape(getCookie(name)) + "; path=/; expires=" + exp.toGMTString();
		}
	else
		{
		if (getCookie(name) != null)
			{
			var subValue = getCookie(name);
			subValue = subValue.replace('|' + subName + '=' + getSubCookie(name, subName), '');
			}
		else
			{
			subValue = '';
			}
		subValue = subValue + '|' + subName + '=' + escape(value);

		document.cookie = name + "=" + escape(subValue) + "; path=/; expires=" + exp.toGMTString();
		}
	}

function getSubCookie(name, subName)
	{
	var aCookie = document.cookie.split("; ");
	for (var i=0; i < aCookie.length; i++)
		{
		var aCrumb = aCookie[i].split("=");
		if (name == aCrumb[0])
			{
			var sCookie = unescape(aCrumb[1]).split("|");
			for (var s=0; s < sCookie.length; s++)
				{
				var sCrumb = sCookie[s].split("=");
				if (subName == sCrumb[0])
					{
					return unescape(sCrumb[1]);
					}
				}
			}
		}
	return null;
	}

function deleteSubCookie(name, subName)
	{
	var subValue = getCookie(name)
	if (subValue != null)
		{
		subValue = subValue.replace('|' + subName + '=' + getSubCookie(name, subName), '');
		//NEVER TESTED THIS FUNCTION!
		//document.cookie = name + "=" + escape(subValue) + "; expires=" + exp.toGMTString();
		document.cookie = name + "=" + escape(subValue) + ";"
		}
	}