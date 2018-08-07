var thename;
var theobj;
var thetext;
var winHeight;
var winWidth;
var boxPosition;
var timerID;
var seconds = 0;
var x = 0;
var y = 0;
var offsetx = 0;
var offsety = 15;
var windowArray = new Array;

iens6 = document.all || document.getElementById;
ns4 = document.layers;
if(ns4)
	{
	document.captureEvents(Event.MOUSEMOVE);
	}
document.onmousemove = getXY;

function addWindow(id, windowText)
	{
	windowArray[id] = windowText;
	}

function setObj(id, inwidth, inheight, boxpos)
	{
	    //clearTimeout(timerID);

	boxPosition = boxpos;
	winWidth = inwidth;
	winHeight = inheight;
	thetext = windowArray[id];
	if (boxPosition == "bottomR") // Right
		{
		x = x + offsetx;
		y = y + offsety;
		}
	if (boxPosition == "bottomL") // Left
		{
		x = x - (offsetx + 2) - winWidth;
		y = y - offsety;
		}
	if (boxPosition == "topR") // Top
		{
		x = x + offsetx;
		y = y + offsety - winHeight;
		}
	if (boxPosition == "topL") // Top
		{
		x = x - (offsetx + 2) - winWidth;
		y = y + offsety - winHeight;
		}
	if (iens6)
		{
		thename = "viewer";
		theobj = document.getElementById? document.getElementById(thename):document.all.thename;
		theobj.style.width = winWidth;
		theobj.style.height = winHeight;
		theobj.style.left = x;
			if (iens6 && document.all)
				{
				theobj.style.top = document.body.scrollTop + y;
				theobj.innerHTML = "";
				theobj.insertAdjacentHTML("BeforeEnd", thetext);
				}
			if (iens6 && !document.all)
				{
				theobj.style.top = window.pageYOffset + y;
				theobj.innerHTML = "";
				theobj.innerHTML = thetext;
				}
		}
	if (ns4)
		{
		thename = "nsviewer";
		theobj = eval("document." + thename);
		theobj.left = x;
		theobj.top = y;
		theobj.width = winWidth;
		theobj.clip.width = winWidth;
		theobj.height = winHeight;
		theobj.clip.height = winHeight;
		theobj.document.write(thetext);
		theobj.document.close();
		}
	viewIt()
	}

function viewIt()
	{
	if (iens6)
		{
		theobj.style.visibility="visible";
		}
	if (ns4)
		{
		theobj.visibility = "visible";
		}
	}

function stopIt()
	{
	if(theobj)
		{
		if(iens6)
			{
			theobj.innerHTML = "";
			theobj.style.visibility = "hidden";
			}
		if(ns4)
			{
			theobj.document.write("");
			theobj.document.close();
			theobj.visibility = "hidden";
			}
		}
	}

function timer(sec)
	{
	seconds = parseInt(sec);
	if (seconds > 0)
		{
		seconds--;
		timerID = setTimeout("timer(seconds)", 1000);
		}
	else
		{
		stopIt();
		}
	}

function getXY(e)
	{
	if (ns4)
		{
		x = 0;
		y = 0;
		x = e.pageX; 
		y = e.pageY;
		}
	if (iens6 && document.all)
		{
		x = 0;
		y = 0;
		x = event.x; 
		y = event.y;
		}
	if (iens6 && !document.all)
		{
		x = 0;
		y = 0;
		x = e.pageX; 
		y = e.pageY;
		}
	}