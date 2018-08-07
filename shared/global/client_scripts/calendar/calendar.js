//Note: Throughout the code the month often has 1 subtracted from it...this is to corrispond to the month array (which starts at 0)

var zz, zv, d, fTSR;
var gBF=false;					//don't think this is used
var g_MINY = 1601;
var g_MAXY = 4500;
var g_month = 0;
var g_day = 0;
var g_year = 0;
var g_yLow = 1951;				//for dealing with a 2 digit date.  the 2 digits represent the 100 years starting with this value
var g_eC=null;
var g_eCV="";
var g_dFmt=GetDateFmt();		//set's date format...haven't figured this out yet
var g_fnCB=null;
var shortYear=true;			//Added by Brian true returns 2 digit year false returns 4 digit year.

var rgMC = new Array(12);		//sets an array of the number of days in the corrasponding month (starts on 0)
rgMC[0] = 31;rgMC[1] = 28;rgMC[2] = 31;rgMC[3] = 30;rgMC[4] = 31;rgMC[5] = 30;rgMC[6] = 31;rgMC[7] = 31;rgMC[8] = 30;rgMC[9] = 31;rgMC[10] = 30;rgMC[11] = 31;

d = new Date();
fTSR=0;
zv = d.getTime();
zz = "&zz="+zv;

function CloseCal() {document.all.CalFrame.style.display="none";}

function GetDowStart() {return 0;}				//day of week to start on.  0=sun, 1=mon, 2=tue, etc.
function GetDateFmt() {return "mmddyy";}		//date format
function GetDateSep() {return "/";}				//sets the character seperating dd, mm, and yy

function GetInputDate(t,f){						//function for figuring out what date is in input box...returns false if unrecognizable
  var l = t.length;
  if(0 == l) return false;
  var cSp = '\0';
  var sSp1 = "";
  var sSp2 = "";
  for(var i = 0; i < t.length; i++){
    var c = t.charAt(i);
    if(c == ' ' || isdigit(c)) continue;
    else if(cSp == '\0' && (c == '/' || c == '-' || c == '.')){				//any other seperators you want to recognize should be included here
      cSp = c;
      sSp1 = t.substring(i+1,l);
    }
    else if(c == cSp) sSp2 = t.substring(i+1,l);
    else if(c != cSp) return false;
  }
  if(0 == sSp1.length) return false;
  var m;
  var d;
  var y; 
  if(f=="mmddyy"){
    m = atoi(t);
    d = atoi(sSp1);
    if(0 != sSp2.length) y = atoi(sSp2);
    else y = DefYr(m,d);
  }
  else if(f=="ddmmyy"){
    m = atoi(sSp1);
    d = atoi(t);
    if(0 != sSp2.length) y = atoi(sSp2);
    else y = DefYr(m,d);
  }
  else{
    if(0 == sSp2.length) return false;
    m = atoi(sSp1);
    d = atoi(sSp2);
    y = atoi(t);
  }
  if(y < 100){
    y = 1900+y;
    while(y < g_yLow) y = y+100;
  }
  if(y < g_MINY || y > g_MAXY || m < 1 || m > 12) return false;
  if(d < 1 || d > GetMonthCount(m,y)) return false;
  g_month = m;
  g_day = d;
  g_year = y;
  return true;
}

function DefYr(m,d){						//decides what year it is. (converts 2 digit if need be)
  var dt = new Date();
  var yCur = (dt.getYear() < 1000) ? 1900+dt.getYear() : dt.getYear();
  if(m-1 < dt.getMonth() || (m-1 == dt.getMonth() && d < dt.getDate())) return 1+yCur;
  else return yCur;
}

function atoi(s){
  var t = 0;
  for(var i = 0; i < s.length; i++){
    var c = s.charAt(i);
    if(!isdigit(c)) return t;
    else t = t*10 + (c-'0');
  }
  return t;
}

function isdigit(c) {return(c >= '0' && c <= '9');}

function GetMonthCount(m,y){				//returns month count (adds one to feb if leap year)
  var c = rgMC[m-1];
  if((2 == m) && IsLeapYear(y)) c++;
  return c;
}
function IsLeapYear(y){
  if(0 == y % 4 && ((y % 100 != 0) || (y % 400 == 0))) return true;
  else return false;
}


//  inputs for ShowCalendar: 
//	eP = calender image/button
//	eD = Date input box
//	eDP = Default object to recieve the date from if ED is not valid 
//	dmin = Minimum of date range - this will be default date if eD and eDP are not valid
//	dmax = Maximum of Date Range
//  fnCB = Additional function to run when date is picked (optional)

function ShowCalendar(eP,eD,eDP,dmin,dmax,fnCB){   
  var docFrame=document.all.CalFrame;
  var wF=window.frames.CalFrame;
  docFrame.allowTransparency = false;

  if(null==wF.g_fCalLoaded || false==wF.g_fCalLoaded){		//makes sure that iframe calender.htm loaded properly
    alert("Unable to load popup calendar.\r\nPlease reload the page.");
    return;
  }

  wF.SetMinMax(new Date(dmin),new Date(dmax));
  g_fnCB=fnCB;

  if(eD==g_eC && "block"==docFrame.style.display){
    if(g_eCV != eD.value && GetInputDate(eD.value,g_dFmt)){
      wF.SetInputDate(g_day,g_month,g_year);
      wF.SetDate(g_day,g_month,g_year);
      g_eCV=eD.value;
    }
    else
      docFrame.style.display="none";
  }
  else{
    if(GetInputDate(eD.value,g_dFmt)){
      wF.SetInputDate(g_day,g_month,g_year);
      wF.SetDate(g_day,g_month,g_year);
    }
    else if(null != eDP && GetInputDate(eDP.value,g_dFmt)){
      wF.SetInputDate(g_day,g_month,g_year);
      wF.SetDate(g_day,g_month,g_year);
    }
    else{
      var dt=new Date(dmin); 
      wF.SetInputDate(-1,-1,-1);
      wF.SetDate(dt.getDate(),dt.getMonth()+1,dt.getFullYear());
    }
    var eL=0;var eT=0;
    for(var p=eP; p&&p.tagName!='BODY'; p=p.offsetParent){
      eL+=p.offsetLeft;
      eT+=p.offsetTop;
    }
    var eH=eP.offsetHeight;
    var dH=docFrame.style.pixelHeight;
    var sT=document.body.scrollTop;
    docFrame.style.left=eL;
    if(eT-dH >= sT && eT+eH+dH > document.body.clientHeight+sT)
      docFrame.style.top=eT-dH;
    else
      docFrame.style.top=eT+eH;
    if("none"==docFrame.style.display)
      docFrame.style.display="block";
    g_eC=eD;
    g_eCV=eD.value;
  }
}

function SetDate(d,m,y){					// Returns date to the input box...any other formatting needs to be set up here.
  var ds=GetDateSep();
  g_eC.focus();
  if(shortYear) y=(""+y).substring(2,4);
  if("mmddyy"==g_dFmt) g_eC.value=m+ds+d+ds+y;
  else if("ddmmyy"==g_dFmt) g_eC.value=d+ds+m+ds+y;
  else g_eC.value=y+ds+m+ds+d;
  g_eCV=g_eC.value;
  if(null != g_fnCB && "" != g_fnCB)
    eval(g_fnCB);
}

function GetDOW2(d,m,y){
  var dt=new Date(y,m-1,d);
  return(dt.getDay()+(7-GetDowStart()))%7;
}

function LoadMonths(n){
  var dt=new Date();
  var m=dt.getMonth()+1;
  var y=dt.getFullYear();
  var rg=new Array(n);
  for(i=0; i < n; i++){
    rg[i]=document.createElement("IMG");
    rg[i].src="/eta/months/w" + GetDOW2(1,m,y) + "d" + GetMonthCount(m,y) + ".gif";
    m++;
    if(12 < m){
      m=1;
      y++;
    }
  }
}

LoadMonths(12);