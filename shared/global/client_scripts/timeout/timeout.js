TO_PERIOD=20;
CHK_PERIOD=1;
WARN_PERIOD=1;
DIFF=TO_PERIOD - WARN_PERIOD;
TO_URL='/Profile/LogOff.asp';
WARN_URL='/shared/global/client_scripts/timeout/timeout_warn.html';
var cw;
gLAT=new Date();
var YPC='ypc',YEC='ysecc',YSC='ysc',YIC='yic',SID='yisi',CSS=',',CISS=':';

function GetC(name) {
	var re=new RegExp('\\b'+name+'\='+'([^;]*)');
	var r=re.exec(document.cookie);
	if(r)
		return unescape(r[1]); }

function GetCookie(name,seg) {
	var v=GetC(name);
	if(v) {
		var re=new RegExp('\\b'+seg+CISS+'[^'+CSS+']*');
		var x=re.exec(v);
		if(x) {
			var y=x[0].replace(new RegExp(seg+CISS),'');
			return y; } } }

function SetCookie(name,seg,val,ses) {
	var old=GetC(name)||'';
	if(old)
		old=old.replace(new RegExp('\\b'+seg+CISS+'[^'+CSS+']*'+CSS),'');
	var val=old+seg+CISS+val+CSS;
	var exp='';
	if(!ses)
		exp=';expires=Mon,18 Jan 2010 14:33:21 GMT';
	document.cookie=name+'='+escape(val)+";path=/"+exp; }

function ato() {
	sid=GetCookie(YEC,SID);
	if(sid==''||sid=='timedout')
		return;
	now=new Date();
	tm=now.getTime()-gLAT.getTime();
	if(cw && tm>=1000*60*TO_PERIOD) {
		cw.close();
		location.replace(TO_URL); }
	else if(!cw && tm>=1000*60*DIFF) {
		document.onmousemove=null;
		cw=show_alert(); }
	setTimeout('ato()',1000*60*CHK_PERIOD); }

function show_alert() {
	return window.open(WARN_URL,'cw','width=320,height=210,top=194,left=235'); }

if(document.captureEvents)
	document.captureEvents(Event.MOUSEMOVE);
document.onmousemove=function(){};
ato();