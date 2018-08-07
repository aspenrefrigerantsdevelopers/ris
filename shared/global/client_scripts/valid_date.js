/**
 * DHTML date validation script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/datevalidation.asp)
 */
// Declaring valid date character, minimum year and maximum year
var dtCh= "/";
var minYear=2000;
var maxYear=2099;

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12)
	//var pos1=dtStr.indexOf("/")
	var pos1=2
	//var pos2=dtStr.indexOf(dtCh,pos1+1)
	var pos2=4		
	var strMonth=dtStr.substring(0,pos1)
	//var strDay=dtStr.substring(pos1+1,pos2)
	var strDay=dtStr.substring(pos1,pos2)
	//var strYear=dtStr.substring(pos2+1)		
	var strYear=dtStr.substring(pos2)
		
	//strYr=strYear
	strYr = "20"+strYear
		
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		alert("The date format should be : mmddyy")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		alert("Please enter a valid month")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		alert("Please enter a valid day")
		return false
	}
	if (strYear.length != 2 || year==0 || year<minYear || year>maxYear){
		alert("Please enter a valid 2 digit year between "+minYear+" and "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		alert("Please enter a valid date")
		return false
	}
return true
}

function Validate(dt){
	//var dt=document.frmSample.txtDate
	
	var vDate = document.getElementById("txtDateValidate").value
	
	if (dt.value==vDate)
    {
		return
	}
	

	if (dt.value.length!=6){
        dt.value=vDate;	
		dt.focus();
		alert("The date is too long or too short, \n format should be : mmddyy");
		return false
	}	
	
	if (isDate(dt.value)==false){
        dt.value=vDate;	
		dt.focus();
		return false
	}
	
	if (isToday(dt.value)==false){
        dt.value=vDate;	
		dt.focus();
		return false
	}
	
    return true
 }
 

function isToday(strDate) {

    var c_mon = parseInt('')
    var c_dt = parseInt('')
    var c_yr = parseInt('')
    //var strDate = "02-23-2011"; // Add any of the date from any control

    var mon = parseInt(strDate.substring(0, 2),10);  
    var dt = parseInt(strDate.substring(2, 4),10);  
    var yr = parseInt(strDate.substring(4),10)+2000;

    var Entrydate = new Date(yr, mon-1, dt);
    Entrydate.setHours(0,0,0,0);    
 
    var Currentdate = new Date();
    Currentdate.setHours(0,0,0,0);
    
    if (Entrydate<Currentdate){
        alert("You can only enter current or future dates")
        return false;           
    }        

    if (Entrydate.toString()==Currentdate.toString()){

        var rbln = confirm("Are you sure the customer needs \n this product delivered today?");
        if (!rbln){
            return false;  
        }
    }

    return true;
        
}

/////================================================================================
//// Date Validation Javascript
//// copyright 30th October 2004, 31st December 2009 by Stephen Chapman
//// http://javascript.about.com

//// You have permission to copy and use this javascript provided that
//// the content of the script is not changed in any way.



//function valDateFmt(datefmt) {myOption = -1;
//for (i=0; i<datefmt.length; i++) {if (datefmt[i].checked) {myOption = i;}}
//if (myOption == -1) {alert("You must select a date format");return ' ';}
//return datefmt[myOption].value;}



//function valDateRng(daterng) {myOption = -1;
//for (i=0; i<daterng.length; i++) {if (daterng[i].checked) {myOption = i;}}
//if (myOption == -1) {alert("You must select a date range");return ' ';}
//return daterng[myOption].value;}



//function stripBlanks(fld) {var result = "";var c=0;for (i=0; i<fld.length; i++) {
//if (fld.charAt(i) != " " || c > 0) {result += fld.charAt(i);
//if (fld.charAt(i) != " ") c = result.length;}}return result.substr(0,c);}
//var numb = '0123456789';



//function isValid(parm,val) {if (parm == "") return true;
//for (i=0; i<parm.length; i++) {if (val.indexOf(parm.charAt(i),0) == -1)
//return false;}return true;}



//function isNumber(parm) {return isValid(parm,numb);}
//var mth = new Array(' ','january','february','march','april','may','june','july','august','september','october','november','december');
//var day = new Array(31,28,31,30,31,30,31,31,30,31,30,31);



//function validateDate(fld,fmt,rng) {
//var dd, mm, yy;var today = new Date;var t = new Date;fld = stripBlanks(fld);
//if (fld == '') return false;var d1 = fld.split('\/');
//if (d1.length != 3) d1 = fld.split(' ');
//if (d1.length != 3) return false;
//if (fmt == 'u' || fmt == 'U') {
//  dd = d1[1]; mm = d1[0]; yy = d1[2];}
//else if (fmt == 'j' || fmt == 'J') {
//  dd = d1[2]; mm = d1[1]; yy = d1[0];}
//else if (fmt == 'w' || fmt == 'W'){
//  dd = d1[0]; mm = d1[1]; yy = d1[2];}
//else return false;
//var n = dd.lastIndexOf('st');
//if (n > -1) dd = dd.substr(0,n);
//n = dd.lastIndexOf('nd');
//if (n > -1) dd = dd.substr(0,n);
//n = dd.lastIndexOf('rd');
//if (n > -1) dd = dd.substr(0,n);
//n = dd.lastIndexOf('th');
//if (n > -1) dd = dd.substr(0,n);
//n = dd.lastIndexOf(',');
//if (n > -1) dd = dd.substr(0,n);
//n = mm.lastIndexOf(',');
//if (n > -1) mm = mm.substr(0,n);
//if (!isNumber(dd)) return false;
//if (!isNumber(yy)) return false;
//if (!isNumber(mm)) {
//  var nn = mm.toLowerCase();
//  for (var i=1; i < 13; i++) {
//    if (nn == mth[i] ||
//        nn == mth[i].substr(0,3)) {mm = i; i = 13;}
//  }
//}
//if (!isNumber(mm)) return false;
//dd = parseFloat(dd); mm = parseFloat(mm); yy = parseFloat(yy);
//if (yy < 100) yy += 2000;
//if (yy < 1582 || yy > 4881) return false;
//if (mm == 2 && (yy%400 == 0 || (yy%4 == 0 && yy%100 != 0))) day[1]=29;else day[1]=28;
//if (mm < 1 || mm > 12) return false;
//if (dd < 1 || dd > day[mm-1]) return false;
//t.setDate(dd); t.setMonth(mm-1); t.setFullYear(yy);
//if (rng == 'p' || rng == 'P') {
//if (t > today) return false;
//}
//else if (rng == 'f' || rng == 'F') {
//if (t < today) return false;
//}
//else if (rng != 'a' && rng != 'A') return false;
//return true;
//}