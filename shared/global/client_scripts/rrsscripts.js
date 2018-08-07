// JavaScript Document
var xmlHttp
var BCLength
var SerialLength
var cylinderMode; //is barcode searching for cylinder or cylinder being updated to barcode
var arg2//txtSerialNumber or txtBarCode
var arg3//opposite of arg2 txtSerialNumber or txtBarcode
var processForm//set in stateChanged_BCorSerialExists
var vCylSearch="" //value of either the serial number or barcode before it is changed
var rowCount = 2
var prevSelLength // previous length of select email list	
var formType //add or non-add form 
var thisUpOrderID = 0 // orderId set when an order is clicked from the results select list during an non-add request
var removeIndex0 = 0
var qu = 0//item quantity on tbl_update
var browser=navigator.appName
var b_version=navigator.appVersion
var version=parseFloat(b_version)
var appV = navigator.appVersion
var platform = navigator.platform
var userAgent = navigator.userAgent
var rc = 0

function chgDefaultAlert() 
{
    arg = arguments[0];
    switch (arg.id)
    {
    case "txt_branchID":
        alert("Please check and select the appropriate return location.");
        break;

    case "sel_ReturnLoc":
        alert("Please check and select the appropriate branch.");
        getReturnLoc_RetLoc(arg);
        break;

    }
}

function chkShipViaM()
{
	arg = arguments[0]
	argID = arg.id
	argV = arg.value
	if(argV == 3)
	{
		document.getElementById("sel_ShipViaMisc").style.visibility = "visible"
	}else{
		document.getElementById("sel_ShipViaMisc").style.visibility = "hidden"
	}
}

function chgCustOwn()
{
	arg = arguments[0]
	argID = arg.id
	argV = arg.value
	cylProcess = document.getElementById("hdnCylProcess").value
	siteuser = document.getElementById("hdnSiteuser").value

	if(cylProcess == 'add')
	{
		if(argV == "Y")
		{
			document.getElementById("td_txtCustID").innerHTML = "<input type='text' id='txtCustSearch' name='custID' size=10 maxlength=10 onChange='chkCylMaint_CustID(this)'>";
			document.getElementById("td_lblCustNFO").innerHTML = '';
		}else{
			var url="ajaxscript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=27"
			url=url+"&siteuser="+siteuser
		
			var chgCustOwn_r = createRequest();	
			chgCustOwn_r.onreadystatechange= function()
			{
			//(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
			if (chgCustOwn_r.readyState==4 || chgCustOwn_r.readyState=="complete")
				{
					var rt = chgCustOwn_r.responseText
					custID = rt.substring(0,rt.indexOf('|0b|'))
					custNFO = rt.substring(rt.indexOf('|0b|')+4,rt.indexOf('|1b|'))
					document.getElementById("td_txtCustID").innerHTML = custID;
					document.getElementById("td_lblCustNFO").innerHTML = custNFO;
				}
			}
			chgCustOwn_r.open("GET",url,true)
			chgCustOwn_r.send(null)	
		}
	}
}

function selOrderType()
{
    var arg = arguments[0];
	var orderFormType = arguments[1].replace(/ /gi,'')
	var argID = arg.id
	var argV = arg.value
	
	//alert(arg.value)

	if (argV != 'PK')
	{
		document.getElementById("td_quotelink").style.visibility = 'hidden'
	}
	else
	{
		document.getElementById("td_quotelink").style.visibility = 'visible'
	}

	if (argV == 'SH')
	{
		document.getElementById("locNameLbl").innerHTML = 'Ship Location:'
	}
	else
	{
		document.getElementById("locNameLbl").innerHTML = 'Return Location:'
	}

	var custNum = document.getElementById("txt_custNum").value
	var custNum_len = custNum.length
	var sendToRIS = document.getElementById('hdn_SendToRIS').value;
	
	if(custNum_len == 0)
	{
		custNum = 0	
	}

    if(sendToRIS == 'Y' && argV == 'WH')
    {
        alert('This order has been sent to RIS and can not be changed to order type warehouse transfer. Please contact support for additional assistance.')
		window.location = "orderForm.asp?menuID=5&lmt=0&formType=Add"
    }

	var url="ajaxscript.asp"
	url=url+"?sid="+Math.random()
	url=url+"&qry=26"
	url=url+"&orderType="+argV
	url=url+"&custNum="+custNum

	var selOrderType_r = createRequest();	
	selOrderType_r.onreadystatechange= function()
		{
		//(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
		if (selOrderType_r.readyState==4 || selOrderType_r.readyState=="complete")
			{
				var rt = selOrderType_r.responseText
				
				if(rt == 0)
				{
					alert('This customer is not allowed for this order type. Please check your order')
					window.location = "orderForm.asp?menuID=5&lmt=0&formType=Add"
				}else{
				
					if(rt=='receive')
					{
						document.getElementById("td_returnPacket").style.visibility = "visible"
						document.getElementById("sel_returnPkt").style.visibility = "visible"
						document.getElementById("td_paymentTitle").style.visibility = "visible"
						document.getElementById("td_payment").style.visibility = "visible"
						document.getElementById("sel_payment").style.visibility = "visible"
						document.getElementById("tbl_email").style.visibility = "visible"
						document.getElementById("sel_email").style.visibility = "visible"
						document.getElementById("txa_newEmail").style.visibility = "visible"
						document.getElementById("btn_newEmail").style.visibility = "visible"
						document.getElementById("td_WS").style.visibility = "visible"
						document.getElementById("sel_WholeSpot").style.visibility = "visible"
												
						//document.getElementById("td_SendToRIS").style.visibility = "visible"
						//document.getElementById("sel_SendToRIS").style.visibility = "visible"
						
						/* if (ordertype is receive and page is Add) or order is a warehouse transfer*/
						/* then hide those two elements. */
						/* if page is not Add then show or if order is not warehouse transfer*/
						/*if ( (orderFormType == 'Add') ||(argV == 'WH') )
						{
						    document.getElementById("td_SendToRIS").style.visibility = "hidden"
						    document.getElementById("sel_SendToRIS").style.visibility = "hidden"
						}
						else
						{
						    document.getElementById("td_SendToRIS").style.visibility = "visible"
						    document.getElementById("sel_SendToRIS").style.visibility = "visible"
						}*/
					}else{
						document.getElementById("td_returnPacket").style.visibility = "hidden"
						document.getElementById("sel_returnPkt").style.visibility = "hidden"
						document.getElementById("td_paymentTitle").style.visibility = "hidden"
						document.getElementById("td_payment").style.visibility = "hidden"
						document.getElementById("sel_payment").style.visibility = "hidden"
						document.getElementById("tbl_email").style.visibility = "hidden"
						document.getElementById("sel_email").style.visibility = "hidden"
						document.getElementById("txa_newEmail").style.visibility = "hidden"
						document.getElementById("btn_newEmail").style.visibility = "hidden"
						document.getElementById("td_WS").style.visibility = "hidden"
						document.getElementById("sel_WholeSpot").style.visibility = "hidden"
						//document.getElementById("td_SendToRIS").style.visibility = "hidden"
						//document.getElementById("sel_SendToRIS").style.visibility = "hidden"
					}
					
				}
			}
		}
	selOrderType_r.open("GET",url,true)
	selOrderType_r.send(null)	
	ChkSendToRIS();
}

function chkProNumber()
{
	arg = arguments[0]
	arID = arg.id
	argV = arg.value
	h = document.getElementById("hdnCylOrder").value //hdnCylOrder : are these cylinders or non-cylinders
	
	dv_chkRcvComments(arg)

	if(h == 0)
	{
		document.getElementById("btnQuantity").disabled = false
	}else{
		document.getElementById("btnShipCyl").disabled = false
	}
}

function submitQuantity()
{
	m = document.getElementById("hdnMenuID").value
	lineID = document.getElementById("lineID").value
	if(m == 7 )
	{
		q = document.getElementById("txtQty").value
		p = 0
		c = 0
	}else{
		q = document.getElementById("sp_QtyStage").innerHTML
		p = document.getElementById("txtPro").value //proNumber
		c = document.getElementById("sel_ShipViaMisc").value //carrier
	}
		
	var url="ajaxscript.asp"
	url=url+"?sid="+Math.random()
	url=url+"&qry=25"
	url=url+"&quantity="+q
	url=url+"&lineID="+lineID
	url=url+"&menuID="+m
	url=url+"&pronumber="+p
	url=url+"&carrier="+c

	var submitQuantity_r = createRequest();	
	submitQuantity_r.onreadystatechange= function()
		{
		//(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
		if (submitQuantity_r.readyState==4 || submitQuantity_r.readyState=="complete")
			{
				var rt = submitQuantity_r.responseText
				//document.write(rt)
				lineCnt = rt.substring(0,rt.indexOf('|0b|'))
				msg = rt.substring(rt.indexOf('|0b|')+4,rt.indexOf('|1b|'))
				alert(msg)
				if(lineCnt == 1)
				{
					window.location = "index.asp"
				}else{
					arg_orderID = document.getElementById("hdnOrderID")
					arg_lineID = document.getElementById("hdnLineID")
					arg_hdnNet = document.getElementById("hdnNet")
					getCylinder_OrderID(arg_orderID)
					getCylinder_LineID(arg_hdnNet)
					//window.location = "form.asp?menuID=7&process=ship"
				}
			}
		}
	submitQuantity_r.open("GET",url,true)
	submitQuantity_r.send(null)	
}

function chkStageQty()
{
	arg = arguments[0]
	argID = arg.id
	argV = arg.value
	argL = argV.length
	
	var q = parseInt(document.getElementById("span_qtyOrd").innerHTML)
	var n = dv_chkNumber_CE(arg)
	if (n == 0)
	{
		var qSt = parseInt(document.getElementById("sp_QtyStage").innerHTML)
		var qSh = parseInt(document.getElementById("sp_QtyShip").innerHTML)
		var t = parseInt(argV) + qSh

		if(t > q)
		{
			alert('Quantity ordered exceeded.')
			arg.value = ''
			document.getElementById("btnQuantity").disabled = true;
		}else{
		
			var s = parseInt(q) - parseInt(qSh)
	
			//if(s > q)
			if(argV <= s)
			{
				if(argL == 0)
					{
						arg.value = 0
					}
				document.getElementById("btnQuantity").disabled = false;
			}else{
				alert('Quantity ordered exceeded.')
				arg.value = ''
				document.getElementById("btnQuantity").disabled = true;	
			}	
		}
	}
}

function closeSearch()
{
	document.getElementById("td_orderSearch").innerHTML =''
}

function lookUpCustomer()
{
	custN = document.getElementById("txtCustSearch").value
	billTo = document.getElementById("txtBillToSearch").value
	shipTo = document.getElementById("txtShipToSearch").value
	shipAddr = document.getElementById("txtShipAddrSearch").value
	
	if(custN.length == 0)
	{ 
		custN = 0
	}
	if(billTo.length == 0)
	{ 
		billTo = 0
	}
	if(shipTo.length == 0)
	{ 
		shipTo = 0
	}
	if(shipAddr.length == 0)
	{ 
		shipAddr = 0
	}
	
	var url="ajaxscriptB.asp"
	url=url+"?sid="+Math.random()
	url=url+"&qry=7"
	url=url+"&custN="+custN
	url=url+"&billTo="+billTo
	url=url+"&shipTo="+shipTo
	url=url+"&shipAddr="+shipAddr

	var lookUpCustomer_r = createRequest();	
	lookUpCustomer_r.onreadystatechange= function()
		{
		//(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
		if (lookUpCustomer_r.readyState==4 || lookUpCustomer_r.readyState=="complete")
			{
				var rt = lookUpCustomer_r.responseText
				if(rt == 0)
				{
					alert('No records found');
				}else{
					document.getElementById("td_selCust").innerHTML = rt
				}
			}
		}
	lookUpCustomer_r.open("GET",url,true)
	lookUpCustomer_r.send(null)
}

function createRequest()
{ 
	var request = null
	try
	{
		request = new XMLHttpRequest();
	}catch (trymicrosoft){
		try
		{
			request = new ActiveXObject("Msxml2.XMLHTTP");	
		}catch (othermicrosoft){
			try
			{
				request = new ActiveXObject("Microsoft.XMLHTTP");	
			}catch (failed){
				request = null;
			}
		}
	}
	if(request == null)
	{
		alert("error creating request object!")	
	}else{
		return request	
	}
}

function shipReset()
{
	menuID = document.getElementById("hidMenuID").value
	window.location = "form.asp?menuID="+menuID+"&process=ship"
}

function shipCylinders()
{
	arg = arguments[0]
	p = document.getElementById("txtPro").value //proNumber
	c = document.getElementById("sel_ShipViaMisc").value //carrier
//alert('206 ' + p + ' c ' + c)	
	var url="ajaxscript.asp"
	url=url+"?sid="+Math.random()
	url=url+"&qry=21"
	url=url+"&lineID="+arg //cylinder table ctkey
	url=url+"&carrier="+c
	url=url+"&pronumber="+p

	var shipCylinders_r = createRequest();	
	shipCylinders_r.onreadystatechange= function()
		{
			//(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
			if (shipCylinders_r.readyState==4 || shipCylinders_r.readyState=="complete")
			{
				//alert('Cylinder status changed to Shipped')
				//window.location = "index.asp"
				var rt = shipCylinders_r.responseText
				lineCnt = rt.substring(0,rt.indexOf('|0b|'))
				msg = rt.substring(rt.indexOf('|0b|')+4,rt.indexOf('|1b|'))
				alert(msg)
				if(lineCnt == 0)
				{
					window.location = "index.asp"
				}else{
					arg_orderID = document.getElementById("hdnOrderID")
					arg_lineID = document.getElementById("hdnLineID")
					arg_hdnNet = document.getElementById("hdnNet")
					getCylinder_OrderID(arg_orderID)
					getCylinder_LineID(arg_hdnNet)
				}
			}
		}
	shipCylinders_r.open("GET",url,true)
	shipCylinders_r.send(null)
}

function removeCylBarcode()
{
	arg = arguments[0]
	del = confirm("Delete Cylinder")
	if (del == true)
	{
		var x = document.getElementById("selCylBarcode");
		var lineID = document.getElementById("txtLineID").value

		var url = "ajaxscript.asp"
		url = url + "?sid=" + Math.random()
		url = url + "&qry=20"
		url = url + "&ctKey=" + arg.value //cylinder table ctkey
		url = url + "&lineID=" + lineID

		var removeCylBarcode_r = createRequest();
		removeCylBarcode_r.onreadystatechange = function ()
		{
			//(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
			if (removeCylBarcode_r.readyState == 4 || removeCylBarcode_r.readyState == "complete")
			{
				qtyStage = removeCylBarcode_r.responseText
				span_CylStaged.innerHTML = qtyStage
			}
		}
		removeCylBarcode_r.open("GET", url, true)
		removeCylBarcode_r.send(null)

		x.remove(x.selectedIndex);
		document.getElementById("txtCylBarcode").focus()
	}
	else
	{
		document.getElementById('selCylBarcode').selectedIndex = -1;
		document.getElementById("txtCylBarcode").focus()
	}
}

function removeAllStagedCylBarcodes_validate() 
{
	var numCyls = document.getElementById("selCylBarcode").length
	var confirmMsg
	
	if ( numCyls == 0 )
	{
		alert("There are no staged cylinders to un-stage")
		return false;
	}
	
	if ( numCyls == 1 ) 
	{
		confirmMsg = "Press [ok] to un-stage 1 cylinder"
	}
	else
	{
		confirmMsg = "Press [ok] to un-stage all " + numCyls + " cylinders"
	}
	
	if ( !confirm( confirmMsg ))
	{
		return false;
	}
	
	return true;
}

function removeAllStagedCylBarcodes()
{
	if ( removeAllStagedCylBarcodes_validate() )
	{
		var cylList
		var i
		var numCyls = document.getElementById("selCylBarcode").length
	
		cylList = ""
	
		for ( i = 0; i < numCyls; i++ )
		{
			cylList = cylList + document.getElementById("selCylBarcode").options[i].value + ','
		}
		
		var lineID = document.getElementById("txtLineID").value
		
		var url="ajaxscript.asp"
		url=url+"?sid="+Math.random()
		url=url+"&qry=29"
		url=url+"&ctKeys="+cylList //cylinder table ctkeys
		url=url+"&lineID="+lineID
		
		var removeCylBarcodes_r = createRequest();	
		removeCylBarcodes_r.onreadystatechange= function()
			{
				//(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
				if (removeCylBarcodes_r.readyState==4 || removeCylBarcodes_r.readyState=="complete")
				{
					qtyStage = removeCylBarcodes_r.responseText
					span_CylStaged.innerHTML = qtyStage
					
					while ( document.getElementById("selCylBarcode").length > 0 )
					{
						document.getElementById("selCylBarcode").remove(0)
					}
				}
			}
		removeCylBarcodes_r.open("GET",url,true)
		removeCylBarcodes_r.send(null)
	}
}

function chk4ValidCyl()
{
	arg = arguments[0]
	argValue = arg.value
	argID = arg.id

	totStaged = parseInt(span_CylStaged.innerHTML)
	qtyShip = parseInt(span_CylShipped.innerHTML)
	qtyOrd = parseInt(span_qtyOrd.innerHTML)
	
	var allOk = 1

	if ( allOk == 1 )
	{
		if( (totStaged+qtyShip) >= qtyOrd)
		{
			allOk = 0
			arg.value = "" //clear value	
			alert('Quantity ordered reached. No more cylinders can be added');
		}
	}	

	if ( allOk == 1 )
	{
		lineID = document.getElementById("txtLineID").value
		barcode = addZero(argID)// addZero() takes argID value and add zeros if necessary
		
		var chk4ValidCyl_r = createRequest();
		var url="ajaxScript.asp"
		url=url+"?sid="+Math.random()
		url=url+"&qry=19"
		url=url+"&barcode="+barcode
		url=url+"&lineID="+lineID
	
		chk4ValidCyl_r.open("GET",url,true)
		chk4ValidCyl_r.onreadystatechange= function()
		{
			if ((chk4ValidCyl_r.readyState==4 || chk4ValidCyl_r.readyState=="complete"))
			{
				document.getElementById("txtCylBarcode").value=""
				validCyl = chk4ValidCyl_r.responseText
				
				var x=document.getElementById("selCylBarcode");
				allOk = validCyl.substring(0,validCyl.indexOf("|a|"))
				if ( allOk == 0 )
				{
					var errMsg = validCyl.substring(validCyl.indexOf("|a|")+3,validCyl.indexOf("|b|"))
					alert( errMsg )
					
					document.getElementById("txtCylBarcode").focus()
				}
				else
				{
					var ctkeyValue = validCyl.substring(validCyl.indexOf("|a|")+3,validCyl.indexOf("|b|"))
				
					if(platform == 'WinCE')
					{
						var submit_r = createRequest();
						var url="mobileAjax.asp"
						url = url+"?sid="+Math.random()
						url = url+"&qry=stageCyl"
						url = url+"&lineID="+lineID
								
						submit_r.onreadystatechange= function()
						{
							if (submit_r.readyState==4 || submit_r.readyState=="complete")
							{
								var rt = submit_r.responseText
								document.getElementById("span_SelCylBarcode").innerHTML = rt
								addStage = span_CylStaged.innerHTML
								addStage = parseInt(addStage) + 1
								span_CylStaged.innerHTML = addStage
							}
						}
										
						submit_r.open("GET",url,true)
						submit_r.send(null)
					}
					else
					{
						try
						{
							var y=document.createElement('option');
							y.text = barcode
							y.value = ctkeyValue
							x.add(y,0); // standards compliant
						
							addStage = span_CylStaged.innerHTML
							addStage = parseInt(addStage) + 1
							span_CylStaged.innerHTML = addStage
						}
						catch(ex)
						{
							x.add(y); // IE only
						}
						document.getElementById("txtCylBarcode").focus()
					}
				} // allOk == 1 (returned from server addCylinderStage_lineID )
			}  // end of ((chk4ValidCyl_r.readyState==4 || chk4ValidCyl_r.readyState=="complete"))
		} // end of chk4ValidCyl_r.onreadystatechange= function()
		
		chk4ValidCyl_r.send(null)
	} 
}

function getCylinder_OrderID()
{
	arg = arguments[0]
	argValue = arg.value
	argID = arg.id
	siteUser = arguments[1]

	var ok = 0
	var cylProcess = document.getElementById("hdnCylProcess").value

	if(platform == 'WinCE')
	{
		ok = dv_chkNumber_CE(arg)
	}
	
	if(ok == 0)
	{
		menuID = document.getElementById("hidMenuID").value
		var getCylinder_LineID_r = createRequest();
		var url="ajaxscript.asp"
		url=url+"?sid="+Math.random()
		url=url+"&qry=24"
		url=url+"&OrderID="+argValue
		url=url+'&menuID='+menuID
		url=url+"&cylProcess="+cylProcess
		url=url+"&siteUser="+siteUser

		var r = createRequest();	
			r.onreadystatechange= function()
			{
				if (r.readyState==4 || r.readyState=="complete")
				{
					var rt =  r.responseText;
					var e = rt.substring(0,rt.indexOf('|a|'));
					if(e == 0)
					{
						var msg = rt.substring(rt.indexOf('|a|')+3,rt.indexOf('|b|'))
						alert(msg)
						window.location = "form.asp?menuID="+menuID+"&process=ship";
					}else{
						document.getElementById("td_txtLineID").innerHTML = rt.substring(rt.indexOf('|a|')+3,rt.indexOf('|b|'))
					}
				}
			}
		r.open("GET",url,true)
		r.send(null)
	}
}

function getCylinder_LineID()
{
	var arg = arguments[0]
	var argValue = arg.value
	var argID = arg.id

    siteUser = arguments[1]

	var ok = 0
	
	if(platform == 'WinCE')
	{
		ok = dv_chkNumber_CE(arg)
	}
	
	if(argID == 'hdnNet')
	{
		document.getElementById("td_lineInfo").innerHTML = ''
	}else{
		if(ok == 0)
		{
			menuID = document.getElementById("hidMenuID").value
			var getCylinder_LineID_r = createRequest();
			var url="ajaxscript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=18"
			url=url+"&lineID="+argValue
			url=url+'&menuID='+menuID
			url=url+"&siteUser="+siteUser
		
			var r = createRequest();	
				r.onreadystatechange= function()
				{
					if (r.readyState==4 || r.readyState=="complete")
					{
						var rt =  r.responseText

						var e=rt.substring(0, rt.indexOf("|a|"))

						if(e == "0")
						{
							var errMsg = rt.substring(rt.indexOf("|a|")+3, rt.indexOf("|b|"))
							alert(errMsg)
							window.location = "form.asp?menuID="+menuID+"&process=process"
						}else{
							document.getElementById("td_lineInfo").innerHTML = rt
							var orderID = (rt.substring(rt.indexOf(":")+2,rt.indexOf("Item :") ))
							document.getElementById("td_txtOrderID").innerHTML = orderID;
							document.getElementById("sp_OrderID").innerHTML = '';
						}
					}
				}
			r.open("GET",url,true)
			r.send(null)
		}
	}
}

function getLineID()
{
	arg=arguments[0]
	var r = createRequest();
	var url="ajaxScript.asp"
	url=url+"?sid="+Math.random()
	url=url+"&qry=13"
	url=url+"&orderID="+arg
	
	r.open("GET",url,true)
		
	r.onreadystatechange= function()
	{
		if ((r.readyState==4 ) && (r.status=="200"))
		{
			document.getElementById("hd_LineID").value = r.responseText
		}
	}
	r.send(null)
}

function menuOption()
{
	url = window.location.search
	s = url.lastIndexOf("=")
	formType = url.substring(s+1)
	
	var menuOption_r = createRequest();
	
	switch(formType)
	{
		case "Add":
			document.getElementById("td_orderSearch").innerHTML = ''
			chkFormtype(formType)
		break;
		
		default :
		
			var url="ajaxScript.asp"
			url= url+"?sid="+Math.random()
			url= url+"&qry=22"
			url= url+"&formType="+formType

			menuOption_r.onreadystatechange= function()
				{
					if (menuOption_r.readyState==4 || menuOption_r.readyState=="complete")
					{
						var menuOption_rStr = menuOption_r.responseText//Update - Item - nothing returned
						document.getElementById("td_orderSearch").innerHTML = menuOption_rStr
						chkFormtype(formType)
					}
				}
				
			menuOption_r.open("GET",url,true)
			menuOption_r.send(null)
		break;
	}
}

function menuReload()
{
	arg = arguments[0]
	argV = arg.value
	lmt = document.getElementById("hdn_lmt").value
	if(argV == 'Customer')
	{
		var menuReload_r = createRequest();
		
		var url="ajaxScript.asp"
			url= url+"?sid="+Math.random()
			url= url+"&qry=23"
		
			menuReload_r.onreadystatechange= function()
				{
					if (menuReload_r.readyState==4 || menuReload_r.readyState=="complete")
					{
						var menuOption_rStr = menuReload_r.responseText//Update - Item - nothing returned
						document.getElementById("td_orderSearch").innerHTML = menuOption_rStr
					}
				}
				
			menuReload_r.open("GET",url,true)
			menuReload_r.send(null)
	}else{
		window.location = "orderForm.asp?lmt="+lmt+"&formType="+arg.value
	}
}

function getOrderInfo_ID()
{
	var args = arguments[0]
    var disOrderType = 0;
   
	xmlHttp=GetXmlHttpObject()	
	thisUpOrderID = args.value
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	}

	var r = createRequest();
	var orderType_cr = createRequest();
	
	var r_url="ajaxScript.asp"
	r_url=r_url+"?sid="+Math.random()
	r_url=r_url+"&qry=13"
	r_url=r_url+"&orderID="+args.value

	var url="ajaxScript.asp"
	url=url+"?sid="+Math.random()
	url=url+"&qry=12"
	url=url+"&OrderID="+args.value
	
	var url_orderType = "ajaxScript.asp"
	url_orderType = url_orderType+"?sid="+Math.random()
	url_orderType = url_orderType+"&qry=28"
	url_orderType = url_orderType+"&OrderID="+args.value	
	
	document.getElementById("hd_findOrderID").value = args.value
	document.getElementById("sel_Email").style.visibility='hidden';
	document.getElementById("txa_newEmail").disabled=false;
	xmlHttp.open("GET",url,true)

	xmlHttp.onreadystatechange = function () {
	    if (xmlHttp.readyState == 4 || xmlHttp.readyState == "complete")
	    ////(r.readyState==4 || r.readyState=="complete") ((r.readyState==4 ) && (r.status=="200"))
	    {
	        // set select list flag. flag determines whether to remove index[0] from select lists after intial selection of order

	        str = xmlHttp.responseText;
	        //document.write(str)			
	        // the arrTD array must be in the same sequence as the called asp stored procedure in qry 12 ajaxscript.asp
	        //if form element is changed, dropped, added you must match the change and adjust the stored proc in qry 12
	        var arrTD = new Array(["sel_SendToRIS"], ["sel_OrderType"], ["td_order"], ["txt_orderDate"], ["txt_branchID"], ["txt_custNum"], ["txt_poNum"], ["td_billTo"], ["sel_Tickets"], ["txt_ShipToName"], ["sel_ReturnLoc"], ["td_returnLoc"], ["txt_ShipToAddr1"], ["sel_returnPkt"], ["txt_ShipToAddr2"], ["txt_proNumber"], ["hdn_ShipToAddr3"], ["sel_FreightOut"], ["sel_Payment"], ["txa_newEmail"], ["txt_shipDate"], ["sel_ShipVia"], ["sel_ShipViaMisc"], ["sel_WholeSpot"], ["sel_Print"], /*["sel_pkPrinter"],*/["txa_Comments"], ["hdn_CusClass"], ["td_OrderStatus"], ["td_CSR"], ["sel_salesmanID"],
["txt_ShipToCity"], ["txt_ShipToState"], ["txt_ShipToZip"], ["txt_CalledInBy"], ["txt_CalledInByPh"], ["td_RISStatus"], ["td_ShippingInfo"], ["txt_RIS_BillNotes"], ["txt_RIS_ShipNotes"], ["td_OnDockDate"], ["hd_LineCount"])

	        frm = str.indexOf("|", 0)
	        radTD = str.substring(0, frm)

	        for (i = 0; i < arrTD.length; i++) {
	            clName = document.getElementById(arrTD[i]).className
	            clID = document.getElementById(arrTD[i]).id

	            if (clID == 'sel_SendToRIS') {
	                switch (radTD) {
	                    case "Y":
	                        disOrderType = 1;
	                        document.getElementById('hdn_SendToRIS').value = 'Y';
	                        //document.getElementById('txt_custNum').disabled = true;
	                        document.getElementById('sel_SendToRIS').disabled = true;
	                        //var ih = "<input type='hidden' id='sel_SendToRIS' name='sendToRIS' value='Y'>"
	                        //document.getElementById('td_RISStatus').innerHTML = 'Sent to RIS'; 
	                        break;
	                    case "V":
	                        disOrderType = 0;
	                        //document.getElementById('txt_custNum').disabled = false;
	                        document.getElementById('hdn_SendToRIS').value = 'V';
	                        document.getElementById('sel_SendToRIS').disabled = true;
	                        document.getElementById('td_RISStatus').innerHTML = 'Does not send to RIS';
	                        break;
	                    case "N":
	                        disOrderType = 0;
	                        //document.getElementById('txt_custNum').disabled = false;
	                        document.getElementById('hdn_SendToRIS').value = 'N';
	                        document.getElementById('sel_SendToRIS').disabled = false;
	                        document.getElementById('td_RISStatus').innerHTML = ''
	                        break;
	                }
	            }

	            switch (clName) {
	                case "td":
	                    document.getElementById(arrTD[i]).innerHTML = radTD
	                    break;
	                case "formField":
	                    document.getElementById(arrTD[i]).value = radTD
	                    chgReq(document.getElementById(arrTD[i]), "black")
	                    break;
	                case "formField_NN":
	                    document.getElementById(arrTD[i]).value = radTD
	                    //chgReq(document.getElementById(arrTD[i]),"black")
	                    break;
	                case "reqSel":
	                    if (clID == 'sel_ReturnLoc') {
	                        t = radTD.indexOf(";", 0)
	                        v = radTD.substring(0, t)
	                        q = radTD.indexOf(";", t + 1)
	                        d = radTD.substring(t + 1, q)
	                        // note: this will be 1 if this is a WH order and the customer has a defined location in custMisc.
	                        //       otherwise it will be 0.
	                        //       It is only used if the formType is 'Update' (see below)
	                        disableWHReturnLoc = radTD.substring(q + 1)
	                    }
	                    else {
	                        t = radTD.indexOf(";", 0);
	                        v = radTD.substring(0, t);
	                        d = radTD.substring(t + 1);
	                    }

	                    if (clID == 'sel_OrderType') {
	                        var orderType = v;
	                        orderType_cr.open("GET", url_orderType, true)//get Line info
	                        orderType_cr.onreadystatechange = function () {
	                            if ((orderType_cr.readyState == 4) && (orderType_cr.status == "200")) {
	                                var results_orderType = orderType_cr.responseText
	                                document.getElementById("td_selOrderTypes").innerHTML = results_orderType

	                                url = window.location.search
	                                s = url.lastIndexOf("formType=")
	                                formType = url.substring(s + 9)

	                                /*if( (formType == 'Update' && disOrderType == 0) || (formType == 'Add' ) )
	                                {	
	                                document.getElementById("sbt_submitForm").style.visibility='visible';
	                                document.getElementById("sel_OrderType").disabled=false;
	                                }else{
	                                document.getElementById("sel_OrderType").disabled=true;
	                                }*/
	                                switch (formType) {
	                                    case "Update":
	                                        getOrderType_orderID_chkQC(args.value);
	                                        /*if( disOrderType == 0)
	                                        {	
	                                        document.getElementById("sel_OrderType").disabled=false;
	                                        }else{
	                                        document.getElementById("sel_OrderType").disabled=true;
	                                        }*/
	                                        document.getElementById("sbt_submitForm").style.visibility = 'visible';
	                                        break;
	                                    case "Add":
	                                        document.getElementById("sbt_submitForm").style.visibility = 'visible';
	                                        break;
	                                    default:
	                                        document.getElementById("sel_OrderType").disabled = true;
	                                }
	                            }
	                        }
	                        orderType_cr.send(null)
	                    }

	                    var s = document.getElementById(arrTD[i]) //select list by ID

	                    //get current option[0]
	                    var s0text = s.options[0].text
	                    var s0value = s.options[0].value
	                    //set selectedIndex to value from database - new option[0]
	                    s.options[s.selectedIndex].text = d
	                    s.options[s.selectedIndex].value = v

	                    if (removeIndex0 == 0) {
	                        //add initial result to options[1] of select list
	                        var opt = document.createElement('option');
	                        opt.text = s0text
	                        opt.value = s0value
	                        try {
	                            s.add(opt, 1); // standards compliant
	                        }
	                        catch (ex) {
	                            s.add(opt); // IE only
	                        }
	                    } else {
	                        //after initial selection of an order id from the order search result list reset index 0 to new values from query
	                        s.options[0].text = d
	                        s.options[0].value = v
	                    }

	                    if (clID == 'sel_ReturnLoc') {
	                        if ((formType == 'Update') && (disableWHReturnLoc == '1')) {
	                            document.getElementById("sel_ReturnLoc").disabled = true
	                        }
	                        retLocSelectedIndex = document.getElementById("sel_ReturnLoc").selectedIndex
	                        document.getElementById("hdnReturnLoc").value = document.getElementById("sel_ReturnLoc").options[retLocSelectedIndex].value
	                    }
	                    break;
	                case "reqHidden":
	                    document.getElementById(arrTD[i]).value = radTD
	                    break;
	            }
	            frm = frm + 1
	            to = str.indexOf("|", frm)
	            radTD = str.substring(frm, to)
	            frm = to
	        }


	        if (orderType == 'PK') {
	            document.getElementById("td_quotelink").innerHTML = "<a href='javascript:openQuoteform(" + args.value + "," + "&quot;" + formType + "&quot;" + ");'>Order Gas Quotes</a>"
	        }


	        if (orderType == 'SH') {
	            document.getElementById("td_returnPacket").style.visibility = "hidden"
	            document.getElementById("sel_returnPkt").style.visibility = "hidden"
	            document.getElementById("td_paymentTitle").style.visibility = "hidden"
	            document.getElementById("td_payment").style.visibility = "hidden"
	            document.getElementById("sel_payment").style.visibility = "hidden"
	            document.getElementById("tbl_email").style.visibility = "hidden"
	            document.getElementById("sel_email").style.visibility = "hidden"
	            document.getElementById("txa_newEmail").style.visibility = "hidden"
	            document.getElementById("btn_newEmail").style.visibility = "hidden"
	            document.getElementById("sel_WholeSpot").style.visibility = "hidden"
	            document.getElementById("td_WS").style.visibility = "hidden"
	        } else {
	            document.getElementById("td_returnPacket").style.visibility = "visible"
	            document.getElementById("sel_returnPkt").style.visibility = "visible"
	            document.getElementById("td_paymentTitle").style.visibility = "visible"
	            document.getElementById("td_payment").style.visibility = "visible"
	            document.getElementById("sel_payment").style.visibility = "visible"
	            document.getElementById("tbl_email").style.visibility = "visible"
	            document.getElementById("sel_email").style.visibility = "visible"
	            document.getElementById("txa_newEmail").style.visibility = "visible"
	            document.getElementById("btn_newEmail").style.visibility = "visible"
	            document.getElementById("sel_WholeSpot").style.visibility = "visible"
	            document.getElementById("td_WS").style.visibility = "visible"
	        }

	        numOfRows = document.getElementById("hd_LineCount").value;
	        if (removeIndex0 == 0) {
	            removeIndex0 = 1
	        }
	        //delete previous rows from previous orderID select
	        //trCount = document.getElementById('tbl_items').getElementsByTagName("tr").length
	        trCount = document.getElementById('tbl_items').rows.length

	        if (trCount != 2) {
	            for (dr = 0; dr < (trCount - 2); dr++) {
	                document.getElementById('tbl_items').deleteRow(2)
	            }
	        }

	        sv = document.getElementById("sel_ShipVia").value
	        if (sv == 3) {
	            document.getElementById("sel_ShipViaMisc").style.visibility = "visible"
	        } else {
	            document.getElementById("sel_ShipViaMisc").style.visibility = "hidden"
	        }

	        r.open("GET", r_url, true)//get Line info
	        r.onreadystatechange = function () {
	            if ((r.readyState == 4) && (r.status == "200")) {
	                document.getElementById("hd_LineID").value = r.responseText
	                line = r.responseText //line IDs 

	                f2 = line.indexOf(";", 0)
	                lineID = line.substring(0, f2)//(from,to)

	                for (a = 0; a < numOfRows; a++) {
	                    addRow(numOfRows, lineID)//lineID from getLineID query ? huh ?args.value = incoming orderID
	                    f2 = f2 + 1
	                    t2 = line.indexOf(";", f2)
	                    lineID = line.substring(f2, t2)
	                    f2 = t2
	                }

	            }
	        }
	        r.send(null)
	    }
	}
xmlHttp.send(null)
}

function getOrderType_orderID_chkQC() {
    var orderID = arguments[0];

    var url = "ajaxScript.asp"
    url = url + "?sid=" + Math.random()
    url = url + "&qry=31"
    url = url + "&OrderID=" + orderID

    xmlHttp.open("GET",url,true)

    xmlHttp.onreadystatechange = function () {
        if (xmlHttp.readyState == 4 || xmlHttp.readyState == "complete") {
            str = xmlHttp.responseText;
            if (str != 0) {
                document.getElementById("sel_OrderType").disabled = true;
                alert("Order contains cylinders in QC status. Order type cannot be changed");
            }
        }
    }
    xmlHttp.send(null)
}


function getOrderResults()
{
	xmlHttp=GetXmlHttpObject()	
	
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	}
	
	url = window.location.search
	s = url.lastIndexOf("=")
	formType = url.substring(s+1)
	
	branchID = document.getElementById("txt_getbranchID").value
	custNum = document.getElementById("txt_getcustNum").value
	orderID = document.getElementById("txt_getOrderNum").value
	poNum = document.getElementById("txt_getPONum").value
	billTo = document.getElementById("txt_getBillName").value
	shipTo = document.getElementById("txt_getShipName").value
	shipToAddr = document.getElementById("txt_getShipToAddr").value
	orderStatus = document.getElementById("txt_getOrderStatus").value	
	salesmanID = document.getElementById("sel_getSalesMan").value
	csr = document.getElementById("sel_getCSRList").value

	var url="ajaxScript.asp"
	url=url+"?sid="+Math.random()
	url=url+"&qry=11"
	url=url+"&branchID="+branchID
	url=url+"&custNum="+custNum
	url=url+"&orderID="+orderID
	url=url+"&poNum="+poNum
	url=url+"&billTo="+billTo
	url=url+"&shipTo="+shipTo
	url=url+"&shipToAddr="+shipToAddr
	url=url+"&orderStatus="+orderStatus	
	url=url+"&salesmanID="+salesmanID
	url=url+"&csr="+csr
	url=url+"&formType="+formType

	xmlHttp.onreadystatechange= function()
	{
		if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")//(deleteRow_cr.readyState==4) && (deleteRow_cr.status=="200")
		{
			document.getElementById("td_searchResults").innerHTML = xmlHttp.responseText
			document.getElementById("sel_searchResults").style.width='100%'
		}
	}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)	
}

function printInqForm()
{
    inqOrderID = document.getElementById("td_order").innerHTML
    risID = document.getElementById("td_RISStatus").innerHTML;
    risID = risID.substr(103, 6);

    if (risID.length == 0) {
        alert(inqOrderID + " requires a RIS order number for printing. Please contact your system administrator.");
    } else {
        prt = document.getElementById("sel_pkPrinter").value
        var identifer = inqOrderID;
        var printQstring = "?print=Y&printType=order&printerID=" + prt + "&reportParameterValue=" + identifer;
        var printQstring = printQstring + "&dbserver=Argedsrefsql001";
        var printQstring = printQstring + "&prefix=RRS_&reportName=RRSOrder&reportFolder=RRS&reportParameter=OrderNumber"
        var url = "http://argedsrefsql001/pdfbuilder/printforms.aspx" + printQstring;

        var printLabel_r = createRequest();
        printLabel_r.onreadystatechange = function () {
            if (printLabel_r.readyState == 4 || printLabel_r.readyState == "complete") {
                var rt = printLabel_r.responseText;
                rt = rt.substr(0, 1);
                if (rt != 1) {
                    alert("Print Error - scripts.js line 1214");
                }
            }
        }
        printLabel_r.open("GET", url, true);
        printLabel_r.send(null);
    }
}

function chkFormtype()
{
	args = arguments[0];
	xmlHttp=GetXmlHttpObject()	
	
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	}	

	switch(args)
	{
		case "Add":
			//order form
			x = document.getElementById("fm_form")
			//document.getElementById("tbl_update").style.visibility = 'hidden';
			//search form
			document.getElementById("td_OrderTitle").innerHTML = '<b>Add Order Form</b>';
			for (var i=0;i<x.length;i++)
			{	
				t = x.elements[i].type
				switch(t)
				{
					case "text":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = false
					break;
					case "select-one":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = false
					break;
					case "textarea":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = false
					break;
				}
			}
//Code to show "quote link" to add quotes to new orders with no orderid
  
            var strHTML = "<a href='javascript:openQuoteform("+document.getElementById("hd_tempOrderQuoteid").value+","+"&quot;"+formType+"&quot;"+");'>Add Order Gas Quotes</a>"
			document.getElementById("td_quotelink").innerHTML =strHTML
//--------
			
			document.getElementById("td_SendToRIS").style.visibility='visible';
			document.getElementById("sel_SendToRIS").style.display='block'
			document.getElementById("tbl_OrderForm").style.visibility ="visible";
			document.getElementById("sbt_submitForm").style.visibility='visible';
			document.getElementById("txa_newEmail").style.visibility = 'hidden';
			document.getElementById("sel_Print").style.visibility = 'hidden';
			document.getElementById("sel_pkPrinter").style.visibility = 'hidden';
			//document.getElementById("txt_findOrderID").style.visibility='hidden';
			orderT = document.getElementById("sel_OrderType").value
			if(orderT != 'SH')
			{
				document.getElementById("td_WS").style.visibility = "visible"
				document.getElementById("sel_WholeSpot").style.visibility = "visible"					
			}
			document.getElementById("hd_formType").value='Add';
			
			//get item table with header rows
			var url="ajaxScript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=10"
			url=url+"&addRow=Y"
		
			xmlHttp.onreadystatechange= function()
			{
				if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
				{
					document.getElementById("td_items").innerHTML = xmlHttp.responseText
				}
			}
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)
			break;
		case "Inquiry":
			// order form
			x = document.getElementById("fm_form")
			x.reset();
			document.getElementById("td_OrderTitle").innerHTML = '<b>Inquiry Order Form</b>';
			document.getElementById("tbl_OrderForm").style.visibility ="visible";
		
			var inqPrint = "<input id='btn_inqPrint' type='button' name='inqPrint' value='Print' onClick='printInqForm()'>"
			document.getElementById("td_submitOrderForm").innerHTML = inqPrint//.style.visibility='hidden';
			document.getElementById("hd_formType").value='Inquiry';
			document.getElementById("td_items").innerHTML="";
			
			//search form
			document.getElementById("txt_getbranchID").style.visibility='visible';
			document.getElementById("txt_getcustNum").style.visibility='visible';
			document.getElementById("txt_getOrderNum").style.visibility='visible';
			document.getElementById("txt_getPONum").style.visibility='visible';
			document.getElementById("txt_getBillName").style.visibility='visible';
			document.getElementById("txt_getShipName").style.visibility='visible';
			document.getElementById("txt_getShipToAddr").style.visibility='visible';
			document.getElementById("txt_getOrderStatus").style.visibility='visible';
			document.getElementById("sel_getSalesMan").style.visibility='visible';
			document.getElementById("sel_getCSRList").style.visibility='visible';
			document.getElementById("btn_Search").style.visibility='visible';
			document.getElementById("sel_searchResults").style.visibility="visible";
			document.getElementById("sel_searchResults").style.width="100%";
			for (var i=0;i<x.length;i++)
			{
				t = x.elements[i].type
				switch(t)
				{
					case "text":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
					case "select-one":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
					case "textarea":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
				}				
			}
			document.getElementById("sel_pkPrinter").disabled = false
			//document.getElementById("td_SendToRIS").style.visibility='visible';
			//document.getElementById("sel_SendToRIS").style.visibility='visible';
			//document.getElementById("sel_SendToRIS").disabled = true;
			//get Item table header rows
			var url="ajaxScript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=10"
			url=url+"&addRow=N"
			xmlHttp.onreadystatechange= function()
			{
				if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
				{
					//draw table items
					document.getElementById("td_items").innerHTML = xmlHttp.responseText
				}
			}
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)
			break;
		case "Update":
			// order form
			x = document.getElementById("fm_form")
			x.reset();
			document.getElementById("td_OrderTitle").innerHTML = '<b>Update Order Form</b>';
			document.getElementById("tbl_OrderForm").style.visibility ="visible";
			document.getElementById("sbt_submitForm").style.visibility='hidden';
			document.getElementById("sbt_submitForm").value='Update Order';
			document.getElementById("hd_formType").value='Update';
			document.getElementById("td_items").innerHTML="";
		
			//search form - make visible to begin search
			document.getElementById("txt_getbranchID").style.visibility='visible';
			document.getElementById("txt_getcustNum").style.visibility='visible';
			document.getElementById("txt_getOrderNum").style.visibility='visible';
			document.getElementById("txt_getPONum").style.visibility='visible';
			document.getElementById("txt_getBillName").style.visibility='visible';
			document.getElementById("txt_getShipName").style.visibility='visible';
			document.getElementById("txt_getShipToAddr").style.visibility='visible';
			document.getElementById("txt_getOrderStatus").style.visibility='visible';
			document.getElementById("sel_getSalesMan").style.visibility='visible';
			document.getElementById("sel_getCSRList").style.visibility='visible';
			document.getElementById("td_SendToRIS").style.visibility='visible';
			document.getElementById("sel_SendToRIS").style.display='block';
			document.getElementById("btn_Search").style.visibility='visible';
			document.getElementById("searchResults").style.visibility="visible";
			document.getElementById("searchResults").style.width = "100%";
			for (var i=0;i<x.length;i++)
			{
				t = x.elements[i].type
				switch(t)
				{
					case "text":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = false
					break;
					case "select-one":
					    var selObj = document.getElementById(x.elements[i].id);
						selObj.style.visibility = "visible"
						selObj.disabled = false
					break;
					case "textarea":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = false
					break;
				}				
			}

document.getElementById("sel_Print").style.visibility = 'hidden';
document.getElementById("sel_pkPrinter").style.visibility = 'hidden';
			//get Item table header rows
			var url="ajaxScript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=10"
			url=url+"&addRow=Y"
			xmlHttp.onreadystatechange= function()
			{
				if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
				{
					document.getElementById("td_items").innerHTML = xmlHttp.responseText
				}
			}
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)
			//disable orderID and Order Date fields; hide select emails
			//document.getElementById("txt_findOrderID").disabled = true
			document.getElementById("txt_orderDate").disabled = true
			document.getElementById("sel_Email").style.visibility="hidden"
			
			break;
			
		case "Close":
			// order form
			x = document.getElementById("fm_form")
			x.reset();
			document.getElementById("td_OrderTitle").innerHTML = '<b>Close Order Form</b>';
			document.getElementById("tbl_OrderForm").style.visibility ="visible";
			document.getElementById("sbt_submitForm").style.visibility='visible'//='hidden';
			document.getElementById("sbt_submitForm").value = 'Close Order';
			document.getElementById("hd_formType").value='Close';
			document.getElementById("td_items").innerHTML="";
			
			//search form
			document.getElementById("txt_getbranchID").style.visibility='visible';
			document.getElementById("txt_getcustNum").style.visibility='visible';
			document.getElementById("txt_getOrderNum").style.visibility='visible';
			document.getElementById("txt_getPONum").style.visibility='visible';
			document.getElementById("txt_getBillName").style.visibility='visible';
			document.getElementById("txt_getShipName").style.visibility='visible';
			document.getElementById("txt_getShipToAddr").style.visibility='visible';
			document.getElementById("txt_getOrderStatus").style.visibility='visible';	
			document.getElementById("sel_getSalesMan").style.visibility='visible';
			document.getElementById("sel_getCSRList").style.visibility='visible';		
			document.getElementById("btn_Search").style.visibility='visible';
			document.getElementById("searchResults").style.visibility="visible";
			document.getElementById("searchResults").style.width="100%";
			for (var i=0;i<x.length;i++)
			{
				t = x.elements[i].type
				switch(t)
				{
					case "text":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
					case "select-one":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
					case "textarea":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
				}				
			}
			//get Item table header rows
			//document.getElementById("td_SendToRIS").style.visibility='visible';
			//document.getElementById("sel_SendToRIS").style.visibility='visible';
			//document.getElementById("sel_SendToRIS").disabled = true;
						
			var url="ajaxScript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=10"
			url=url+"&addRow=N"
			xmlHttp.onreadystatechange= function()
			{
				if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
				{
					document.getElementById("td_items").innerHTML = xmlHttp.responseText
				}
			}
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)
			break;
			
		case "Cancel":
			// order form
			x = document.getElementById("fm_form")
			x.reset();
			document.getElementById("td_OrderTitle").innerHTML = '<b>Cancel Order Form</b>';
			document.getElementById("tbl_OrderForm").style.visibility ="visible";
			document.getElementById("sbt_submitForm").style.visibility='visible'//='hidden';
			document.getElementById("sbt_submitForm").value = 'Cancel Order';
			document.getElementById("hd_formType").value='Cancel';
			document.getElementById("td_items").innerHTML="";
			
			//search form
			document.getElementById("txt_getbranchID").style.visibility='visible';
			document.getElementById("txt_getcustNum").style.visibility='visible';
			document.getElementById("txt_getOrderNum").style.visibility='visible';
			document.getElementById("txt_getPONum").style.visibility='visible';
			document.getElementById("txt_getBillName").style.visibility='visible';
			document.getElementById("txt_getShipName").style.visibility='visible';	
			document.getElementById("txt_getShipToAddr").style.visibility='visible';
			document.getElementById("txt_getOrderStatus").style.visibility='visible';
			document.getElementById("sel_getSalesMan").style.visibility='visible';
			document.getElementById("sel_getCSRList").style.visibility='visible';			
			document.getElementById("btn_Search").style.visibility='visible';
			document.getElementById("searchResults").style.visibility="visible";			
			document.getElementById("searchResults").style.width="100%";
			//document.getElementById("td_SendToRIS").style.visibility='hidden';
			//document.getElementById("sel_SendToRIS").style.visibility='hidden';
			for (var i=0;i<x.length;i++)
			{
				t = x.elements[i].type
				switch(t)
				{
					case "text":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
					case "select-one":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
					case "textarea":
						document.getElementById(x.elements[i].id).style.visibility = "visible"
						document.getElementById(x.elements[i].id).disabled = true
					break;
				}				
			}
			//get Item table header rows
			//document.getElementById("td_SendToRIS").style.visibility='visible';
			//document.getElementById("sel_SendToRIS").style.visibility='visible';
			//document.getElementById("sel_SendToRIS").disabled = true;		
			var url="ajaxScript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=10"
			url=url+"&addRow=N"
			xmlHttp.onreadystatechange= function()
			{
				if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
				{
					document.getElementById("td_items").innerHTML = xmlHttp.responseText
				}
			}
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)
			break;	
	}
}

function deleteRow()
{
	r = arguments[0]
	var i=r.parentNode.parentNode.rowIndex//indicates which row this is
	if(formType == "Update")
	{
		var deleteRow = confirm("Delete Line Item");
		if(deleteRow == true)
		{
			var deleteRow_cr = createRequest();
			lineID = r.value;

			var url="ajaxScript.asp"
			url=url+"?sid="+Math.random()
			url=url+"&qry=14" //case SearchCylwithBarcodeGetSerialNumber
			url=url+"&lineID="+lineID
			
			deleteRow_cr.open("GET",url,true)	
			deleteRow_cr.onreadystatechange=function()
			{
				if ((deleteRow_cr.readyState==4) && (deleteRow_cr.status=="200"))
				{ 	
					document.getElementById('tbl_items').deleteRow(i)
				}
			}
	
			deleteRow_cr.send(null)	
		}else{
			r.checked = false
		}
	}else{
		document.getElementById('tbl_items').deleteRow(i)	
	}
}
function addEmail()
{
	arg = arguments[0]
	
	eAddress = document.getElementById("txa_newEmail").value
	custNum = document.getElementById("txt_custNum").value
	
	eLen = eAddress.length
	lastChar = eAddress.charAt(eLen-1)
	if(lastChar != ';')
	{
		eAddress = eAddress+';'	
	}
	
	xmlHttp=GetXmlHttpObject()	
	
		if (xmlHttp==null)
		{
			alert ("Browser does not support HTTP Request")
			return
		}
		
		var url="ajaxScript.asp"
		url=url+"?email="+eAddress	
		url=url+'&custNum='+custNum
		url=url+"&sid="+Math.random()
		url=url+"&qry=9"

		xmlHttp.onreadystatechange= function()
		{
			if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
			{				
				var radTD = xmlHttp.responseText
				sel = document.getElementById("sel_Email")
				selN = sel.name
				sLen = sel.length
				for(j=1;j<sLen;j++)
				{
					sel.remove(1)//cylces through loop and removes option in index[1]
				}
				
				var y = document.createElement('option');
				y.text = "Enter new email address"
				y.value = "New;"
				try
					{
						sel.add(y,null); // standards compliant
					}
					catch(ex)
					{
						sel.add(y); // IE only
					}
				var eMailList = radTD.split(";")
				var eX
				for (eX in eMailList)
				{
					var y = document.createElement('option');
					y.text = eMailList[eX]
					y.value = eMailList[eX]
					try
						{
							sel.add(y,null); // standards compliant
						}
						catch(ex)
						{
							sel.add(y); // IE only
						}
				}
			}
			document.getElementById("txa_newEmail").value = ""	
			document.getElementById("sp_NewEmail").style.visibility = 'hidden'
		}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function chkMail()
{
	// when enter new email address is selected the second txtarea box and add email button is visible
	arg = arguments[0]
	if(arg.value == 'New;')
	{
		document.getElementById("sp_NewEmail").style.visibility = 'visible'
		document.getElementById("txa_newEmail").style.visibility = 'visible'
	}else{
		document.getElementById("sp_NewEmail").style.visibility = 'hidden'
		document.getElementById("txa_newEmail").style.visibility = 'hidden'
	}
}

function chgReq()
{
	arg = arguments[0]
	color = arguments[1]
	
	myTD = (arg).parentNode;
	prevTD = myTD.previousSibling;
	prevTD.style.color = color
}

function getBranchInfo_branchID()
{
	args = arguments[0]
	xmlHttp=GetXmlHttpObject()	
	
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	}
	
	var url="ajaxScript.asp"
	url=url+"?branchID="+args.value	
	url=url+"&sid="+Math.random()
	url=url+"&qry=8" //case SearchCylwithBarcodeGetSerialNumber

	xmlHttp.onreadystatechange= function()
	{
		if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
		{ 
			var str = xmlHttp.responseText
			frm = str.indexOf("|",0)
			radTD = str.substring(0,frm)//(from,to)
			if(radTD==0)
			{
				chgReq(args,"red")
				
				alert("No Records Found");
				document.getElementById("txt_branchID").value = "";
				document.getElementById("td_branchName").innerHTML = "&nbsp;";
				document.getElementById("txt_branchID").focus();
			}else{
				
				chgReq(args,"black")
				
				frm = frm + 1
				to = str.indexOf("|",frm)
				radTD = str.substring(frm,to)
				document.getElementById("td_branchName").innerHTML = radTD;
			}
		}
	}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)	
}

function chkSwap()
{
	arg = arguments[0]
	rowID = (arg.id).substring((arg.id).lastIndexOf("_")+1)// get row number
	cOwn = document.getElementById("sel_custOwn_"+rowID).value
	// onChange check if the cylinder is customer owned. if no, then swap is no. if yes, swap can be yes or no. 
	if(cOwn == 'N')
	{
		alert("Check Customer Owned value")
		arg.options[arg.selectedIndex].text = "Yes"
		arg.options[arg.selectedIndex].value = "Y"
		arg.remove(0)//cylces through loop and removes option in index[1]
		arg.remove(1)//cylces through loop and removes option in index[1]
		var y = document.createElement('option');
		var n = document.createElement('option');
		y.text = "Yes"
		y.value = "Y"
		n.text = "No"
		n.value = "N"		
		try
			{
				arg.add(y,null); // standards compliant
				arg.add(n,null); // standards compliant
			}
			catch(ex)
			{
				arg.add(y); // IE only
				arg.add(n); // IE only
			}
	}else{
		upOrderLine_Values(arg)	
	}
}

function saveQty()
{
	q_arg = arguments[0]
	qu = q_arg.value
}

function chkQtyOrdered()
{
	arg = arguments[0]	
	rowID = (arg.id).substring((arg.id).lastIndexOf("_")+1)
	uom = document.getElementById("td_uom_"+rowID).innerHTML
	if(uom == 'CL')
	{
		dv_chkNumber(arg)
	}else{
			dv_chkChar(event.keyCode,arg,'QtyOrd')
	}
}

function upOrderLine_Values()
{
	arg = arguments[0]
	rowID = (arg.id).substring((arg.id).lastIndexOf("_")+1)// find last '_' get number after '_' = get row number
	lineID = document.getElementById("rd_delete_"+rowID).value// OrderLine table lineID
	
	//check if when formType = update and user has selected an item and entered quantity w/o entering an order id
	if(isNaN(lineID) == true)
	{
		var lineID = 0
	}
//	custOwn = document.getElementById("sel_custOwn_"+rowID).value
//	swap = document.getElementById("sel_swap_"+rowID).value
	price = document.getElementById("txt_price_"+rowID).value
	quantity = document.getElementById("txt_quantity_"+rowID).value
	uom = document.getElementById("td_uom_"+rowID).innerHTML
	ship = document.getElementById("hdn_qtyShipped_"+rowID).value
	stage = document.getElementById("hdn_qtyStage_"+rowID).value

	var tot = parseInt(ship) + parseInt(stage)
	//q = (arg.id).search('txt_quantity')

	if(isNaN(tot))
	{
		tot = 0	
	}
	
	if (quantity < tot && uom == 'CL')
	{
		alert('Quantity ordered can not be less than shipped and staged.')
		// reset txt quantity ordered to original value
		document.getElementById("txt_quantity_"+rowID).value = ""
	}else{
		var upOrderLine_Values_cr = createRequest();
		url = "ajaxscript.asp"
		url=url+"?lineID="+lineID
		url=url+"&sid="+Math.random()
		url=url+"&qry=16"
		//url=url+"&custOwn="+custOwn+"&swap="+swap+"&price="+price+"&quantity="+quantity
		url=url+"&price="+price+"&quantity="+quantity
		//document.write(url)
		upOrderLine_Values_cr.onreadystatechange= function()
		{
			if (upOrderLine_Values_cr.readyState==4 || upOrderLine_Values_cr.readyState=="complete")
			{
				var str = upOrderLine_Values_cr.responseText//Update - Item - nothing returned
				//document.write(str)
			}
		}
			
		upOrderLine_Values_cr.open("GET",url,true)
		upOrderLine_Values_cr.send(null)
	}
}

function chkCylMaint_CustID()
{ 

	arg = arguments[0]
	argID = arg.id
	argV = arg.value
	var ok = dv_chkAlphaN_CE(arg)
// Author = LH, 10/11/07	
	if(argV == '')
	{
		document.getElementById("txtCustSearch").value = ''
		document.getElementById("td_lblCustNFO").innerHTML = ''
	}
// end Author = LH	
	else if(ok == 0)	
//	if(ok == 0) changed 101207 4pm
	{
		var chkCylMaint_CustID_cr = createRequest();
		url = "ajaxscriptB.asp"
		url=url+"?sid="+Math.random()
		url=url+"&qry=9"
		url=url+"&custID="+argV

		chkCylMaint_CustID_cr.onreadystatechange= function()
		{
			if (chkCylMaint_CustID_cr.readyState==4 || chkCylMaint_CustID_cr.readyState=="complete")
			{
				var str = chkCylMaint_CustID_cr.responseText
				if(str == 0)
				{
					alert('Customer ID not found. Please enter a valid Customer ID')
					document.getElementById("txtCustSearch").value = ''
					document.getElementById("td_lblCustNfo").innerHTML = ''
				}else{
					document.getElementById("td_lblCustNFO").innerHTML = str
				}
				//alert(str)
			}
		}
			
		chkCylMaint_CustID_cr.open("GET",url,true)
		chkCylMaint_CustID_cr.send(null)	
	}
}

function getCylItemInfo()
{
	arg = arguments[0] //select object; arg.id = select object id = row number
	rowID = (arg.id).substring((arg.id).lastIndexOf("_")+1)// find last '_' get number after '_' = get row number

	document.getElementById("txt_quantity_"+rowID).disabled = false

	var getCylItemInfo_cr = createRequest();
	if(formType == "Update")
	{
		//lineID is zero if adding a new row to an order; lineID is a value if updating an existing line	
		lineID = document.getElementById("rd_delete_"+rowID).value
		orderID = document.getElementById("td_order").innerHTML
		itemV = document.getElementById("sel_Item_"+rowID).value
		var upCylinder_Item_cr = createRequest();
		url = "ajaxscript.asp"
		url=url+"?orderID="+orderID
		url=url+"&lineID="+lineID
		url=url+"&sid="+Math.random()
		url=url+"&qry=15"
		url=url+"&itemV="+itemV
	
		upCylinder_Item_cr.onreadystatechange= function()
		{
			if (upCylinder_Item_cr.readyState==4 || upCylinder_Item_cr.readyState=="complete")
			{
				var str = upCylinder_Item_cr.responseText//Update - Item - nothing returned
//document.write('1597 ' + str)				
				document.getElementById("rd_delete_"+rowID).value = str
			}
		}
		
		upCylinder_Item_cr.open("GET",url,true)
		upCylinder_Item_cr.send(null)
	}
	var url="ajaxScript.asp"
	url=url+"?supItem="+arg.value
	url=url+"&rowID="+rowID
	url=url+"&sid="+Math.random()
	url=url+"&qry=7"
	
	getCylItemInfo_cr.onreadystatechange= function()
	{
		if (getCylItemInfo_cr.readyState==4 || getCylItemInfo_cr.readyState=="complete")
		{
			var str = getCylItemInfo_cr.responseText

			//var arrTD = new Array("sp_desc_"+rowID,"td_uom_"+rowID)
			var arrTD = new Array("td_uom_"+rowID)
			//document.getElementById("td_desc_"+rowID).innerHTML = str
			frm = str.indexOf("|",0)
			radTD = str.substring(0,frm)//(from,to)
			for (i=0;i<arrTD.length;i++)
			{
				document.getElementById(arrTD[i]).innerHTML=radTD
				frm = frm + 1
				to = str.indexOf("|",frm)
				radTD = str.substring(frm,to)
				frm = to
			}
		}
	}
	getCylItemInfo_cr.open("GET",url,true)
	getCylItemInfo_cr.send(null)
}

function getReturnLoc_RetLoc()
{
	arg = arguments[0]	
	xmlHttp=GetXmlHttpObject()	
	
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	}
	
	var url="ajaxScript.asp"
	url=url+"?retLoc="+arg.value	
	url=url+"&sid="+Math.random()
	url=url+"&qry=6" 

	xmlHttp.onreadystatechange= function()
	{
		if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
		{ 
			var str = xmlHttp.responseText
			frm = str.indexOf("|",0)
			radTD = str.substring(0,frm)//(from,to)
			if(radTD==0)
			{
				chgReq(arg,"red")
				alert("Please select a valid return location");
				//document.getElementById("txt_return").value = ""
				document.getElementById("td_returnLoc").innerHTML = "&nbsp;"
				document.getElementById("td_returnLoc").focus();
			}else{
				chgReq(arg,"black")
				frm = frm + 1
				to = str.indexOf("|",frm)
				radTD = str.substring(frm,to)
				document.getElementById("td_returnLoc").innerHTML = radTD;
				var retLocSelectedIndex = document.getElementById("sel_ReturnLoc").selectedIndex
				document.getElementById("returnLoc").value = document.getElementById("sel_ReturnLoc")[retLocSelectedIndex].value
			}
		}
	}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function fillCustInfo_fromSearch()
{
	arg = arguments[0]
	txtCust = document.getElementById("txt_custNum")
	txtCust.value = arg.value
	getCustInfo_OrderEntry(txtCust)
}

function getCustInfo_OrderEntry()
{
	arg = arguments[0];
	xmlHttp=GetXmlHttpObject()	
	var orderType = document.getElementById("sel_OrderType").value
	
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	}
	var url="ajaxScript.asp"
	url=url+"?custNo="+arg.value	
	url=url+"&sid="+Math.random()
	url=url+"&qry=5" //case SearchCylwithBarcodeGetSerialNumber
	url=url+"&ordertype="+orderType

	xmlHttp.onreadystatechange= function()
	{
		if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
		{ 
			var str = xmlHttp.responseText

			var arrTD = new Array("hd_recCount","td_billTo","txt_ShipToName","txt_ShipToAddr1","txt_ShipToAddr2","hdn_ShipToAddr3","hdn_CusClass", "sel_salesmanID", "sel_WholeSpot", "sel_ReturnLoc", "sel_Email", "txt_ShipToCity", "txt_ShipToState", "txt_ShipToZip","hdn_custSendToRIS")//,"sel_salesmanID")//keep sel_Email as last element in array;
			frm = str.indexOf("|",2) // a|b| then rest of records if found

			var outCount = str.substring(0,1)
			var custMatch = str.substring(2,3)
			
			radTD = str.substring(0,frm)//(from,to)
		
			if(outCount==0 || custMatch == 0)
			{
				if(outCount == 0)
				{
					alert("No Records Found");
				}else{				
					alert('This customer is not allowed for this order type. Please check your order')
				}
				
				chgReq((arg),"red") //change td title to red
				
				document.getElementById("sp_NewEmail").style.visibility = 'hidden'
				for (i=0;i<arrTD.length;i++)
				{
					clName = document.getElementById(arrTD[i]).className
					
					switch(clName)
					{
					case "td":
							document.getElementById(arrTD[i]).innerHTML="&nbsp;"
						break;
					case "formField":
							document.getElementById(arrTD[i]).value=""
							chgReq(document.getElementById(arrTD[i]),"red") //change td title to red							
					break;
					case "formField_NN":
						document.getElementById(arrTD[i]).value=""
					break;
					case "reqSel":
						var s = document.getElementById(arrTD[i])
						if(s.id == 'sel_Email')
							{
								sel = document.getElementById(arrTD[i])
								selN = sel.name
								sLen = sel.length
								for(j=1;j<sLen;j++)
								{
									sel.remove(1)//cylces through loop and removes option in index[1]
								}
							}
					break;						
					}
				}
				document.getElementById("txt_custNum").value = ""
				document.getElementById("txt_custNum").focus();
			}else{
				chgReq((arg),"black")//change to black
				for (i=0;i<arrTD.length;i++)
				{
					clName = document.getElementById(arrTD[i]).className
					
					switch(clName)
					{
						case "td":
							document.getElementById(arrTD[i]).innerHTML=radTD
						break;
						case "formField": //was reqText
							document.getElementById(arrTD[i]).value=radTD
							chgReq(document.getElementById(arrTD[i]),"black") //change td title to black
						break;
						case "formField_NN":
							document.getElementById(arrTD[i]).value=radTD
						break;
						case "custSendToRIS":
						    document.getElementById(arrTD[i]).value=radTD
						    ChkSendToRIS()
						break;
						case "reqSel":
						//alert('1857 fix email trouble')
							var s = document.getElementById(arrTD[i])
							switch(s.id)
							//if(s.id == 'sel_salesmanID')
							{
								case 'sel_WholeSpot':
									t = radTD.indexOf(";",0);
									v = radTD.substring(0,t);
									d = radTD.substring(t+1);
									
									//get current option[0]
									var s0text = s.options[0].text
									var s0value = s.options[0].value
									//set selectedIndex to value from database - new option[0]
									if(d == 'W')
									{
										d = 'Wholesale'
										v = 'W'
									}else{
										d = 'Spot'  
										v = 'S'
									}

									s.options[s.selectedIndex].text = d
									s.options[s.selectedIndex].value = v
									
									if(removeIndex0 == 0)
									{
										//add initial result to options[1] of select list
										var opt=document.createElement('option');
										opt.text = s0text
										opt.value = s0value
										try
											{
											s.add(opt,1); // standards compliant
											}
										  catch(ex)
											{
											s.add(opt); // IE only
											}
									}else{
										//after initial selection of an order id from the order search result list reset index 0 to new values from query
										s.options[0].text = d
										s.options[0].value = v
									}
									break;
								case 'sel_salesmanID':
									//alert(radTD)
									t = radTD.indexOf(";",0);
									v = radTD.substring(0,t);
									d = radTD.substring(t+1);
									
									//get current option[0]
									var s0text = s.options[0].text
									var s0value = s.options[0].value
									//set selectedIndex to value from database - new option[0]
									s.options[s.selectedIndex].text = d
									s.options[s.selectedIndex].value = v
									
									if(removeIndex0 == 0)
									{
										//add initial result to options[1] of select list
										var opt=document.createElement('option');
										opt.text = s0text
										opt.value = s0value
										try
											{
											s.add(opt,1); // standards compliant
											}
										  catch(ex)
											{
											s.add(opt); // IE only
											}
									}else{
										//after initial selection of an order id from the order search result list reset index 0 to new values from query
										s.options[0].text = d
										s.options[0].value = v
									}
									break;
									
								case 'sel_ReturnLoc':
									// radTD will be 0;0 unless this is a WH transfer and there is a defined return location
									// for the selected customer.
									var pos1 = radTD.indexOf(";",0)
									
									var disable_enable = radTD.substring(0,pos1)
									pos1 = pos1+1
									
									var pos2 = radTD.indexOf(";",pos1)			
									var selReturnLocValue = radTD.substring( pos1, pos2)
									
									var returnLocDesc = radTD.substring(pos2+1)

									if ( selReturnLocValue != "0" ) 
									{
										var retLocLength = s.length
										var foundLocation = false
										var returnLocIndex
										
										for ( returnLocIndex = 0; returnLocIndex < retLocLength; returnLocIndex++)
										{
											if ( s[returnLocIndex].value == selReturnLocValue )
											{
												s.selectedIndex = returnLocIndex
											}
										}
										
										document.getElementById("td_returnLoc").innerHTML = returnLocDesc
									}
									
									if ( disable_enable == 'disable' )
									{
										s.disabled = true
									}
									else
									{
										s.disabled = false
									}
									
									var retLocSelectedIndex = document.getElementById("sel_ReturnLoc").selectedIndex
									document.getElementById("returnLoc").value = document.getElementById("sel_ReturnLoc")[retLocSelectedIndex].value
									
									break;
									
								case 'sel_Email':
									prevSelLength  = s.length
									for(j=1;j<prevSelLength;j++)
									{
										//alert(s.length)
										s.remove(1)//cylces through loop and removes option in index[1]
									}	
									
									document.getElementById(arrTD[i]).multiple = true;
									
									var y = document.createElement('option');
										y.text = "Enter new email address"
										y.value = "New;"
										try
											{
												s.add(y,null); // standards compliant
											}
											catch(ex)
											{
												s.add(y); // IE only
											}
				
									var semiC = 0
									//count how many email address delimited by ';'
									for(a=0;a<radTD.length;a++)
									{
										if(radTD.charAt(a) == ";")
										{
											semiC = semiC + 1
										}
									}
									//parse email string delimited by ';'
									t = radTD.indexOf(";",0)
									email = radTD.substring(0,t);
									f = t + 1;
									for(b=0;b<semiC;b++)
									{
										var y = document.createElement('option');
										y.text = email
										y.value = email+";"
										try
											{
												s.add(y,null); // standards compliant
											}
											catch(ex)
											{
												s.add(y); // IE only
											}								
										t = radTD.indexOf(";",f)//(look for, from)
										email = radTD.substring(f,t)//(start, end)
										f = t + 1							
									}
									break;
							}						
						break;
					}
					frm = frm + 1
					to = str.indexOf("|",frm)
					radTD = str.substring(frm,to)
					frm = to
				}
			}
		}	
	}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)		
}

function ChkSendToRIS()
{
    var hdn_custSendToRIS = document.getElementById("hdn_custSendToRIS").value;
    var sel_OrderType = document.getElementById("sel_OrderType").value;
    var hd_formType = document.getElementById("hd_formType").value;
    var orderTypeOK;
    var canSendToRIS;
    switch(sel_OrderType)
	{
		case "PK":
			orderTypeOK = 'Y';
			break;
		case "PB":
			orderTypeOK = 'Y';
			break;	
	    case "PR":
	        orderTypeOK = 'Y';
	        break;
	    case "SH":
	        orderTypeOK = 'Y';
	        break;
	    case "WH":
	       	orderTypeOK = 'Y';
	       	break;
	       default:
	        orderTypeOK = 'V';
	        break;		
	}
    if ((orderTypeOK == 'Y') &&( hdn_custSendToRIS == 'Y' ))
    {
        document.getElementById('hdn_SendToRIS').value='N';
        canSendToRIS = 'Y'
    }
    else
    {
        document.getElementById('hdn_SendToRIS').value='V';
        canSendToRIS = 'V'
    }
    switch (hd_formType)
    {
        case 'Update':
            if(canSendToRIS == 'V')
            {
                document.getElementById("sel_SendToRIS").disabled = true;
                document.getElementById('td_RISStatus').innerHTML = 'Can not send to RIS'
            }else{
                document.getElementById("sel_SendToRIS").disabled = false;
                document.getElementById('td_RISStatus').innerHTML = ''
            }    
            break;
    }    
}

function chgHdnSendToRIS()
{
    var arg = arguments[0]
    document.getElementById('hdn_SendToRIS').value= arg.value;
}

function ajax()
{
	xmlHttp=GetXmlHttpObject()	
	
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	}
	
	var url="ajaxScript.asp"
	url=url+"?rowCount="+rowCount	
	url=url+"&sid="+Math.random()
	url=url+"&qry=4" //case SearchCylwithBarcodeGetSerialNumber
	xmlHttp.onreadystatechange= function()
	{
		if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
		{ 
			alert(url + ' ' + xmlHttp.responseText)
		}		
	}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)	
}

function addRow(){
	//three types of request for add row. one is new order (ADD) and add rows from an existing order (Update). an add row to existing order (Update) is allowed. in that case the addRow() call is similiar to addRow in an ADD order.
	arg = arguments[0]//number of items
	lineID = arguments[1]
	var addRow_cr = createRequest();

	

	var url="ajaxScript.asp"
	url=url+"?fType="+ formType 	
	url=url+"&sid="+Math.random()
	url=url+"&qry=4" //case SearchCylwithBarcodeGetSerialNumber
	url=url+"&lineID="+lineID //if lineID = 0 then adding new row for either 'ADD' or Update else getting existing orderline for lineID
	url=url+'&orderID='+thisUpOrderID// if thisUpOrderID = 0 then 'ADD' else adding new row for an 'Update'
	url=url+'&rowID='+getNextRowID()//in common.js file

	//addRow_cr.open("GET",url,true)	
	addRow_cr.onreadystatechange=function()
	{
		if ((addRow_cr.readyState==4) && (addRow_cr.status=="200"))//(xmlHttp.readyState==4 || xmlHttp.readyState=="complete")//
		{ 
			str = addRow_cr.responseText

			frm = str.indexOf("|",0)
			radTD = str.substring(0,frm)//(from,to)
		
			rowCount = radTD
			var itemTR = document.getElementById('tbl_items').insertRow(2)
			itemTR.id = 'tr_Line_'+rowCount
			//create TDs
			var radTD = itemTR.insertCell(0)
			radTD.id = 'td_rad_'+rowCount
			radTD.innerHTML = '&nbsp;'
			
			var itemTD = itemTR.insertCell(1)
			itemTD.id = 'td_item_'+rowCount
		
			var uomTD = itemTR.insertCell(2)
			uomTD.id = 'td_uom_'+rowCount
			
			var qtyTD = itemTR.insertCell(3)
			qtyTD.id = 'td_qty_'+rowCount
			
			var qtyShTD = itemTR.insertCell(4)
			qtyShTD.id = 'td_qtySh_'+rowCount
			
			var qtyStageTD = itemTR.insertCell(5)
			qtyStageTD.id = 'td_qtyStage_'+rowCount
			
			var priceTD = itemTR.insertCell(6)
			priceTD.id = 'td_price_'+rowCount
			
			//cylinder info
			var itemCyl_Row = document.getElementById('tbl_items').insertRow(3)
			itemCyl_Row.id = 'tr_CylLine_'+rowCount
			//create TDs
			var cylTD0 = itemCyl_Row.insertCell(0)
			cylTD0.id = 'td_cyl0_'+rowCount
			cylTD0.innerHTML = '&nbsp;'
			
			var cylTD1 = itemCyl_Row.insertCell(1)
			cylTD1.id = 'td_cyl1_'+rowCount
			document.getElementById("td_cyl1_"+rowCount).colSpan="6"
			cylTD1.innerHTML = '&nbsp;'
			
			/*(var custOwnTD = itemTR.insertCell(8)
			custOwnTD.id = 'td_custOwn_'+rowCount
			
			var swapTD = itemTR.insertCell(9)
			swapTD.id = 'td_swap_'+rowCount*/
			
			//remember the elements need to be in the same order as the result set from the query. do not let element order not match the result set order
			//var arrTD = new Array("td_rad_","td_item_","td_desc_","td_uom_","td_qty_","td_qtySh_","td_qtyStage_","td_price_","td_custOwn_","td_swap_")
			var arrTD = new Array("td_rad_","td_item_","td_uom_","td_qty_","td_qtySh_","td_qtyStage_","td_price_","td_cyl1_")
			for (i=0;i<arrTD.length;i++)
			{
				frm = frm + 1
				to = str.indexOf("|",frm)
				radTD = str.substring(frm,to)
				document.getElementById(arrTD[i]+rowCount).innerHTML=radTD 
				frm = to
			}
			// turn quantity on or off depending on whether form is update or add
			urlB = window.location.search
			s = urlB.lastIndexOf("=")
			formType = urlB.substring(s+1)		
			if(formType == 'Update')
			{
				document.getElementById("txt_quantity_"+rowCount).disabled = false
			}
	
		}
	}
	addRow_cr.open("GET",url,true)	
	addRow_cr.send(null)
}

function showTrCylInfo()
{
	arg = arguments[0]
	arg2 = arguments[1]

	var trCylInfo = document.getElementById("trCylInfo_"+arg)
	var tdCylInfo = document.getElementById("tdCylInfo_"+arg) 
	if(arg2 == 'H')
	{
		trCylInfo.style.display="none"
		tdCylInfo.innerHTML = "<a href='javascript:showTrCylInfo(&quot;"+arg+"&quot;,&quot;S&quot;)'>Show</a>"
	}else{
		trCylInfo.style.display="block"
		tdCylInfo.innerHTML = "<a href='javascript:showTrCylInfo(&quot;"+arg+"&quot;,&quot;H&quot;)'>Hide</a>"
	}
}

function stateChanged_AddItemRow() 
{ 
if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{ 
		str = xmlHttp.responseText
		var arrTD = new Array("td_rad_","td_item_","td_desc_","td_uom_","td_qty_","td_qtySh_","td_price_","td_custOwn_","td_swap_")
		var pos1 = 0
		frm = str.indexOf("|",0)
		radTD = str.substring(0,frm)//(from,to)
		for (i=0;i<arrTD.length;i++)
		{
			document.getElementById(arrTD[i]+rowCount).innerHTML=radTD
			frm = frm + 1
			to = str.indexOf("|",frm)
			radTD = str.substring(frm,to)
			frm = to
		}
	}
	 	 
} 

function ResetCylForm()
{
	cylinderMode = "Search"	
	location.href="form.asp?menuID=1"
	document.getElementById("rrsForm").elements[0].focus();
}
function rssFormEn()
{
	document.getElementById("sbSubmit").disabled = false;
}
function rssFormDis()
{
	document.getElementById("sbSubmit").disabled = true;	
}

function chkOrderForm2()
{
	//check processForm js variable?
	//document.getElementById("sbSubmit").disabled = true;
	//var x = arguments[0];
	sbt_value = document.getElementById("sbt_submitForm").value
	
	_true = true
	
	switch(sbt_value)
	{
		case "Update Order":
			var upOrder = confirm("Update Order");
			if(upOrder != true)
			{		
				_true = false;
			}
			break;
		case "Cancel Order":
			var upOrder = confirm("'Cancel Order");
			if(upOrder != true)
			{		
				_true = false;
			}
			break;	
	}

	//UN-REM LATER FOR FORM CHECK
	/*for (var i=0;i<x.length;i++)
	{
		if(x.elements[i].className == "reqText" && (x.elements[i].value).length == 0)
		{
			alert("The form is incomplete. Please fill in all required fields " + x.elements[i].id);
			document.getElementById(x.elements[i].id).focus
			_true = false;
			break;
		}
	}*/
	document.getElementById("hd_rowCount").value = rowCount
	//alert(rowCount);
	return _true;
}

function orderEntry_chkForm()
{
	arg = arguments[0]
	argID = arg.id
	argV = arg.value
	
	var formOK = 0

	subForm = confirm(argV);

	if(subForm == true)
	{
		
		oType = document.getElementById("sel_OrderType").value
		cusClass = document.getElementById("hdn_CusClass").value
		
		if(oType == 'WH' && cusClass != 'W')
		{
			formOK = 1
			msg = 'Only warehouse customers are allowed when order type is warehouse transfer';
		}
		
		//check for red fields containing data errors
		var tds = document.getElementsByTagName("td")
		for (i=0;i<tds.length;i++)
		{
			if((tds[i]).style.color == 'red')
			{
				alert(tds[i].innerHTML + ' is invalid. Enter valid ' + tds[i].innerHTML);
				formOK = 1
				break;
			}
		}
		
		//check for blank fields
		var txtbox = document.getElementsByTagName("input")
		for (i=0;i<txtbox.length;i++)
		{
			
			if(txtbox[i].id == 'txt_ShipToAddr1' || txtbox[i].id == 'txt_ShipToAddr2')
			{
				v = txtbox[i].value
				len = v.length
				if(len == 0)
				{
					formOK = 1;	
					//txtbox[i].focus();
					msg = 'Please enter a valid shipping address'		
				}else{
					formOK = 0; 
					break;
				}
			}
			else
			{
			if( (txtbox[i].className=='formField')  && ((txtbox[i].value).length==0) )
				{
					myTD = txtbox[i].parentNode;
					prevTD = myTD.previousSibling;
					alert('Enter valid ' + prevTD.innerHTML);
					txtbox[i].focus();
					formOK = 1;
					break;
				}
			}
			
			switch (txtbox[i].className)
			{
				case "reqQuantity" :
					if(( (txtbox[i].value).length == 0 ) || ( (txtbox[i].value) == 0) )
					{
						msg = 'Invalid Quantity'
						//txtbox[i].focus();
						formOK = 1;
						break;
					}
				break;
				case "reqPrice" :
					if( (txtbox[i].value).length == 0)
					{
						txtbox[i].value = '0'
						break;
					}
				break;
			}
		}
		
		//check select lists for value = 0
	
		var sel = document.getElementsByTagName("select")
		for (i=0;i<sel.length;i++)
		{
			if( ( (sel[i].className=='reqSel') ) && ( sel[i].value == 9999999 ) )
			{
				myTD = sel[i].parentNode;
	            prevTD = myTD.previousSibling;
				alert('Select a valid ' + prevTD.innerHTML);
				sel[i].focus();
				formOK = 1;
				break;
			}
		}
		t_rows = document.getElementById('tbl_items').rows.length
		if(t_rows == 2) //there are two table rows by default
		{
			formOK = 1
			msg = 'Orders must have at least one item. Please add an item for this order.'
		}
		if(formOK == 0)
		{
		    document.getElementById("sel_OrderType").disabled=false;
		    document.getElementById("txt_CustNum").disabled=false;
			x = document.getElementById("fm_form")
			x.submit()
			//alert('submitted')
		}else{
			alert(msg)	
		}        		
	}
}

function chkCylForm()
{
	//check processForm js variable?
	//document.getElementById("sbSubmit").disabled = true;
	var x = arguments[0];
	sb_value = document.getElementById("sbSubmit").value
	
	_true = true
	
	/*switch(sb_value)
	{
		case "Update Order":
			var upOrder = confirm("Update Order");
			if(upOrder != true)
			{		
				_true = false;
			}
			break;
		case "Delete Order":
			var upOrder = confirm("'Delete Order");
			if(upOrder != true)
			{		
				_true = false;
			}
			break;	
	}*/
	/*tagName_td = document.getElementsByTagName("td")
	for(var i=0;i<tagName_td.length;i++)
	{
		ih = (tagName_td[i].innerHTML).style.color
		alert(ih);
	}*/
	//UN-REM LATER FOR FORM CHECK
	/*for (var i=0;i<x.length;i++)
	{
		if(x.elements[i].className == "reqText" && (x.elements[i].value).length == 0)
		{
			alert("The form is incomplete. Please fill in all required fields " + x.elements[i].id);
			document.getElementById(x.elements[i].id).focus
			_true = false;
			break;
		}
	}*/
	//document.getElementById("hd_rowCount").value = rowCount
	//alert(rowCount);
	return _true;
}

function GetXmlHttpObject()
{ 
	var objXMLHttp=null
	if (window.XMLHttpRequest)
	{
		objXMLHttp=new XMLHttpRequest()
	}
	else if (window.ActiveXObject)
	{
		objXMLHttp=new ActiveXObject("Microsoft.XMLHTTP")
	}
	return objXMLHttp
}

function clickMe()
{
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	} 
	
	var url="ajaxScript.asp"
	url=url+"?branchID="+1
	url=url+"&sid="+Math.random()
	xmlHttp.onreadystatechange=stateChanged_e
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)	
}

function stateChanged_e() 
{ 
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
		{ 
			document.getElementById("td1").innerHTML = xmlHttp.responseText
		}
	 	 
} 

function chkV()//vCylSearch 
{
	if(vCylSearch.length!=0)
	{
		document.getElementById("sbSubmit").disabled = true
	}
}

function ckBCSerialLength()
{
	barcode = document.getElementById("txtBarCode"); //barcode object
	BCLength = (barcode.value).length
	serial = document.getElementById("txtSerialNumber"); //serial object
	SerialLength = (serial.value).length 
	if(BCLength == 0 && SerialLength == 0)
	{
		cylinderMode = "Search"
	}
}

function ckCylForm()
{
	arg = arguments[0]; //barcode object
	if(arg.id == "txtBarCode")
	{
		arg1="barcode"
		arg2="txtSerialNumber"
	}else{
		arg1="serial"
		arg2="txtBarCode"
	}
	arg3 = arg
	//document.getElementById(arg.id).disabled=true
	xmlHttp=GetXmlHttpObject()
	xmlHttp2=GetXmlHttpObject()	
	
	if (xmlHttp==null)
	{
		alert ("Browser does not support HTTP Request")
		return
	} 
	
	//var url="RSSajaxScript.asp"
	switch(cylinderMode)
	{
		case "Add Cylinder" :
			var url="ajaxScript.asp"
			url=url+"?arg1="+arg.value
			url=url+"&searchBy="+arg1		
			url=url+"&sid="+Math.random()
			url=url+"&qry=3" //case SearchCylwithBarcodeGetSerialNumber
			xmlHttp.onreadystatechange=stateChanged_BCorSerialExists
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)
			break;
		case "Update Cylinder" :
			//document.getElementById(arg.id).disabled=true
			var url="ajaxScript.asp"
			url=url+"?arg1="+arg.value
			url=url+"&searchBy="+arg1		
			url=url+"&sid="+Math.random()
			url=url+"&qry=3" //case SearchCylwithBarcodeGetSerialNumber
			//xmlHttp.onreadystatechange=stateChanged_BCorSerialExists
			xmlHttp.onreadystatechange=function()
			{
			if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
					{ 
						var found = xmlHttp.responseText
						if(found == 1)
						{
							alert("Cylinder "+arg1+" number exists. Please enter a new "+arg1+" number")
							document.getElementById(arg3.id).value = vCylSearch;
							document.getElementById(arg3.id).focus;
							rssFormDis()
							processForm = "No"
						}else{
							rssFormEn()
							processForm = "Yes"	
						}
					}				
			}
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)			
			break;
		default : 
			//get serial number when searching with a barcode or get barcode when searching with a serial number
			var url="ajaxScript.asp"
			url=url+"?arg1="+arg.value//can be a serial or barcode
			url=url+"&searchBy="+arg1		
			url=url+"&sid="+Math.random()
			url=url+"&qry=1" //if arg1 = barcode then serial returned; if arg1 = serial then barcode retuned
			xmlHttp.onreadystatechange=stateChanged_SearchCylinder
			xmlHttp.open("GET",url,true)
			xmlHttp.send(null)	
			//search for cylinder info after getting barcode (if serial sent) or serial (if barcode sent)
			var url="ajaxScript.asp"
			url=url+"?arg1="+arg.value//can be a serial or barcode
			url=url+"&searchBy="+arg1
			url=url+"&sid="+Math.random()
			url=url+"&qry=2" //

			xmlHttp2.onreadystatechange=stateChanged_SearchCylinderValues
			xmlHttp2.open("GET",url,true)
			xmlHttp2.send(null)
		break;
	}	
}

function stateChanged_BCorSerialExists() 
{ 
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
		{ 
			var found = xmlHttp.responseText
			if(found == 1)
			{
				alert("Cylinder "+arg1+" number exists. Please enter a new "+arg1+" number")
				document.getElementById(arg3.id).value = vCylSearch;
				document.getElementById(arg3.id).focus;
				rssFormDis()
				processForm = "No"
			}else{
				rssFormEn()
				processForm = "Yes"	
			}
		}
}

function stateChanged_SearchCylinder() 
{ 
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
		{ 
			document.getElementById(arg2).value= xmlHttp.responseText
			vCylSearch = xmlHttp.responseText// either serial(if barcode search) or barcode(if serial search)
			//rssFormDis()
			alert('1445 ' + vCylSearch)
		}
}

function stateChanged_SearchCylinderValues() 
{ 
	if (xmlHttp2.readyState==4 || xmlHttp2.readyState=="complete")
		{ 
			document.getElementById("td_results").innerHTML=xmlHttp2.responseText
			cylinderMode = document.getElementById("sbSubmit").value; //cylinder is either being added or updated
			//rssFormDis()
		}

}

function openQuoteform(orderid,formType)
{
    if(orderid==0) {
      
        var num = Math.round(Math.random() * 100000) * -1
  

        var mywindow = window.open("http://rrs/ordQuote_form.asp?menuid=67&ordid=" + num + "&formType=Add", "GasQuotes", "scrollbars=1,width=650,height=600");  
         

        }
    
    else
        {

        var mywindow = window.open("http://rrs/ordQuote_form.asp?menuid=67&ordid=" + orderid + "&formType=" + formType + "", "GasQuotes", "scrollbars=1,width=650,height=600");  
        //var myWindow = window.open("ordQuote_form.asp?menuid=67&process=quote_init&ordid="+orderid+"&formType="+formType+"", "GasQuotes");        

        }
        
  
    
}
      
        