<!-- #include file = subs.asp -->
<!-- #include file = conSQL.asp -->

<%
sub SearchCylwithBarcodeGetSerialNumber(barcode)
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.getCylInfo"
		.Parameters.Append oraComm.CreateParameter("@inBarcode",200,&H0001,50,arg1)	'can be a serial or sub
		.Parameters.Append oraComm.CreateParameter("@inSearchBy",200,&H0001,9,searchBy)	
		.Parameters.Append oraComm.CreateParameter("@outCount",200,adParamOutput,255,"")
		outCount = .parameters("@outCount")			
	end with
	set rs = oraComm.execute	
	do while not rs.eof
		response.Write(rs(0))'rs(0) is either serial (arg1 = a barcode) or a barcode (arg1 = a serial)
		rs.movenext
	loop
end sub

sub CylMaintenance_SearchBarcode(barcode)		
	response.Write("<table>")	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.getCylInfo"
		.Parameters.Append oraComm.CreateParameter("@inBarcode",200,&H0001,50,arg1)'can be a serial or barcode
		.Parameters.Append oraComm.CreateParameter("@inSearchBy",200,&H0001,9,searchBy)	
		.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
		outCount = .parameters("@outCount")	
	end with
	set rs = oraComm.execute
	recCount = rs.recordcount
	if (recCount = 0)then
		call entryForm(4) 'menuID is 4 - empty cylinder info form elements
		
		response.Write("<tr>")
			response.Write("<td colspan='2'>")
				response.Write("<input type='hidden' id='hdNumber' name='numberType' value='"&arg1&"'>")
				response.Write("<input type='hidden' id='hdProcess' name='process' value='AddUpCyl'>")
				response.Write("<input type='hidden' id='hdMenuID' name='menuID' value='9999'>")
				response.Write("<input type='hidden' id='hdAddUp' name='AddUp' value='AddCyl'>")
				response.Write("<input id='sbSubmit' name='cylSubmit' type='submit' value='Add Cylinder' disabled>")
				response.Write("<input type='reset' value='Reset Form'  onClick='ResetCylForm()' />")
			response.Write("</td>")
		response.Write("</tr>")	
		
	else
		call CylInfoValues_Barcode(barcode)'in subs.asp; gets cylinder info. can search by barcode or serial number
		response.Write("<tr>")
			response.Write("<td colspan='2'>")
				response.Write("<input type='hidden' id='hdNumber' name='numberType' value='"&searchBy&"'>")
				response.Write("<input type='hidden' id='hdProcess' name='process' value='AddUpCyl'>")
				response.Write("<input type='hidden' id='hdMenuID' name='menuID' value='9999'>")
				response.Write("<input type='hidden' id='hdAddUp' name='AddUp' value='UpCyl'>")
				response.Write("<input id='sbSubmit' name='cylSubmit' type='submit' value='Update Cylinder' disabled>")
				response.Write("<input type='reset' value='Reset Form' onClick='ResetCylForm()'/>")				
			response.Write("</td>")
		response.Write("</tr>")			
		
	end if
end sub

sub BCorSerialExists 
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getCylInfo"
			.Parameters.Append oraComm.CreateParameter("@inBarcode",200,&H0001,9,arg1)
			.Parameters.Append oraComm.CreateParameter("@inSearchBy",200,&H0001,9,searchBy)	
			.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
			'outCount = .parameters("@outCount")	
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outCount")			
		response.Write(outCount)		
	end with
end sub

sub addItemRow(fType, lineID, orderID)
'addItemRow either allows for an add (new order or update order) or it retreives the info for existing order
'if the order has existing item rows, the item row info is retrieved from the db. in the case of the select list, the order's
'cylinder type is the first option in the select list along with the avail cylinder items. if a new row is added, then a list of cylinder items is returned
'check sproc getCylItems for intial values and how the item row is built
	'rowID = 0
'	Randomize
'	max = 1000000
'	min = 1
'	'rowID = (Int((max-min+1)*Rnd+min))
'
'	rowID = Int((Rnd*999999999998)+1) 
	rowID = request.QueryString("rowID")
	if(lineID <> 0 and fType <> "Update")then 'not an Add or Update / disable if Close or Inquiry Request
		dis = "disabled"
	end if
	response.Write(rowID&"|")
		set oraComm = Server.CreateObject("ADODB.Command")
		with oraComm
			with oraComm
				.activeconnection = conn
				.CommandType = 4 'adCmdStoredProc            
				.CommandText = "dbo.getCylItems"
				.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,6,lineID)	
				.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,6,orderID)	
				.Parameters.Append oraComm.CreateParameter("@out_desc",200,&H0002,255,"")
				.Parameters.Append oraComm.CreateParameter("@out_uom",200,&H0002,255,"")
				.Parameters.Append oraComm.CreateParameter("@out_lineID",200,&H0002,255,"")
				.Parameters.Append oraComm.CreateParameter("@outTot",200,&H0002,255,"")
				'outCount = .parameters("@outCount")
			end with
			set rs = oraComm.execute	
			out_desc = .parameters("@out_desc")
			out_uom = .parameters("@out_uom")
			out_lineID = .parameters("@out_lineID")
			outTot = .parameters("@outTot") 'totat of staged and shipped cylinders for line
		end with
		
		if (outTot = "0") then
			delItem = "<input type=checkbox id='rd_delete_"&rowID&"' name='delete_"&rowID&"' onClick='deleteRow(this)' value='"&out_lineID&"' "&dis&" >"
		else
			delItem = "<input type='hidden' id='rd_delete_"&rowID&"' name='delete_"&rowID&"' value='"&out_lineID&"'>"
		end if
		
		response.Write(delItem)
		response.Write("|")
		response.Write("<select id='sel_Item_"&rowID&"' name='item_"&rowID&"' style='width:490' onChange=getCylItemInfo(this) "&dis&">")
		'remember options including intial "pick item" option are built from getCylitems sproc
			do while not rs.eof
				itemMiscID = rs(0)
				supItemN = rs(1)&"-"&rs(2)
				desc = rs(3)
				custOwn = rs(5)
				response.Write("<option value = '"&itemMiscID&"'>{"&supItemN&"} "&desc&" "&custOwn&" "&"</option>")
				spaceC = ""
			rs.movenext
			loop
		response.Write("</select>")	
	response.Write("<br>")				
	response.Write("<span id='sp_desc_"&rowID&"'>")
		'response.Write(out_desc)
		response.Write("</span>")
	response.Write("|")
		response.Write(out_uom)
	response.Write("|")
	
	'get qtyOrdered, qtyShipped, customer owned, swap if lineID not 0 meaning this is not an add
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.getLineValues_lineID"
		.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,6,out_lineID)	
	end with
	set rs = oraComm.execute	
	c = 0
	do while not rs.eof
		qtyOrd = rs(0)
		qtyShip = rs(1)
		qtyStage = rs(2)
		price = rs(3)
		receive = rs(4)
		c = c + 1
	rs.movenext
	loop
	if(len(price) = 0) then
		price = 0
	end if
	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.cylinder_fromlineID"
			.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,6,out_lineID)	
			.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@outOrderType",200,&H0002,255,"")
		end with
		set rsCyl = oraComm.execute	
		outCount = .parameters("@outCount") 'totat of staged and shipped cylinders for line
		outOrderType = .parameters("@outOrderType")
	end with	

	if(outCount = 0) then
		trCyl = ""
	else
		if(outOrderType = "ship")then		
			trCyl = "<tr>"
				trCyl = trCyl & "<td id='tdCylInfo_"&rowID&"' width='5%'><a href='javascript:showTrCylInfo(&quot;"&rowID&"&quot;,&quot;H&quot;)'>Hide</a></td>"
				trCyl = trCyl & "<td width='15%'>Barcode</td><td width='30%'>Serial</td><td width='15%'>Status</td><td width='30%'>Last Scan Date</td>"
			trCyl = trCyl & "</tr>"
			cols = 5
		else		
			trCyl = "<tr>"
				trCyl = trCyl & "<td id='tdCylInfo_"&rowID&"' width='5%'><a href='javascript:showTrCylInfo(&quot;"&rowID&"&quot;,&quot;H&quot;)'>Hide</a></td>"
				trCyl = trCyl & "<td width='10%'>Barcode</td><td width='10%'>Serial</td><td width='5%'>Status</td><td width='15%'>Last Scan Date</td>"
				trCyl = trCyl & "<td width='10%'>Gas Name</td><td width='10%'>Primary</td><td width='5%'>Purity</td><td width='5%'>Net Boil Off</td>"
				trCyl = trCyl & "<td width='10%'>Cust Owned</td>"
			trCyl = trCyl & "</tr>"	
			cols = 10	
		end if
	end if

	trCyl = trCyl & "<tr id='trCylInfo_"&rowID&"'><td colspan="&cols&">"
	trCyl = trCyl & "<table width='100%' border=0>"
	
	do while not rsCyl.eof
		barcode = rsCyl(0)
		serial = rsCyl(1)
		statusName = rsCyl(2)
		lsDate = rsCyl(3)
		custOwn = rsCyl(4)
		purity = rsCyl(6)
		boilOff = rsCyl(7)
		repairNeeded = rsCyl(8)
		if(repairNeeded = 1)then
			repair = "Y"
		else
			repair = "N"
		end if
		gasName = rsCyl(9)
		primaryGasName = rsCyl(10)
		netBoilOff = rsCyl(11)
	
		if(outOrderType = "ship")then
			trCyl = trCyl & "<tr>"
				trCyl = trCyl & "<td width='5%'>&nbsp;</td>"
				trCyl = trCyl &"<td width='15%'>"&barcode&"</td><td width='30%'>"&serial&"</td><td width='15%'>"&statusName&"</td><td width='30%'>"&lsDate&"</td>"
			trCyl = trCyl & "</tr>"
		else
			trCyl = trCyl & "<tr>"
			trCyl = trCyl & "<td width='5%'>&nbsp;</td>"
			trCyl = trCyl &"<td width='10%'>"&barcode&"</td><td width='10%'>"&serial&"</td><td width='5%'>"&statusName&"</td><td width='15%'>"&lsDate&"</td>"
			trCyl = trCyl &"<td width='10%'>"&gasName&"</td><td width='10%'>"&primaryGasName&"</td><td width='5%'>"&FormatNumber(Cdbl(purity)*100,2,-1)&"</td>"
			
			if netBoilOff = "" then
				trCyl = trCyl & "<td width='5%'>&nbsp;</td>"
			else
				trCyl = trCyl & "<td width='5%'>"&FormatNumber(Cdbl(netBoilOff),0,,,0)&"</td>"
			end if 			
			
			trCyl = trCyl &"<td width='10%'>"&custOwn&"</td>"			
			trCyl = trCyl & "</tr>"		
		end if	
		rsCyl.movenext
	loop	
	
	trCyl = trCyl & "</table>"	
	
		
response.Write("<input type='text' id='txt_quantity_"&rowID&"' name='quantity_"&rowID&"' value='"&qtyOrd&"' size='5' disabled  onChange='upOrderLine_Values(this)' onKeyUp='chkQtyOrdered(this)' class='reqQuantity' onFocus='saveQty(this)'>")		
	response.Write("|")
		response.Write("<input type='hidden' id='hdn_qtyShipped_"&rowID&"' name='qtyShipped_"&rowID&"' value='"&qtyShip&"' size='5' disabled>")
		response.Write(qtyShip&"/"&receive)	
	response.Write("|")
		response.Write("<input type='hidden' id='hdn_qtyStage_"&rowID&"' name='qtyStage_"&rowID&"' value='"&qtyStage&"' size='5' disabled>")	
		response.Write(qtyStage)
	response.Write("|")
		response.Write("<input type='text' id='txt_price_"&rowID&"' name='price_"&rowID&"' value='"&price&"' "&dis&" size='5' onChange='upOrderLine_Values(this)' onKeyUp='dv_chkDecimal(this)' class='reqPrice' >")	
	response.Write("|")
	with response
		.Write("<table width='100%'>")
			.Write(trCyl)
		.Write("</table> ")
	end with
	response.Write("|")
				
end sub

sub getCustInfo_OrderEntry
	custNo = request.QueryString("custNo")
	orderType = request.QueryString("orderType")
	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getCustInfo_OrderEntry"
			.Parameters.Append oraComm.CreateParameter("@inCustNum",200,&H0001,6,custNo)
			.Parameters.Append oraComm.CreateParameter("@inOrderType",200,&H0001,10,orderType)
			.Parameters.Append oraComm.CreateParameter("@inUserID",200,&H0001,10,session("username"))
			.Parameters.Append oraComm.CreateParameter("@outRecCount",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@outCustMatch",200,&H0002,255,"")
			'.Parameters.Append oraComm.CreateParameter("@outOrderCount",200,&H0002,255,"")	
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outRecCount")	
		outCustMatch = .parameters("@outCustMatch")		
		response.Write(outCount &"|")	
		response.Write(outCustMatch &"|")	
	end with
	
	do while not rs.eof
		'br = rs(0)
		'brName = rs(1)
		btName = rs(0)
		stName = rs(1)
		stAdd1 = rs(2)
		stAdd2 = rs(3)
		stAdd3 = rs(4)
		cusClass = rs(5)
		email = rs(6)
		salesman = rs(9)
		wholespot = rs(10)
		returnLocInfo = rs(11)
		st_City = rs("st_city")
		st_State = rs("st_st")
		st_zip = rs("st_zip")
		custSendToRIS = rs("custSendToRIS")
		rs.movenext
	loop
	
	response.Write(btName&"|"&stName&"|"&stAdd1&"|"&stAdd2&"|"&stAdd3&"|"&cusClass&"|"&salesman&"|"&wholespot&"|"&returnLocInfo&"|"&email&"|"&st_City&"|"&st_State&"|"&st_zip&"|"&custSendToRIS&"|")

end sub

sub getReturnLoc_RetLoc
	retLoc = request.QueryString("retLoc")

	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getReturnLoc_RetLoc"
			.Parameters.Append oraComm.CreateParameter("@inRetLoc",200,&H0001,3,retLoc)
			.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
			'.Parameters.Append oraComm.CreateParameter("@outOrderCount",200,&H0002,255,"")	
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outCount")			
		response.Write(outCount &"|")
		do while not rs.eof
			branchName = rs(0)
			branchID = rs(1)
			response.Write(branchName&" - "&branchID&"|")
		rs.movenext
		loop		
	end with
end sub

sub getCylItemInfo
	supItem = request.QueryString("supItem")
	rowID = request.QueryString("rowID")
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.getCylItems_itemID"
		.Parameters.Append oraComm.CreateParameter("@inItemID",200,&H0001,8,supItem)
	end with
	set rs = oraComm.execute
	do while not rs.eof
		desc = rs(0)
		uom = rs(1)
		'response.Write(desc&"|"&uom&"|")
		response.Write(uom&"|")
	rs.movenext
	loop
end sub

sub getBranchInfo_branchID
	branchID = request.QueryString("branchID")
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getBranchInfo_branchID"
			'.Parameters.Append oraComm.CreateParameter("@getBranchInfo_branchID",200,&H0001,25,branchID)
			'.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
			'.Parameters.Append oraComm.CreateParameter("@outOrderCount",200,&H0002,255,"")	
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outCount")			
		response.Write("1" &"|")
		do while not rs.eof
			branchID = rs(0)
			branchName = rs(1)
			response.Write(branchName&" - "&branchID&"|")
		rs.movenext
		loop		
	end with
end sub

sub addEmail
	email = request.QueryString("email")
	custNum = request.QueryString("custNum")
	set oraComm = Server.CreateObject("ADODB.Command")	
	with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.updateEmail"
		.Parameters.Append oraComm.CreateParameter("@inEmail",200,&H0001,255,email)
		.Parameters.Append oraComm.CreateParameter("@incustNum",200,&H0001,255,custNum)
		'.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
		'.Parameters.Append oraComm.CreateParameter("@outOrderCount",200,&H0002,255,"")	
	end with
	set rs = oraComm.execute
	do while not rs.eof
		response.Write(rs(0))
	rs.movenext
	loop
end sub

sub getItems
	addRow = request.QueryString("addRow")
	js = "&nbsp;"
	menuChoice = request.QueryString("menuChoice") 'add or maintenance menu choice
	if(addRow = "Y")then
		js = "<a href='javascript:addRow(&quot;0&quot;,&quot;0&quot;)'>Add Item row</a>"
	end if
	with response
		.Write("<table width='100%' border=0 id='tbl_items' name='tableItems'>")
		.Write("<tr>")
			'.write("<td>&nbsp;</td>")
			.write("<td colspan=10>"&js&"</td>")
		.write("</tr>")
		.Write("<tr id='tr_itemHeader'>")
			.Write("<td width='10%'>Click to delete</td>")
			.Write("<td width='50%'>Item</td>")
			'.write("<td width='10%'>Description</td>")
			.write("<td width='5%'>UOM</td>")
			.write("<td width='10%' >Qty Ordered</td>")
			.write("<td width='8%'>Ship/Receive</td>")
			.write("<td width='10%'>Qty Staged</td>")
			.write("<td width='5%' >Price</td>")
'			.write("<td width='10%'>Customer Owned</td>")
'			.Write("<td width='10%'>Swap</td>")
		.Write("</tr>")
		.Write("</table>")
	end with
end sub

sub getOrderResults
	branchID = request.QueryString("branchID")
	custNum = request.QueryString("custNum")
	orderID = request.QueryString("orderID")
	poNum = request.QueryString("poNum")
	billTo = request.QueryString("billTo")
	shipTo = request.QueryString("shipTo")
	shipToAddr = request.QueryString("shipToAddr")
	orderStatus = request.QueryString("orderStatus")
	salesmanID = request.QueryString("salesmanID")
	csr = request.QueryString("csr")
	formType = request.QueryString("formType")
		
	if(branchID = "")then
		branchID = NULL
	end if
	
	if(custNum = "")then
		custNum = NULL
	end if
	
	if(orderID = "")then
		orderID = NULL
	end if
	
	if(poNum = "")then
		poNum = NULL
	end if
	
	if(billTo = "")then
		billTo = NULL
	end if
	
	if(shipTo = "")then
		shipTo = NULL
	end if	
	
	if(shipToAddr = "")then
		shipToAddr = NULL
	end if
	
	if(orderStatus = "")then
		orderStatus = NULL
	end if	
	
	if(salesmanID = "")then
		salesmanID = NULL
	end if
	
	if(csr = "")then
		csr = NULL
	end if			

	response.Write("<select id='sel_searchResults' name='searchResults' size='3' style='visibility:visible;width=100%' onChange='getOrderInfo_ID(this)'>")
	
	set oraComm = Server.CreateObject("ADODB.Command")	
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getOrderResults"
			.Parameters.Append oraComm.CreateParameter("@inBranchID",200,&H0001,255,branchID)
			.Parameters.Append oraComm.CreateParameter("@inCustN",200,&H0001,255,custNum)
			.Parameters.Append oraComm.CreateParameter("@inOrderId",200,&H0001,255,orderID)
			.Parameters.Append oraComm.CreateParameter("@inPONum",200,&H0001,255,poNum)
			.Parameters.Append oraComm.CreateParameter("@inBillTo",200,&H0001,255,billTo)
			.Parameters.Append oraComm.CreateParameter("@inShipTo",200,&H0001,255,shipTo)
			.Parameters.Append oraComm.CreateParameter("@inFormType",200,&H0001,255,formType)
			.Parameters.Append oraComm.CreateParameter("@inShipToAddr",200,&H0001,255,shipToAddr)			
			.Parameters.Append oraComm.CreateParameter("@orderStatus",200,&H0001,255,orderStatus)
			.Parameters.Append oraComm.CreateParameter("@inSalesmanID",200,&H0001,255,salesmanID)			
			.Parameters.Append oraComm.CreateParameter("@inCSR",200,&H0001,255,csr)			
			.Parameters.Append oraComm.CreateParameter("@outParam",200,&H0002,500,"")
			'.Parameters.Append oraComm.CreateParameter("@outOrderCount",200,&H0002,255,"")	
		end with
		set rs = oraComm.execute
		outParam = .parameters("@outParam")	
	end with

	do while not rs.eof
		orderID = rs(0)
		custNo = rs(1)
		custPO = rs(2)
		orderDate = rs(3)
		ordStatus = rs(4)
		shipInfo = rs(5)
		billTo = rs(6)
		response.Write("<option value = '"&orderID&"'>"&orderID&" "&custNo&" "&custPO&" "&orderDate&" Bill To: "&billTo&" Ship To: "&shipInfo&" Status: "&ordStatus&"</option>")
	rs.movenext
	loop	
	
	response.Write("</select>")
end sub

sub getOrderInfo_ID
	orderID = request.QueryString("orderID")
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm	
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getOrderInfo_ID"
			.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,255,OrderID)
			'.Parameters.Append oraComm.CreateParameter("@incustNum",200,&H0001,255,custNum)
			.Parameters.Append oraComm.CreateParameter("@outLineCount",200,&H0002,255,"")
			'.Parameters.Append oraComm.CreateParameter("@outOrderCount",200,&H0002,255,"")	
		end with
		set rs = oraComm.execute
		outLineCount = .parameters("@outLineCount")	
	end with
	'this order must correspond & exactly match the result sequence from the stored procedure
	do while not rs.eof
		orderType = rs(0)
		orderID = rs(1)	
		orderDate = rs(2)
		branchID = rs(3)
		custNo = rs(4)
		custPO = rs(5)
		bt_name = rs(6)
		tickets = rs(7)
		st_name = rs(8)
		retLocID = rs(9)
		retLocDesc = rs(10)
		st_add1 = rs(11)
		returnPkt = rs(12)
		st_add2 = rs(13)
		proNumber = rs(14)
		st_add3 = rs(15)
		freightOut = rs(16)			
		payment = rs(17)		
		email = rs(18)		
		shipDate = rs(19)
		shipVia = rs(20)
		shipViaMisc = rs(21)
		wholeSpot = rs(22)		
		printTkt = rs(23)
		'printer = rs(24)
		comments = rs(25)
		cusClass = rs(26)	
		ordStatus = rs(27)	
		orderWho = rs(28)
		slsID = rs(29)
		shipToCity = rs("shipToCity")
		shipToSt = rs("shipToSt")
		shipToZip = rs("shipToZip")
		calledInBy = rs("calledInBy")
		calledInByPh = rs("calledInByPh")
		sendToRIS = rs("sendToRIS")
		ris_id = "<a href = 'http://intranet.refron.com/apps/order_invoice_lookup/order.asp?OrderID=" & rs("ris_id") & "'  target='_blank'>" & rs("ris_id")&"</a>"
		shipNFO = rs("shipNFO")
		billNotes = rs("billNotes")
		shipNotes = rs("shipNotes")
		onDockDate = rs("onDockDTime")
	rs.movenext
	loop
	'this order must correspond & exactly match the do while sequence 
	response.Write(sendToRIS&"|")
	response.Write(orderType&"|"&orderID&"|"&orderDate&"|"&branchID&"|"&custNo&"|"&custPO&"|")
	response.Write(bt_name&"|"&tickets&"|"&st_name&"|"&retLocID&"|"&retLocDesc&"|")
	response.Write(st_add1&"|"&returnPkt&"|"&st_add2&"|"&proNumber&"|"&st_add3&"|")
	response.Write(freightOut&"|"&payment&"|"&email&"|"&shipDate&"|"&shipVia&"|")
	response.Write(shipViaMisc&"|"&wholeSpot&"|")
	'response.Write(printTkt&"|"&printer&"|"&comments&"|"&cusClass&"|"&ordStatus&"|"&orderWho&"|"&outLineCount&"|")
	response.Write(printTkt&"|"&comments&"|"&cusClass&"|"&ordStatus&"|"&orderWho&"|"&slsID&"|")
	response.Write(shipToCity&"|"&shipToSt&"|"&shipToZip&"|"&calledInBy&"|"&calledInByPh&"|")
	response.Write(ris_id&"|"&shipNFO&"|"&billNotes&"|"&shipNotes&"|"&onDockDate&"|")
	response.Write(outLineCount&"|")
end sub

sub getLineID
	orderID = request.QueryString("orderID")
	set oraComm = Server.CreateObject("ADODB.Command")
	'with oraComm	
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getLineID"
			.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,255,OrderID)
			'.Parameters.Append oraComm.CreateParameter("@incustNum",200,&H0001,255,custNum)
			'.Parameters.Append oraComm.CreateParameter("@outLineCount",200,&H0002,255,"")
			'.Parameters.Append oraComm.CreateParameter("@outOrderCount",200,&H0002,255,"")	
		end with
		set rs = oraComm.execute
		'outLineCount = .parameters("@outLineCount")
		do while not rs.eof
			lineID = rs(0)
			response.Write(lineID&";")
		rs.movenext
		loop	
	'end with
end sub

sub deleteOrderLine
	lineID = request.QueryString("lineID")
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.deleteOrderLine"
		.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,18,lineID)
	end with
	set rs = oraComm.execute
end sub

sub upCylinder_Item
	itemV = request.QueryString("itemV")
	orderID = request.QueryString("orderID")
	lineID = request.QueryString("lineID")
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.upCylinder_Item"
			.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,18,orderID)
			.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,10,lineID)
			.Parameters.Append oraComm.CreateParameter("@inItemV",200,&H0001,6,itemV)
			.Parameters.Append oraComm.CreateParameter("@outLineID",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outLineID = .parameters("@outLineID")
		response.Write(outLineID)	
	end with
end sub

sub upOrderLine_Values
	'custOwn = request.QueryString("custOwn")
	'swap = request.QueryString("swap")
	price = request.QueryString("price")
	quantity = request.QueryString("quantity")
	lineID = request.QueryString("lineID")
	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.upOrderLine_Values"
		'.Parameters.Append oraComm.CreateParameter("@inCustOwn",200,&H0001,1,custOwn)
		'.Parameters.Append oraComm.CreateParameter("@inSwap",200,&H0001,1,swap)
		.Parameters.Append oraComm.CreateParameter("@inPrice",200,&H0001,10,price)
		.Parameters.Append oraComm.CreateParameter("@inQuantity",200,&H0001,10,quantity)
		.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,10,lineID)
	end with
	set rs = oraComm.execute
	response.Write("545")
end sub

sub getLines_OrderID
	orderID = request.QueryString("orderID")
	set oraComm = Server.CreateObject("ADODB.Command")	
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getOLCount_orderID"
			.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,100,orderID)
		end with
		set rs = oraComm.execute
		do while not rs.eof
			lCount = rs(0)
			response.Write(lCount)
			rs.movenext
		loop
end sub

sub getCylinder_LineID
	lineID = request.QueryString("lineID")
	menuID = request.QueryString("menuID")
	process = request.QueryString("process")
	siteUser = request.QueryString("siteUser")
	
	set oraComm = Server.CreateObject("ADODB.Command")	
	with oraComm	
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getLineCount_lineID"
			.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,255,lineID)			
			'.Parameters.Append oraComm.CreateParameter("@inUserID",200,&H0001,255,siteUser)
			.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@outCntOrder",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@outErrorMessage",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outCount")
		outCntOrder = .parameters("@outCntOrder")
		errorMessage = .parameters("@outErrorMessage")
	end with	

	if(outCntOrder = 0) then
		response.Write("0|a|" & errorMessage & "|b|")
	else	
		do while not rs.eof
			orderID = rs(0)
			itemN = rs(2)
			itemD = rs(3)
			shipTo = rs(4)
			shipAdd1 = rs(5)
			shipAdd2 = rs(6)
			shipAdd3 = rs(7)
			qtyOrd = rs(8)
			qtyShip = rs(9)
			qtyStage = rs(10)
			orderType = rs(11)
			returnLoc = rs(12)
			LocName = rs(13)
			LocAddr1 = rs(14)
			LocAddr2 = rs(15)
			LocAddr3 = rs(16)
			
			response.Write("<span id='sp_OrderID'>OrderID : " & orderID & "</span>")
			response.Write("<input type='hidden' id='hdnOrderID' value='"&orderID&"'>")
			response.Write("<input type='hidden' id='hdnLineID' value='"&lineID&"'>")
			response.Write("<br>Item : "&itemN&"<br>Item Description : "&itemD&"<br>")
			if ( orderType = "WH" ) then
				response.Write("Ship To : "&returnLoc + " " + LocName&"<br>Ship Address : ")
				if ( LocAddr1 <> "" ) then
					response.write (LocAddr1&"<br>")
				end if 
				if ( LocAddr2 <> "" ) then 
					response.write( LocAddr2&"<br>")
				end if 
				if ( LocAddr3 <> "" ) then
					response.write( LocAddr3&"<br>")
				end if 
			else
				response.Write("Ship To : "&shipTo&"<br>Ship Address : "&shipAdd1&"<br>"&shipAdd2&"<br>"&shipAdd3&"<br>")
			end if 
			response.Write("Quantity Ordered : <span id='span_qtyOrd'>"&qtyOrd&"</span><br>")
			response.Write("Quantity Staged : <span id='sp_QtyStage'>" & qtyStage & "</span><br>")
			response.Write("Quantity Shipped : <span id='sp_QtyShip'>" & qtyShip & "</span>")
			response.Write("<br>")
			'response.Write("Swap Allowed : " & swap & "<br>")
		rs.movenext
		loop
	
	
	
		if(outCount<>0)then 'Cylinder Order or Line
			if(menuID = 7) then 'staging cylinder
				response.Write("Enter Cylinder's Barcode<br>")
				response.Write("<input type=text id='txtCylBarcode' name='cylBarcode' onChange='chk4ValidCyl(this)' style='width:250' onKeyUp='dv_chkBarCode(this)'>")
				response.Write("<br>")
				jscript = "onChange='removeCylBarcode(this)'"
			else 'shipping cylinder
				shipButton = "<input type=button id='btnShipCyl' name='shipCyl' value='Ship Cylinders' onClick='shipCylinders(&quot;"&lineID&"&quot;)' disabled>"
			end if
			set oraComm = Server.CreateObject("ADODB.Command")	
				with oraComm
					with oraComm
						.activeconnection = conn
						.CommandType = 4 'adCmdStoredProc            
						.CommandText = "dbo.getCylinderStage_LineID"
						.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,255,lineID)
						.Parameters.Append oraComm.CreateParameter("@outSCount",200,&H0002,500,"")
						.Parameters.Append oraComm.CreateParameter("@outShipped",200,&H0002,500,"")
					end with
					set rs = oraComm.execute
					outSCount = .parameters("@outSCount")	
					outShipped = .parameters("@outShipped")	
				end with
				if(outSCount < 4)then
					outSCount = 5
				end if
			response.Write("<span id='span_SelCylBarcode'>")	
				response.Write("<select id='selCylBarcode' name='cylBarcode' style='width:170' size="&outSCount&" "&jscript&">")
					do while not rs.eof
						ctkey = rs(0)
						barcode = rs(1)
						serial = rs(2)
						custId = rs(3)
						lastDate = rs(4)
						lastWho = rs(5)
						response.Write("<option value = '"&ctKey&"'>"&barcode&"</option>")
					rs.movenext
					loop		
				response.Write("</select>")
			response.Write("</span>")
			
	'		response.Write("<br>")
	'		response.Write("Pro Number :<br>")
	'		response.Write("<input type='text' id='txtPro' name='proNumber' onChange='chkProNumber(this)' >")
	'		response.Write("<br>")
	'		response.Write("Carrier : <br>")
	'		set oraComm = Server.CreateObject("ADODB.Command")
	'		with oraComm
	'			.activeconnection = conn
	'			.CommandType = 4 'adCmdStoredProc            
	'			.CommandText = "dbo.shipVMList"
	'			.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,500,orderID)
	'		end with
	'		set rs = oraComm.execute			
	'		response.Write("<select id='sel_ShipViaMisc' name='shipViaMisc'>")
	'		do while not rs.eof
	'			shipViaID = rs(0)
	'			shipViaM = rs(1)
	'			desc = rs(2)
	'			response.Write("<option value='"&shipViaID&"'>"&shipViaM & " - " & desc&"</option>")
	'		rs.movenext
	'		loop
	'		response.Write("</select>")		
			
			response.Write("<p id='p_CylinderShipped'>Cylinders Shipped : ")
				response.Write("<span id='span_CylShipped'>"&qtyShip&"</span>")
			response.Write("</p>")
			response.Write("<p id='p_cylToBeStaged'>Total Staged : ")
				response.Write("<span id='span_CylStaged'>"&qtyStage&"</span>")
			response.Write("</p>")
	'		response.Write("<input type='button' id='bt_reset' name='btnReset' value='Reset' onClick='shipReset()'>")
	'		response.Write("<input type='hidden' id='hdnCylOrder' name='cylOrder' value='0'>")
	'		response.Write(shipButton)' shipButton defined 'round 602
		else 'NON-Cylinder
			if(menuID = 7) then
				response.Write("Quantity Staged : ")
				response.Write("<input type='text' id='txtQty' name='quantity' onChange='chkStageQty(this)' >")
				response.Write("<input type='hidden' id='hdnMenuID' name='menuID' value=7>")
				btn = "disabled"			
	'		else
	'			response.Write("<input type='hidden' id='hdnMenuID' name='menuID' value=8>")
	'			response.Write("<br>")
	'			response.Write("Pro Number : ")
	'			response.Write("<input type='text' id='txtPro' name='proNumber' onChange='chkProNumber(this)' >")
	'			response.Write("<br>")
	'			response.Write("Carrier : &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;")
	'			set oraComm = Server.CreateObject("ADODB.Command")
	'			with oraComm
	'				.activeconnection = conn
	'				.CommandType = 4 'adCmdStoredProc            
	'				.CommandText = "dbo.shipVMList"
	'				.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,500,orderID)
	'			end with
	'			set rs = oraComm.execute			
	'			response.Write("<select id='sel_ShipViaMisc' name='shipViaMisc'>")
	'			do while not rs.eof
	'				shipViaID = rs(0)
	'				shipViaM = rs(1)
	'				desc = rs(2)
	'				response.Write("<option value='"&shipViaID&"'>"&shipViaM & " - " & desc&"</option>")
	'			rs.movenext
	'			loop
	'			response.Write("</select>")
	'			btn = "disabled"
			end if
	'		response.Write("<input type='hidden' id='hdnLineID' name='lineID' value="&lineID&">")
	'		response.Write("<p><input type='button' id='btnQuantity' name='btnQuantity' value='Click to Submit' onClick='submitQuantity(this)' "&btn&"></p>")		
		end if
		
		if(menuID = 8)then
			response.Write("<input type='hidden' id='hdnMenuID' name='menuID' value=8>")
	'		response.Write("<br>")
			response.Write("Pro Number : ")
			response.Write("<input type='text' id='txtPro' name='proNumber' onChange='chkProNumber(this)' >")
			response.Write("<br>")
			response.Write("Carrier : &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;")
			set oraComm = Server.CreateObject("ADODB.Command")
			with oraComm
				.activeconnection = conn
				.CommandType = 4 'adCmdStoredProc            
				.CommandText = "dbo.shipVMList"
				.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,500,orderID)
			end with
			set rs = oraComm.execute			
			response.Write("<select id='sel_ShipViaMisc' name='shipViaMisc'>")
			do while not rs.eof
				shipViaID = rs(0)
				shipViaM = rs(1)
				desc = rs(2)
				response.Write("<option value='"&shipViaID&"'>" & desc & "</option>")
			rs.movenext
			loop
			response.Write("</select>")
			btn = "disabled"	
		end if
		
		if(outCount<>0)then 'Cylinder Order or Line
			response.Write("<p>")	
			if menuID = 7 then
				response.Write("<input type='button' id='bt_reset' name='btnReset' value='Un-Stage All' onClick='removeAllStagedCylBarcodes()'>")
			end if 
			response.Write("<input type='hidden' id='hdnCylOrder' name='cylOrder' value='1'>")
			'response.Write(shipButton)	
			response.Write("<input type=button id='btnShipCyl' name='shipCyl' value='Ship Cylinders' onClick='shipCylinders(&quot;"&lineID&"&quot;)' disabled>")
			response.Write("</p>")
		else
			response.Write("<input type='hidden' id='hdnLineID' name='lineID' value="&lineID&">")
			response.Write("<input type='hidden' id='hdnCylOrder' name='cylOrder' value='0'>")	
			response.Write("<p><input type='button' id='btnQuantity' name='btnQuantity' value='Click to Submit' onClick='submitQuantity(this)' "&btn&"></p>")	
			response.Write("<br>")
		end if
	end if	
end sub

sub chk4ValidCyl
	barcode = request.QueryString("barcode")
	lineID = request.QueryString("lineID")
	siteUser = session("username")
	
	set oraComm = Server.CreateObject("ADODB.Command")
	
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.addCylinderStage_LineID"
			.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,10,lineID)
			.Parameters.Append oraComm.CreateParameter("@inBC",200,&H0001,6,barcode)
			.Parameters.Append oraComm.CreateParameter("@inScanWho",200,&H0001,255,siteUser)
			.Parameters.Append oraComm.CreateParameter("@outAllOk",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@outCtKey",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@outMessage",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outAllOk = .parameters("@outAllOk")
		outCtKey = .parameters("@outCtKey")		
		outMessage = .parameters("@outMessage")		
	end with

	' outAllOk is 1 if everything's ok, is 0 if there's an error.	
	if outAllOk = "1" then
		response.Write( outAllOk & "|a|" & outCtKey & "|b|" )
	else
		response.Write( outAllOk & "|a|" & outMessage & "|b|" )
	end if 
	
end sub

sub removeCylBarcode_makeAvail
	ctkey = request.QueryString("ctkey")
	lineID = request.QueryString("lineID")
'response.Write("794 " & ctkey)	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.makeCylStage_Avail"
			.Parameters.Append oraComm.CreateParameter("@inCTKey",200,&H0001,6,ctkey)
			.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,6,lineID)
			.Parameters.Append oraComm.CreateParameter("@outQtyStage",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outQtyStage = .parameters("@outQtyStage")
	end with
	response.Write(outQtyStage)
end sub

sub remove_ALL_CylBarcodes_makeAvail
	ctkeys = request.QueryString("ctkeys")
	lineID = request.QueryString("lineID")
'response.Write("794 " & ctkey)	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.makeCylStage_ALL_Avail"
			.Parameters.Append oraComm.CreateParameter("@inCTKeys",200,&H0001,8000,ctkeys)
			.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,6,lineID)
			.Parameters.Append oraComm.CreateParameter("@outQtyStage",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outQtyStage = .parameters("@outQtyStage")
	end with
	response.Write(outQtyStage)
end sub

sub shipCylinders
	lineID = request.QueryString("lineID")
'   Author = LH, 10/04/07	
'	siteUser = request.serverVariables("Remote_User")
	siteUser = session("username")
	carrier = request.QueryString("carrier")
	pronumber = request.QueryString("pronumber")

	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm	
		with oraComm
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.shipCylinders"
		.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,10,lineID)
		.Parameters.Append oraComm.CreateParameter("@inSiteUser",200,&H0001,255,siteUser)
		.Parameters.Append oraComm.CreateParameter("@inPronumber",200,&H0001,255,pronumber)
		.Parameters.Append oraComm.CreateParameter("@inCarrier",200,&H0001,255,carrier)
		.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outCount")
	end with		
	response.Write(outCount &"|0b|")
	response.Write("Cylinders Shipped|1b|")	
end sub

sub orderSearch
orderformType = request.QueryString("formType")

response.Write("<table id='tbl_update' width='100%' border=0 style='visibility:collapse;'>")
	response.Write("<tr>")
		response.Write("<td>Branch</td>")
		response.Write("<td id='td_getbranch' class='input'>")
			response.Write("<input type=text id='txt_getbranchID' name=getbranchID style='visibility:hidden;' onKeyUp='dv_chkNumber(this)'>")
		response.Write("</td>")
		response.Write("<td>Customer Number</td>")
		response.Write("<td><input type=text id='txt_getcustNum' name='getcustNum' style='visibility:hidden;' onKeyUp='dv_chkAlphaNumber(this)'></td>")
		'response.Write("<td colspan=5>&nbsp;</td>")
		response.Write("<td>Order Number</td>")
		response.Write("<td><input type=text id='txt_getOrderNum' name='getOrderNum' style='visibility:hidden;' onKeyUp='dv_chkNumber(this)'></td>")
		'response.Write("<td colspan=5>&nbsp;</td>")
	response.Write("</tr>")
	response.Write("<tr>")
		response.Write("<td>P.O. Number</td>")
		response.Write("<td><input type=text id='txt_getPONum' name='getPONum' style='visibility:hidden;' onKeyUp='dv_chkNumber(this)'></td>")
		'response.Write("<td colspan=5>&nbsp;</td>")
		response.Write("<td>Bill To Name</td>")
		response.Write("<td><input type=text id='txt_getBillName' name='getBillName' style='visibility:hidden;' onKeyUp='dv_chkApostrophe(this)'></td>")
		'response.Write("<td colspan=5>&nbsp;</td>")
		response.Write("<td>Ship To Name</td>")
		response.Write("<td><input type=text id='txt_getShipName' name='getShipName' style='visibility:hidden;' onKeyUp='dv_chkApostrophe(this)'></td>")
		'response.Write("<td colspan=5>&nbsp;</td>")
	response.Write("</tr>")
	
	response.Write("<tr>")
		response.Write("<td>Ship To Address</td>")
		response.Write("<td><input type=text id='txt_getShipToAddr' name='getShipToAddr' style='visibility:hidden;' onKeyUp='dv_chkApostrophe(this)'></td>")

		response.Write("<td>Status</td>")
		response.Write("<td>")
			set oraComm = Server.CreateObject("ADODB.Command")
			with oraComm
				.activeconnection = conn
				.CommandType = 4 'adCmdStoredProc            
				.CommandText = "dbo.getOrderStatus_fromOrders"		
			end with
			set rs = oraComm.execute			
			response.Write("<select id='txt_getOrderStatus' name='getOrderStatus' style='visibility:hidden';>")
				if(orderformtype = "Inquiry") then	
					response.Write("<option>ALL</option>")			
					do while not rs.eof
						orderStatus = rs(0)
							response.Write("<option value = '"&orderStatus&"'>"&orderStatus&"</option>")
						rs.movenext
					loop
				else
					response.Write("<option value = 'Open'>Open</option>")
				end if	
			response.Write("</select>")
		response.Write("</td>")
		
		response.Write("<td>Salesman</td>")
		response.Write("<td>")
			set oraComm = Server.CreateObject("ADODB.Command")
			with oraComm
				.activeconnection = conn
				.CommandType = 4 'adCmdStoredProc            
				.CommandText = "dbo.getSalesman_CustID"		
				.Parameters.Append oraComm.CreateParameter("@inCustID",200,&H0001,25,"NULL")
			end with
			set rs = oraComm.execute			
			response.Write("<select id='sel_getSalesMan' name='getSalesman' style='visibility:hidden';>")
				response.Write("<option>ALL</option>")
				do while not rs.eof
					salesmanID = rs(0)
					salesmanName = rs(1)
					response.Write("<option value = '"&salesmanID&"'>"&salesmanName&"</option>")
					rs.movenext
				loop
			response.Write("</select>")
		response.Write("</td>")
		
	response.Write("</tr>")		
	
	with response
		.write("<tr>")
			.Write("<td>CSR</td>")
			.Write("<td>")
			set oraComm = Server.CreateObject("ADODB.Command")
			with oraComm
				.activeconnection = conn
				.CommandType = 4 'adCmdStoredProc            
				.CommandText = "dbo.getCSRList"		
			end with
			set rs = oraComm.execute			
				.Write("<select id='sel_getCSRList' name='getCSRList' style='visibility:hidden';>")
					.Write("<option>ALL</option>")
					do while not rs.eof
						csr = rs(0)
							.Write("<option value = '"&csr&"'>"&csr&"</option>")
						rs.movenext
					loop
				.Write("</select>")
			.Write("</td>")
		.write("</tr>")
	end with
			
	with response
		.Write("<tr>")
		.Write("<td align='center' colspan='6'>")
			.write("<input type='button' id='btnCloseSrch' name='closeSrch' value='End Search' onClick='closeSearch(this)'>")
			.write ("<input type=button id='btn_Search' name='search' value='Search for Order' style='visibility:hidden;' onClick='getOrderResults(this)'>")
		.write("</td>")
		.Write("</tr>")	
		.write("<tr>")
			.Write("<td id='td_searchResults' valign=top colspan='6'>")
				.Write("<select id='sel_searchResults' name='searchResults' multiple='multiple' style='visibility:hidden';>")
				.Write("</select>")
			.Write("</td>")
		.write("</tr>")
	end with		
response.Write("</table>")
end sub

sub customerSearch
'Customer Search table cell	
with response
	.write("<table width='100%'>")
		.write("<tr>")
			.write("<td colspan='4'>Customer Search</td>")
		.write("</tr>")
		.write("<tr>")
			.write("<td>Customer Number : </td>")
			.write("<td>")
				.write("<input type='text' id='txtCustSearch' name='custSearch' size='40' onKeyUp='dv_chkAlphaNumber(this)'>")
			.write("</td>")
			.write("<td>Bill To : </td>")
			.write("<td>")
				.write("<input type='text' id='txtBillToSearch' name='billToSearch' size='40' onKeyUp='dv_chkApostrophe(this)'>")
			.write("</td>")
		.write("</tr>")
		.write("<tr>")
			.write("<td>Ship To : </td>")
			.write("<td>")
				.write("<input type='text' id='txtShipToSearch' name='shipToSearch' size='40' onKeyUp='dv_chkApostrophe(this)'>")
			.write("</td>")
			.write("<td>Ship Address : </td>")
			.write("<td>")
				.write("<input type='text' id='txtShipAddrSearch' name='shipAddrSearch' size='40' onKeyUp='dv_chkApostrophe(this)'>")
			.write("</td>")
		.write("</tr>")
		.write("<tr>")
			.write("<td colspan='4' align='center'>")
				.write("<input type='button' id='btnCloseSrch' name='closeSrch' value='End Search' onClick='closeSearch(this)'>")
				.write("<input type='button' id='btnCustSrch' name='custSrch' value='Search' onClick='lookUpCustomer(this)'>")
			.write("</td>")
		.write("</tr>")
		.write("<tr>")
			.write("<td colspan='4' id='td_selCust'>")
				.Write("&nbsp;")
			.write("</td>")
		.write("</tr>")
	.write("</table>")
end with
'End customer search
end sub

sub getCylinder_OrderID
	orderID = request.QueryString("orderID") 
	menuID = request.QueryString("menuID")
	cylProcess = request.QueryString("cylProcess")
	siteUser = request.QueryString("siteUser")

	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getLineID_orderID"
			.Parameters.Append oraComm.CreateParameter("@inOrderID",200,&H0001,50,orderID)
			.Parameters.Append oraComm.CreateParameter("@inCylProcess",200,&H0001,50,cylProcess)
			.Parameters.Append oraComm.CreateParameter("@outCount",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@outMessage",200,&H0002,255,"")
			.Parameters.Append oraComm.CreateParameter("@inMenuID",200,&H0001,50,menuID)
			'.Parameters.Append oraComm.CreateParameter("@inUserID",200,&H0001,50,siteUser)			
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outCount")
		outMessage = .parameters("@outMessage")
	end with	
if(outCount = 0)then
	response.Write("0|a|" & outMessage & "|b|")
else	
	with response
		.Write("1|a|")
		.Write("<select id='txtLineID' name='lineID' onChange='getCylinder_LineID(this)'>")
		.write("<option value='0'>Select Line</option>")
		do while not rs.eof
			lineID = rs(0)
			supItem = rs(1)
			shipTo = rs(2)
			shipAdd1 = rs(3)
			shipAdd2 = rs(4)
			shipAdd3 = rs(5)
			custNo = rs(6)
			bt_name = rs(7)
			orderID = rs(8)
			.Write("<option value = '"&lineID&"'>"&lineID &" - "&supItem&"</option>")
		rs.movenext
		loop
		.Write("</select>")
		.write("<input type='hidden' id='hidMenuID' name='menuID' value='"&menuID&"'>")
		.Write("|b|")
	end with
end if		
end sub

sub submitQuantity
	quantity = request.QueryString("quantity")
	lineID = request.QueryString("lineID")
	menuID = request.QueryString("menuID")
	pronumber = request.QueryString("pronumber")
	carrier = request.QueryString("carrier")

	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm	
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.submitQuantity"
			.Parameters.Append oraComm.CreateParameter("@inLineID",200,&H0001,50,lineID)	
			.Parameters.Append oraComm.CreateParameter("@inQuantity",200,&H0001,50,quantity)
			.Parameters.Append oraComm.CreateParameter("@inMenuID",200,&H0001,50,menuID)
			.Parameters.Append oraComm.CreateParameter("@inPronumber",200,&H0001,50,pronumber)
			.Parameters.Append oraComm.CreateParameter("@inCarrier",200,&H0001,50,carrier)
			.Parameters.Append oraComm.CreateParameter("@outLineCnt",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outCount = .parameters("@outLineCnt")
	end with	
	response.Write(outCount &"|0b|")
	response.Write("Quantity updated|1b|")
	
end sub

sub selOrderType
	orderType = request.QueryString("orderType")
	custNum = request.QueryString("custNum")
	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.selOrderType"
			.Parameters.Append oraComm.CreateParameter("@inOrder",200,&H0001,50,orderType)	
			.Parameters.Append oraComm.CreateParameter("@inCustNum",200,&H0001,50,custNum)
			.Parameters.Append oraComm.CreateParameter("@outCustMatch",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outCustMatch = .parameters("@outCustMatch")
	end with
	
	if(outCustMatch = 0)then
		response.Write(0)
	else
		do while not rs.eof
			oType = rs(0)
			response.Write(oType)
		rs.movenext
		loop
	end if
	
end sub

sub chgCustOwn
	siteuser = request.QueryString("siteuser")
	set oraCommB = Server.CreateObject("ADODB.Command")
	with oraCommB
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.getWarehoseCust"
		.Parameters.Append oraCommB.CreateParameter("@inUserID",200,&H0001,255,siteuser)
	end with
	set rsB = oraCommB.execute
	do while not rsB.eof		
		warehouseID = rsB(0)	
		btName = rsB(1)
		btAdd = rsB(2) & "<br>" & rsB(3) & "<br>" & rsB(4) & ", " & rsB(5) & " " & rsB(6)	
		rsB.movenext
	loop	
	
	custInput = warehouseID 
	custNFO = btName & " " & btAdd
	custNFO = custNFO & " " & "<input type='hidden' id='txtCustSearch' name='custID' value = '"& warehouseID &"'>"	
	response.Write(custInput&"|0b|")
	response.Write(custNFO&"|1b|")
			
end sub

sub	selectOrderType
	set oraCommB = Server.CreateObject("ADODB.Command")
	with oraCommB
		.activeconnection = conn
		.CommandType = 4 'adCmdStoredProc            
		.CommandText = "dbo.getOrderType_orderID"
		.Parameters.Append oraCommB.CreateParameter("@inOrderID",200,&H0001,255,orderID)
	end with
	set rs = oraCommB.execute			
	response.Write("<select id='sel_OrderType' name='orderType' class='reqSel' onChange='selOrderType(this, "" "&orderFormType&" "")'>")
	do while not rs.eof
		orderType = rs(0)
		desc = rs(1)
		response.Write("<option value='"&orderType&"'>"&desc&"</option>")
	rs.movenext
	loop
	response.Write("</select>")
end sub

sub getOrderType_orderID_chkQC
	orderID = request.QueryString("orderID")
	
	set oraComm = Server.CreateObject("ADODB.Command")
	with oraComm
		with oraComm
			.activeconnection = conn
			.CommandType = 4 'adCmdStoredProc            
			.CommandText = "dbo.getOrderType_orderID_chkQC"
			.Parameters.Append oraComm.CreateParameter("@inOrder",200,&H0001,50,orderID)	
			.Parameters.Append oraComm.CreateParameter("@outQCCount",200,&H0002,255,"")
		end with
		set rs = oraComm.execute
		outQCCount = .parameters("@outQCCount")
	end with
	
	if(outQCCount = 0)then
		response.Write(0)
	else
		response.Write(1)
	end if
	
end sub

sub getTempQuoteOrderID

Dim tempid
Dim str
Dim qtUser

        qtUser = Session("username")
         
        set oraComm = Server.CreateObject("ADODB.Command")

        with oraComm
	        .activeconnection = conn
	        .CommandType = 4 'adCmdStoredProc            
	        .CommandText = "rrs.dbo.getOrderQuoteTempID"        
            '.Parameters.Append oraComm.CreateParameter("@qtUser",200,&H0001,30,qtUser) 
                                                                                                        
	        set rs = .execute

        end with
        
        str = "|a|" & rs("tempid")                  
        str = str & "|b|"

        response.write str        
         	
end sub


qry = request.QueryString("qry")
'   Author = LH, 10/04/07	
'	siteUser = request.serverVariables("Remote_User")
siteUser = session("username")

select case qry 
	case 1
		arg1 = request.QueryString("arg1")
		searchBy = request.QueryString("searchBy")
		call SearchCylwithBarcodeGetSerialNumber(barcode) 'from subs.asp
	case 2
		arg1 = request.QueryString("arg1")
		searchBy = request.QueryString("searchBy")
		call CylMaintenance_SearchBarcode(barcode) 'from subs.asp
	case 3
		arg1 = request.QueryString("arg1")
		searchBy = request.QueryString("searchBy")
		call BCorSerialExists 'from subs.asp
	case 4
		fType = request.QueryString("fType")
		lineID = request.QueryString("lineID")
		orderID = request.QueryString("orderID")
		call addItemRow(fType, lineID, orderID)		
	case 5
		call getCustInfo_OrderEntry
	case 6
		call getReturnLoc_RetLoc
	case 7
		call getCylItemInfo
	case 8
		call getBranchInfo_branchID	
	case 9
		call addEmail
	case 10
		call getItems	
	case 11
		call getOrderResults
	case 12
		call getOrderInfo_ID
	case 13
		call getLineID	
	case 14
		call deleteOrderLine
	case 15 
		call upCylinder_Item	
	case 16
		call upOrderLine_Values	
	case 18
		call getCylinder_LineID	
	case 19
		call chk4ValidCyl
	case 20
		call removeCylBarcode_makeAvail
	case 21
		call shipCylinders	
	case 22
		call orderSearch	
	case 23
		call customerSearch	
	case 24
		call getCylinder_OrderID	
	case 25
		call submitQuantity	
	case 26
		call selOrderType	
	case 27
		call chgCustOwn	
	case 28
		orderID = request.QueryString("orderID")
		call selectOrderType
	case 29
		call remove_ALL_CylBarcodes_makeAvail
		
	case 30
	    call getTempQuoteOrderID
    case 31
        call getOrderType_orderID_chkQC
		
end select
'sample code
'Set objWord = CreateObject("Word.Application")
'	Set objDoc = objWord.Documents.Open("c:\cais\rrs\RecLabel.doc",0,1)
'	Set myMMo = objWord.ActiveDocument.MailMerge
'	Call myMMo.opendatasource( "c:\cais\rrs\RecDetail\RecOrderID.doc", 4, DontConfirmConversions,OpenReadOnly)
'
'	objWord.ActiveDocument.MailMerge.Destination = 0
'	objWord.ActiveDocument.MailMerge.Execute
'	objWord.ActiveDocument.SaveAs("RecDetail.doc")
'
'	lprintername = printer
'	lsaveprinter = objWord.ActivePrinter
'	objWord.ActivePrinter = lprintername
'	'Set objDoc = objWord.Documents.Open(resultpage,0,1)
'	objWord.ActiveDocument.Printout(FALSE)
'	objWord.ActivePrinter = lsaveprinter
'
'	call objWord.ActiveDocument.Close(0,0,FALSE)
'	call objWord.Quit(0,0,False)
'	SET objword  = nothing
%>

