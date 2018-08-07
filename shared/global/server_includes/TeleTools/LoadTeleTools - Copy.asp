<script type="text/vbscript" src="/shared/global/server_includes/TeleTools/Constants.txt" language="VBScript">msgbox(1)</script>

<!--
<object
    id="etLPK"
    style="display:none;"
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
    <param name="LPKPath" value="/shared/global/server_includes/TeleTools/TeleTools40.lpk">
</object>
-->
<object
    id="etDeploy"
    style="display:none;"
    classid="CLSID:C9CAE3E4-23CC-407A-A834-DC88D81B629A"
	codebase="/shared/global/server_includes/TeleTools/etTT40.cab#version=4.0.9.7">
</object>
<object
    id="etDevice"
    style="display:none;"
    classid="CLSID:37BB095B-0D65-499C-8C01-B03183A861B6">
</object>
<object
    id="etCall"
    style="display:none;"
    classid="CLSID:49675FC6-BD7B-4D57-9CBE-4BFD1487F4EC">
</object>
<object
    id="etLine"
    style="display:none;"
    classid="CLSID:FC61F000-B4E1-4D23-9902-6E124E215333">
</object>
<!--
-->

<script type="text/javascript" language="JavaScript">
<!--

	var myBody, myObject, myParam;
	myBody = document.getElementsByTagName('body')[0];
	myObject = document.createElement('form');
	myObject.setAttribute('id', 'PhoneLogForm');
	myObject.setAttribute('method', 'post');
	myObject.setAttribute('action', '/shared/global/server_includes/TeleTools/DisplayLog.asp');
	myObject.setAttribute('target', 'PhoneLogWindow');
	//myObject.style.margin = 'none';
	myParam = document.createElement('input');
	myParam.setAttribute('type', 'hidden');
	myParam.setAttribute('id', 'Log');
	myParam.setAttribute('name', 'Log');
	myObject.appendChild(myParam);
	myBody.appendChild(myObject);

/*
	myObject = document.createElement('object');
	myObject.setAttribute('id', 'etLPK');
	myObject.setAttribute('classid', 'clsid:5220cb21-c88d-11cf-b347-00aa00a28331');
	myObject.style.display = 'none';
	myParam = document.createElement('param');
	myParam.setAttribute('name', 'LPKPath');
	myParam.setAttribute('value', '/shared/global/server_includes/TeleTools/TeleTools40.lpk');
	myObject.appendChild(myParam);
	myBody.appendChild(myObject);

	myObject = document.createElement('object');
	myObject.setAttribute('id', 'etDeploy');
	myObject.setAttribute('classid', 'clsid:C9CAE3E4-23CC-407A-A834-DC88D81B629A');
	myObject.setAttribute('codebase', '/shared/global/server_includes/TeleTools/TeleTools40.cab');
	myObject.style.display = 'none';
	myBody.appendChild(myObject);

	myObject = document.createElement('object');
	myObject.setAttribute('id', 'etDevice');
	myObject.setAttribute('classid', 'clsid:37BB095B-0D65-499C-8C01-B03183A861B6');
	myObject.style.display = 'none';
	myBody.appendChild(myObject);

	myObject = document.createElement('object');
	myObject.setAttribute('id', 'etCall');
	myObject.setAttribute('classid', 'clsid:49675FC6-BD7B-4D57-9CBE-4BFD1487F4EC');
	myObject.style.display = 'none';
	myBody.appendChild(myObject);

	myObject = document.createElement('object');
	myObject.setAttribute('id', 'etLine');
	myObject.setAttribute('classid', 'clsid:FC61F000-B4E1-4D23-9902-6E124E215333');
	myObject.style.display = 'none';
	myBody.appendChild(myObject);
*/

//-->
</script>

<input type="hidden" name="TeleToolsActive" id="TeleToolsActive" value="False">

<script type="text/vbscript" language="VBScript">
<!--

    'Global Variables
    Dim InCall : InCall = False
    Dim CallNumber : CallNumber = 0
    Dim PhoneLog, PhoneLogWindow
    Dim PhoneLink, PhoneLinkDict
    Set PhoneLinkDict = CreateObject("Scripting.Dictionary")
    PhoneLinkDict.Add "backgroundColor", ""
    PhoneLinkDict.Add "fontWeight", ""
    PhoneLinkDict.Add "innerHTML", ""

    Sub Window_OnLoad()
    
        On Error Resume Next
        etDeploy.etEnabled = True
        etDeploy.etSerialNumber = "4z1i-214p-z914-62k0"
        If Err.number = 0 Then
            On Error GoTo 0
            LoadTeleTools()
        Else
            On Error GoTo 0
            'WriteStatus("Error: TeleTools is not installed")
        End If
    
    End Sub
    
    Sub LoadTeleTools()

        etLine.Enabled = True
        
        Dim i
        For i=0 To etLine.DeviceCount - 1
            If Left(etLine.DeviceList(i), 8) = "Hi-Phone" Then
                etLine.DeviceID = i
                etLine.PrivilegeMonitor = True
                etLine.PrivilegeOwner = True
                etLine.DeviceActive = True
                TeleToolsActive.Value = "True"
                WriteStatus("Ready")
            End If
        Next

    End Sub
    
    Sub Window_OnUnload()

        etLine.DeviceActive = False

    End Sub
    
    Function FormatPhone(strTxt, strFormat)

	    strTxt = StripPhone(strTxt)
	    If strTxt = "" Or strTxt = "0" Then FormatPhone = "" : Exit Function
	    
	    Dim strExt
	    Select Case strFormat
		    Case "(PPP) PPP-PPPP"
			    Select Case Len(strTxt)
				    Case 1,2,3,4:
					    FormatPhone = strTxt
				    Case 5,6,7:
					    FormatPhone = Left(strTxt, Len(strTxt)-4) & "-" & Right(strTxt, 4)
				    Case 8,9,10:
					    FormatPhone = "(" & Left(strTxt, Len(strTxt)-7) & ") " & Mid(strTxt, Len(strTxt)-6, 3) & "-" & Right(strTxt, 4)
				    Case Else
					    strExt = Mid(strTxt,11)
					    strTxt = Left(strTxt,10)
					    FormatPhone = "(" & Left(strTxt, Len(strTxt)-7) & ") " & Mid(strTxt, Len(strTxt)-6, 3) & "-" & Right(strTxt, 4) & " x" & strExt
			    End Select
		    Case "PPP-PPP-PPPP"
			    Select Case Len(strTxt)
				    Case 1,2,3,4:
					    FormatPhone = strTxt
				    Case 5,6,7:
					    FormatPhone = Left(strTxt, Len(strTxt)-4) & "-" & Right(strTxt, 4)
				    Case 8,9,10:
					    FormatPhone = Left(strTxt, Len(strTxt)-7) & "-" & Mid(strTxt, Len(strTxt)-6,3) & "-" & Right(strTxt, 4)
				    Case Else
					    strExt = Mid(strTxt, 11)
					    strTxt = Left(strTxt, 10)
					    FormatPhone = Left(strTxt, Len(strTxt)-7) & "-" & Mid(strTxt,Len(strTxt)-6, 3) & "-" & Right(strTxt, 4) & " x" & strExt
			    End Select
	    End Select

    End Function
    
    Function StripPhone(strIn)

	    If IsNull(strIn) Then StripPhone = "" : Exit Function

	    Dim strTxt : strTxt = strIn
	    strTxt = Replace(strTxt, "-", "")
	    strTxt = Replace(strTxt, "(", "")
	    strTxt = Replace(strTxt, ")", "")
	    strTxt = Replace(strTxt, " ", "")
	    strTxt = Replace(strTxt, "/", "")
	    strTxt = Replace(strTxt, "\", "")
	    strTxt = Replace(strTxt, ".", "")
	    strTxt = Replace(strTxt, "x", "")
	    strTxt = Replace(strTxt, "X", "")
	    strTxt = Replace(strTxt, "ext", "")
	    strTxt = Replace(strTxt, "Ext", "")
	    If Left(strTxt, 1) = "1" Then strTxt = Mid(strTxt, 2)
	    StripPhone = strTxt

    End Function

    Sub Dial(PhoneNumber)

        If TeleToolsActive.Value = "True" Then
            etLine.CallPhoneNumber = "91" & Left(StripPhone(PhoneNumber), 10)
            If InCall Then
				Hangup()
			End If
            etLine.CallDial()
        Else
            WriteStatus("Auto Dial Disabled")
	        SetTimeout "ClearStatus()", 3000
        End If

    End Sub
    
    Sub WriteLog(Text)
        PhoneLog = PhoneLog & vbNewLine & Now() & " - " & Text
    End Sub
    
    Sub DisplayLog()
		PhoneLogForm.Log.value = "<ol>" & Replace(PhoneLog, vbNewLine, "<li>") & "</ol>"
		PhoneLogForm.submit()
        'PhoneLogWindow = Window.open("/shared/global/server_includes/TeleTools/DisplayLog.asp?Log=<ol><li>" & Replace(PhoneLog, vbNewLine, "<li>") & "</ol>", "PhoneLogWindow")
    End Sub

    Sub ClearLog()
        PhoneLog = ""
    End Sub

    Sub WriteStatus(Text)

        On Error Resume Next
        CallStatus.innerText = Text
        On Error GoTo 0

    End Sub

    Sub ClearStatus(ID)

        On Error Resume Next
        If CallNumber = ID And IsEmpty(PhoneLink) Then
	        CallStatus.innerText = "Ready"
	    End If
        On Error GoTo 0

    End Sub

    Sub Hangup()

        WriteLog("Hangup")
        WriteStatus("Hanging Up")
        If etLine.CallHandle <> 0 And etLine.CallState <> LINECALLSTATE_IDLE Then
            If Not etLine.CallHangup Then       
                WriteLog("Error = " & etLine.ErrorText & " The line is already hung up")
            End If
        Else
            WriteLog("There are no active calls or the Call State is already idle")  
        End If

    End Sub

    Sub etLine_OnOffering(CallHandle)
        WriteLog("etLine_OnOffering")
        WriteStatus("Incoming Call")
    End Sub

    Sub etLine_OnRing(Count, RingMode)
        WriteLog("etLine_OnRing [" & Count & ", " & RingMode & "]")
    End Sub

    Sub etLine_OnProceeding(CallHandle)
        WriteLog("etLine_OnProceeding")
        WriteStatus("Connecting")
    End Sub

    Sub etLine_OnRedirectingID(CallHandle)
        WriteLog("etLine_OnRedirectingID")
        WriteLog("RedirectingIDName = " & etLine.CallRedirectingIDName)
    End Sub

    Sub etLine_OnRedirectionID(CallHandle)
        etLine.CallHandle = CallHandle
        WriteLog("etLine_OnRedirectionID")
        WriteLog("RedirectionIDName = " & etLine.CallRedirectionIDName)
        WriteLog("RedirectionIDNumber = " & etLine.CallRedirectionIDNumber)
    End Sub

    Sub etLine_OnRingBack(CallHandle)
        WriteLog("etLine_OnRingBack")
        WriteStatus("Ringing Back")
    End Sub

    Sub etLine_OnSpecialInfo(CallHandle)
        etLine.CallHandle = CallHandle
        WriteLog("etLine_OnSpecialInfo")
        WriteLog("etLine.CallSpecialInfo = " & etLine.StringLINESPECIALINFO(etLine.CallSpecialInfo))
        WriteStatus("Special Info Message")
        SetTimeout "Hangup()", 10000
    End Sub

    Sub etLine_OnDialing(CallHandle)
        WriteLog("etLine_OnDialing")
        WriteStatus("Dialing " & FormatPhone(etLine.CallPhoneNumber, "(PPP) PPP-PPPP"))
    End Sub

    Sub etLine_OnDialtone(CallHandle)
        WriteLog("etLine_OnDialtone")
        WriteStatus("Dialtone")
    End Sub

    Sub etLine_OnConnected(CallHandle)
        'Most modems report the connected state immediately after dialing the phone number. You may not truly be connected!
        WriteLog("etLine_OnConnected")
        WriteStatus("Connected")
        If (InStr(1, etLine.TAPITSP, "Modem", 1) > 0) Then
            If (Not etLine.PrivilegeNone) Then
                'Using a voice modem
                etLine.CallMonitorDigitsActive = True
                WriteLog("etLine.CallMonitorDigitsActive = True")
            End If
        Else
            If (etLine.AddressCapabilitiesCallFeatures And LINECALLFEATURE_MONITORDIGITS) <> 0 Then
                etLine.CallMonitorDigitsActive = True
                WriteLog("etLine.CallMonitorDigitsActive = True")
            End If
        End If
        WriteLog("etLine.CallOrigin = " & etLine.StringLINECALLORIGIN(etLine.CallOrigin))
    End Sub

    Sub etLine_OnDigitReceived(CallHandle, Digit, Tag)
        WriteLog("etLine_OnDigitReceived [" & Chr(Digit) & "]")
    End Sub

    Sub etLine_OnCallerID(CallHandle)
        WriteLog("etLine_OnCallerID")
        WriteLog("CallerIDName = " & etLine.CallCallerIDName)
        WriteLog("CallerIDNumber = " & etLine.CallCallerIDNumber)
        If etLine.CallCallerIDBlocked Then
            WriteLog("Blocked")
            WriteStatus("Incoming call from blocked number")
        ElseIf etLine.CallCallerIDOutOfArea Then
            WriteLog("Out of area")
            WriteStatus("Incoming call from out of area number")
        ElseIf etLine.CallCallerIDName = "" Then
            'Do nothing - do not change the status from "incoming call" unless we know more info
        ElseIf StripPhone(etLine.CallCallerIDName) = StripPhone(etLine.CallCallerIDNumber) Then
            WriteStatus("Incoming call from " & FormatPhone(etLine.CallCallerIDNumber, "(PPP) PPP-PPPP"))
        Else
            WriteStatus("Incoming call from " & etLine.CallCallerIDName & " (" & FormatPhone(etLine.CallCallerIDNumber, "PPP-PPP-PPPP"))
        End If
        PhoneSearch = Window.open("/apps/phone_search/home.asp?SearchPhone=" & StripPhone(etLine.CallCallerIDNumber) & "&chkType=1&chkType=2&chkType=3&chkType=4&chkType=5&chkType=6&chkType=0", "PhoneSearch")
    End Sub

    Sub etLine_OnDisconnected(CallHandle)
        WriteLog("etLine_OnDisconnected")
        WriteLog("etLine.CallDisconnectMode = " & etLine.StringLINEDISCONNECTMODE(etLine.CallDisconnectMode))
        WriteStatus("Disconnected")
        Hangup()
    End Sub

    Sub etLine_OnError()
        ' You can process errors in your other routines or use the etLine "OnError" event handler here
		InCall = False
        WriteLog("etLine_OnError")
        WriteLog("etLine.ErrorText = " & etLine.ErrorText)
        Select Case etLine.ErrorText
			Case "LINEERR_UNITIALIZED"
				WriteStatus("Error: Phone line not found")
			Case Else
		        WriteStatus("Error: " & etLine.ErrorText)
		End Select
        If Not IsEmpty(PhoneLink) And PhoneLinkDict.Item("innerHTML") <> "" Then
			PhoneLink.style.backgroundColor = PhoneLinkDict.Item("backgroundColor")
			PhoneLinkDict.Item("backgroundColor") = ""
			PhoneLink.style.fontWeight = PhoneLinkDict.Item("fontWeight")
			PhoneLinkDict.Item("fontWeight") = ""
			PhoneLink.innerHTML = PhoneLinkDict.Item("innerHTML")
			PhoneLinkDict.Item("innerHTML") = ""
			PhoneLink = Empty
		End If
        SetTimeout "ClearStatus(CallNumber)", 10000
    End Sub

    Sub etLine_OnIdle(CallHandle)
        etLine.CallHandle = CallHandle
        WriteLog("etLine_OnIdle")
        WriteStatus("Idle")
    End Sub

    Sub etLine_OnBusy(CallHandle)
        WriteLog("etLine_OnBusy")
        WriteStatus("Busy")
        Hangup()
    End Sub

    Sub etLine_OnCallBegin(CallHandle)
		InCall = True
		CallNumber = CallNumber + 1
        WriteLog("etLine_OnCallBegin")
        WriteStatus("Call Begin")
        If Not IsEmpty(PhoneLink) Then
			PhoneLinkDict.Item("backgroundColor") = PhoneLink.style.backgroundColor
			PhoneLink.style.backgroundColor = "red"
			PhoneLinkDict.Item("fontWeight") = PhoneLink.style.fontWeight
			PhoneLink.style.fontWeight = "bold"
			PhoneLinkDict.Item("innerHTML") = PhoneLink.innerHTML
			PhoneLink.innerHTML = "&nbsp;" & PhoneLink.innerHTML & "&nbsp;"
		End If
    End Sub

    Sub etLine_OnCalledID(CallHandle)
        etLine.CallHandle = CallHandle  ' needs to be here for a device that handles multiple calls
        WriteLog("etLine_OnCalledID")
        WriteLog("CalledIDName = " & etLine.CallCalledIDName)
        WriteLog("CalledIDNumber = " & etLine.CallCalledIDNumber)
    End Sub

    Sub etLine_OnCallEnd(CallHandle)
		InCall = False
        WriteLog("etLine_OnCallEnd")
        WriteStatus("Call End")
        If Not IsEmpty(PhoneLink) Then
			PhoneLink.style.backgroundColor = PhoneLinkDict.Item("backgroundColor")
			PhoneLinkDict.Item("backgroundColor") = ""
			PhoneLink.style.fontWeight = PhoneLinkDict.Item("fontWeight")
			PhoneLinkDict.Item("fontWeight") = ""
			PhoneLink.innerHTML = PhoneLinkDict.Item("innerHTML")
			PhoneLinkDict.Item("innerHTML") = ""
			PhoneLink = Empty
		End If
        SetTimeout "ClearStatus(CallNumber)", 3000
    End Sub

'-->
</script>


