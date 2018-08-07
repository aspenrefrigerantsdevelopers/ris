<script type="text/vbscript" language="VBScript">
<!--

    Dim TeleToolsFrame

    Sub Window_OnLoad()

'		On Error Resume Next
'		If top.name = self.name Then
'            If opener.parent.frames(0).TeleToolsActive.Value = "True" Then LinkPhoneNumbers()
'		Else
'            If parent.frames(0).TeleToolsActive.Value = "True" Then LinkPhoneNumbers()
'		End If
'		On Error GoTo 0

		On Error Resume Next
        Set TeleToolsFrame = parent.frames(0)
        Set TeleToolsFrame = parent.parent.frames(0)
        Set TeleToolsFrame = parent.opener.parent.frames(0)
        Set TeleToolsFrame = opener.parent.frames(0)
		On Error GoTo 0
		
		If IsObject(TeleToolsFrame) Then
            Call LinkPhoneNumbers()
		End If

	End Sub

	Sub LinkPhoneNumbers()

		For Each Element In Document.All
		    If Element.ClassName = "PhoneLink" And Element.innerHTML <> "" Then
				'Element.style.color = "blue"
				'Element.style.textdecoration = "underline"
				Element.style.cursor = "hand"
				Element.innerHTML = Element.innerHTML & " <img src=""/shared/images/icons/phone.gif"" style=""vertical-align:middle; margin-bottom:-1;"">"
				Element.title = "Click here to dial number"
		        Call Element.AttachEvent("onclick", GetRef("CallDial"))
		    End If
		Next
	
	End Sub
    
    Sub CallDial()

'		If top.name = self.name Then
'			If Not opener.parent.frames(0).InCall Then
'				Set opener.parent.frames(0).PhoneLink = window.event.srcElement
'				Call opener.parent.frames(0).Dial(window.event.srcElement.innerHTML)
'			End If
'		Else
'			If Not parent.frames(0).InCall Then
'				If window.event.srcElement.tagName = "SPAN" Then
'					Set parent.frames(0).PhoneLink = window.event.srcElement
'					Call parent.frames(0).Dial(window.event.srcElement.innerHTML)
'				Else
'					Set parent.frames(0).PhoneLink = window.event.srcElement.parentElement
'					Call parent.frames(0).Dial(window.event.srcElement.parentElement.innerHTML)
'				End If
'			End If
'		End If
	
		If Not TeleToolsFrame.InCall Then
			If window.event.srcElement.tagName = "SPAN" Then
				Set TeleToolsFrame.PhoneLink = window.event.srcElement
				Call TeleToolsFrame.Dial(window.event.srcElement.innerHTML)
			Else
				Set TeleToolsFrame.PhoneLink = window.event.srcElement.parentElement
				Call TeleToolsFrame.Dial(window.event.srcElement.parentElement.innerHTML)
			End If
		End If

	End Sub

//-->
</script>