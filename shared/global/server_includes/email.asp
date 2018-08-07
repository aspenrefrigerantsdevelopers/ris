<%
'---- Mail configuration settings ----
Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const cdoSMTPServerPickupDirectory =  "http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory"
Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const cdoSendUsingPickup = 1
Const cdoSendUsingPort = 2
Const cdoSendServer = "mail.refron.com"
Const cdoSendPort = 25
Const cdoPickupDirectory =  "C:\Inetpub\Mailroot\Pickup"

'---- Creates a message configuration object for use by CDO ----
Dim msgConfig
Sub SetMsgConfig()
	Set msgConfig = Server.CreateObject("CDO.Configuration")
	msgConfig.Fields.Item(cdoSMTPServer) = cdoSendServer
	msgConfig.Fields.Item(cdoSMTPServerPort) = cdoSendPort
	msgConfig.Fields.Item(cdoSendUsingMethod) = cdoSendUsingPort	'Change this only
	msgConfig.Fields.Item(cdoSMTPServerPickupDirectory) = cdoPickupDirectory
	msgConfig.Fields.Update
End Sub

Function IsEmailAddress(strEmail)
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern ="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$" 
	regEx.IgnoreCase = True
	IsEmailAddress = regEx.Test(strEmail)
End Function
%>