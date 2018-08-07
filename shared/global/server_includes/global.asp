<%@ Language=VBScript%>

<%
Option Explicit

If Application("AFS_is_Down") = "" then
    Application("AFS_is_Down") = False
End if

'---- Define the environment (live or dev)
Const gIsLive = True
Const liveServerName = "172.20.92.145"

'---- Sets page path for reference ----
Sub SetPath()
	Dim PagePath, PagePathTime
	PagePath = Request.ServerVariables("PATH_INFO")
	If PagePath <> "/apps/mail/mailbox_list.asp" Then
		If Request.ServerVariables("QUERY_STRING") <> "" Then
			PagePath = PagePath & "?" &  Request.ServerVariables("QUERY_STRING")
		End If
		Session("PagePath5Time") = Session("PagePath4Time")
		Session("PagePath5") = Session("PagePath4")
		Session("PagePath4Time") = Session("PagePath3Time")
		Session("PagePath4") = Session("PagePath3")
		Session("PagePath3Time") = Session("PagePath2Time")
		Session("PagePath3") = Session("PagePath2")
		Session("PagePath2Time") = Session("PagePath1Time")
		Session("PagePath2") = Session("PagePath1")
		Session("PagePath1Time") = Session("PagePathTime")
		Session("PagePath1") = Session("PagePath")
		Session("PagePathTime") = Now()
		Session("PagePath") = PagePath
	End If
End Sub

Function TogglePrint
	Application("togglePrint") = Not CBool(Application("togglePrint"))
	TogglePrint = Application("togglePrint")
End Function

'---- Include debug class ----
%><!--#include virtual="/shared/global/classes/clsDebug.asp"--><%
Dim Debug
Set Debug = New clsDebug

'---- Include local global functions file ----
%><!--#include virtual="/global/local_global.asp"--><%

'---- Include custom style functions ----
%><!--#include virtual="/global/styles.asp"--><%

'---- Include custom style functions ----
%><!--#include virtual="/shared/global/server_includes/as400.asp"-->