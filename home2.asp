<!--#include virtual="/shared/global/server_includes/global.asp"-->

<%

Dim Page
Set Page = Server.CreateObject("RefronIntranetPage.WSC")

Page.Title = "Refron Intranet"
Page.Stylesheet = ""
Page.PrintHeader()

Dim strMainPage
If Request.QueryString("URL") <> "" Then
	strMainPage = Request.QueryString("URL")
Else
	strMainPage = "/menu.asp?p=101"
End If
%>

<frameset rows="20,*" frameborder="0" framespacing="0" border="0">
	<frame name="banner" src="global/banner.asp" target="main" scrolling="no" noresize marginwidth="0" marginheight="0" tabindex="3" frameborder="0">

	<frameset cols="120,*">
		<frame name="menu" src="global/menu.asp" target="main" scrolling="no" noresize marginwidth="0" marginheight="0" tabindex="2" frameborder="0">
		<frame name="main" src="<%=strMainPage%>" target="_self" scrolling="yes" noresize tabindex="1" frameborder="0">
	</frameset>

	<noframes>
		<body>
		<p>This page uses frames, but your browser doesn't support them.</p>
		</body>
	</noframes>
</frameset>


<%
Set Page = Nothing
%>

</html>