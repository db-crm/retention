<%
thispage = "showsite" 'Transfer to and from this page to enter top of menu with showsite changed
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
If Session("showsite") = "TRUE" Then
	Session("showsite") = "FALSE"
Else
	Session("showsite") = "TRUE"
End If
Response.Clear
Server.Transfer version & "_menu.asp" %>	

