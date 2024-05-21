<%
thispage = "help"
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="version.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Turn Help On/Off</title>
</head>
<body>

<% 'Toggle help and return to the menu
If Session("help") = "FALSE" Then
	Session("help") = "TRUE"
Else
	Session("help") = "FALSE"
End If 
Response.Clear
Server.Transfer version & "_menu.asp"
%>

</body>
</html>
