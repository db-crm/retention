<%
thispage = "quit"
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="version.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Quit</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<% Session.abandon %>

<p><strong>Thank you for logging off the records retention system</strong></p>
<p><strong>To log in again, open the <a href="<%= version %>_login.asp">records login page</strong></a></p>
		
</div>
</body>
</html>
