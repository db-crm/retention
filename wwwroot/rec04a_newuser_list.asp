<%
thispage = "newuser_list"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Display User Info</title>
</head>
<body>
<div align="center">

<p><a href="<%= version %>_newuser.asp" title="Return to Manage Users"><strong><%= Session("client") %> Users, <%= Date %></a>
<p>
<table width="100%" border="0" cellspacing="0" cellpadding="3">
	<tr>
		<td align="left">		
		<strong>User Name</strong>
		<td align="left">
		<strong>Access Type</strong>
		<td align="left">		
		<strong>Default Site</strong>	
		<td align="left">
		<strong>Real Name</strong>
		<td align="left">
		<strong>Organization</strong>		
		<td align="left">
		<strong>E-Mail</strong>				
		<td align="left">
		<strong>Contact</strong>
		
<%
SQL = "SELECT username,accesstype,site,organization,realname,email,usercontact FROM userinfo ORDER BY username"
%>
<!--#Include File="db_select.asp"-->				
<%
Do Until rs.eof %>

	<tr>
		<td align="left">		
		<%= rs("username") %>
		<td align="left">
		<%= rs("accesstype") %>
		<td align="left">
		<%= rs("site") %>		
		<td align="left">
		<%= rs("realname") %>	
		<td align="left">
		<%= rs("organization") %>		
		<td align="left">
		<%= rs("email") %>		
		<td align="left">
		<%= rs("usercontact") %>
		
<%	rs.movenext
Loop %>

</table>

</div>
</body>
</html>
