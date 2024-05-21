<%
thispage = "labels"
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
site = Session("site") %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Display Labels Reference</title>
</head>
<body>
<div align="center">

<a href="<%= version %>_menu.asp" title="Return to the Menu"><strong>Return to the Menu</strong></a>
<p>
<table border="0" cellspacing="0" cellpadding="3">
	<tr>
		<td align="right"><strong>Site</strong>
		<td align="left">&nbsp;<strong>Site Name</strong>
		<td align="left">&nbsp;<strong>Location Key</strong>
		
<% 'Display the location labels
SQL = "SELECT site,longname,loclabel FROM loclabel INNER JOIN reflist ON loclabel.site = reflist.shortname WHERE list = 'SITE' ORDER BY site"
%>
<!--#Include File="db_select.asp"-->
<%
Do Until rs.eof %>

	<tr>
		<td align="right"><%= rs("site") %>
		<td align="left">&nbsp;<%= rs("longname") %>
		<td align="left">&nbsp;<%= rs("loclabel") %>

<%	rs.movenext
Loop%>

</table>
<p>
<table border="0" cellspacing="0" cellpadding="3"> 

<% 'Display the owner header
SQL = "SELECT ownername,ownerlong FROM ownerinfo WHERE ownerno = -1"
%>
<!--#Include File="db_select.asp"-->	
<%
Do Until rs.eof %>

	<tr>
		<td align="right"><strong><%= rs("ownername") %></strong>
		<td align="left"><strong>&nbsp;<%= rs("ownerlong") %></strong>
		
<%	If Session("accesstype") = "SUPER" Or Session("accesstype") = "ADMIN" Then %>

		<td align="left">&nbsp;<strong>Contact</strong>
				
<%	End If
	rs.movenext
Loop %>
<% 'Display the owner names
SQL = "SELECT ownername,ownerlong,contact FROM ownerinfo WHERE ownerno > 0 ORDER BY ownername"
%>
<!--#Include File="db_select.asp"-->	
<%
Do Until rs.eof %>

	<tr>
		<td align="right"><%= rs("ownername") %>
		<td align "left">&nbsp;<%= rs("ownerlong") %>
		
<%	If Session("accesstype") = "SUPER" Or Session("accesstype") = "ADMIN" Then %>

		<td align="left">&nbsp;<%= rs("contact") %>
				
<%	End If
	rs.movenext
Loop %>

</table>

</div>
</body>
</html>
