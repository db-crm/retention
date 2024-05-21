<%
thispage = "ref_select"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->
<%
reflist = Session("reflist")
If lastpage = thispage Then
	Session("listmember") = Request("listmember")
	Select Case Request("action")
		Case "Add a Member"
			Session("refaction") = "ADD"
		Case "Activate/Deactivate/Edit the Member"
			Session("refaction") = "EDIT"
		Case "Set the Sort Order"
			Session("refaction") = "SORT"
	End Select
	Response.Clear
	Server.Transfer version & "_ref_update.asp"		
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Modify the Site Reference List</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<p><strong>Reference List Members for List <font color="#008000"><%= Session("reflist") %></font></strong>

<form action="<%= version %>_ref_select.asp" method="post">
			
<%'Display the existing site reference list members

SQL = "SELECT shortname,longname,status,sortorder FROM reflist WHERE list = '" & reflist &  "' and change = 'TRUE' ORDER BY status,sortorder,shortname"
%>
<!--#Include File="db_select.asp"-->
	
	<p>
	<select name="listmember">
	
<%
Do While Not rs.eof %> 
							
		<option value=<%= rs("shortname") %>><%= rs("status") %>&nbsp;,&nbsp;<%= rs("shortname") %>&nbsp;,&nbsp;<%= rs("longname") %>
		
<%	rs.movenext
Loop %>
			
	</select>
	<p>
	<input type="submit" name="action" value="Add a Member">
	<p>
	<input type="submit" name="action" value="Activate/Deactivate/Edit the Member">
	<p>
	<input type="submit" name="action" value="Set the Sort Order">	
</form>
		
<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/expired_help.asp"--></font>
<%
End If %>
		
</div>
</body>
</html>
