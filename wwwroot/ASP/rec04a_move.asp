<%
thispage = "move"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
site = Session("site") %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Select an Owner for a Box to Move</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<p><strong>Move a Box</strong>
<p>
<form method="POST" action="<%= version %>_move_select.asp">

<%'Make the Owner select list
selected_option = Session("ownername")
actowner = "TRUE"
longname = "TRUE"
listall = "FALSE" %>

	<p>Owner&nbsp;
	<!--#Include File="cd_ownerlist.asp"-->
	
<% 'Make the site selection list if show site is TRUE
If Session("showsite") = "TRUE" Then
	SQL = "SELECT DISTINCT location.site,reflist.longname,sortorder FROM location INNER JOIN reflist ON location.site = reflist.shortname WHERE list = 'SITE' AND reflist.status = 'ACTIVE' ORDER BY sortorder"
%>
<!--#Include File="db_select.asp"-->
<%
 %>

	&nbsp;&nbsp;to Destination Site&nbsp;
	<select name = "site">

<%
	Do While Not rs.eof
		If site = rs("site") Then
			selected = "selected"
		Else 
			selected = ""
		End If %> 
								
		<option value="<%= rs("site") %>" <%= selected %>><%= rs("site") & "&nbsp;&nbsp;" & rs("longname") %>	
			
	<%	rs.movenext
	Loop %>
			
	</select>
	
<%
Else 'Use the default site %>

	<input type="hidden" name="site" value="<%= site %>">
	
<%
End If %>	
	
	<p>
	<input type="submit" value="Continue">
</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/move_help.asp"--></font>
<%
End If %>
		
</div>
</body>
</html>
