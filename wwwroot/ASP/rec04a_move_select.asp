<%
thispage = "move_select"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
ownername = Request("ownername")
Session("ownername") = ownername
site = Request("site")
Session("site") = site
SQL = "SELECT longname FROM reflist WHERE list = 'SITE' AND shortname = '" & site & "'"
%>
<!--#Include File="db_select.asp"-->
<%
sitelong = rs("longname")
Session("sitelong") = sitelong

'Get the owner number and long name.
%>
<!--#Include File="cd_ownerno.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Select a Box to Move</title> 
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Move a Box</strong>
<form method="POST" action="<%= version %>_move_insert.asp">
	<p>Owner&nbsp;<strong><%= ownername %></strong>&nbsp;Box&nbsp;
		
<%'Make the Box select list.
SQL="SELECT recordno,box,boxid FROM record WHERE ownerno = " & ownerno & " AND status = 'ACTIVE' ORDER BY box"
%>
<!--#Include File="db_select.asp"-->	

	<select name="recordno">
					
<%
Do While Not rs.eof %>
	
		<option value=<%= rs("recordno") %>><%= rs("box")%>&nbsp;<%= rs("boxid") %>		
		
<%	rs.movenext
Loop %>
	
	</select>
	&nbsp;New Location&nbsp;	
		
<%'Make the Location select list.
SQL="SELECT location FROM location WHERE site = '" & site & "' AND recordno = 0 ORDER BY location"
%>
<!--#Include File="db_select.asp"-->	
<%
If rs.eof Then %>

<font color="#FF0000">No Available Locations</font>	at Site <%= site %>&nbsp;&nbsp;<%= sitelong %>		
					
<%
Else %>

	<select name="new_location" size="1">

<%	Do While Not rs.eof %>
						
		<option><%= rs("location") %>		
		
<%		rs.movenext
	Loop %>
	
	</select>&nbsp;
	at Site <%= site %>&nbsp;&nbsp;<%= sitelong %>
	<p>
	<input type="submit" value="Move">
	
<%
End If %>
	
</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/move_select_help.asp"--></font>
<%
End If %>
		
</div>
</body>
</html>
