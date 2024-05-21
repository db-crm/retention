<%
thispage = "vacant_display"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="../includes/version.asp"-->
<%
site = Session("site") %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Display Vacant Locations</title>
</head>
<body>
<div align="center">

<% 'Get the count of vacant and active locations
SQL = "SELECT COUNT(*) AS loccount FROM location WHERE site = '" & site & "' AND recordno = 0"
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
If Not rs.eof Then
	vacantcount = rs("loccount")
End If 
SQL = "SELECT COUNT(*) as loccount FROM location WHERE site = '" & site & "' AND recordno > 0"
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
If Not rs.eof Then
	activecount = rs("loccount")
End If %>

<p><a href="<%= version %>_menu.asp" title="Return to the Menu"><strong>There are&nbsp;<%= vacantcount %>&nbsp;Vacant Locations and&nbsp;<%= activecount %>&nbsp;Active Locations at Site <%= site %></strong></a>

<%
If vacantcount = 0 Then %>

<p>
<strong>There are no Vacant Locations to Display</strong>

<%
Else 'Show the vacant locations
	SQL = "SELECT loclabel FROM loclabel WHERE site = '" & site & "'"
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
	If Not rs.eof Then
		loclabel = rs("loclabel")
	End If
	SQL = "SELECT location FROM location WHERE site = '" & site & "' AND recordno = 0 ORDER BY location"
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
	If Not rs.eof Then %>

	<p>
	<table border="1" cellspacing="0" cellpadding="2">
		<tr align="center">
			<td><%= loclabel %>
				
<%		Do While Not rs.eof %>
			
		<tr align="center">
			<td><%= rs("location") %>
			
<%			rs.movenext
		Loop %>
		
	</table>
		
<%	End If
End If %>	

</div>
</body>
</html>
