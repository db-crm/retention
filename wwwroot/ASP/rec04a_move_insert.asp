<%
thispage = "move_insert"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
recordno = CInt(Request("recordno"))
location = Request("new_location")
site = Session("site")

'Get entry date for the record to be moved
SQL = "SELECT entrydate FROM record WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_select.asp"-->
<%
If Not rs.eof Then
	entrydate = FormatDateTime(rs("entrydate"),2)
End If

histaction = "MOVE" 'Make an entry in the history table
%>
<!--#Include File="cd_history.asp"-->	
<%

'Get the old location
SQL = "SELECT site,location FROM location WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_select.asp"-->
<%	
old_location = rs("location")
old_site =rs("site")

'Make the old location vacant
SQL = "UPDATE location SET recordno = 0 WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_execute.asp"-->
<%

'Move the box by putting record number in the new location
SQL = "UPDATE location SET recordno = " & recordno & " WHERE location = '" & location & "'"
%>
<!--#Include File="db_execute.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Move a Box and Make a New Label</title>	
</head>
<body>
<div align="center">

<p><a href="<%= version %>_move.asp" title="Return to Move a Box">The box location moved from <%= old_site %>&nbsp;<%= old_location %> on <%= Date %>.</a>&nbsp;&nbsp;<strong>Record Number&nbsp;<%= recordno %>&nbsp;Created&nbsp;<%= entrydate %></strong>

<p>
<!--#Include File="cd_box_label.asp"-->

</div>
</body>
</html>
