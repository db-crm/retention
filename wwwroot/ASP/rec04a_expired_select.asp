<%
thispage ="expired_select"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
'Form for selecting delete or update for expired records.  Disposition is ignored if 
'update or no action is selected

'Make arrays for disposition and disposition long
arydispolist = Split(Session("dispolist"),",")
arydispolistlong = Split(Session("dispolistlong"),",")

'Get the owner number and long name.
ownername = Request("ownername")
%>
<!--#Include File="cd_ownerno.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Select Expired Records for Deletion or Update</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<%
If Session("showsite") = "TRUE" Then
	lochead = "Site" & "/" & "Location"
Else
	lochead = "Location"
End If %>		

<p><strong>Owner&nbsp;&nbsp;<%= ownername & "&nbsp;&nbsp;" & ownerlong %></strong>
<form method=POST action="<%= version %>_expired_process.asp">
	<p>
	<input type=submit value="Submit">
	<p>
	<table border="1" cellspacing="0" cellpadding="2" bordercolorlight="#000000">
		<tr align="center">
			<td>Box
			<td><%= lochead %>
			<td>Review Date
			<td>Delete
			<td>Update
			<td>No Action
			<td>Disposition			
			
<% 'Get the expired records for the selected owner selected and reviewdate specified by expired_query
SQL="SELECT record.recordno,site,location,box,boxid,reviewdate FROM record INNER JOIN location ON record.recordno = location.recordno WHERE ownerno = " & ownerno & " AND record.status = 'ACTIVE' AND " & Session("expired_query") & " ORDER BY box"
%>
<!--#Include File="db_select.asp"-->	
<%
Do While Not rs.eof
	location = rs("location")
	If Session("showsite") = "TRUE" Then
		location = rs("site") & "/" & location
	End If %>

		<tr align="center">
			<td><%= rs("box") %>&nbsp;<%= rs("boxid") %>	
			<td><%= location %>
			<td><%= rs("reviewdate") %>
			<td>
			<input type="radio" name="rr<%= rs("recordno") %>" value="rrdelete">
			<td>
			<input type="radio" name="rr<%= rs("recordno") %>" value="rrupdate">
			<td>
			<input type="radio" checked name="rr<%= rs("recordno") %>" value="rrnone">
			<td>
			<select name="rc<%= rs("recordno") %>">
		
<%	i = 0
	Do Until i > Ubound(arydispolist) %>

				<option value="<%= arydispolist(i) %>"><%= arydispolistlong(i) %>
				
<%		i = i + 1
	Loop %>			
				
			</select>				
	
<%	rs.movenext
Loop %> 	
				
	</table>
</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/expired_select_help.asp"--></font>
<%
End If %>

</div>
</body>
</html>
