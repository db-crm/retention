<%
thispage = "expired_destruct"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
ownername = Request("ownername")
Session("ownername") = ownername

'Get the owner number, longname and contact info
%>
<!--#Include File="cd_ownerno.asp"-->
	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Destruction Form</title>	
</head>
<body>
<div align="center">

<%
If Session("showsite") = "TRUE" Then
	lochead = "Site" & "/" & "Location"
Else
	lochead = "Location"
End If %>	

<table width="100%" border="0" cellspacing="0" cellpadding="5">
	<tr>
		<td><a href="<%= version %>_expired.asp">Expired Records, <%= Session("expired_label") %></a>&nbsp;&nbsp;<strong><%= ownername %>&nbsp;&nbsp;<%= ownerlong %>&nbsp;&nbsp;<%= Date %>&nbsp;</strong>
	<tr>
		<td>Contact: <%= contact %>	
	<tr>
		<td>The records listed below have met or exceeded the required retention as specified by the <%= Session("client") %>.  These records serve no further value to my unit and therefore I authorize that they be destroyed in accordance with the <%= Session("client") %> records destruction policy.
	<tr>
		<td height="45" valign="bottom">
		Department Head__________________________________________Date____________
	<tr>
		<td height="45" valign="bottom">
		Division Head_____________________________________________Date____________
	<tr>			
		<td height="45" valign="bottom">
		City Clerk________________________________________________Date____________
	<tr>
		<td>If there are records which you do not want destroyed, initial the 'Save' column and enter a new retention period in years.  New retention period = Old retention period + Additional period.	
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="2" bordercolorlight="#000000">
	<tr align="center">
		<td rowspan="2" valign="top">Save
		<td colspan="2">Retention
		<td rowspan="2" valign="top">Box
		<td rowspan="2" valign="top">Content
		<td rowspan="2" valign="top"><%= lochead %>
	<tr align="center">
		<td>New
		<td>Old
		
<%'Get the records ready for destruction for the selected owner and reviewdate specified by expired_query
SQL = "SELECT site,location,content,box,boxid,retention FROM record INNER JOIN location ON record.recordno = location.recordno WHERE ownerno = " & ownerno & " AND record.status = 'ACTIVE' AND " & Session("expired_query") & " ORDER BY box"
%>
<!--#Include File="db_select.asp"-->	
<%
Do While Not rs.eof
	location = rs("location")
	If Session("showsite") = "TRUE" Then
		location = rs("site") & "/" & location
	End If %>

	<tr align="center">
		<td><br>
		<td><br>
		<td valign="top"><%= rs("retention") %>
		<td valign="top"><%= rs("box") %>&nbsp;<%= rs("boxid") %>
		<td align="left"><%= rs("content") %>
		<td valign="top"><%= location %>

<%	rs.movenext
Loop %> 	
		
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="5">
	<tr>
		<td>The above listed records belonging to the <%= Session("client") %>, were destroyed in accordance with the <%= Session("client") %> record destruction policy on _______________,____________.
	<tr>
		<td height="45" valign="bottom">
		Records Management_______________________________________Date____________
</table>

</div>
</body>
</html>
