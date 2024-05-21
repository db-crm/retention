<%
thispage = "expired_update"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
datedelim = Session("datedelim")

If lastpage <> thispage Then 'Get original update list
	Select Case Session("update_mode")
		Case "expired"
			updatelist = Session("updatelist")
		Case "report"
			updatelist = Request("updatelist")
		Case Else 'This case should not occur
			Response.Clear
			Server.Transfer version & "_menu.asp"
	End Select
Else 'Perform the update
	updatelist = Request("updatelist") 'Get update list made by this page
	recordno = Request("recordno")
	recorddate = Request("recorddate")
	retention = Request("new_retention")
	content = Request("content")
	boxid = Trim(Request("boxid"))
	If Not IsEmpty(Request("new_content")) Then 'No empty content field
		new_content = UCase(Replace(Request("new_content"),"'","''"))
	Else
		new_content = content
	End If
	histaction = "UPDATE"
%>
<!--#Include File="cd_history.asp"-->	
<%

'Make the Review date always the first of the month and ten years from now if retention is 100
If retention = 100 Then
	reviewdate = DatePart("m",Date) & "/1/" & (DatePart("yyyy",Date) + 10)
Else
	reviewdate = DatePart("m",recorddate) & "/1/" & (DatePart("yyyy",recorddate) + retention)
End If

	SQL="UPDATE record SET retention = " & retention & ",boxid = '" & boxid & "',content = '" & new_content & "',reviewdate = " & datedelim & reviewdate & datedelim & ",updatedate = " & datedelim & Date & datedelim & " WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_execute.asp"-->	
<%
End If

If updatelist = "none" Then 'If updatelist finished, return to the calling page
	Response.Clear
	Server.Transfer Session("returnto")
Else 'Parse update list for leftmost record number; then remove from list
	delimpos = Instr(updatelist,",")
	If delimpos = 0 Then 'No delimeter, only one record number
		recordno = updatelist
		updatelist = "none"
	Else
		recordno = Left(updatelist,delimpos-1)
		updatelist = Right(updatelist,Len(updatelist)-delimpos)
	End If
End If

'Get the ownernames, box, old retention period, record date, and content for the record to be updated
SQL = "SELECT ownername,ownerlong,box,boxid,content,retention,recorddate FROM record INNER JOIN ownerinfo ON record.ownerno = ownerinfo.ownerno WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_select.asp"-->	
<%
'Old values
ownername = rs("ownername")
ownerlong = rs("ownerlong")
box = rs("box")
boxid = rs("boxid")
retention = rs("retention")
content = rs("content")
recorddate = rs("recorddate") %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Update Records</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	
	
<p><strong>Update for Owner <%= ownername & "&nbsp;&nbsp;" & ownerlong %>, Box <%= box %>&nbsp;<%= boxid %></strong>
<form action="<%= version %>_expired_update.asp" method="post">

<%
If Session("useboxid") = "TRUE" Then %>

	<p>Box Identifier&nbsp;
	<input type="text" name="boxid" value="<%= boxid %>" size="16" maxlength="16">
	&nbsp;Optional
	
<%
Else %>

	<input type="hidden" name="boxid" value="<%= boxid %>">
	
<%	
End If %>
	
	<p>Retention Period&nbsp;
	<select name="new_retention">
	
<%
If Session("update_mode") = "expired" Then
	i = CInt(retention)
Else
	i = 0
End If

Do Until i > 30
	If retention = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
		<option <%=selected%>><%= i %>
	
<%	i = i + 1
Loop %> 
	
<%  i = 35
Do Until i > 50
	If retention = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
		<option <%=selected%>><%= i %>
	
<%	i = i + 5
Loop %>
	
<% i = 60
Do Until i > 100
	If retention = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
		<option <%=selected%>><%= i %>
	
<%	i = i + 10
Loop %>
		 	
	</select>
	&nbsp;in Years from the Record Date
	<p>
	<textarea cols="80" rows="5" name="new_content"><%= content %></textarea>
	<p>
	<input type="submit" value="Update">
	<input type="hidden" name="updatelist" value="<%= updatelist %>">
	<input type="hidden" name="recordno" value="<%= recordno %>">
	<input type="hidden" name="recorddate" value="<%= recorddate %>">
	<input type="hidden" name="retention" value="<%= retention %>">
	<input type="hidden" name="content" value="<%= content %>">	
</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/expired_update_help.asp"--></font>
<%
End If %>

</div>
</body>
</html>
