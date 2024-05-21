<%
thispage ="cleanup_edit"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
errors = 0 
recordno = Session("recordno")
ownerno = Session("ownerno")
ownername = Session("ownername")
box = Session("box")
If lastpage = thispage Then 'Set new values for the record
	boxid = Trim(Request("boxid"))
	retention = CInt(Request("retention"))
	recorddate = Trim(Request("recorddate"))
	content = UCase(Trim(Request("content")))
Else 'Set existing values
	SQL = "SELECT boxid,retention,recorddate,content FROM record WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_select.asp"-->	
<%
	If Not rs.eof Then
		boxid = rs("boxid")
		retention = rs("retention")
		recorddate = rs("recorddate")
		content = rs("content")
	End If
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Edit a Record</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Edit the Record for <%= ownername %>&nbsp;<%= box %>&nbsp;<%= boxid %></strong>
<form method="POST" action="<%= version %>_cleanup_edit.asp">
	<p>
	<table width="100%" border="0" cellspacing="0" cellpadding="3">
	
<% 'Make a field to enter the box identifier if boxid preference set true

If Session("useboxid") = "TRUE" Then
	labelcolor = "Black" %>

	<tr>			
		<td align="right"><font color=<%= labelcolor %>>Box Identifier
		<td>
		<input type="text" name="boxid" value="<%= boxid %>" size="16" maxlength="16">
		&nbsp;Optional
		
<%
End If %>		
			
<% 'Make the Retention Period select list. Preselect if already submitted %>

		<tr>
			<td align="right">Retention
			<td>
			<select name="retention">
	
<% 
i = 0
Do Until i > 30	
	If retention = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
				<option <%= selected %>><%= i %>
	
<%	i = i + 1
Loop 

i = 35
Do Until i > 50
	If retention = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
				<option <%= selected %>><%= i %>
	
<%	i = i + 5
Loop

i = 60
Do Until i > 100
	If retention = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
				<option <%= selected %>><%= i %>
	
<%	i = i + 10
Loop %>
		 	
			</select>
			&nbsp;Years
		
<%'Make the Record Date entry field.  Preload existing date. Label red if bad date entered
labelcolor = "Black"
If lastpage = thispage Then
	If IsDate(recorddate) Then
		If DateDiff("d",recorddate,Date) < 0 Then
			errors = errors + 1
			labelcolor = "Red"
		End If
	ELse
		errors = errors + 1
		labelcolor = "Red"
	End If
End If %>

	<tr>			
		<td align="right"><font color=<%= labelcolor %>>Record Date</font>
		<td>
		<input type="text" name="recorddate" value="<%= recorddate %>" size="10" maxlength="10">

<%'Make the Content text area.  Load text already submitted. Label red if blank
labelcolor = "Black"
If lastpage = thispage Then
	If Len(content) = 0 Then
		errors = errors + 1
		labelcolor = "Red"
	End If
End If %>

	<tr>		
		<td align="right" valign="top"><font color=<%= labelcolor %>>Content</font>
		<td>
		<textarea cols="80" rows="5" name="content" wrap="soft"><%= content %></textarea>
</table>
	
<%
If lastpage = thispage Then 
	If errors <> 0 Then 'Invalid entries%>
	
	<p><font color="Red"><strong>Correct the fields labeled in red and Validate again</strong></font>

		
<%	Else 'Everything OK. Last chance to change
		If Request("action") = "Finished" Then
			Response.Clear
			Server.Transfer version & "_cleanup_update.asp"
		End If %>
 			
<p><font color="Green"><strong>The data is valid.  To update the record in the database, click Finished</strong></font>

	<p>
	<input type="submit" name="action" value="Finished">
	<p><strong>If you make any changes to this page, click Validate</strong>

<%	End If
End If %>

	<p>
	<input type="submit" value="Validate">
</form>

</div>
</body>
</html>
