<%
thispage ="add"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
errors = 0
site = Session("site") %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Add a Record</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<%
SQL = "SELECT longname FROM reflist WHERE list = 'SITE' AND shortname = '" & site & "'"
%>
<!--#Include File="db_select.asp"-->
<%
Session("sitelong") = rs("longname") %>

<p><strong>Add a Record</strong>
<form method="POST" action="<%= version %>_add.asp">
	<p>
	<table width="100%" border="0" cellspacing="0" cellpadding="3">
	
<%'Make the Location select list.  Preselect if already submitted. Label red if no spaces %>

		<tr>		
			<td width="15%" align="right">Site
			<td width="75%"><%= site %>&nbsp;&nbsp;<%= Session("sitelong") %>
		<tr>	
			<td align="right">Location
			<td>			
			
<%
labelcolor = "Black"			
If lastpage = thispage Then
	selected_option = Request("location")
Else 
	selected_option = ""
End If %>			
			
			<select name="location">

<%
SQL="SELECT location FROM location WHERE recordno = 0 AND site = '" & site & "' ORDER BY location"
%>
<!--#Include File="db_select.asp"-->	
<%
If rs.eof Then
	errors = errors + 1
	labelcolor = "Red" %>
			
				<option><font color="<%= labelcolor %>">No Available Spaces</font>
					
<%
Else
	Do While Not rs.eof
		optionvalue = rs("location")
		optionlabel = rs("location")
%>
<!--#Include File="cd_selectoption.asp"-->
		
<%		rs.movenext
	Loop 
End If %>
       
			</select>
			&nbsp;<%= Session("loclabel") %>
		
<%'Make the Owner select list. Preselect if already submitted %>

		<tr>		
			<td align="right">Owner
			<td>
			
<%
If lastpage = thispage Then
	selected_option = Request("ownername")
Else 
	selected_option = ""
End If
listall = "FALSE"
actowner = "TRUE"
longname = "TRUE"
%>			
<!--#Include File="cd_ownerlist.asp"-->

<% 'Make a field to enter the box identifier if boxid preference set true

If Session("useboxid") = "TRUE" Then
	labelcolor = "Black"
	If lastpage = thispage Then
		lastvalue = Trim(Request("boxid"))
	Else 
		lastvalue = ""
	End If %>

	<tr>			
		<td align="right"><font color=<%= labelcolor %>>Box Identifier
		<td>
		<input type="text" name="boxid" value="<%= lastvalue %>" size="16" maxlength="16">
		&nbsp;Optional
		
<%
End If %>		
			
<% 'Make the Retention Period select list. Preselect if already submitted %>

		<tr>
			<td align="right">Retention
			<td>
			
<%
If lastpage = thispage Then
	selected_option = CInt(Request("retention"))
Else 
	selected_option = ""
End If %>
			
			<select name="retention">
	
<% 
i = 0
Do Until i > 30	
	If selected_option = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
				<option <%= selected %>><%= i %>
	
<%	i = i + 1
Loop 

i = 35
Do Until i > 50
	If selected_option = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
				<option <%= selected %>><%= i %>
	
<%	i = i + 5
Loop

i = 60
Do Until i > 100
	If selected_option = i Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
				<option <%= selected %>><%= i %>
	
<%	i = i + 10
Loop %>
		 	
			</select>
			&nbsp;Years
		
<%'Make the Record Date entry field.  Preload today date. Label red if bad date entered
labelcolor = "Black"
If lastpage = thispage Then
	lastvalue = Trim(Request("recorddate"))
	If IsDate(lastvalue) Then
		If DateDiff("m",lastvalue,Date) > 60 Or DateDiff("d",lastvalue,Date) < 0 Then
			errors = errors + 1
			labelcolor = "Red"
		End If
	ELse
		errors = errors + 1
		labelcolor = "Red"
	End If
Else 
	lastvalue = Date
End If %>

	<tr>			
		<td align="right"><font color=<%= labelcolor %>>Record Date</font>
		<td>
		<input type="text" name="recorddate" value="<%= lastvalue %>" size="10" maxlength="10">
		&nbsp;Enter the Record Date if it is not Today's Date

<%'Make the Content text area.  Load text already submitted. Label red if blank
labelcolor = "Black"
If lastpage = thispage Then
	lastvalue = UCase(Trim(Request("content")))
	If Len(lastvalue) = 0 Then
		errors = errors + 1
		labelcolor = "Red"
	End If
Else 
	lastvalue = Empty
End If %>

	<tr>		
		<td align="right" valign="top"><font color=<%= labelcolor %>>Content</font>
		<td>
		<textarea cols="80" rows="5" name="content" wrap="soft"><%= lastvalue %></textarea>
</table>
	
<%
If lastpage = thispage Then 
	If errors <> 0 Then 'Invalid entries%>
	
	<p><font color="Red"><strong>Correct the fields labeled in red and Validate again</strong></font>

		
<%	Else 'Everything OK. Last chance to change
		If Request("action") = "Finished" Then
			Response.Clear
			Server.Transfer version & "_add_insert.asp"
		End If %>
 			
<p><font color="Green"><strong>The data is valid.  To add the record to the database, click Finished</strong></font>

	<p>
	<input type="submit" name="action" value="Finished">
	<p><strong>If you make any changes to this page, click Validate</strong>

<%	End If
End If %>

	<p>
	<input type="submit" value="Validate">
</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/add_help.asp"--></font>
<%
End If %>

</div>
</body>
</html>
