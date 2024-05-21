<%
thispage = "vacant_add"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->
<%
errormsg = ""
plan_case = Session("plan_case")
plan_length = Session("plan_length")
plan_levels = Session("plan_levels")
plan_maxloc = Session("plan_maxloc")
plan_delim = Session("plan_delim")
site = Session("site")

If UCase(plan_delim) = "SPACE" Then
	dlm = "&nbsp;"
ElseIf UCase(plan_delim) = "NONE" Then
	dlm = ""
Else
	dlm = plan_delim
End If

If lastpage = thispage Then

	r = 1 'Input level
	Do Until r = plan_levels + 1 'Make comma delimited lists level1 through level(plan_levels)
		c = 1
		level = ""
		Do Until c = plan_maxloc + 1
			cell = r & "x" & c
			value = Trim(Replace(Request(cell),",","")) 'Trim spaces and remove commas
			value = Replace(value,plan_delim,"")
			If plan_case = "UPPER" Then value = UCase(value) 'Upper case if upper set
			If Len(value) > 0 Then 'Add to comma delimited list
				level = level & "," & value
			End If
			c = c + 1
		Loop
		If Len(level) > 0 Then 'Remove the leading comma
			level = Right(level,Len(level) - 1)
		Else
			errormsg = "A Level cannot be empty"
		End If
		levelname = "level" & r
		Session(levelname) = level
		str = levelname & " = level"
		Execute str
	r = r + 1
	Loop

	loclabel = Request("loclabel")
	Session("loclabel") = loclabel
	If Not Len(loclabel) > 0 Then
		errormsg = "Location Label cannot be blank"
	End If
	
	If Request("action") <> "Validate" And errormsg = "" Then
		Response.Clear
		Server.Transfer version & "_vacant_process.asp"
	End If	
Else
	SQL = "SELECT loclabel FROM loclabel WHERE site = '" & site & "'"
%>
<!--#Include File="db_select.asp"-->
<%
	If rs.eof Then
		loclabel = ""
	Else
		loclabel = rs("loclabel")
	End If
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention View and Add and Remove Vacant Locations</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Add and Remove Vacant Locations</font></strong>
<p><font color="#FF0000"><%= errormsg %></font>
<form method="POST" action="<%= version %>_vacant_add.asp">
	<p>Location Label&nbsp;
	<input type="text" name="loclabel" value="<%= loclabel %>" size="64" maxlength="64">
	&nbsp;at Site&nbsp;<strong><%= site %></strong>
	<p>
	<table border="0" cellspacing="0" cellpadding="3">
		<tr>
			<td>Level
			<td>&nbsp;&nbsp;
			<td colspan="<%= plan_maxloc %>">Location Identifiers
 	
<%
r = 1	
Do Until r = plan_levels + 1 %>

		<tr>
			<td align="right"><%= r %>
			<td>

<%	If lastpage = thispage Then 'Make an array from the levelr variable
		str = "level = level" & r 
		Execute str
		arylevel = Split(level,",")
	Else 'Empty array first time through
		arylevel = Array("")
	End If
	c = 1
	Do Until c = plan_maxloc + 1
		i = c - 1 'Array base is 0
		If i > UBound(arylevel) Then 'Blanks if fewer than maxloc at this level
			locvalue = ""
		Else
			locvalue = arylevel(i)
		End If %>
	
			<td>
			<input type="text" name="<%= r %>x<%= c %>" value="<%= locvalue %>" size="<%= plan_length %>" maxlength="<%= plan_length %>">

<%		c = c + 1
	Loop
	r = r + 1
Loop %>
	
	</table>
	<p>
	<input type="submit" name="action" value="Validate">&nbsp;
		
<%
If lastpage = thispage And errormsg = "" Then
	If Session("vacant_action") = "ADD" Then %>	
	
	<input type="submit" name="action" value="Add">
	
<%	Else %>
	
	<input type="submit" name="action" value="Remove">
	
<%	End If
End If %>

</form>

<p>Example for Location Delimiter '<strong><%= plan_delim %></strong>'<br>
&nbsp;<br>
A&nbsp;&nbsp;B<br>
1&nbsp;&nbsp;2&nbsp;&nbsp;3<br>
&nbsp;<br>
will <%= Session("vacant_action") %> locations&nbsp;&nbsp;<strong>A<%= dlm %>1&nbsp;&nbsp;A<%= dlm %>2&nbsp;&nbsp;A<%= dlm %>3&nbsp;&nbsp;B<%= dlm %>1&nbsp;&nbsp;B<%= dlm %>2&nbsp;&nbsp;B<%= dlm %>3&nbsp;&nbsp;at Site <%= site %></strong>

</div>
</body>
</html>
