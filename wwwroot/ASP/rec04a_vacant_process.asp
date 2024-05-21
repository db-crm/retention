<%
thispage = "vacant_process"
pageaccess = "ADMIN"
%>
<!--#Include File="../includes/version.asp"-->
<%
plan_levels = Session("plan_levels")
plan_delim = Session("plan_delim")
site = Session("site")
If UCase(plan_delim) = "SPACE" Then
	dlm = "&nbsp;"
ElseIf UCase(plan_delim) = "NONE" Then
	dlm = ""
Else
	dlm = plan_delim
End If

If Len(Trim(Session("loclabel"))) > 0 Then 'Update the location label for this site
	SQL = "DELETE FROM loclabel WHERE site = '" & site & "'"
%>
<!--#Include File="../includes/db_execute.asp"-->	
<%
	SQL = "INSERT into loclabel(site,loclabel) VALUES('" & site & "','" & Session("loclabel") & "')"
%>
<!--#Include File="../includes/db_execute.asp"-->	
<%End If

'This page uses level1 through level(plan_levels) comma delimited lists of location elements created by
'the vacant_add page.
'The lists are passed in Session("level1"), etc.  Locations are created by making cross products between
'a level and the level above starting with the lowest level child (highest numeric level).  The process
'is repeated between the resultant product and the next highest level until the highest parent level (lowest 'numeric level) which is the completed location.  The delimiter specified by plan_delim is used when making
'the products.  Each finished location is inserted or removed from the location table depending on the 
'value of Session("vacant_action").  Duplicates are not inserted and active locations are not removed.
'All locations are site specific and identical locations may exist at different sites.

addcount = 0
removecount = 0
dupcount = 0
activecount = 0
i = CInt(Session("plan_levels"))
result = Session("level" & i) 'First lower array is highest level
i = i - 1
Do Until i = 0
	aryupper = Split(Session("level" & i),",") 'Upper array is one above lower array
	arylower = Split(result,",") 'Next lower array is result of cross product
	result = "" 'Clear result
	j = 0
	Do Until j = UBound(aryupper) + 1
		k = 0
		Do Until k = UBound(arylower) + 1 'Result is cross product of upper and lower arrays
			location = aryupper(j) & dlm & arylower(k)
			If i = 1 Then 'Completed location, add or remove from the database 
				If Session("vacant_action") = "ADD" Then 'See if the location already exists
					SQL = "SELECT * FROM location WHERE site = '" & site & "' AND location = '" & location & "'"
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
					If rs.eof Then 'No, add the location
						SQL = "INSERT into location(site,location,recordno) VALUES('" & site & "','" & location & "',0)"
%>
<!--#Include File="../includes/db_execute.asp"-->	
<%
						addcount = addcount + 1
					Else 'Yes, increment the duplicate count
						dupcount = dupcount + 1
					End If
					msg = addcount & " new vacant locations were added to the database<br>" & dupcount & " duplicate locations were not added to the database"
				Else 'See if the location to be used is active
					SQL = "SELECT * FROM location WHERE site = '" & site & "' AND location = '" & location & "' AND recordno > 0"
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
					If rs.eof Then 'No, remove the location
						SQL = "DELETE FROM location WHERE site = '" & site & "' AND location = '" & location & "'"
%>
<!--#Include File="../includes/db_execute.asp"-->	
<%
						removecount = removecount + 1
					Else 'Yes, increment the active count
						activecount = activecount + 1
					End If
					msg = removecount & " vacant locations were removed from the database<br>" & activecount & " active locations were not removed from the database"
				End If
			Else 'Make the result list
				result = result & "," & location
			End If	
			k = k + 1
		Loop
		j = j + 1
	Loop
	If i <> 1 Then 'No result list when i = 1
		result = Right(result,Len(result) - 1) 'Remove leading comma
	End If
	i = i - 1 'Move up one level	
Loop %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Process Vacant Locations Add and Remove</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="../includes/header.asp"-->

<p><strong><%= msg %></strong>
<p><strong>Vacant Locations for Site <%= site %></strong>
<p>
<table width="100%" border="0" cellspacing="0" cellpadding="3">

<%
SQL = "SELECT location FROM location WHERE site = '" & site & "' AND recordno = 0 ORDER BY location"
%>
<!--#Include File="../includes/db_select.asp"-->
<%
Do %>

	<tr>
	
<%	c = 1
	Do Until c = 9
		If rs.eof Then Exit Do %>
	
		<td align="center"><%= rs("location") %>
		
<%		rs.movenext
		c = c + 1
	Loop
	If rs.eof Then Exit Do
Loop %>

</table>	

</div>
</body>
</html>
