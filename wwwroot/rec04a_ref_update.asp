<%
thispage = "ref_update"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->
<%
errormsg = ""
refaction = Session("refaction")
reflist = Session("reflist")
listmember = UCase(Trim(Request("listmember")))

If lastpage = thispage Then
	longname = Trim(Request("longname"))
	status = Request("status")
	Select Case refaction
		Case "ADD"
			If Len(listmember) > 0 And Len(longname) > 0 Then 'OK, see if the member already exists
				SQL = "SELECT * FROM reflist WHERE list = '" & reflist & "' AND shortname = '" & listmember & "'"
%>
<!--#Include File="db_select.asp"-->	
<%
				If Not rs.eof Then 'No duplicates
					errormsg = "List member " & listmember & " already exists"
				Else 'OK insert the new member
					SQL = "INSERT into reflist(list,shortname,longname,status,sortorder,change) values('" & reflist & "','" & listmember & "','" & longname & "','ACTIVE',NULL,'TRUE')"
%>
<!--#Include File="db_execute.asp"-->	
<%
					Session("refaction") = "SORT"
					Response.Clear
					Server.Transfer version & "_ref_select.asp"
				End If
			Else
				errormsg = "The Short Name and Long Name cannot be blank"
			End If
		Case "EDIT"
			If Len(longname) > 0 Then
				SQL = "UPDATE reflist SET longname = '" & longname & "',status = '" & status & "' WHERE list = '"  & reflist & "' AND shortname = '" & listmember & "'"
%>
<!--#Include File="db_execute.asp"-->	
<%
				Response.Clear
				Server.Transfer version & "_ref_select.asp"
			Else
				errormsg = "The Long Name cannot be blank"
			End If
		Case "SORT"
			SQL = "SELECT shortname FROM reflist WHERE list = '" & reflist & "'"
%>
<!--#Include File="db_select.asp"-->	
<%
			Do Until rs.eof
				shortname = rs("shortname")
				sortorder = Trim(Request("sort" & shortname))
				If IsNumeric(sortorder) Then
					sort = " sortorder = " & sortorder & " "
				Else
					sort = " sortorder = NULL "
				End If
				SQL = "UPDATE reflist SET " & sort & " WHERE list = '"  & reflist & "' AND shortname = '" & shortname & "'"
%>
<!--#Include File="db_execute.asp"-->	
<%
				rs.movenext
			Loop
			Response.Clear
			Server.Transfer version & "_ref_select.asp"
	End Select
Else 'Initial valus
	If refaction = "ADD" Then 'Blank for ADD
		listmember = ""
		longname = ""
	Else 'Get the current values
		SQL = "SELECT shortname,longname,status FROM reflist WHERE list = '" & reflist & "' AND shortname = '" & listmember & "'"
%>
<!--#Include File="db_select.asp"-->	
<%
		listmember = rs("shortname")
		longname = rs("longname")
	End If
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Update or Add a Reference List Member</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<%
If Len(errormsg) > 0 Then %>

<p><font color="#FF0000"><%= errormsg %></font>

<%
End If %>

<form action="<%= version %>_ref_update.asp" method="post">
	<p>
	
<%
Select Case refaction
	Case "ADD" %>
	
	<strong>Add a Member to Reference List <font color="#008000"><%= reflist %></font></strong>
	<p>	
	<table width="100%" border="0" cellspacing="0" cellpadding="3">	
		<tr>
			<td width="35%" align="right">Short Name
			<td width="65%">
			<input type="text" name="listmember" value="<%= listmember %>" size="16" maxlength="16">
		<tr>
			<td align="right">Long Name
			<td>
			<input type="text" name="longname" value="<%= longname %>" size="64" maxlength="64">
			<input type="hidden" name="status" value="ACTIVE">
	</table>
	<p>
	<input type="submit" value="Submit">	
	<p>
	<font color="#0000FF">The short name is an abbreviation for the list member.  It is converted to<br>
	uppercase and is used on the report page to reduce column width.  Eg 'CH'.<br>
	The long name is the full name for the list member.  It is stored as entered and<br>
	used throughout the program with the short name.  Eg 'City Hall'.  Neither can be blank.<br>
	After entering a new list member, run Sort Order to order the new member in the list.<br>
	Quit and Login to activate new or updated list members.</font>   			

	
<%	Case "EDIT" %>	
	
	<strong>Update Reference List <font color="#008000"><%= reflist %></font> Member <font color="#008000"><%= listmember %></font></strong>
	<p>
	<input type="hidden" name="listmember" value="<%= listmember %>">		
	<table width="100%" border="0" cellspacing="0" cellpadding="3">	
		<tr>
			<td width="35%" align="right">Long Name
			<td width="65%">
			<input type="text" name="longname" value="<%= longname %>" size="64" maxlength="64">
		<tr>
			<td align="right">Status
			<td>			
			<select name="status">
				<option value="ACTIVE">Active
				<option value="INACTIVE">Inactive
			</select>
	</table>
	<p>
	<input type="submit" value="Submit">	
	<p>
	<font color="#0000FF">You may change the long name or the staus for a list member.  You may<br>
	make a list member inactive, but you cannot remove a list member from the database.<br></font>
	<font color="#FF0000">Warning!</font>  <font color="#0000FF">An inactive list member will not be available for use with any new entries<br>
	to the database and the long name is changed for all occurences in the database.<br>
	You may re-activate a list member at any time.<br>
	Quit and Login to activate new or updated list members.</font>
	
<%	Case "SORT"
		SQL = "SELECT shortname,longname,status,sortorder FROM reflist WHERE list = '" & reflist & "' ORDER BY status,sortorder,shortname"
%>
<!--#Include File="db_select.asp"-->

	<p><strong>Set the Sort Order for Reference List <font color="#008000"><%=  reflist %></font></strong>
	<p>	
	<table width="100%" border="0" cellspacing="0" cellpadding="3">	
		<tr>
			<td width="50%" align="right">Sort Order
			<td width="50%">&nbsp;<%= reflist %>			

<%		Do Until rs.eof %>

		<tr>
			<td align="right">
			<input type="text" name="<%= "sort" & rs("shortname") %>" value="<%= rs("sortorder") %>" size="2" maxlength="2">
			<td>&nbsp;<%= rs("shortname") %>&nbsp;&nbsp;<%= rs("longname") %>
			
	
<%			rs.movenext
		Loop %>


	</table>
	<p>	
	<input type="submit" value="Submit">	
	<p>
	<font color="#0000FF">Sort order is used to order list members on select lists and reports.<br>
	Number the list members in the order you want to see them.  To order alphabetically,<br>
	use blanks (or, all the same number).  If mixed, blanks will order before numbers<br>
	Quit and Login to activate new or updated list members.</font>
	
<%	
End Select %>	
	 
</form>

</div>
</body>
</html>
