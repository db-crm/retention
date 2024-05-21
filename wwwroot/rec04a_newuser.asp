<%
thispage = "newuser"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->
<%
errormsg = ""
actionmsg = Session("newusermsg")
Session("newusermsg") = ""

If lastpage = thispage Then
	newuser = Request("newuser")
	Select Case Request("action")
		Case "Add"
			Session("newuser_action") = "ADD"
			Response.Clear
			Server.Transfer version & "_newuser_edit.asp"
		Case "Edit"
			Session("newuser") = newuser
			Session("newuser_action") = "EDIT"
			Response.Clear
			Server.Transfer version & "_newuser_edit.asp" 				
		Case "Remove"
			If Not newuser = Session("username") Then 'Can't remove yourself
				SQL = "DELETE FROM userinfo WHERE username = '" & newuser & "'"
%>
<!--#Include File="db_execute.asp"-->
<%
				actionmsg = "User " & newuser & " removed"
			Else 
				errormsg = "You cannot remove yourself"
			End If
		Case "Show Users"
			Response.Clear
			Server.Transfer version & "_newuser_list.asp"
	End Select
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Maintain the User Info Table</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Manage Users</strong>
<p><font color="#008000"><%= actionmsg %></font>
<p><font color="#FF0000"><%= errormsg %></font>	
<form method="POST" action="<%= version %>_newuser.asp">
	<p>
	<select name="newuser">

<%
SQL = "SELECT username,realname FROM userinfo ORDER BY username"
%>
<!--#Include File="db_select.asp"-->	
<%
Do Until rs.eof %>

		<option value = "<%= rs("username") %>"><%= rs("username") %>, <%= rs("realname") %>

<%	rs.movenext
Loop %>

	</select>
	<p>
	<input type="submit" name="action" value="Add">&nbsp;
	<input type="submit" name="action" value="Edit">&nbsp;
	<input type="submit" name="action" value="Show Users">
	<p>
	<p><font color="#FF0000">Clicking Remove will permanently remove the user from the database.</font>
	<p>	
	<input type="submit" name="action" value="Remove">		
</form>

</div>
</body>
</html>
