<%
thispage = "newuser_edit"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->
<%
errormsg = ""
actionmsg = ""
userlength = CInt(Session("userlength"))
If Session("passwordmix") = "TRUE" Then
	mixmsg = "Case Sensitive"
Else
	mixmsg = "Not Case Sensitive"
End If

If lastpage = thispage Then
	newuser = UCase(Trim(Request("newuser")))
	accesstype = Request("accesstype")
	usersite = Request("usersite")
	password = Trim(Request("password"))
	realname = Trim(Request("realname"))
	organization = Trim(Request("organization"))
	email = Trim(Request("email"))
	usercontact = Trim(Request("usercontact"))
	'Check for valid username and password, short existing values are ok
	If (newuser = Session("olduser") Or Len(newuser) >= userlength) And (password = Session("oldpassword") Or Len(password) >= userlength) Then
		'Check for existing user
		SQL = "SELECT username FROM userinfo WHERE username = '" & newuser & "'"
%>
<!--#Include File="db_select.asp"-->	
<%
		If rs.eof Then 'Newuser does not exist
			userexists = "FALSE"
			exuser = ""
		Else
			userexists = "TRUE"
			exuser = rs("username")
		End If	
		Select Case Session("newuser_action")
			Case "ADD"
				If userexists = "FALSE" Then 'No duplicate, do the add
					SQL = "INSERT INTO userinfo(username,password,accesstype,site,realname,organization,email,usercontact) VALUES('" & newuser & "','" & password & "','" & accesstype & "','" & usersite & "','" & realname & "','" & organization & "','" & email & "','" & usercontact & "')"
%>
<!--#Include File="db_execute.asp"-->	
<%
					Session("newusermsg") = newuser & " has been added"
					Response.Clear
					Server.Transfer version & "_newuser.asp" 
				Else
					errormsg = "User " & newuser & "already exists"
				End If
			Case "EDIT"
				If exuser = Session("olduser") Or userexists = "FALSE" Then 'No duplicate or editing the selected name, do the update
					SQL = "UPDATE userinfo SET username = '" & newuser & "',userinfo.password = '" & password & "',accesstype = '" & accesstype & "',realname = '" & realname & "',organization = '" & organization & "',email = '" & email & "',usercontact = '" & usercontact & "'  WHERE username = '" & Session("olduser") & "'"
%>
<!--#Include File="db_execute.asp"-->	
<%
					Session("newusermsg") = newuser & " has been edited"
					Response.Clear
					Server.Transfer version & "_newuser.asp"
				Else
					errormsg = "Changing username to " & newuser & " would replace another existing user."
				End If
		End Select
	Else
		errormsg = "New usernames and passwords must be " & userlength & " or more characters"
	End If
ElseIf Session("newuser_action") = "EDIT" Then 'First time through for Edit action
	newuser = Session("newuser")
	Session("olduser") = newuser 'Remember old username, even if edited
	SQL = "SELECT password,accesstype,site,realname,organization,email,usercontact FROM userinfo WHERE username = '" & newuser & "'"
%>
<!--#Include File="db_select.asp"-->	
<%
	If Not rs.eof Then
		password = rs("password")
		Session("oldpassword") = password 'Remember old password
		accesstype = rs("accesstype")
		usersite = rs("site")
		realname = rs("realname")
		organization = rs("organization")
		email = rs("email")
		usercontact = rs("usercontact")
	End If
Else 'First time through for Add action
	newuser = ""
	Session("olduser") = ""
End If
str = accesstype & "SEL = " & Chr(34) & "selected" & Chr(34) 'Make selected accesstype 'selected'
Execute str %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Perform a User Addition or Edit</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->
	
<p><font color="#008000"><%= actionmsg %></font>
<p><font color="#FF0000"><%= errormsg %></font>
<form method="POST" action="<%= version %>_newuser_edit.asp">
	<p>
	<table width="640" border="0" cellspacing="0" cellpadding="3">
		<tr>
			<td>		
			<td align="left">
			<strong><%= Session("newuser_action") %> </strong>
		<tr>
			<td align="right">		
			User Name
			<td align="left">
			<input type="text" name="newuser" value="<%= newuser %>" size="16" maxlength="16">&nbsp;
			New Usernames Minimum <%= userlength %> Characters
		<tr>
			<td align="right">		
			Password
			<td align="left">
			<input type="text" name="password" value="<%= password %>" size="16" maxlength="16">&nbsp;
			New Passwords Minimum <%= userlength %> Characters, <%= mixmsg %>
		<tr>	
			<td align="right">
			Access Type
			<td>
			<select name="accesstype">
			
<%
If Session("olduser") <> Session("username") Then 'ADMINs cannot change their own accesstype %>
	
				<option <%= GROUPSEL %> >GROUP
				<option <%= USERSEL %> >USER
				<option <%= SUPERSEL %> >SUPER
				
<%
End If %>
				
				<option <%= ADMINSEL %> >ADMIN
			</select>

		<tr>	
			<td align="right">
			Default Site
			<td>
			<select name="usersite"> 

<%
SQL = "SELECT shortname,longname,sortorder FROM reflist WHERE list = 'SITE' AND status = 'ACTIVE' ORDER BY sortorder"
%>
<!--#Include File="db_select.asp"-->
<%
Do Until rs.eof
	If usersite = rs("shortname") Then
		selected = "selected"
	Else
		selected = ""
	End If %> 
								
				<option value="<%= rs("shortname") %>" <%= selected %> ><%= rs("shortname") %>&nbsp;&nbsp;<%= rs("longname") %>	
			
<%	rs.movenext
Loop %>

			</select>			
		<tr>	
			<td align="right">
			Real Name
			<td>
			<input type="text" name="realname" value="<%= realname %>" esize="32" maxlength="32">
		<tr>	
			<td align="right">
			Organization
			<td>
			<input type="text" name="organization" value="<%= organization %>" esize="32" maxlength="32">			
		<tr>	
			<td align="right">
			E-Mail
			<td>
			<input type="text" name="email" value="<%= email %>" esize="48" maxlength="48">			
		<tr>	
			<td align="right">
			Contact
			<td>
			<input type="text" name="usercontact" value="<%= usercontact %>" size="64" maxlength="64">

		<tr>	
			<td align="right">
		<tr>	
			<td>
			<td>
			<input type="submit" value="Submit">								
	</table
</form>

<p>
<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">Username is not case sensitive.   Username is required for all actions and new usernames must be between the minimum (set by Preferences) and 16 characters.</font> 
	<tr>
		<td><font color="#0000FF">Password may be case sensitive (set by Preferences) and has the same length constraint as username.</font>
		<td>&nbsp;
	<tr>
		<td>&nbsp;	
	<tr>
		<td><font color="#0000FF">Select an Access Type when Adding a user or Changing the Access Type.<br>
 A GROUP user has only one password which cannot be changed by a user.  GROUP actions are limited to viewing Reports, viewing the Reference Lists, and viewing Help.  The GROUP type is intended to be used (where security is not an issue) to allow groups of users to share a common username and password for viewing reports.<br>
A USER can perform GROUP actions and can change their own password.  The USER type is intended for viewing reports by individuals using password security.<br>
A SUPER can perform the USER actions and perform all the actions required to operate the Records Retention SETUP.<br>
An ADMIN can perform SUPER actions and perform actions required to manage users, add new sites and locations, and change preferences and reference lists.<br>
A username cannot belong to more than one access type.  If a single person is to have more than one access type, such as SUPER for operations and ADMIN for administration, they must have two usernames.<br>
An ADMIN cannot remove or change the accesstype for their own ID.  This avoids being locked out of the system.  Create or use another ADMIN to remove or change the accesstype.</font>
	<tr>
		<td>&nbsp;
	<tr>	
		<td><font color="#0000FF">The default site is the site most often used by the user if there are multiple sites.  All types except GROUP may change the default site for the current session using the menu.  This page must be used to permanently change the default site for a user.</font>
	<tr>
		<td>&nbsp;
	<tr>
		<td><font color="#0000FF">The program imposes no formatting on the Real Name, Organization, EMail or Contact information.  It is recommended that the customer implement an entry format that will allow these items to order consistently if selected using queries against the underlying database.</font>   
</table>

</div>
</body>
</html>
