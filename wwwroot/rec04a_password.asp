<%
thispage = "password"
pageaccess = "USERSUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
usersite = Request("usersite")
password = Trim(Request("password"))
username = Session("username")
site = Session("site")
userlength = CInt(Session("userlength")) %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change Password</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<%
'If the submitted password isn't blank, update the database and session, notify of the change, and go to the menu or quit.  If not, or first time through, do again.

Select Case Request("action")
	Case "Change Password"
		If Len(password) >= userlength Then
			SQL = "UPDATE userinfo SET userinfo.password = '" & password & "' WHERE username = '" & username & "'"
%>
<!--#Include File="db_execute.asp"-->	
<%
  		  Session("password") = password %> 		
	
<p><strong>Your new Password is <font color="Blue"><%= password %></font></strong>

<%		Else %>

<p><strong><font color="#FF0000"><%= password %> is Invalid</font></strong>
		
<%		End If

	Case "Change Default Site"
		SQL = "UPDATE userinfo SET site = '" & usersite & "' WHERE username = '" & username & "'"
%>
<!--#Include File="db_execute.asp"-->	
<%
    Session("site") = usersite %>
	
<p><strong>Your new Default Site is <font color="Blue"><%= usersite %></font></strong>	
	
<%	
End Select %>

<form method="POST" action="<%= version %>_password.asp">
	<p><strong>Enter a New Password of at least <%= userlength %> Characters</strong>
	
<%
If Session("passwordmix") = "TRUE" Then %>

	<p><strong>Passwords are case sensitive</strong>

<%
End If %>
	 
	<p>Password&nbsp;
	<input type="text" name="password" size="16" maxlength="16">
	<p>
	<input type="submit" name="action" value="Change Password">
		
<%
If Session("showsite") <> "NEVER" Then %>

	<p>Default Storage Site&nbsp;		
	<select name="usersite"> 

<%	SQL = "SELECT shortname,longname,sortorder FROM reflist WHERE list = 'SITE' AND status = 'ACTIVE' ORDER BY sortorder"
%>
<!--#Include File="db_select.asp"-->
<%
	Do Until rs.eof %> 
								
		<option value="<%= rs("shortname") %>"><%= rs("shortname") & "&nbsp;&nbsp;" & rs("longname") %>	
			
<%		rs.movenext
	Loop %>

	</select>
	<p>
	<input type="submit" name="action" value="Change Default Site">
		
<%
End If
%>	

</form>
	
<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/password_help.asp"--></font>
<%
End If
%>

</div>
</body>
</html>
