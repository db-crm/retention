<%
thispage = "login"
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
If lastpage <> thispage Then 'First time thru, set session preference items and location and owner labels

	Session("help") = "FALSE" 'Initial help display state is off
	Session("content_size") = "full" 'Initial content size is full

	SQL = "SELECT item,itemvalue,itemtype FROM preference" 
%>
<!--#Include File="db_select.asp"-->	
<%
	Do While Not rs.eof 'Set the value of the session items defined in preference
		str = "Session(" & Chr(34) & rs("item") & Chr(34) & ") = " & Chr(34) & rs("itemvalue") & Chr(34)
		Execute str
		rs.movenext
	Loop

	Select Case Session("database") 'Set database specific items
		Case "JET"
			Session("datedelim") = "#"
		Case Else
			Session("datedelim") = "'"
	End Select

	arybackground = Split(Session("background"),",") 'Parse background into red, green, and blue decimal values and convert to hex for bgcolor of body attribute. Session(background) id decimal and Session(bgcolor) is hex.
	Session("bgcolor") = "#" & Hex(arybackground(0)) & Hex(arybackground(1)) & Hex(arybackground(2))

	'Get the owner labels, initial owner name is blank.
	Session("ownername") = ""
	SQL = "SELECT ownername,ownerlong FROM ownerinfo WHERE status = 'LABEL'" 
%>
<!--#Include File="db_select.asp"-->	
<%
	If Not rs.eof Then
		Session("ownerlabel") = rs("ownername")
		Session("ownerlabellong") = rs("ownerlong")
	End If

	'Get the default site
	SQL = "SELECT shortname,longname FROM reflist WHERE list = 'SITE' AND status = 'ACTIVE'"
%>
<!--#Include File="db_select.asp"-->	
<%
	If Not rs.eof Then 'Pick the first site which should be the only site if a single site client
		Session("site") = rs("shortname")
		Session("sitelong") = rs("longname")
	End If

	'Get the disposition values and make arrays for disposition short name and long name, exclude active
	SQL = "SELECT shortname,longname FROM reflist WHERE list = 'DISPOSITION' AND shortname <> 'ACTIVE' AND status = 'ACTIVE' ORDER BY sortorder" 
%>
<!--#Include File="db_select.asp"-->	
<%
	If Not rs.eof Then
		dispolistlength = Len(rs("shortname"))
		dispolist = rs("shortname")
		dispolistlong = rs("longname")
		rs.movenext
		Do While Not rs.eof
			If Len(rs("shortname")) > dispolistlength Then
				dispolistlength = Len(rs("shortname"))
			End If 
			dispolist = dispolist & "," & rs("shortname")
			dispolistlong = dispolistlong & "," & rs("longname")
			rs.movenext
		Loop
		Session("dispolist") = dispolist
		Session("dispolistlong") = dispolistlong
		Session("dispolistlength") = dispolistlength
	End If
	
Else 'If not the first time, process for a valid user.  If yes, set session variables and go to the Menu

	username = Trim(Request("username"))
	password = Trim(Request("password"))

	SQL = "SELECT username,accesstype,password,site FROM userinfo WHERE username = '" & username & "' AND password = '" & password & "'"
%>
<!--#Include File="db_select.asp"-->	
<%
	If Not rs.eof Then 'Got a user, check for password exact match if mixcase set
		If  Session("passwordmix") = "TRUE" And password <> rs("password") Then
			validuser = "FALSE"
		Else 'Valid user, set session info and go to the Menu.
			Session("site") = rs("site")
			Session("accesstype") = UCase(rs("accesstype"))
			If Session("accesstype") = "GROUP" Then	Session("showsite") = "TRUE"
			Session("username") = rs("username")
			Session("password") = rs("password")
			Session.Timeout = 20
			Session("start_time") = Timer
			Response.Clear
			Server.Transfer version & "_menu.asp"
		End If
	Else
		validuser = "FALSE"
	End If
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Log In</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->
<%
If validuser = "FALSE" Then %>

<p><font color="Red"><strong>The submitted Username and Password are not a valid Login</strong></font>

<%	If Session("passwordmix") = "TRUE" Then %>

<p><font color="Red"><strong>Passwords are case sensitive</strong></font>

<%	End If
End If %> 

<form method="POST" action="<%= version %>_login.asp">
	<p>
	<table width="100%" border="0" cellspacing="0" cellpadding="2">
		<tr>
		   <td width="50%" align="right"><strong>Username</strong>
		   <td width="50%">
		   <input type="text" name="username" value="<%=username%>" size="16" maxlength="16">
		<tr>
		   <td align="right"><strong>Password</strong>
		   <td>
		   <input type="password" name="password" size="16" maxlength="16">
	</table>
	<p>	
	<input type="submit" value="Login">
</form>

<!--#Include File="footer.asp"-->

</div>
</body>
</html>
