<%
thispage = "pref_userlength"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->
<%
SQL = "SELECT itemvalue FROM preference WHERE item = 'userlength'"
%>	
<!--#Include File="db_select.asp"-->	
<%
userlength = CInt(rs("itemvalue")) %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change User and Password Minimum Length</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Username and Password Length</strong>	

<form method="POST" action="<%= version %>_pref_userlength.asp">

<%
If thispage = lastpage Then
	userlength = CInt(Request("userlength"))
	item = "USERLENGTH"	
	itemvalue = "'" & userlength & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

	<p>
	<select name="userlength">
<%
i = 4
Do Until i > 16
	If i = userlength Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
		<option <%=selected%>><%= i %>
	
<%	i = i + 1
Loop %>
 	
	</select>
	<p>
	<input type="submit" value="Change the Username and Password Minimum Length">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">Enter the minimum length for the username and password.  The maximum length is 16 characters.  Existing usernames and passwords are not effected.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
