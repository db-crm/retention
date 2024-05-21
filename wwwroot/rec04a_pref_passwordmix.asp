<%
thispage = "pref_passwordmix"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "PASSWORDMIX"	
	itemvalue = "'" & Request("passwordmix") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Password Case Sensitivity</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Password Case Sensitivity</strong>	

<form method="POST" action="<%= version %>_pref_passwordmix.asp">
	<p>
	<select name="passwordmix">
		<option value="TRUE">Passwords are Mixed Case
		<option value="FALSE">Passwords are Not Case Sensitive
	</select>
	<p>
	<input type="submit" value="Change Password Case Sensitivity">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">Change to Mixed Case to require an exact case match for passwords.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
