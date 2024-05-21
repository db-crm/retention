<%
thispage = "pref_site"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "SHOWSITE"	
	itemvalue = "'" & Request("showsite") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Show Storage Site</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Show Storage Sites</strong>
	
<form method="POST" action="<%= version %>_pref_site.asp">
	<p>
	<select name="showsite">
		<option value="TRUE">Show the Storage Sites
		<option value="FALSE">Hide the Storage Sites
		<option value="NEVER">Single Site Only
	</select>
	<p>
	<input type="submit" value="Change Show Storage Sites">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">Select Show to set default to show all sites and Hide to hide all sites except the user default site.  Users may also change the default site from the menu during their session.  Select Single if you have only one storage site.  If the sites are hidden, they cannot be managed by the program. After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
