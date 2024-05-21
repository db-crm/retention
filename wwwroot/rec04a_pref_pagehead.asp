<%
thispage = "pref_pagehead"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "GRAPHICHEAD"	
	itemvalue = "'" & Request("graphichead") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Page Header</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Page Header</strong>
	
<form method="POST" action="<%= version %>_pref_pagehead.asp">
	<p>
	<select name="graphichead">
		<option value="TRUE">Image in Page Header
		<option value="FALSE">Text Only Page Header
	</select>
	<p>
	<input type="submit" value="Change Page Header">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">Change the page header to Text Only to remove the image from the header.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
