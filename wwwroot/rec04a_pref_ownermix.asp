<%
thispage = "pref_ownermix"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "OWNERMIX"	
	itemvalue = "'" & Request("ownermix") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Owner Short Name Case</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Owner Short Name Case</strong>	

<form method="POST" action="<%= version %>_pref_ownermix.asp">
	<p>
	<select name="ownermix">
		<option value="TRUE">Owner Short Names are Mixed Case
		<option value="FALSE">Owner Short Names are Upper Case
	</select>
	<p>
	<input type="submit" value="Change Owner Short name Case">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">Change to Upper to force the Owner Short Name to upper case.  Mixed case will store the Owner Short Name exactly as entered when added or edited.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
