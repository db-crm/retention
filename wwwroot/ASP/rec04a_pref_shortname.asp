<% thispage = "pref_shortname" %>

<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "SHORTNAMEUPPER"
	itemvalue = "'" & Request("shortnameupper") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Short Name Case</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">

<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Owner and Location Short Name Case</strong>	

<form method="POST" action="<%= version %>_pref_shortname.asp">

	<select name="shortnameupper">
		<option value="TRUE">Upper Case
		<option value="FALSE">Mixed Case
	</select>
	
	<p><input type="submit" value="Change Short Name Case">

</form>

<p><font color="Blue">Change the short name for Owner to Upper Case (eg DEPT/DIV)<br>
or Mixed Case (eg Dept/Div).<br>
After changing a preference, you must Login again.</font>
		
</div>
</body>
</html>
