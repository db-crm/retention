<% thispage = "pref_shortlabel" %>

<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "SHORTLABELUPPER"
	itemvalue = "'" & Request("shortlabelupper") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Short Name Label Case</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">

<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Owner and Location Short Name Label Case</strong>	

<form method="POST" action="<%= version %>_pref_shortlabel.asp">

	<select name="shortlabelupper">
		<option value="TRUE">Upper Case
		<option value="FALSE">Mixed Case
	</select>
	
	<p><input type="submit" value="Change Short Label Case">

</form>

<p><font color="Blue">Change the short name label for Owner and Location to Upper Case (eg DEPT/DIV)<br>
or Mixed Case (eg Dept/Div).  The short name label is only used on page<br>
and report headers.  The short name (not the label) is always upper case.<br>
After changing a preference, you must Login again.</font>
		
</div>
</body>
</html>
