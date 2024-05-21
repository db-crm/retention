<%
thispage = "pref_boxid"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "USEBOXID"	
	itemvalue = "'" & Request("useboxid") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Use of Box Identifier</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>User Assigned Box Identifier</strong>
	
<form method="POST" action="<%= version %>_pref_boxid.asp">
	<p>
	<select name="useboxid">
		<option value="TRUE">Allow Assignment of Box Identifiers
		<option value="FALSE">Only Use Box Numbers Assigned by the Program 
	</select>
	<p>
	<input type="submit" value="Submit">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">The program automatically assigns a box number when the record is added to the database. You may assign an additional, uncontrolled, user defined box identifier by selecting 'Allow Assignment of Box Identifiers'. The box identifier field is added to the New Record entry page and the program places no constraints on the content of the field. If entered, the box identifier will be displayed as a suffix to the box number. After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
