<%
thispage = "pref_box"
pageaccess = "ADMINSYSTEM"
%>
<!--#Include File="version.asp"-->

<%
If thispage = lastpage Then
	item = "AUTOBOX"	
	itemvalue = "'" & Request("autobox") & "'"
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Automatic Box Number Assignment</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Automatic Box Number Assignment</strong>
	
<form method="POST" action="<%= version %>_pref_box.asp">
	<p>
	<select name="autobox">
		<option value="TRUE">Box Numbers are Assigned Only by the Program
		<option value="FALSE">Box Numbers May be Assigned by the Operator
	</select>
	<p>
	<input type="submit" value="Submit">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">When box numbers are assigned only by the program, the numbers are assigned sequentially and the operator has no control over the numbers.  If you want to assign numbers according to another scheme, select the option to let the operator assign the number.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
