<%
thispage = "pref_reporthead"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<% errors = 0
If thispage <> lastpage Then 'Get the current header spacings
	SQL = "SELECT itemvalue FROM preference WHERE item = 'longheadcount'"
%>	
<!--#Include File="db_select.asp"-->	
<%
	longheadcount = CInt(rs("itemvalue"))
	SQL = "SELECT itemvalue FROM preference WHERE item = 'shortheadcount'"
%>	
<!--#Include File="db_select.asp"-->	
<%
	shortheadcount = CInt(rs("itemvalue"))
	
Else 'Update the header spacings if Reset requested or no errors
	If Request("action") = "Reset to Default" Then
		longheadcount = "itemdefault"
		shortheadcount = "itemdefault"
	Else
		longheadcount = Trim(Request("longheadcount"))
		shortheadcount = Trim(Request("shortheadcount"))
		If IsNumeric(longheadcount) And IsNumeric(shortheadcount) And longheadcount > 0 And shortheadcount > 0 Then
		longheadcount = "'" & longheadcount & "'"
		shortheadcount = "'" & shortheadcount & "'"	
		Else
			errors = errors + 1
		End If
	End If
	
	If errors = 0 Then 'Update the data header spacings
		SQL = "UPDATE preference SET itemvalue = " & longheadcount & " where item = 'LONGHEADCOUNT'"
%>
<!--#Include File="db_execute.asp"-->	
<%
		SQL = "UPDATE preference SET itemvalue = " & shortheadcount & " where item = 'SHORTHEADCOUNT'"
%>
<!--#Include File="db_execute.asp"-->	
<%
		Response.Clear
		Server.Transfer version & "_login.asp"
	End If
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change Report Header Spacings</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Report Header Spacing</strong>

<%
If errors > 0 Then %>

<font color="#FF0000"><p>Header Spacings Must be Positive Integers</font>

<%
End If %>	

<form method="POST" action="<%= version %>_pref_reporthead.asp">
	<p>Number of report lines between headers for full content&nbsp;	
	<input type="text" name="shortheadcount" value=<%= shortheadcount %> size="3" maxlength="3">
	<p>Number of report lines between headers for short or no content&nbsp;
	<input type="text" name="longheadcount" value=<%= longheadcount %> size="3" maxlength="3">
	<p>
	<input type="submit" name="action" value="Change the Report Header Spacings">
	<p>
	<input type="submit" name="action" value="Reset to Default">	
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">The report header spacings are the number of lines between the rolling header on the report page.  Long spacing is used when short (truncated) content is selected. Short spacing is used when full content is selected.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
