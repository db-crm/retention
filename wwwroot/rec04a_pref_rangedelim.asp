<%
thispage = "pref_rangedelim"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->
<%

If thispage = lastpage Then
	item = "RANGEDELIM"	
	If Request("action") = "Reset to Default" Then
		itemvalue = "itemdefault"
	Else
		itemvalue = "'" & Request("rangedelim") & "'"
	End If
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
Else
	SQL = "SELECT itemvalue,itemdefault FROM preference WHERE item = 'RANGEDELIM'"
%>	
<!--#Include File="db_select.asp"-->	
<%
	rangedelim = rs("itemvalue")
	defdelim = rs("itemdefault")
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Range Delimiter</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Range Delimiter</strong>	
<p>The current delimiter is <%= rangedelim %>.  The Default delimiter is <%= defdelim %>.
<form method="POST" action="<%= version %>_pref_rangedelim.asp">
	<p>

<%
selectname = "rangedelim"
aryoptionvalue = Array(">","<","^","&","*","%","$","@","!","?","#",":",";")
aryoptionlabel = Array(">","<","^","&","*","%","$","@","!","?","#",":",";")
%>
<!--#Include File="cd_selectarray.asp"-->

	<p>
	<input type="submit" name="action" value="Change Range Delimiter">
	<p>
	<input type="submit" name="action" value="Reset to Default">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">The range delimiter is a single character used to identify or separate a numeric or date range on the report query page.  The program limits selection of the range delimiter to characters which would not occur in common date formats as determined by the script interpreter.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
