<%
thispage = "summary"
pageaccess = "SUPERADMIN"
errors = 0
delim = Session("rangedelim")
%>
<!--#Include File="version.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Summary Counts</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<%
If lastpage = thispage Then
	rsummarydate = "Y"
	fieldname = "summarydate"
	fieldtype =  "date"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	If summarydate_error = "N" Then 'No errors, display the results %>
	
<p><strong>Records Retention Activity Summary for <%= Request("summarydate") %></strong>
	
<%		qdate = Replace(qsummarydate,"summarydate","entrydate")
		SQL = "SELECT COUNT(*) AS entrycount FROM record WHERE 1=1" & qdate & ""
%>	
<!--#Include File="db_select.asp"-->	
		
<p><strong><%= rs("entrycount") %>&nbsp;New Records</strong>

<%		qdate = Replace(qsummarydate,"summarydate","updatedate")		
	 	SQL = "SELECT COUNT(*) AS updatecount FROM record WHERE 1=1" & qdate & ""
%>
<!--#Include File="db_select.asp"-->	
		
<p><strong><%= rs("updatecount") %>&nbsp;Updated Records</strong>

<%		qdate = Replace(qsummarydate,"summarydate","deletedate")		
		SQL = "SELECT COUNT(*) AS deletecount FROM record WHERE 1=1" & qdate & ""
%>
<!--#Include File="db_select.asp"-->	
		
		<p><strong><%= rs("deletecount") %>&nbsp;Deleted Records</strong>
		
<%	Else %>	

<p><font color="#FF0000"><strong>The entered date range is invalid</strong></font>

<%	End If
End If

If (lastpage <> thispage) Or (errors > 0) Then %>

<p><strong>Summary Count</strong>

<form method="POST" action="<%= version %>_summary.asp">
	<p>
	<table width=30% border="0" cellspacing="0" cellpadding="2">
	
<%
rsummarydate = "Y"
fieldname = "summarydate"
fieldlabel = "Date Range"
fieldsize = 21
%>
<!--#Include File="cd_field_entry.asp"-->
	
		<input type="submit" value="Submit">
	</table>
</form>

<p>Use '<%= Session("rangedelim") %>' to make the range

<%
	If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/summary_help.asp"--></font>
<%
	End If
End If %>	
		
</div>
</body>
</html>
