<%
thispage = "cleanup_update"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
datedelim = Session("datedelim")
recordno = Session("recordno")
recorddate = Request("recorddate")
retention = Request("retention")
content = Replace(Request("content"),"'","''")
boxid = Request("boxid")

histaction = "CLEANUP" 'Make an entry in the history table
%>
<!--#Include File="cd_history.asp"-->	
<%

'Make the Review date always the first of the month and ten years from now if retention is 100
If retention = 100 Then
	reviewdate = DatePart("m",Date) & "/1/" & (DatePart("yyyy",Date) + 10)
Else
	reviewdate = DatePart("m",recorddate) & "/1/" & (DatePart("yyyy",recorddate) + retention)
End If

'Update the Record.
SQL = "UPDATE record SET recorddate = " & datedelim & recorddate & datedelim & ",reviewdate = " & datedelim & reviewdate & datedelim & ",retention = " & retention & ",boxid = '" & boxid & "',content = '" & content & "',updatedate = " & datedelim & Date & datedelim & "  WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_execute.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Update a Record</title>	
</head>
<body>
<div align="center">
	
<p><a href="<%= version %>_cleanup.asp" title="Return to Edit a Record">The record has been edited in the database.</a>&nbsp;&nbsp;<strong>Record Number&nbsp;<%= recordno %>&nbsp;Updated&nbsp;<%= Date %></strong>

<p>
<!--#Include File="cd_box_label.asp"-->

</div>
</body>
</html>
