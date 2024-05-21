<%
thispage = "add_insert"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
ownername = Request("ownername")
Session("ownername") = ownername
datedelim = Session("datedelim")
site = Session("site")
sitelong =Session("sitelong")
location = Request("location")
recorddate = Request("recorddate")
retention = CInt(Request("retention"))
content = Replace(Request("content"),"'","''")
boxid = Trim(Request("boxid"))

'Get the owner number and long name.
%>
<!--#Include File="cd_ownerno.asp"-->
<%

'Make the Review date always the first of the month and ten years if retention is 100
If retention = 100 Then
	retyears = 10
Else
	retyears = retention
End If
reviewdate = DatePart("m",recorddate) & "/1/" & (DatePart("yyyy",recorddate) + retyears)

'See if there are any records for the owner.
SQL = "SELECT count(*) as boxcount FROM record WHERE ownerno = " & ownerno & ""
%>
<!--#Include File="db_select.asp"-->	
<%
If rs("boxcount") > 0 Then 'Get the box number for the new record.
	SQL = "SELECT max(box) as maxbox FROM record WHERE ownerno = " & ownerno & ""
%>
<!--#Include File="db_select.asp"-->	
<%
	box = rs("maxbox") + 1
Else 'If no boxes for the owner, box = 1.
	box = 1
End If

'Get the new record number.
SQL = "SELECT max(recordno) as maxrecord FROM record"
%>
<!--#Include File="db_select.asp"-->	
<%
recordno = rs("maxrecord") + 1
'Insert into Record.
SQL = "INSERT into record(recordno,ownerno,box,boxid,content,retention,status,disposition,entrydate,updatedate,recorddate,reviewdate,deletedate) values(" & recordno & "," & ownerno & "," & box & ",'" & boxid & "','" & content & "'," & retention & ",'ACTIVE','ACTIVE'," & datedelim & Date & datedelim & ",NULL," & datedelim &  recorddate & datedelim & "," & datedelim & reviewdate & datedelim & ",NULL)" 
%>
<!--#Include File="db_execute.asp"-->

<%
'Update the vacant location with the new record number.
SQL = "UPDATE location SET recordno = " & recordno & " WHERE site = '" & site & "' AND location = '" & location & "'"
%>
<!--#Include File="db_execute.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Add a New Record</title>	
</head>
<body>
<div align="center">
	
<p><a href="<%= version %>_add.asp" title="Return to Add a Record">The record has been added to the database.</a>&nbsp;&nbsp;<strong>Record Number&nbsp;<%= recordno %>&nbsp;Created&nbsp;<%= Date %></strong>

<p>
<!--#Include File="cd_box_label.asp"-->

</div>
</body>
</html>
