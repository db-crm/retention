<%
thispage = "cleanup_select"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
errormsg = ""
ownername = Session("ownername")
ownerno = Session("ownerno")

If thispage = lastpage Then 'Get the record number and boxid from ownerno and box
	box = Request("box")
	SQL = "SELECT recordno,boxid FROM record WHERE ownerno = " & ownerno & " AND box = " & box & ""
%>
<!--#Include File="db_select.asp"-->	
<%
	If Not rs.eof Then
		recordno = rs("recordno")
		boxid = rs("boxid")
	End If
	If Request("action") = "Remove" Then 'Put the record in recordhistory, delete from record, set recordno in location to zero, and return to cleanup
		histaction = "CLEANUP"
%>
<!--#Include File="cd_history.asp"-->	
<%
		SQL = "DELETE FROM record WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_execute.asp"-->	
<%
		SQL = "UPDATE location SET recordno = 0 WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_execute.asp"-->	
<%
		Session("cleanupmsg") = "Owner " & ownername & ", Box " & box & " " & boxid & " has been Deleted"
		Response.Clear
		Server.Transfer version & "_cleanup.asp"		
	Else 'Go to the edit page
		Session("recordno") = recordno
		Session("box") = box
		Response.Clear
		Server.Transfer version & "_cleanup_edit.asp"
	End If
Else 'Get the ownername from the previous page and set session ownername and ownerno
	ownername = Request("ownername")
%>
<!--#Include File="cd_ownerno.asp"-->
<%
	Session("ownername") = ownername
	Session("ownerno") = ownerno
End If %>	

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Select a Record to Edit or Remove</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<p><strong>Edit or Remove a Record</strong>
<p><font color="#FF0000"><%= errormsg %></font>
<p>
<form method="POST" action="<%= version %>_cleanup_select.asp">
	
<%'Make the Box select list.
SQL="SELECT recordno,box FROM record WHERE ownerno = " & ownerno & " AND status = 'ACTIVE' ORDER BY box"
%>
<!--#Include File="db_select.asp"-->	

	Owner&nbsp;<%= ownername %>&nbsp;Box&nbsp;
	<select name="box">
					
<%
Do While Not rs.eof %>
	
		<option><%= rs("box")%>		
		
<%	rs.movenext
Loop %>
	
	</select>
	<p>
	<input type="submit" name="action" value="Edit">
	<p>
	<p><font color="#FF0000">Clicking Remove will permanently remove the record from the database.</font>
	<p>
	<input type="submit" name="action" value="Remove">
</form>

</div>
</body>
</html>
