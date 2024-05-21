<%
thispage = "cleanup"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
actionmsg = Session("cleanupmsg")
Session("cleanupmsg") = "" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Edit or Remove a Record</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<p><strong>Edit or Remove a Record</strong>
<p><font color="#008000"><%= actionmsg %></font>
<p>
<form method="POST" action="<%= version %>_cleanup_select.asp">

<%'Make the Owner select list
selected_option = Session("ownername")
actowner = "TRUE"
longname = "TRUE"
listall = "FALSE" %>

	<p>Owner&nbsp;
	<!--#Include File="cd_ownerlist.asp"-->
	<p>
	<input type="submit" name="action" value="Continue">
</form>
		
</div>
</body>
</html>
