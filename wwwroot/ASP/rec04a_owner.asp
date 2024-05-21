<%
thispage = "owner"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Add an New Owner</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<p><strong>Add or Edit an Owner</strong>
<p><font color="#0000FF">Select an Owner to Edit or to use as a Template for a New Owner</font>
<form method="POST" action="<%= version %>_owner_add.asp">

<%'Make the Owner select list
If Session("accesstype") = "ADMIN" Then 'Add the label to the list if Admin
	where = " ownerno <> 0"
Else
	where = " ownerno > 0"
End If
SQL="SELECT ownerno,ownername,ownerlong,contact,status FROM ownerinfo WHERE status <> 'INACTIVE' AND " & where & " ORDER BY status,ownername"
%>
<!--#Include File="db_select.asp"-->
	
	<p>Owner&nbsp;
	<select name="ownerno">			

<%
Do While Not rs.eof %>
					
		<option value=<%= rs("ownerno") %>><%= rs("ownername") & "&nbsp;&nbsp;" & rs("ownerlong") %>	
			
<%	rs.movenext
Loop %> 
	        
	</select>&nbsp;&nbsp;Action&nbsp;
	<select name="owner_action">
		<option value="ADD">Add a New Owner
		<option value="EDIT">Edit an Existing Owner
	</select>
	<p>
	<input type="submit" value="Continue">
</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/move_help.asp"--></font>
<%
End If %>
		
</div>
</body>
</html>
