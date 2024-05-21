<%
thispage = "expired"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
Session("returnto") = version & "_expired.asp"
Session("update_mode") = "expired"

Select Case Request("goto")
	Case "Destruction Form"
		Response.Clear
		Server.Transfer version & "_expired_destruct.asp" 
	Case "Delete and Update"
		Response.Clear
		Server.Transfer version & "_expired_select.asp" 
End Select %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Expired Records</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<p><strong>Expired Records, <%= Session("expired_label") %></strong>
<form action="<%= version %>_expired.asp" method="post">
			
<%'Display and select the owner and number of records with reviewdate specified by expired_query
SQL = "SELECT ownername,record.ownerno,count(*) AS excount FROM record INNER JOIN ownerinfo ON record.ownerno = ownerinfo.ownerno WHERE record.status = 'ACTIVE' AND " & Session("expired_query") & " GROUP BY ownername,record.ownerno ORDER BY ownername"
%>
<!--#Include File="db_select.asp"-->	
<%
If rs.eof Then %>
			
	<p><strong><font color="Green">There are no expired records for the period selected</font></strong>
				
<%
Else %>

	<p>Owner/Records
	<select name="ownername">
			
<% 	Do While Not rs.eof %> 
							
		<option value=<%= rs("ownername") %>><%= rs("ownername") & "/" & rs("excount") %>	
		
<%		rs.movenext
	Loop %>
			
	</select>
	<p>
	<input type="submit" name="goto" value="Destruction Form">
	<p>	
	<input type="submit" name="goto" value="Delete and Update">
</form>
		
<%
End If

If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/expired_help.asp"--></font>
<%
End If %>
		
</div>
</body>
</html>
