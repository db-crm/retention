<%
thispage = "owner_add"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
errormsg = ""
ownerno = Request("ownerno")

If lastpage = thispage Then
	ownername = Trim(Request("ownername"))
	If Session("ownermix") <> "TRUE" Then ownername = UCase(ownername)
	ownerlong = Trim(Request("ownerlong"))
	contact = Trim(Request("contact"))
	status = Request("status")
	If status = "CHECKED" Then 'Make the owner inactive if checked 
		status = "INACTIVE"
		checked = "checked"
	Else
		status = "ACTIVE"
		checked = ""
	End If
	If Len(ownername) = 0 Or Len(ownerlong) = 0 Or Len(contact) = 0 Then
			errormsg = "Owner Names and Contact cannot be blank"
	Else
		Select Case Request("action")
			Case "Add the New Owner" 'Add the new owner
				SQL = "SELECT * FROM ownerinfo WHERE ownername = '" & ownername & "'"
%>
<!--#Include File="db_select.asp"-->
<%
				If rs.eof Then 'No duplicate, get a new owner number and add the owner
					SQL = "SELECT MAX(ownerno) AS maxowner FROM ownerinfo"
%>
<!--#Include File="db_select.asp"-->
<%
					ownerno = rs("maxowner") + 1
					SQL = "INSERT INTO ownerinfo(ownerno,ownername,ownerlong,status,contact) VALUES(" & ownerno & ",'" & ownername & "','" & ownerlong & "','ACTIVE','" & contact & "')"
%>
<!--#Include File="db_execute.asp"-->
<%
					Response.Clear
					Server.Transfer version & "_owner.asp"
				Else
					errormsg = "Owner " & ownername & " already exists"
				End If
			Case "Submit the Changes" 'Update an existing owner
				SQL = "SELECT * FROM ownerinfo WHERE ownerno <> " & ownerno & " AND ownername = '" & ownername & "'"
%>
<!--#Include File="db_select.asp"-->
<%
				If rs.eof Then 'No duplicate, get a new owner number and add the owner
					SQL = "UPDATE ownerinfo SET ownername = '" & ownername & "',ownerlong = '" & ownerlong & "',status = '" & status & "',contact = '" & contact & "' WHERE ownerno = " & ownerno & ""
%>
<!--#Include File="db_execute.asp"-->
<%
					Response.Clear
					Server.Transfer version & "_owner.asp"
				Else
					errormsg = "Owner " & ownername & " already exists"
				End If
		End Select
	End If
Else 'First time thru, get owner info from ownerno
Session("owner_action") = Request("owner_action")
%>
<!--#Include File="cd_ownername.asp"-->
<%
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention View and Add or Edit an Owner</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<p><strong>Add or Edit Owner</font></strong>
<p><font color="#FF0000"><%= errormsg %></font>
<form method="POST" action="<%= version %>_owner_add.asp">
	<p>Owner Short Name&nbsp;
	<input type="text" name="ownername" value="<%= ownername %>" size="32" maxlength="32">
	<p>Owner Long Name&nbsp;
	<input type="text" name="ownerlong" value="<%= ownerlong %>" size="64" maxlength="64">
	<p>Contact&nbsp;
	<textarea cols="64" rows="2" name="contact"><%= contact %></textarea>
	
<% 'Inactive checkbox only if action is edit
If Session("owner_action") = "EDIT" Then %>

	<p>Make the Owner Inactive&nbsp;
	<input type="checkbox" name="status" value="CHECKED" <%= checked %>>
	
<%
End If %>	
	
	<p>
	<input type="submit" name="action" value="Validate">&nbsp;
	<input type="hidden" name="ownerno" value="<%= ownerno %>">
		
<%
If lastpage = thispage And errormsg = "" Then
	If Session("owner_action") = "ADD" Then  %>	
	
	<input type="submit" name="action" value="Add the New Owner">
	
<%	Else %>

	<input type="submit" name="action" value="Submit the Changes">
	
<%	End If
End If %>

</form>

</div>
</body>
</html>
