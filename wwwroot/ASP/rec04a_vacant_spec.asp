<%
thispage = "vacant_spec"
pageaccess = "ADMIN"
%>
<!--#Include File="../includes/version.asp"-->
<%
errormsg = ""
site = Session("site")

If lastpage = thispage Then
	site = Request("site")
	plan_delim = Trim(Replace(Request("plan_delim"),",",""))
	If IsNumeric(Request("plan_levels")) And IsNumeric(Request("plan_length")) And IsNumeric(Request("plan_maxloc")) And Len(plan_delim) > 0 Then
		Session("plan_case") = Request("plan_case")
		Session("plan_levels") = Request("plan_levels")
		Session("plan_length") = Request("plan_length")
		Session("plan_maxloc") = Request("plan_maxloc")
		Session("plan_delim") = Request("plan_delim")
		Session("site") = site
		Response.Clear
		Server.Transfer version & "_vacant_add.asp"
	Else
		errormsg = "Levels, Length, and Locations must be Positive Integers and Delimiter Non-blank"
	End If
End If %>	

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Location Plan Specifications</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="../includes/header.asp"-->	

<p><strong>Location Plan Specifications</strong>
<p><font color="#FF0000"><%= errormsg %></font>
<form action="<%= version %>_vacant_spec.asp" method="post">

<%
SQL = "SELECT shortname,longname,sortorder FROM reflist WHERE list = 'SITE' AND status = 'ACTIVE' ORDER BY sortorder"
%>
<!--#Include File="../includes/db_select.asp"-->

	<p>Site&nbsp;
	<select name = "site">

<%
Do Until rs.eof
	If site = rs("shortname") Then
		selected = "selected"
	Else
		selected = ""
	End If %> 
								
		<option value="<%= rs("shortname") %>" <%= selected %> ><%= rs("shortname") & "&nbsp;&nbsp;" & rs("longname") %>	
			
<%	rs.movenext
Loop %>
			
	</select>&nbsp;&nbsp;
	<p>
	Case&nbsp;
	<select name="plan_case">
			<option value="UPPER">Upper Case	
			<option value="MIX">Mixed Case
	</select>
	<p>Number of Levels&nbsp;
	<input type="text" name="plan_levels" size="2" maxlength="2">
	<p>Maximum Length of Each Location&nbsp;
	<input type="text" name="plan_length" size="2" maxlength="2">
	<p>Maximum Number of Locations per Level&nbsp;
	<input type="text" name="plan_maxloc" size="2" maxlength="2">
	<p>Location Delimiter&nbsp;
	<input type="text" name="plan_delim" size="5" maxlength="5">&nbsp;Enter 'space' for space or 'none' for no delimiter 
	<p>
	<input type="submit" value="Continue">
</form>
		
</div>
</body>
</html>
