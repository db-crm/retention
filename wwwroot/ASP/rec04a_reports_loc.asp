<%
thispage = "reports_loc"
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="../includes/version.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Get Site and Location for Report</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="../includes/header.asp"-->	

<form method="POST" action="<%= version %>_reports_loc.asp">

<%
If lastpage = thispage Then
	site = Request("site")
	gotsite = "TRUE"
	If Request("action") = "Display Report" Then
		location = Request("location")
		If location = "ALL" Then
			where = ""
		Else
			where = " AND location = '" & location & "'"
		End If
		Session("returnto") = version & "_reports_loc.asp"
		Session("report_title") = "Report by Location"
		Session("report_where") = " WHERE site = '" & site & "' " & where
		Session("report_orderby") = " ORDER BY site,location"
		Response.Clear
		Server.Transfer version & "_reports_display.asp"
	End If
Else
	site = Session("site")
	If	Session("showsite") = "NEVER" Then
		gotsite = "TRUE"
	Else
		site_select = "site" %>
	
	<p><strong>Select the Site for the Location you want to Report</strong> 	
	<p>Site&nbsp;
	
<!--#Include File="../includes/cd_site_select.asp"-->

	<p>
	<input type="submit" value="Continue">

<%
	End If
End If

If gotsite = "TRUE" Then %>

	<p><strong>Select the Location you want to Report at Site <%= site %></strong> 
	<p>Location&nbsp;
	<select name="location">
	
<%	optionvalue = "ALL"
	optionlabel = "All"
%>
<!--#Include File="../includes/cd_selectoption.asp"-->
<%
	SQL="SELECT location FROM location WHERE site = '" & site & "' AND recordno <> 0 ORDER BY location"
%>
<!--#Include File="../includes/db_select.asp"-->	



<%	Do While Not rs.eof
		optionvalue = rs("location")
		optionlabel = rs("location")
%>
<!--#Include File="../includes/cd_selectoption.asp"-->
<%
		rs.movenext
	Loop %>
					       
	</select>&nbsp;
	<input type="hidden" name="site" value="<%= site %>">
	<input type="submit" name="action" value="Display Report">		
						
<%
End If %>

</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<!--#Include File="../help/reports_loc_help.asp"-->
<%
End If %>
		
</div>
</body>
</html>
