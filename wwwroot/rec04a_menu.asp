<%
thispage = "menu"
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="version.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Menu</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->

<% 'Clear menu_select if coming from another page
If thispage = lastpage Then
	menu_select = Request("menu_select")
Else
	menu_select = " "
End If %>	

<form method=POST action="<%= version %>_menu.asp" >

<%
If Session("accesstype") = "ADMIN" Then

	menu_label = "Vacant Locations"
	If menu_select = menu_label Then
		Session("vacant_action") = Request("vacant_action")
		Response.Clear
		Server.Transfer version & "_vacant_spec.asp"
	End If %>
	
	<p>
	<input type="submit" name="menu_select" value="<%=menu_label%>">
	&nbsp;
	<select name = "vacant_action">
		<option value="ADD">Add Locations				
		<option value="REMOVE">Remove Locations
	</select>
	
<%	menu_label = "Add and Remove Users, Change Passwords"
	gotofile = "newuser"
%>
<!--#Include File="cd_menu.asp"-->
<%
End If

If Session("accesstype") = "ADMIN" Or Session("accesstype") = "" Then

	menu_label = "Modify a Reference List"
	If menu_select = menu_label Then
		Session("reflist") = Request("reflist")
		Response.Clear
		Server.Transfer version & "_ref_select.asp"
	End If %>
	
	<p>
	<input type="submit" name="menu_select" value="<%=menu_label%>">
	&nbsp;
	<select name = "reflist">
			
<%	SQL = "SELECT DISTINCT list FROM reflist WHERE change = 'TRUE' ORDER BY list"
%>
<!--#Include File="db_select.asp"-->	
<%
	Do While Not rs.eof %> 
								
		<option><%= rs("list") %>	
			
	<%	rs.movenext
	Loop %>
			
	</select>

<%	menu_label = "Change Preference Items"
	If menu_select = menu_label Then
		preference = Request("preference")
		Response.Clear
		Server.Transfer version & "_pref_" & preference & ".asp"
	End If %>	
	
	<p>
	<input type="submit" name="menu_select" value="<%=menu_label%>">
	&nbsp;
	<select name="preference">
		<option value="color">Background Color
		<option value="pagehead">Page Header	
		<option value="rangedelim">Range Delimiter
		<option value="reporthead">Report Header Spacing
		<option value="site">Show Storage Sites
		<option value="userlength">Username and Password Length
		<option value="passwordmix">Password Case Sensitivity
		<option value="ownermix">Owner Short Name Case
		<option value="boxid">Use Box Identifier					
	</select>
<%
End If

menu_label = "Report"
If menu_select = menu_label Then
	Session("report_type") = Request("report_type")
	Response.Clear
	Server.Transfer version & "_reports.asp"
End If %>

	<p>
	<input type="submit" name="menu_select" value="<%=menu_label%>">
	&nbsp;<strong>by</strong>&nbsp;
	
<%
selectname = "report_type"
aryoptionvalue = Array("active","deleted","all","location")
aryoptionlabel = Array("Active Records","Deleted Records","All Records","Location")
%>
<!--#Include File="cd_selectarray.asp"-->
<%
If Session("accesstype") = "SUPER" Or Session("accesstype") = "ADMIN" Then

	menu_label = "Add Records"
	site_select = "add_site"
	If menu_select = menu_label Then
		If Session("showsite") <> "NEVER" Then
			Session("site") = Request(site_select)
		End If	
		Response.Clear
		Server.Transfer version & "_add.asp"
	End If %>	

	<p>
	<input type=submit name=menu_select value="<%=menu_label%>">
	
<%	If	Session("showsite") <> "NEVER" Then %>
	
	&nbsp;<strong>at Site</strong>&nbsp;
	
<!--#Include File="cd_site_select.asp"-->
<%
	End If

	menu_label = "Expired Records"
	If menu_select = menu_label Then
		Session("expired") = Request("expired")
		month_first = Session("datedelim") & DatePart("m",Date) & "/1/" & DatePart("yyyy",Date) & Session("datedelim")
		If DatePart("m",Date) = 12 Then
			month_next = Session("datedelim") & "1/1/" & (DatePart("yyyy",Date) + 1) & Session("datedelim")
		Else
			month_next = Session("datedelim") & (DatePart("m",Date) + 1) & "/1/" & DatePart("yyyy",Date) & Session("datedelim")
		End If
		Select Case Request("expired")
			Case "month"
				Session("expired_label") = "In " & MonthName(DatePart("m",Date)) & " " & DatePart("yyyy",Date)
				Session("expired_query") = " reviewdate >= " & month_first & " and reviewdate < " & month_next & " "
			Case "before"
				Session("expired_label") = "Before " & MonthName(DatePart("m",Date)) & " " & DatePart("yyyy",Date)
				Session("expired_query") = " reviewdate < " & month_first & " "
			Case "all"
				Session("expired_label") = "All" 
				Session("expired_query") = " reviewdate < " & month_next & " "
		End Select
		Response.Clear
		Server.Transfer version & "_expired.asp"
	End If %>
		
	<p>
	<input type=submit name=menu_select value="<%=menu_label%>">
	&nbsp;
	
<%	selectname = "expired"
	aryoptionvalue = Array("month","before","all")
	aryoptionlabel = Array("This Month","Before This Month","All")
%>
<!--#Include File="cd_selectarray.asp"-->
<%
	menu_label = "Edit or Remove Records"
	gotofile = "cleanup"
%>
<!--#Include File="cd_menu.asp"-->
<%
menu_label = "Move Boxes"
	gotofile = "move"
%>
<!--#Include File="cd_menu.asp"-->
<%
	menu_label = "Display Vacant Locations"
	site_select = "display_site"
	If menu_select = menu_label Then
		If Session("showsite") <> "NEVER" Then
			Session("site") = Request(site_select)
		End If
		Response.Clear
		Server.Transfer version & "_vacant_display.asp"
	End If %>	

	<p>
	<input type=submit name=menu_select value="<%=menu_label%>">
	
<%	If	Session("showsite") <> "NEVER" Then %>
	
	&nbsp;<strong>at Site</strong>&nbsp;
	
<!--#Include File="cd_site_select.asp"-->
<%
	End If

	menu_label = "Add/Edit Owners"
	gotofile = "owner"
%>
<!--#Include File="cd_menu.asp"-->
<%
	menu_label = "Summary Count"
	gotofile = "summary"
%>
<!--#Include File="cd_menu.asp"-->
<%
End If

If Session("accesstype") <> "GROUP" Then 'Groups can't change the password

	If Session("showsite") = "NEVER" Then
		menu_label = "Change Password"
	Else
		menu_label = "Change Password/Default Site"	
	End If
	gotofile = "password"
%>
<!--#Include File="cd_menu.asp"-->
<%
End If

If Session("showsite") <> "NEVER" Then
	menu_label = "Show All Sites On/Off"
	gotofile = "showsite"
%>
<!--#Include File="cd_menu.asp"-->
<%
End If

menu_label = "Site, Location, and Owner Reference Lists"
gotofile = "labels"
%>
<!--#Include File="cd_menu.asp"-->
<%
menu_label = "Help On/Off"
gotofile = "help"
%>
<!--#Include File="cd_menu.asp"-->
	
</form>

<%
If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<font color="Blue"><!--#Include File="help/menu_help.asp"--></font>
<%
End If %>

</div> 
</body>
</html>
