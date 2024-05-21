<%
thispage = "report_display"
pageaccess = "GROUPUSERSUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
Session("update_mode") = "report"
Session("content_size") = Request("content_size")
Session("showwhere") = Request("showwhere")

'Remove the leading AND to make the where clause to be displayed in the report header
headwhere = Session("report_where")
pos = InStr(headwhere,"AND")
If pos = 0  Or Session("showwhere") <> "TRUE" Then 'No display if no where or show where unchecked
	headwhere = " "
Else
	headwhere = "(" & Right(headwhere,Len(headwhere)-pos-3) & ")"
End If

select_fields = "  record.recordno,ownername,box,boxid,retention,record.status,disposition,content,entrydate,updatedate,recorddate,reviewdate,deletedate "

select_from = "  FROM (record INNER JOIN ownerinfo ON record.ownerno = ownerinfo.ownerno) "

If Session("report_type") = "location" Or Session("report_type") = "active" Then
	select_fields = "site,location," & select_fields
	select_from = select_from & " INNER JOIN location ON record.recordno = location.recordno "
End If

'Get the count of records returned 
SQL = "SELECT COUNT(*) AS reportcount " & select_from & Session("report_where")
%>
<!--#Include File="db_select.asp"-->
<%
reportcount = rs("reportcount")
headercount = reportcount & " records"

'Get the data to report
SQL = "SELECT " & select_fields & select_from & Session("report_where") & Session("report_orderby")
%>
<!--#Include File="db_select.asp"-->	

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Display Report Results</title>	
</head>
<body>

<p><font size="+1"><%= Session("client") %>&nbsp;&nbsp;<a href="<%= Session("returnto") %>" title="Return to the Menu"><%= Session("report_title") %>,&nbsp;<%= Date %></font></a><strong>&nbsp;&nbsp;<%= headwhere %>&nbsp;<%= headercount %></strong>

<%
If reportcount > 0 Then %>

<form action="<%= version %>_reports_select">
	<p>
	<table width="100%" border="1" cellspacing="0" cellpadding="2">
		
<% 	If Session("content_size") = "full" Then
		headcount = CInt(Session("shortheadcount"))
	Else
		headcount = CInt(Session("longheadcount"))
	End If		

	Do While Not rs.eof 'Make a header row every 'headcount' rows %>
	
		<tr align="center">
			<td><strong><font color="#008000">Owner/Box</font></strong>
			
<%		If Session("report_type") ="active" Or Session("report_type") = "location" Then %>	
					
			<td><strong><font color="#008000">Location</font></strong>
			
<%		End If %>				
			
			<td><strong><font color="#008000">Retention</font></strong>

<%		If Session("report_type") ="all" Or Session("report_type") = "deleted" Then %>			
			
			<td><strong><font color="#008000">Disposition</font></strong>
			
<%		End If
		If Instr(Request("report_fields"),"recorddate") > 0 Then %>					
			
			<td><strong><font color="#008000">Created</font><strong>
			
<%		End If
		If Instr(Request("report_fields"),"reviewdate") > 0 Then %>					
			
			<td><strong><font color="#008000">Review</font></strong>
			
<%		End If
		If Instr(Request("report_fields"),"deletedate") > 0 Then %>					
			
			<td><strong><font color="#008000">Deleted</font></strong>
			
<%		End If	
		If Instr(Request("report_fields"),"entrydate") > 0 Then %>				
			
			<td><strong><font color="#008000">Entered</font></strong>
			
<%		End If
		If Instr(Request("report_fields"),"updatedate") > 0 Then %>					
			
			<td><strong><font color="#008000">Updated</font></strong>
			
<%		End If
		If Session("content_size") <> "none" Then %>						
			
			<td><strong><font color="#008000">Content</font></strong>
			
<%		End If
		i = 1	
		Do While Not rs.eof
			recorddate = "//"
			If IsDate(rs("recorddate")) Then recorddate = FormatDateTime(rs("recorddate"),2)
			entrydate = "//"
			If IsDate(rs("entrydate")) Then entrydate = FormatDateTime(rs("entrydate"),2)
			reviewdate = "//"
			If IsDate(rs("reviewdate")) Then reviewdate = FormatDateTime(rs("reviewdate"),2)
			updatedate = "//"
			If IsDate(rs("updatedate")) Then updatedate = FormatDateTime(rs("updatedate"),2)
			deletedate = "//"
			If IsDate(rs("deletedate")) Then deletedate = FormatDateTime(rs("deletedate"),2)
			If Session("content_size") = "short" Then
				content = Left(rs("content"),20)
			Else
				content = rs("content")
			End If %> 
	
			<tr align="center" valign="top">
				<td align = "left"><%= rs("ownername") %>/<%= rs("box") %>&nbsp;<%= rs("boxid") %>
				
<%			If Session("report_type") ="active" Or Session("report_type") = "location" Then
				If Session("showsite") = "NEVER" Then
					location = rs("location")
				Else
					location = rs("site") & "/" & rs("location")
				End If %>				
				
				<td align = "left"><%= location %>
				
<%			End If %>					
				
				<td><%= rs("retention") %>
				
<%			If Session("report_type") = "all" Or Session("report_type") = "deleted" Then %>				
				
				<td><%= rs("disposition") %>
				
<%			End If
			If Instr(Request("report_fields"),"recorddate") > 0 Then %>						
								
				<td><%= recorddate %>
				
<%			End If
			If Instr(Request("report_fields"),"reviewdate") > 0 Then %>						
				
				<td><%= reviewdate %>
				
<%			End If
			If Instr(Request("report_fields"),"deletedate") > 0 Then %>						
				
				<td><%= deletedate %>
				
<%			End If
			If Instr(Request("report_fields"),"entrydate") > 0 Then %>						
				
				<td><%= entrydate %>
				
<%			End If
			If Instr(Request("report_fields"),"updatedate") > 0 Then %>						
				
				<td><%= updatedate %>
				
<%			End If
			If Session("content_size") <> "none" Then %>		
				
				<td align="left"><%= content %>
				
<%			End If
			rs.movenext
			i = i + 1
			If i = headcount Then Exit Do
		Loop
	Loop %>
	
	</table>
</form>	
	
<% 
End If %>

</body>
</html>
