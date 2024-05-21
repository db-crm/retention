<%
thispage = "reports"
pageaccess = "GROUPUSERSUPERADMIN"
errors = 0
delim = Session("rangedelim")
%>
<!--#Include File="version.asp"-->
<%

'Set all query input flags to yes.  If the flag is Y, the field is included on the query input screen and the input value is part of the where clause.  If the flag is N, the field is not on the query input screen and all values of the field will be selected.

Session("returnto") = version & "_reports.asp"
rowner = "Y"
rstatus = "Y"
rdisposition = "Y"
rretention = "Y"
rbox = "Y"
rcontent = "Y"
rentrydate = "Y"
rrecorddate = "Y"
rreviewdate = "Y"
rupdatedate = "Y"
rdeletedate = "Y"

'Get the checked report fields.  Fields that are not checked will not be displayed in the report.  They may be used as select criteria even if not displayed.

If thispage = lastpage Then
	report_fields = Request("report_fields")
Else
	report_fields = Session("report_fields")
End If

strfields = Trim(report_fields)
yes = "Y"
If Len(strfields) > 0 Then 'If any fields checked, parse the string
	Do
		pos = Instr(strfields,",") 'Find the first comma
		If pos = 0 Then Exit Do 'If no comma, is last field
		fieldname = Left(strfields,pos-1) 'Get the field name
		strfields = Trim(Right(strfields,Len(strfields)-pos)) 'Remainder of string
		str = "v" & fieldname & " = yes"
		Execute str 'Set vfieldname = "Y"
	Loop
	str = "v" & strfields & " = yes"
	Execute str 'Last, or only, field
End If

'Report type determines which fields may be used as select criteria.  A blank field will return all values of the field

Select Case Session("report_type")
	Case "active" 'Include all input fields except disposition, deletedate, and location
		qstatus = " AND record.status = 'ACTIVE'"
		rdisposition = "N"
		rdeletedate = "N"
		rlocation = "N"
		orderby = "ownername,box"
		Session("report_title") = "Report by Active Records"
	Case "deleted" 'Include all input fields except location
		qstatus = " AND record.status = 'DELETED'"
		rlocation = "N"
		orderby = "disposition,ownername,box"
		Session("report_title") = "Report by Deleted Records"
	Case "all" 'Include all input fields except disposition and location
		qstatus = " "
		rdisposition = "N"
		rlocation = "N"
		orderby = "disposition,ownername,box"
		Session("report_title") = "Report by All Records"
	Case "location"
		Response.Clear
		Server.Transfer version & "_reports_loc.asp"	
	Case Else 'This case should not occur
		Response.Clear
		Server.Transfer version & "_menu.asp"
End Select

'If not the first time through, make the part of the where clause for each query parameter and go to the display page if no errors. Don't include in the where clause if blank or not a query parameter	

If thispage = lastpage Then
	Session("ownername") = Request("ownername")
	If Request("ownername") <> "ALL" And rowner = "Y" Then	
		qownername = " AND ownername = '" & Request("ownername") & "'"
	Else
		qownername = " "
	End If
	Session("disposition") = Request("disposition")
	If Request("disposition") <> "ALL" And rdisposition = "Y" Then 
		qdisposition = " AND disposition = '" & Request("disposition") & "'"
	Else
		qdisposition = " "
	End If
	If Len(Request("searchcontent")) <> 0 And rcontent = "Y" Then 
		qcontent = " AND content LIKE '%" & Trim(Request("searchcontent")) & "%'"
	Else
		qcontent = " "
	End If

	fieldname = "retention"
	fieldtype = "integer"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	fieldname = "box"
	fieldtype = "integer"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	fieldname = "reviewdate"
	fieldtype = "date"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	fieldname = "deletedate"
	fieldtype = "date"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	fieldname = "recorddate"
	fieldtype = "date"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	fieldname = "entrydate"
	fieldtype = "date"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	fieldname = "updatedate"
	fieldtype = "date"
%>
<!--#Include File="cd_field_input.asp"-->	
<%
	If errors = 0 Then 'No errors, make the where clause and go to the report display page
		Session("report_where") = " WHERE 1=1 " & qownername & qstatus & qdisposition & qcontent & qretention & qbox & qentrydate & qrecorddate & qreviewdate & qdeletedate
		Session("report_orderby") = " ORDER BY " & orderby
		Response.Clear
		Server.Transfer version & "_reports_display.asp"
		
	End If
End If %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Enter Report Parameters</title>
</head>
<body bgcolor="<%= Session("bgcolor") %>">
<div align="center">

<!--#Include File="header.asp"-->	

<%
If errors = 0 Then %>

	<p><strong><%= Session("report_title") %></strong><br>

<%
Else %>

	<p><font color="#FF0000"><strong>Correct the fields labeled in red and Submit again<br></strong></font>

<%
End If %>

<form method="POST" action="<%= version %>_reports.asp">
	<p>
	<table width=80% border="0" cellspacing="0" cellpadding="2">

<%
If rowner = "Y" Then 'Make the Owner select list.  Preselect if already submitted %>


		<tr>		
			<td align="right">Owner
			<td>
			
<%	If thispage = lastpage Then
		selected_option = Request("ownername")
	Else
		selected_option = Session("ownername")
	End If	 
	listall = "TRUE"
	longname = "TRUE"
%>	
<!--#Include File="cd_ownerlist.asp"-->		
<%
End If

If rbox = "Y" Then 'Make the Box entry field
	fieldname = "box"
	fieldlabel = "Box Number"
	fieldsize = 11
%>
<!--#Include File="cd_field_entry.asp"-->	
<%
End If

If rretention = "Y" Then 'Make the Retention entry field
	fieldname = "retention"
	fieldlabel = "Retention"
	fieldsize = 6
%>
<!--#Include File="cd_field_entry.asp"-->

&nbsp;Years

<%
End If

If rcontent = "Y" Then 'Make the Search Content entry field.  Load text already submitted. %>

		<tr>			
			<td align="right">Content Search String
			<td>

<%	If lastpage = thispage Then 
		lastvalue = Request("searchcontent")
	Else 
		lastvalue = Empty
	End If %>
	
			<input type="text" name="searchcontent" value="<%=lastvalue%>" size="30" maxlength="30">&nbsp;
			
<%	selectname = "content_size"
	aryoptionvalue = Array("full","short","none")
	aryoptionlabel = Array("Full Content","Short Content","No Content")
%>
<!--#Include File="cd_selectarray.asp"-->
<%
End If

If rdisposition = "Y" Then 'Make the Delete Code select list. 'Make disposition arrays
	arydispolist = Split(Session("dispolist"),",")
	arydispolistlong = Split(Session("dispolistlong"),",")%>

		<tr>	
			<td align="right">Disposition
			<td>
			
<%	If lastpage = thispage Then	
		selected_option = Request("disposition")
	Else 
		selected_option = Session("disposition")
	End If %>
				
			<select name="disposition">
			
<%	optionvalue = "ALL"
	optionlabel = "All"
%>
<!--#Include File="cd_selectoption.asp"-->
<%
	i = 0
	Do Until i > Ubound(arydispolist)
		optionvalue = arydispolist(i)
		optionlabel = arydispolistlong(i)
%>
<!--#Include File="cd_selectoption.asp"-->
<%
		i = i + 1
	Loop %>	
		        
			</select>
				
<%
End If

If rreviewdate = "Y" Then 'Make Reviewdate entry field
	fieldname = "reviewdate"
	fieldlabel = "Review Date"
	fieldsize = 21
%>
<!--#Include File="cd_field_entry.asp"-->	
<%
%>
<!--#Include File="cd_checkbox.asp"-->	
<%
End If

If rdeletedate = "Y" Then 'Make the deletedate entry field
	fieldname = "deletedate"
	fieldlabel = "Delete Date"
	fieldsize = 21
%>
<!--#Include File="cd_field_entry.asp"-->	
<%
%>
<!--#Include File="cd_checkbox.asp"-->	
<%
End If

If rrecorddate = "Y" Then 'Make the Recorddate entry field
	fieldname = "recorddate"
	fieldlabel = "Record Date"
	fieldsize = 21
%>
<!--#Include File="cd_field_entry.asp"-->	
<%
%>
<!--#Include File="cd_checkbox.asp"-->	
<%
End If

If rentrydate = "Y" Then 'Make the Entrydate entry field
	fieldname = "entrydate"
	fieldlabel = "Entry Date"
	fieldsize = 21
%>
<!--#Include File="cd_field_entry.asp"-->	
<%
%>
<!--#Include File="cd_checkbox.asp"-->	
<%
End If

If rupdatedate = "Y" Then 'Make the Udatedate entry field
	fieldname = "updatedate"
	fieldlabel = "Update Date"
	fieldsize = 21
%>
<!--#Include File="cd_field_entry.asp"-->	
<%
%>
<!--#Include File="cd_checkbox.asp"-->	
<%
End If

If Session("accesstype") = "SUPER" Or Session("accesstype") = "ADMIN" Or Session("accesstype") = "" Then 'Make a checkbox to turn the where clause on and off on the report if access ADMIN or above
	If lastpage = thispage Then
		showwhere = Request("showwhere")
		Session("showwhere") = showwhere
	Else
		showwhere = Session("showwhere")
	End If
	If showwhere = "TRUE" Then
		checked = "checked"
	Else
		checked = Empty
	End If %>	
		
		<tr>			
			<td align="right">
			<td><input type="checkbox" name="showwhere" value="TRUE" <%= checked %>>
			&nbsp;Display the Where Clause on the Report
	
<%
End If %>

	</table>
	<p><input type="submit" value="Submit">
</form>

<%
If Not rlocation = "Y" Then %>

<p>Use '<%= Session("rangedelim") %>' to make ranges

<%
End If

If Session("help") = "TRUE" Then 'Append help text if help ON
%>
<!--#Include File="../help/reports_help.asp"-->
<%
End If %>
		
</div>
</body>
</html>
