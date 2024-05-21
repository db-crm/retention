<% 'Get the values for a box label.  Input - recordno

SQL = "SELECT site,location,ownerno,box,boxid,reviewdate,content FROM record INNER JOIN location ON record.recordno = location.recordno WHERE record.recordno = " & recordno & ""
%>
<!--#Include File="db_select.asp"-->	
<%
If Not rs.eof Then 
	site = rs("site")
	location = rs("location")
	box = rs("box")
	boxid = rs("boxid")
	reviewdate = rs("reviewdate")
	content = rs("content")
End If

SQL = "SELECT longname FROM reflist INNER JOIN location ON reflist.shortname = location.site WHERE location.recordno = " & recordno & " AND reflist.list = 'SITE'"
%>
<!--#Include File="db_select.asp"-->	
<%
If Not rs.eof Then 
	sitelong = rs("longname")
End If

SQL = "SELECT ownername,ownerlong FROM record INNER JOIN ownerinfo ON record.ownerno = ownerinfo.ownerno WHERE record.recordno = " & recordno & ""
%>
<!--#Include File="db_select.asp"-->	
<%
If Not rs.eof Then 
	ownername = rs("ownername")
	ownerlong = rs("ownerlong")
End If %>

<table width="90%" border="0" cellspacing="0" cellpadding="10">
	<tr>
		<td width="20%" align="right"><font size="6">Site&nbsp;</font>
		<td width="80%"><font size="6"><%= site %>&nbsp;&nbsp;<%= sitelong %></font>
	<tr>
		<td width="20%" align="right"><font size="6">Location&nbsp;</font>
		<td width="80%"><font size="6"><%= location %></font>
	<tr>
		<td align="right"><font size="6">Owner&nbsp;</font>
		<td ><font size="6"><%= ownername %>&nbsp;&nbsp;<%= ownerlong %></font>
	<tr>
		<td align="right"><font size="6">Box&nbsp;</font>
		<td><font size="6"><%= box %>&nbsp;<%= boxid %></font>
	<tr>
		<td align="right"><font size="6">Review_Date&nbsp;</font>
		<td><font size="6"><%= reviewdate %></font>		
	<tr>
		<td align="right" valign="top"><font size="6">Contents&nbsp;</font>
		<td><%= content %>
</table>