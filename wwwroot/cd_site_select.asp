
<select name = "<%= site_select %>">

<%
If Session("showsite") = "TRUE" Then 'Get the list of all active sites
	SQL = "SELECT shortname,longname,sortorder FROM reflist WHERE list = 'SITE' AND status = 'ACTIVE' ORDER BY sortorder"
%>
<!--#Include File="db_select.asp"-->
<%
	Do Until rs.eof %> 
								
	<option value="<%= rs("shortname") %>"><%= rs("shortname") & "&nbsp;&nbsp;" & rs("longname") %>	
			
<%		rs.movenext
	Loop
Else 'Use the user default site
	SQL = "SELECT shortname,longname FROM userinfo INNER JOIN reflist ON userinfo.site = reflist.shortname WHERE reflist.list = 'SITE' AND username = '" & Session("username") & "'"
%>
<!--#Include File="db_select.asp"-->
<%
	Do Until rs.eof %>
	
	<option value="<%= rs("shortname") %>"><%= rs("shortname") & "&nbsp;&nbsp;" & rs("longname") %>
	
<%		rs.movenext
	Loop
End If %>

</select>