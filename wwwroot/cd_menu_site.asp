<%
If menu_select = menu_label Then
	Session("site") = Request("site")	
	Response.Clear
	Server.Transfer version & "_" & gotofile & ".asp"
End If %>	

<p>
<input type=submit name=menu_select value="<%=menu_label%>">

<%
	SQL = "SELECT shortname,longname,sortorder FROM reflist WHERE list = 'SITE' AND status = 'ACTIVE' ORDER BY sortorder"
%>
<!--#Include File="db_select.asp"-->
<%
If Session("showsite") = "TRUE" Then 'Get the list of active sites %>

&nbsp;<strong>Site</strong>&nbsp;
<select name = "site">

<%
	Do Until rs.eof %> 
								
	<option value="<%= rs("shortname") %>"><%= rs("shortname") & "&nbsp;&nbsp;" & rs("longname") %>	
			
<%		rs.movenext
	Loop %>
			
</select>
	
<%
Else 'Use the first sight determined by sort order as the session site %>

<input type="hidden" name="site" value="<%= rs("shortname") %>">
	
<%
End If %>	