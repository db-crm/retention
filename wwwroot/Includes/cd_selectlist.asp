<%	If lastpage = thispage Then
		selected_option = Request(listname)
	Else 
		selected_option = ""
	End If %>
	
<select name="<%= listname %>" size="1" >
			
<%	SQL = "select shortname from reflist where list = '" & listlabel & "' order by sortorder"
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
	Do While Not rs.eof
		If rs("shortname") = selected_option Then
			selected = "selected"
		Else 
			selected = ""
		End If %>
	
	<option <%= selected %>><%= rs("shortname") %>
		
<%		rs.movenext
	Loop %>	
		
</select>