<% 'Make the owner select list.  Preselect if already selected.

If actowner = "TRUE" Then 'Active owners only
	where = " status = 'ACTIVE' "
Else 'Always exclude the label record
	where = " status <> 'LABEL' "
End If

If listall <> "TRUE" Then
	where = where & "AND ownerno <> 0 "
End If %>

<select name="ownername">			
			
<%
SQL="SELECT ownerno,ownername,ownerlong FROM ownerinfo WHERE " & where & " ORDER BY ownername"
%>
<!--#Include File="db_select.asp"-->	
<%
Do While Not rs.eof
	ownername = rs("ownername")
	ownerlist = ownername
	If longname = "TRUE" Then ownerlist = ownerlist & "&nbsp;&nbsp;" & rs("ownerlong")
	If selected_option = ownername Then
		selected = "selected"
	Else
		selected = ""
	End If %> 
					
	<option <%= selected %> value=<%= ownername %>><%= ownerlist %>	
			
<%	rs.movenext
Loop %> 
	        
</select>