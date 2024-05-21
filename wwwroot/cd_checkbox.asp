<%'Make a checkbox
str = "check = v" & fieldname
Execute str
If check = "Y" Then
	checked = "checked"
Else
	checked = Empty
End If %>
	
&nbsp;<input type="checkbox" name="report_fields" value="<%= fieldname %>" <%= checked %>>
&nbsp;Display