<% 'Use selected_option, optionvalue, and optionlabel to make a select list option
If selected_option = optionvalue Then
	selected = "selected"
Else
	selected = ""
End If %>

<option value="<%= optionvalue %>" <%=selected%>><%= optionlabel %>