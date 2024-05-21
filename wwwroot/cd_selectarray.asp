<%'Use selectname, aryoptionvalue, and aryoptionlabel to make a select list
If lastpage = thispage Then 
	selected_option = Request(selectname)
Else 
	selected_option = Session(selectname)
End If %>			
			
<select name="<%= selectname %>">

<%
i = 0
Do Until i > UBound(aryoptionvalue)
	optionvalue = aryoptionvalue(i)
	optionlabel = aryoptionlabel(i)
	If selected_option = optionvalue Then
		selected = "selected"
	Else
		selected = ""
	End If %>

	<option value="<%= optionvalue %>" <%=selected%>><%= optionlabel %>

<%	i = i + 1
Loop %>	
		
</select>