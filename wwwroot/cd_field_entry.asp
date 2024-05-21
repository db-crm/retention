<% 'Make a text entry field.  Load text already submitted.
'Inputs are fieldname which will be submitted from the form, fieldlabel which will be displayed next
'to the entry field, field size, and the Y/N error flag [fieldname]_error.
'Outputs are a table row with labels and text entry field.  Labels are red if the error flag is set invalid.

str = "field_error = " & fieldname & "_error"
Execute str

labelcolor = "Black"
If lastpage = thispage Then
	lastvalue = Request(fieldname)
	If field_error = "Y" Then labelcolor = "Red"
Else 
	lastvalue = Empty
End If %>
	
<tr>			
	<td align="right"><font color=<%=labelcolor%>><%= fieldlabel %></font>
	<td><input type="text" name="<%= fieldname %>" value="<%=lastvalue%>" size="<%= fieldsize %>" maxlength="<%= fieldsize %>">
				
