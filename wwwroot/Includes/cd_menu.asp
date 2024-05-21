<%
If menu_select = menu_label Then
	Response.Clear
	Server.Transfer version & "_" & gotofile & ".asp"
End If %>	

<p><input type=submit name=menu_select value="<%=menu_label%>">