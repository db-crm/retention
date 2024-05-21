<%
Response.Buffer = True
version = "rec04a"
If Session("page") = "menu" Then Session("returnto") = "menu"
If Len(Session("accesstype")) = 0 Then Session("accesstype") = "NONE"
If thispage = "login" Or InStr(pageaccess,Session("accesstype")) > 0 Then
	lastpage = Session("page")
	Session("page") = thispage
Else
	Response.Clear
	Server.Transfer version & "_login.asp"
End If %>
