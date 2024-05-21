<%'Get the owner number and long name using the owner name.
SQL = "SELECT ownerno,ownerlong,contact FROM ownerinfo WHERE ownername = '" & ownername & "'"
%>
<!--#Include File="db_select.asp"-->
<%
If Not rs.eof Then
	ownerno = rs("ownerno")
	ownerlong = rs("ownerlong")
	contact =rs("contact")
End If %>