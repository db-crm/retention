<%'Get the owner name and long name using the owner number.
SQL = "SELECT ownername,ownerlong,contact FROM ownerinfo WHERE ownerno = " & ownerno & ""
%>
<!--#Include File="db_select.asp"-->
<%
If Not rs.eof Then
	ownername = rs("ownername")
	ownerlong = rs("ownerlong")
	contact =rs("contact")
End If %>