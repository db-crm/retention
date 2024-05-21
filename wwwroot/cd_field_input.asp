<% 'Test an entered text field and return a query and error flag.
'Inputs are fieldname and fieldtype, where fieldname is the name of the field submitted by the form.
'Eg inputdate and date
'Outputs are the query clause q[fieldname] and the Y/N error flag [fieldname]_error.
'Eg qinputdate and inputdate_error.

input = Request(fieldname)
str = "rfield = r" & fieldname
Execute str
qfield = " "
errfield = "N"

If rfield = "Y" And Len(input) > 0 Then 'Not blank and is select criteria, check if valid input and make the query clause

	Select Case fieldtype
		Case "integer"
%>
<!--#Include File="cd_integer_test.asp"-->	
<%
		Case "date"
%>
<!--#Include File="cd_date_test.asp"-->	
<%
	End Select

End If
str = "q" & fieldname & " = qfield"
Execute str
str = fieldname & "_error = errfield"
Execute str %>