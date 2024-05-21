<%
inputnum = 0 'Zero is error indicator
delim_pos = InStr(input,delim) 'Get the location of the delimiter in 'input'
datedelim = Session("datedelim")

If delim_pos = 0 Then 'No delimiter, check for a valid single date
	If IsDate(input) Then inputnum = 1 'Valid single date
ElseIf delim_pos = 1 Then 'Delimiter is first, high end of open range
	input = Right(input,Len(input)-1) 'Strip delimiter
	If IsDate(input) Then inputnum = 3 'Valid high end date
ElseIf delim_pos = Len(input) Then 'Delimiter is last, low end of open range
	input = Left(input,delim_pos-1) 'Strip delimiter
	If IsDate(input) Then inputnum = 4 'Valid low end date
Else 'Parse for begin and end of range
	input1 = Left(input,delim_pos-1)
	input2 = Right(input,Len(input)-delim_pos)
	If IsDate(input1) And IsDate(input2) And InStr(input2,delim) = 0 Then inputnum = 2 'Range ok
End If

Select Case inputnum
	Case 1
		qfield = " AND " & fieldname & " = " & datedelim & input & datedelim
	Case 2
		qfield = " AND " & fieldname & " BETWEEN " & datedelim & input1 & datedelim & " AND " & datedelim & input2 & datedelim
	Case 3
		qfield = " AND " & fieldname & " <= " & datedelim & input & datedelim
	Case 4
		qfield = " AND " & fieldname & " >= " & datedelim & input & datedelim
	Case Else 'Entry no good
		errors = errors + 1
		errfield = "Y"
End Select %>