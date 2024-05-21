<%
inputnum = 0 'Zero is error indicator
delim_pos = InStr(input,delim) 'Single input 
If delim_pos = 0 Then
	If IsNumeric(input) Then inputnum = 1
ElseIf delim_pos = 1 Then 'Less Than single input	
	input = Right(input,Len(input)-1)
	If IsNumeric(input) Then inputnum = 3
ElseIf delim_pos = Len(input) Then 'Greater Than single input
	input = Left(input,delim_pos-1)
	If IsNumeric(input) Then inputnum = 4
Else 'Parse for begin and end of range
	input1 = Left(input,delim_pos-1)
	input2 = Right(input,Len(input)-delim_pos)
	If IsNumeric(input1) And IsNumeric(input2) And InStr(input2,delim) = 0 Then inputnum = 2 'Range ok
End If	

Select Case inputnum
	Case 1
		qfield = " and " & fieldname & " = " & input & ""
	Case 2
		qfield = " and " & fieldname & " between " & input1 & " and " & input2 & ""	
	Case 3
		qfield = " and " & fieldname & " <= " & input & ""
	Case 4
		qfield = " and " & fieldname & " >= " & input & ""
	Case Else 'Entry no good
		errors = errors + 1
		errfield = "Y"
End Select %>