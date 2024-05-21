<%
thispage = "pref_color"
pageaccess = "ADMIN"
%>
<!--#Include File="version.asp"-->

<%
item = "BACKGROUND"
itemvalue = ""
If thispage = lastpage Then 'Color is test color.  Change color if not test.
	red = CInt(Request("red"))
	green = CInt(Request("green"))
	blue = CInt(Request("blue"))
	background = red & "," & green & "," & blue	
	If Request("action") = "Change Color" Then
		itemvalue = "'" & background & "'"
	ElseIf Request("action") = "Reset to Default" Then
		itemvalue = "itemdefault"
	End If
	If Len(itemvalue) > 0 Then
%>
<!--#Include File="cd_pref_change.asp"-->	
<%
	End If
Else 'Color is color from preference table
	arybackground = Split(Session("background"),",")
	red = CInt(arybackground(0))
	green = CInt(arybackground(1))
	blue = CInt(arybackground(2))
End If
bgcolor = "#" & Hex(red) & Hex(green) & Hex(blue) %>	
	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Records Retention Change the Background Color</title>
</head>
<body bgcolor="<%= bgcolor %>">
<div align="center">

<!--#Include File="header.asp"-->	

<p><strong>Background Color</strong>

<form method="POST" action="<%= version %>_pref_color.asp">
	<p>
	Red&nbsp;
	<select name="red">
	
<%
i = 153
Do Until i > 255
	If i = red Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
		<option <%=selected%>><%= i %>
	
<%	i = i + 1
Loop %>
 	
	</select>
	&nbsp;
	Green&nbsp;
	<select name="green">
	
<%
i = 153
Do Until i > 255
	If i = green Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
		<option <%=selected%>><%= i %>
	
<%	i = i + 1
Loop %>
 	
	</select>
	&nbsp;
	Blue&nbsp;
	<select name="blue">
	
<%
i = 153
Do Until i > 255
	If i = blue Then
		selected = "selected"
	Else
		selected = ""
	End If %>
	
		<option <%=selected%>><%= i %>
	
<%	i = i + 1
Loop %>
 	
	</select>
	<p>
	<input type="submit" name="action" value="Test Color">
	<p>
	<input type="submit" name="action" value="Change Color">
	<p>
	<input type="submit" name="action" value="Reset to Default">
</form>

<table width="640" border="0" align="center">
	<tr>
		<td><font color="#0000FF">Background color may be changed from Gray (153,153,153) to White (255,255,255).  The standard color map uses values of 153, 204, and 255.  Intermediate values will be interpreted by the browser.  After changing a preference, you must Login again.</font> 
</table>
		
</div>
</body>
</html>
