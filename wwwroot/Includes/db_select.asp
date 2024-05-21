<!--#Include File="../includes/db_adoconnect.asp"-->	
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 'Use client cursor
rs.Open adoCommand 'Do the select
Set rs.ActiveConnection  = Nothing 'Disconnect record set
%>
<!--#Include File="../includes/db_close.asp"-->	
