<%
SQL = "UPDATE preference SET itemvalue = " & itemvalue & " where item = '" & item & "'"
%>
<!--#Include File="db_execute.asp"-->	
<%
Response.Clear
Server.Transfer version & "_login.asp" %>