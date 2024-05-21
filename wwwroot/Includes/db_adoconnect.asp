<%
Set adoConnect = Server.CreateObject("ADODB.Connection") 'Connect to the database

dbpath="DBQ=" & server.mappath("../../data/rec04a.mdb")
adoConnect.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & dbpath'

Set adoCommand = Server.CreateObject("ADODB.Command") 'Set up for SQL statement
Set adoCommand.ActiveConnection = adoConnect
adoCommand.CommandText = SQL
adoCommand.CommandType = 1
%>