<% 'Put a record into the history table.  Inputs - recordno and histaction
datedelim = Session("datedelim")

SQL = "SELECT site,location FROM location WHERE recordno = " & recordno & ""
%>
<!--#Include File="../includes/db_select.asp"-->
<%
If Not rs.eof Then
	histlocation = rs("site") & "," & rs("location")
Else
	histlocation = "NONE"
End If

SQL = "SELECT ownername,box,boxid,retention,record.status AS status,disposition,entrydate,updatedate,recorddate,reviewdate,deletedate,content FROM record INNER JOIN ownerinfo ON record.ownerno = ownerinfo.ownerno WHERE recordno = " & recordno & ""
%>
<!--#Include File="../includes/db_select.asp"-->	
<%
histownername = rs("ownername")
histbox = rs("box")
histrecord = recordno & "," & rs("boxid") & "," & rs("retention") & "," & rs("status") & "," & rs("disposition") & "," & rs("entrydate") & "," & rs("updatedate") & "," & rs("recorddate") & "," & rs("reviewdate") & "," & rs("deletedate") & "," & rs("content")

SQL = "INSERT INTO recordhistory(ownername,box,cleanupdate,action,location,record) VALUES('" & histownername & "'," & histbox & "," & datedelim & Date & datedelim & ",'" & histaction & "','" & histlocation & "','" & histrecord & "')"
%>
<!--#Include File="../includes/db_execute.asp"-->