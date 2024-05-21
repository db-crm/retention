<%
thispage = "expired_process"
pageaccess = "SUPERADMIN"
%>
<!--#Include File="version.asp"-->
<%
'This page makes a list of record numbers to delete and performs the deletes, if any; makes a list of record numbers to update and transfers to the Update page, or returns to Expired page if no updates.  It generates no client page

'Look through the Form collection for records selected for delete or update

rnum = Request.Form.Count 'Number of items in the collection
updatelist = Empty
datedelim = Session("datedelim")
i = 1
Do Until i = rnum + 1
	Select Case Request.Form.Item(i)
		Case "rrdelete" 'Record selected for delete
			recordno = Right(Request.Form.Key(i),Len(Request.Form.Key(i))-2) 'Record number is right part of name
			j = 1
			Do Until j = rnum + 1
				If Left(Request.Form.Key(j),2) = "rc" And recordno = Right(Request.Form.Key(j),Len(Request.Form.Key(j))-2) Then
'Find the disposition for the record number
					SQL="UPDATE record SET disposition = '" & Request.Form.Item(j) & "',deletedate = " & datedelim & Date & datedelim & ",status = 'DELETED' WHERE recordno = " & recordno & ""
%>
<!--#Include File="db_execute.asp"-->	
<%
'Get the site and location being vacated
					SQL="SELECT site,location FROM location WHERE recordno = " & recordno & "" 
%>
<!--#Include File="db_select.asp"-->
<%
					site = rs("site")
					location = rs("location")
'Put the vacated site and location into the oldlocation table
					SQL="INSERT INTO oldlocation(site,location,recordno) VALUES('" & site & "','" & location & "'," & recordno & ")"
%>
<!--#Include File="db_execute.asp"-->
<%
'Make the location vacant
					SQL="UPDATE location SET recordno = 0 WHERE recordno = " & recordno & "" 
%>
<!--#Include File="db_execute.asp"-->
<%
					Exit Do 'Found the disposition and did the delete, exit the loop
				End If
				j = j + 1
			Loop	
		Case "rrupdate" 'Record selected for update
			recordno = Right(Request.Form.Key(i),Len(Request.Form.Key(i))-2) 'Record number is right part of name (key)			
			updatelist = updatelist & "," & recordno 'Add record number to the update list
	End Select	
	i = i + 1
Loop

If Not IsEmpty(updatelist) Then
	updatelist = Right(updatelist,Len(updatelist)-1) 'Strip leading comma
	Session("updatelist") = updatelist 'Set session variable for update page to use
	Response.Clear
	Server.Transfer version & "_expired_update.asp" 'Go to update page
Else
	Response.Clear
	Server.Transfer version & "_expired.asp" 'Nothing to update, return to Expired page
End If %>

