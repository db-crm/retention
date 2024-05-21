<%
If Session("graphichead") = "FALSE" Then %>

<a href="../ASP/<%= version %>_menu.asp"><strong>MENU</strong></a>&nbsp;&nbsp;<strong><font size="+2">Records Retention</font></strong>&nbsp;&nbsp;<a href="../ASP/<%= version %>_quit.asp"><strong>QUIT</strong></a><br><strong><%= Session("client") %></strong>

<%
Else %>

<img src="../images/crmheader.gif" alt="Records Retention" width="640" height="60" border="0" usemap="#libhead">
<map name="libhead">
<area alt="Menu" coords="6,10,54,50" href="../ASP/<%= version %>_menu.asp">
<area alt="Quit" coords="586,10,634,50" href="../ASP/<%= version %>_quit.asp">
</map>

<%
End If %>