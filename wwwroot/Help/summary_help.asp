<p>
<table width="65%" border="0" align="center">
	<tr>
		<td><font color="#0000FF"><strong>Range Delimiter</strong></font>.  Use a '<%= Session("rangedelim") %>', to enter a range in the Date Range field.  For example, 1/1/99<%= Session("rangedelim") %>1/1/00.  Ranges are inclusive and the start or end of the range may be omitted.  For example, 1/1/01<%= Session("rangedelim") %>, will create a summary for dates greater then 1/1/01 and <%= Session("rangedelim") %>1/1/01) will create a summary for dates less than 1/1/01.  The Date Range field will accept most common date formats. 
</table>

