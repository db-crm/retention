<p>
<table width="65%" border="0" align="center">
	<tr>
		<td><font color="#0000FF"><strong>Selection Criteria</strong></font>.  Enter the selection criteria for the report you want.  A blank value will return records for all values of the field.  Use a '<%= Session("rangedelim") %>' to enter a range in the Retention, Box, or Date fields.  For example, 3<%= Session("rangedelim") %>5.  Ranges are inclusive and the start or end of the range may be omitted.  For example, 3<%= Session("rangedelim") %> will search for items greater than or equal to 3, and <%= Session("rangedelim") %>5 will search for items less than or equal to 5.  Box and Retention must be integers.  Date fields will accept most common date formats.  Fields needing correction will be labeled red.
	<tr>
		<td><font color="#0000FF"><strong>Content</strong></font>.  Enter a string if you want to search for it in the Content field.  Content is stored uppercase and your entry will be converted before searching.  Select 'Full Content' to include the entire Content field in the report.  Select 'Short Content' to include a truncated Content field in the report.  Select 'No Content' for a report with no Content field.
	<tr>
		<td><font color="#0000FF"><strong>Display Dates</strong></font>.  Check the box after the dates you want to be displayed in the report.  A date may be a selection criteria even if it isn't displayed.
	<tr>
		<td><font color="#0000FF"><strong>Selection Method</strong></font>.  Remember that if you enter selection criteria in more than one field, all criteria must be met to select the record for the report.  The select query is 'criteria 1' AND 'criteria 2', not 'criteria 1' OR 'criteria 2'.
	<tr>
		<td><font color="#0000FF"><strong>Report Page</strong></font>.  The report page is text-only to facilitate printing with your browser print command.  Use the print preview and setup functions to change the margins and so forth.  To exit the report page, click the 'report type' link at the top of the page.
</table> 
 

