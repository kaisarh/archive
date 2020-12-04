<!--#include file="inc_show_1.asp"-->

	<%
		sPageURL = "show_guestbook.asp"
		sPageTitle = "Guest Book"

		sSQL = "SELECT * FROM GuestBook"
		if sSrc <> "" then	sSQL = sSQL & _
			" WHERE Name LIKE '%" & sSrc & "%'" & _
			" OR Email LIKE '%" & sSrc & "%'" & _
			" OR Comment LIKE '%" & sSrc & "%'"
	%>

<!--#include file="inc_show_2.asp"-->

	<tr align="middle">
		<td><strong>Name</strong></td>
		<td><strong>E-mail</strong></td>
		<td><strong>Comment</strong></td>
	</tr>
	
	<% while (iNCount <= iNEnd and not oRS.EOF) %>
    	<tr align="middle">
			<td><% =oRS("Name") %></td>
			<td><% =oRS("Email") %></td>
			<td><% =oRS("Comment") %></td>
  		</tr>

<!--#include file="inc_show_3.asp"-->
