<!--#include file="inc_show_1.asp"-->

	<%
		sPageURL = "show_banned.asp"
		sPageTitle = "Banned Users"

		sSQL = "SELECT * FROM BannedUsers"
		if sSrc <> "" then	sSQL = sSQL & _
			" WHERE Login LIKE '%" & sSrc & "%'" & _
			" OR Email LIKE '%" & sSrc & "%'"
	%>

<!--#include file="inc_show_2.asp"-->

	<tr align="middle">
		<td><strong>Login</strong></td>
		<td><strong>E-mail</strong></td>
	</tr>
	
	<% while (iNCount <= iNEnd and not oRS.EOF) %>
    	<tr align="middle">
			<td><% =oRS("Login") %></td>
			<td><% =oRS("Email") %></td>
  		</tr>

<!--#include file="inc_show_3.asp"-->
