<!--#include file="inc_show_1.asp"-->

	<%
		sPageURL = "show_users.asp"
		sPageTitle = "Users"

		sSQL = "SELECT * FROM Users"
		if sSrc <> "" then	sSQL = sSQL & _
			" WHERE UserID LIKE '%" & sSrc & "%'" & _
			" OR Login LIKE '%" & sSrc & "%'" & _
			" OR Name LIKE '%" & sSrc & "%'" & _
			" OR Email LIKE '%" & sSrc & "%'" & _
			" OR BuyCount LIKE '%" & sSrc & "%'" & _
			" OR SellCount LIKE '%" & sSrc & "%'"
	%>

<!--#include file="inc_show_2.asp"-->
	
	<tr align="middle">
        <td><strong>ID</strong></td>
		<td><strong>Login</strong></td>
		<td><strong>Name</strong></td>
		<td><strong>E-mail</strong></td>
		<td><strong>Buy</strong></td>
		<td><strong>Sell</strong></td>
	</tr>
	
	<% while (iNCount <= iNEnd and not oRS.EOF) %>
    	<tr align="middle">
            <td><% =oRS("UserID") %></td>
			<td><% =oRS("Login") %></td>
			<td><% =oRS("Name") %></td>
			<td><% =oRS("Email") %></td>
			<td><% =oRS("BuyCount") %></td>
			<td><% =oRS("SellCount") %></td>
  		</tr>

<!--#include file="inc_show_3.asp"-->
