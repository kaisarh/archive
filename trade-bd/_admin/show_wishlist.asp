<!--#include file="inc_show_1.asp"-->

	<%
		sPageURL = "show_wishlist.asp"
		sPageTitle = "Wish List"

		sSQL = "SELECT * FROM WishList"
		if sSrc <> "" then	sSQL = sSQL & _
			" WHERE UserID LIKE '%" & sSrc & "%'" & _
			" OR Wish LIKE '%" & sSrc & "%'" & _
			" OR Price LIKE '%" & sSrc & "%'"
	%>

<!--#include file="inc_show_2.asp"-->
	
	<tr align="middle">
		<td><strong>User ID</strong></td>
		<td><strong>Wish</strong></td>
		<td><strong>Price</strong></td>
	</tr>
	
	<% while (iNCount <= iNEnd and not oRS.EOF) %>
    	<tr align="middle">
			<td><% =oRS("UserID") %></td>
			<td><% =oRS("Wish") %></td>
			<td><% =oRS("Price") %></td>
   		</tr>

<!--#include file="inc_show_3.asp"-->
