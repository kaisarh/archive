<!--#include file="inc_show_1.asp"-->

	<%
		sPageURL = "show_banned_ip.asp"
		sPageTitle = "Banned IPs"

		sSQL = "SELECT * FROM BannedIP"
		if sSrc <> "" then	sSQL = sSQL & _
			" WHERE IP LIKE '%" & sSrc & "%'"
	%>

<!--#include file="inc_show_2.asp"-->

	<tr align="middle">
		<td><strong>IP</strong></td>
	</tr>
	
	<% while (iNCount <= iNEnd and not oRS.EOF) %>
    	<tr align="middle">
			<td><% =oRS("IP") %></td>
  		</tr>

<!--#include file="inc_show_3.asp"-->
