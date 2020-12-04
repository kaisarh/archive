<!--#include file="../../inc_db.asp"-->
<html>
	<head>
		<title>View Wishes</title>
		<meta http-equiv="Content-Language" content="en-us" />
		<link rel="stylesheet" href="../../style.css" />
	
	</head>
	<body>
		<%
		Dim iNTotal, iNStart, iNEnd, iNCount, iPg
		Dim bFound, sPageURL, sPageTitle, sSrc, str
		
		''sPage is the page name, like "show_items.asp"

		sub page_link(s, pg_num)
			str = "<a href=" & chr(34) & sPageURL & "?pg=" & pg_num
			if sSrc <> "" then str = str & "&src=" & sSrc
			str = str & chr (34) & ">" & s & "</a>"

			response.write(str)
		end sub

		const PAGE_SIZE = 20
		sSrc = Request.form("Search")
		if sSrc = "" then sSrc = Request.querystring("src")

		sPageURL = "wish_view.asp"
		sPageTitle = "Wish List"

		sSQL = "SELECT * FROM WishList"
		if sSrc <> "" then	sSQL = sSQL & _
			" WHERE UserID LIKE '%" & sSrc & "%'" & _
			" OR Wish LIKE '%" & sSrc & "%'" & _
			" OR Price LIKE '%" & sSrc & "%'"
	
		sSQL = sSQL & " ORDER BY InsertedDate DESC"
			oRS.Open sSQL, oConn, 1
			iNTotal = oRS.RecordCount
		%>

			<table border=0 width="640" cellspacing=1 cellpadding=0>
				<tr>
					<td width="100%" colspan="3">
			    		<table border=0 width="640" cellspacing=1 cellpadding=0>
							<tr>
							    <td width="50%" align="left"><b><% =sPageTitle %></b></td>
							    <td align="right">
                                	<% if Session("bLogin") then %>
                                	<a href="../account.asp"
							    	>My Account</a> | 
							    	<% end if %>
							    	<a href="../../main.asp">Main</a> |
                                <a href="wish_add.asp">Add a wish!</a>
							    	</td>
							</tr>
						</table>
					</td>
				</tr>
			  	</tr>
			  	<tr>
			    	<td width="100%" colspan="3"><hr size="1" color="#FFFFFF"></td>
			    </tr>
   				<tr>
					<td width="100%" colspan="3" height="6"></td>
				</tr>
		<%	if not oRS.EOF then %>
				<tr>
					<td width="84%">
				<%
				''this is the current page number
				iPg = Request.querystring("pg")
				if iPg = "" then iPg = 1

				''page size should be set first
				oRS.PageSize = PAGE_SIZE
				oRS.AbsolutePage = iPg

				''other variables to handle
				iNStart = (iPg-1)*PAGE_SIZE + 1
				iNEnd = iNStart + PAGE_SIZE - 1
				if iNEnd > iNTotal then iNEnd = iNTotal
				iNCount = iNStart

				''now we set the Prev and next button	
				Response.write ("Total <b>" & iNTotal & "</b> Entries! ") 
				Response.write ("Displaying <b>" & iNStart & " to " & iNEnd & "</b>")

				''now we write the links
				%>
					</td>
					<td width="8%" align="right">
				<%		if iNStart <> 1 then page_link "Prev", iPg-1 %>
					</td>

					<td align="right">
				<%		if iNTotal > iNEnd then page_link "Next", iPg+1 %>
					</td>
				</tr>
				<tr>
					<td width="100%" colspan="3" height="6"></td>
				</tr>
				<tr>
					<td width="100%" colspan="3">
						<table class=list width="640">
								
								<tr align="middle">
									<th width="184" class=list><b>
                                    <font size="2">Name</font></b>
                                    <font size="2">&nbsp;</font>
                                    <th width="380" class=list><b>
                                    <font size="2">Wish</font></b><font size="2"></td>
                                    </font>
									<th width="156" class=list><b>
                                    <font size="2">Price</font></b><font size="2"></td>
                                    </font>
								</tr>
								
									<th width="120" class=list><font size="2">
                                    Date</font></tr>
								
								<% while (iNCount <= iNEnd and not oRS.EOF)
								sSQL2 = "SELECT * FROM Users WHERE UserID=" & oRS("UserID")
								oRS2.Open sSQL2, oConn, 1
								%>
							    	<tr align="middle">
										<td class=list width="184"><font size="2"><% =oRS2("Name") %>&nbsp;</font></td>
										<td class=list width="380"><font size="2"><% =oRS("Wish") %>&nbsp;</font></td>
										<td class=list width="156"><font size="2"><% =oRS("Price") %>&nbsp;</font></td>
										<td class=list width="120"><font size="2"><% =FormatDateTime(oRS("InsertedDate"),2) %>&nbsp;</font></td>
							   		</tr>
							<% 
						   		oRS2.Close
								oRS.MoveNext
								iNCount = iNCount + 1
							wend %>
						</table>
					</td>
				</tr>
				<tr align="middle">
					<td width="100%" colspan="3">
						<p align="left"><a href="javascript:history.back();">&lt;&lt;Back</a></p>
                        <p><b>Search:</b> Search the Wish List for 
                        specific items</p>
                        <p>
						<form method="POST" action="<%=sPageURL%>">
				  				<p>
				  					<input type="text" name="Search" size="34"> &nbsp;
				  					<input type="submit" value="Search" name="b1">
				  				</p>
						</form>
					</td>
				</tr>
		<% end if %>
		<tr align="middle">
			<td width="100%" colspan="3">
  <p align="center"><font size="2">Copyright © 2002, All rights reserved by
  <a href="../../about.htm" style="color: white">K. Haque</a></font></p>

			</td>
		</tr>
		</table>
		<% oRS.Close
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing %>
	</body>
</html>