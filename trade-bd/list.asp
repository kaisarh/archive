<!--#include file="inc_db.asp"-->
<html>
	<head>
	<title>List of items</title>
	<meta http-equiv="Content-Language" content="en-us" />
	<link rel="stylesheet" href="style.css" />

	</head>
	<body>
	<%	
		sub page_link(s, pg_num)
			str = "<a href=" & chr(34) & "list.asp?iCode=" & iCode & "&iPg=" & pg_num
			if sSort <> "" then str = str & "&sSort=" & sSort
			if sSrc <> "" then str = str & "&sSearch=" & sSrc
			str = str & chr (34) & ">" & s & "</a>"
			response.write(str)
		end sub

		Dim sPUnit, strItem
		
		Dim iNTotal, iNStart, iNEnd, iNCount, iPg
		Dim sSort, sSrc, str
		Dim iDB, iCode

		const PAGE_SIZE = 10

		iCode 	= Request.querystring("iCode")
		sSort 	= Request.querystring("sSort")
		sSrc	= Request.querystring("sSearch")			'from next, prev buttons
		
		if sSrc = "" then sSrc	= Request.form("search")	'first time searching

		iDB = int(iCode/10)

		''first we select the table
			select case iDB
				case 1
					sSql="Alls"
				case 2
					sSql="PCs"
				case 3
					sSql="CDsNBooks"
			end select

			sSql = "SELECT * FROM " & sSql & " WHERE iCode=" & iCode & " AND iStatus<=1"

		''then we add the search string
			if sSrc <> "" then
				if iCode>30 then
						str = " AND (Author LIKE '%" & sSrc & "%' OR Title LIKE '%" & sSrc & "%')"

				elseif iCode=20 then
						str = " AND (Brand LIKE '%" & sSrc & "%' OR Processor LIKE '%" & sSrc & "%'" & _
							"OR Speed LIKE '%" & sSrc & "%' OR RAM LIKE '%" & sSrc & "%'" & _
							"OR HDD LIKE '%" & sSrc & "%' OR MB LIKE '%" & sSrc & "%'" & _
							"OR Display LIKE '%" & sSrc & "%' OR Sound LIKE '%" & sSrc & "%'" & _
							"OR Modem LIKE '%" & sSrc & "%' OR Monitor LIKE '%" & sSrc & "%'" & _
							"OR SoundBit LIKE '%" & sSrc & "%' OR ModemSpeed LIKE '%" & sSrc & "%'" & _
							"OR MonitorSize LIKE '%" & sSrc & "%')"

				else
						str = " AND (Brand LIKE '%" & sSrc & "%'	OR Name LIKE '%" & sSrc & "%'" & _
							"OR Brand LIKE '%" & sSrc & "%' OR Spec LIKE '%" & sSrc & "%'" & _
							"OR Acc LIKE '%" & sSrc & "%')"
				end if
				sSql = sSql & str
			end if

		if sSort <> "" then sSql = sSql & " ORDER BY " & sSort
		oRS.Open sSql, oConn, 1

		iNTotal = oRS.RecordCount
	%>
			<table width="640">
				<tr>
					<td width="100%" colspan="3">
						<% if sSrc <> "" then 
							Response.write ("Search result for '<b>" & sSrc & "</b>': ") 
							if iNTotal > 0 then
								Response.write ("Total <b>" & iNTotal & "</b> item(s) found!") 
							else
								Response.write ("Sorry, no item was found!") 
							end if
						else
							Response.write ("<b>Listing items:</b> ") 
							if iNTotal > 0 then
								Response.write ("Total <b>" & iNTotal & "</b> item(s) available!") 
							else
								Response.write ("Sorry, no item is available!") 
							end if
						end if%>
					</td>
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
						Response.write ("<br>Displaying <b>" & _
							iNStart & " to " & iNEnd & "</b>")
				%></td>
					<td width="8%" align="right">
				<%		if iNStart <> 1 then page_link "Prev", iPg-1 %>
					</td>

					<td align="right">
				<%		if iNTotal > iNEnd then page_link "Next", iPg+1 %>
					</td>
				</tr>
			</table>
				
			<table class=list width="640">
			   		<tr>
				        <th class=list width="38">S/N</th>
				        <th class=list width="467" >Item Description</th>
				        <th class=list width="25" >R</th>
				        <th class=list align="middle" width="82">Price</th>
			      	</tr>

				<%	sCol = Col1
					while (iNCount <= iNEnd and not oRS.EOF) %>
					
						<tr>
							<td class=list align="middle" width="38"><font size=2><% =iNCount %></font></td>
							<td class=list width="467">
	
						<% if iCode>30  then
								strItem = oRS("Author") & " - " & oRS("Title")
						
						elseif iCode=20 then
								strItem = oRS("Processor") & " " & oRS("Speed") & "MHz, " & _
								 oRS("RAM") & " MB RAM, " & oRS("HDD") & " GB HDD, <br>&nbsp;&nbsp;" & _
								 oRS("Display") & ", " & oRS("MonitorSize") & chr(34) & " " & oRS("Monitor") & " Monitor"
						
						else
								strItem = oRS("Brand") & " " & oRS("Name") & "<br>&nbsp;&nbsp;" & oRS("Spec")
								if oRS("Acc")<> "" then strItem = strItem & " with " & oRS("Acc")
						end if

						sSql = "SELECT * FROM Items WHERE ItemID=" & oRS("ItemID") & " AND iCode=" & oRS("iCode")
						oRS2.Open sSql, oConn, 1

						Response.Write("<font size=2>" & strItem & "" & _
							"<a class=table href='account/trade/item_details.asp?iItemID=" & oRS("ItemID") & "&iCode=" & iCode & "'>[Details]</a></font>")

						if oRS2("PriceUnit").value=1 then sPUnit = "Taka" else sPUnit = "US$"

						Response.Write("</td><td class=list align='center'><font size=2>")
						if oRS("iStatus")=1 then
							Response.Write("R")
						else
							Response.Write("&nbsp;")
						end if
						
						Response.Write("</font></td><td class=list valign='middle' align='middle'><font size=2>" & oRS2("Price") & " " & sPUnit & "</font></td></tr>")

						oRS2.Close
						oRS.MoveNext
						iNCount = iNCount + 1
					wend
					oRS.Close %>
						</table>
					</td>
				</tr>
			</table>
			<table width="640">
				<tr>
					<td width="63"><a href="main.asp">&lt;&lt;Back</a>
					</td>
					<td><b>Note: </b>Click on the details to get detials about the items.
					</td>
				</tr>
	<% end if %>
			</table>
			<br>
			<% if not (sSrc="" and iNTotal <1) then %>
			<table width="640">
				<tr>
					<td>
						<b>Search:</b>
								<% if sSrc <> "" then
									Response.write("You can search again with other keywords.")
								else
									Response.write("To minimize the list you can search with some keywords.")
								end if %><p>
						<form method="POST" action="list.asp?iCode=<% =(iCode) %>">
				  			<center>
				  				<p>
				  					<input type="text" name="Search" size="34"> &nbsp;
				  					<input type="submit" value="Search" name="b1">
				  				</p>
								<p><font size="2">Copyright © 2002, All rights reserved by
      <a href="about.htm">K. Haque</a></font></p>
				  			</center>
						</form>
						<p>
					</td>
				</tr>
			</table>
			<% end if %>
	<%
		'oConn.Close
		'Set oRS = Nothing
		Set oRS2 = Nothing
		'Set oConn = Nothing %>
	</body>
</html>