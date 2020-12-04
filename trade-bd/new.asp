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
			str = "<a href=" & chr(34) & "new.asp?iPg=" & pg_num
			str = str & chr (34) & ">" & s & "</a>"
			response.write(str)
		end sub

		Dim sPUnit, strItem
		Dim iNTotal, iNStart, iNEnd, iNCount, iPg
		Dim str, iDB, iCode

		const PAGE_SIZE = 10

		sSql = "SELECT * FROM Items ORDER BY InsertedDate DESC"
		oRS.Open sSql, oConn, 1

		iNTotal = oRS.RecordCount
	%>
<table width="640">
  <tr>
    <td width="100%" colspan="3"><%
							Response.write ("<b>Listing New items:</b> ") 
							if iNTotal > 0 then
								Response.write ("Total <b>" & iNTotal & "</b> item(s) found!") 
							else
								Response.write ("Sorry, no item was found!")
							end if
						%> </td>
  </tr>
  <%	if not oRS.EOF then %>
  <tr>
    <td width="84%"><%
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
    <td width="8%" align="right"><%		if iNStart <> 1 then page_link "Prev", iPg-1 %>
    </td>
    <td align="right"><%		if iNTotal > iNEnd then page_link "Next", iPg+1 %>
    </td>
  </tr>
</table>
<table class="list" width="640">
  <tr>
    <th class="list" width="5%">S/N</th>
    <th class="list" width="20%">Inserted Date</th>
    <th class="list" width="50%">Item Description</th>
    <th class="list" width="8%">Status</th>
    <th class="list">Price</th>
  </tr>
  <%	sCol = Col1
					while (iNCount <= iNEnd and not oRS.EOF)
					
							iCode 	= oRS("iCode")
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

							sSql = "SELECT * FROM " & sSql & " WHERE ItemID=" & oRS("ItemID") & " AND iCode=" & iCode '& " AND iStatus<=1"
							oRS2.Open sSql, oConn, 1

							if iCode>30  then
								strItem = oRS2("Author") & " - " & oRS2("Title")
							elseif iCode=20 then
								strItem = oRS2("Processor") & " " & oRS2("Speed") & "MHz, " & _
								oRS2("RAM") & " MB RAM, " & oRS2("HDD") & " GB HDD, <br>&nbsp;&nbsp;" & _
								oRS2("Display") & ", " & oRS2("MonitorSize") & chr(34) & " " & oRS2("Monitor") & " Monitor"
							else
								strItem = oRS2("Brand") & " " & oRS2("Name") & "<br>&nbsp;&nbsp;" & oRS2("Spec")
								if oRS2("Acc")<> "" then strItem = strItem & " with " & oRS2("Acc")
							end if %>
  <tr>
    <td class="list" align="middle"><font size="2"><% =iNCount %></font></td>
    <td class="list" align="middle"><font size="2"><% =oRS("InsertedDate") %></font></td>
    <td class="list"><% Response.Write("<font size=2>" & strItem & _
							"<a class=table href='account/trade/item_details.asp?iItemID=" & oRS("ItemID") & "&iCode=" & iCode & "'>[Details]</a></font>")

						if oRS("PriceUnit").value=1 then
							sPUnit = "Taka" 
						else 
							sPUnit = "US$"
						end if

						Response.Write("</td><td class=list align='center'><font size=2>")
						if oRS("iStatus")=1 then
							Response.Write("Rsrv")
						elseif oRS("iStatus")=2 then
							Response.Write("Sold")
						else
							Response.Write("&nbsp;")
						end if

						Response.Write("</font></td><td class=list valign='middle' align='middle'><font size=2>" & oRS("Price") & " " & sPUnit & "</font></td></tr>")

						oRS2.Close
						oRS.MoveNext

						iNCount = iNCount + 1
					wend
					oRS.Close %> </td>
  </tr>
</table>
</td>
</tr>
</table>
<table width="640">
  <tr>
    <td width="63"><a href="main.asp">&lt;&lt;Back</a> </td>
    <td><b>Note: </b>Click on the details to get detials about the items. </td>
  </tr>
  <% end if %>
</table>
<br>
<% if not (sSrc="" and iNTotal <1) then %> <% end if %> <%
		'oConn.Close
		'Set oRS = Nothing
		Set oRS2 = Nothing
		'Set oConn = Nothing %>
<table width="640">
  <tr>
    <td width="100%">
    <p align="center"><font size="2">Copyright © 2002, All rights reserved by
    <a href="about.htm">K. Haque</a></font></p>
    </td>
  </tr>
</table>

</body>

</html>