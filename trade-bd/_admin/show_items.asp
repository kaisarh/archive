<!--#include file='inc_admin_login.asp'-->

<html>
	<head>
		<title>
		</title>
	</head>
	<body>

		<%
		
		Dim oConn, oRS, sSQL, oConn2, oRS2
		Dim iNTotal, iNStart, iNEnd, iNCount, iPg
		Dim bFound
		
		Dim strType, strPrice, strDetail

		''//////////////////////////////// Sub //////////////////////////////////////
		sub get_detail(itemID, dbtype)
			
			if bFound=1 then oRS2.Close
			
			strType = "ERROR in DataBase!"
			strDetail = ""
			strPrice = ""
			
			select case dbtype
				case 1:
					sSQL = "SELECT * FROM Alls WHERE ItemID=" & itemID
					oRS2.Open sSQL, oConn2, 1
	
					if not oRS2.EOF then
						select case oRS2("Type")
							case 1:
								str = "PC Accessory"
								sType = "PCAcc"
							case 2:
								str = "Electronic"
								sType = "Elec"
							case 3:
								str = "Car"
								sType = "Car"
							case 4:
								str = "Other"
								sType = "Other"
						end select
		
						strType = "<strong>Type: </strong>" & str
						strDetail = oRS2("Brand") & " " & oRS2("Name") & " " & oRS2("Spec")
					end if
					
				case 2:
					sType = "PC"
					sSQL = "SELECT * FROM PCs WHERE ItemID=" & itemID
					oRS2.Open sSQL, oConn2, 1
	
					if not oRS2.EOF then
						strType = "<strong>Type: </strong>Computer"
						
						''TO DO: correct this details
						
						strDetail = oRS2("Processor") & ", " & oRS2("RAM") & " MB RAM, " & oRS2("HDD") & " GB HDD, " & _
									oRS2("Display") & ", " & oRS2("Monitor") & " Monitor "
					end if
					
				case 3:
					sSQL = "SELECT * FROM CDsNBooks WHERE ItemID=" & itemID
					oRS2.Open sSQL, oConn2, 1
					
					if not oRS2.EOF then
						select case oRS2("Type")
							case 1:
								str = "Audio/Video/MP3 CD"
								sType = "CD"
							case 2:
								str = "Book"
								sType = "Book"
							case 3:
								str = "DVD"
								sType = "DVD"
							case 4:
								str = "Software/Game CD"
								sType = "Soft"
						end select
						
						strType = "<strong>Type: </strong>" & str
						strDetail = oRS2("Author") & " - " & oRS2("Title")
					end if
					
			end select
		
			if not oRS2.EOF then
				strPrice = "<strong>Price: </strong>" & oRS2("Price")
				if oRS2("PriceUnit")=1 then
					strPrice = strPrice & " Taka"
				else
					strPrice = strPrice & " US$"
				end if
				strDetail = "<strong>Details: </strong>" & strDetail
			end if
			
			bFound = 1
		end sub
		''//////////////////////////////// End Sub //////////////////////////////////////

		
		sub page_link(s, pg_num)
			response.write("<a href=" & chr(34) & "show_items.asp?pg=" & pg_num & _
					chr (34) & ">" & s & "</a>")
		end sub

		Response.Expires = -1000
		
		const PAGE_SIZE = 20
		
		''let's create our objects	
		
		Set oConn = Server.CreateObject("ADODB.Connection")
		Set oConn2 = Server.CreateObject("ADODB.Connection")
		Set oRS = Server.CreateObject( "ADODB.Recordset" )
		Set oRS2 = Server.CreateObject( "ADODB.Recordset" )
		
		oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/data.mdb")
		oConn2.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("\tradebd\db\mabnfopiw12_a3.mdb")) ' & "; PWD=harry")
		
		sSQL = "SELECT * FROM Items ORDER BY ItemID"
		
		oRS.Open sSQL, oConn, 1
		iNTotal = oRS.RecordCount
	%>
		<center>
			<table border=0 width="640" cellspacing=1 cellpadding=0>
				<tr>
				    <td width="50%" colspan="3"><p align="left">Admininstrator's Page: <strong>Items</strong></td>
			  	</tr>
			  	<tr>
			    	<td width="100%" colspan="3"><hr size="1"></td>
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
				
				Response.write ("Total <Strong>" & iNTotal & "</Strong> Entries! ") 
				Response.write ("Displaying <Strong>" & _
					iNStart & " to " & iNEnd & "</Strong>")
		
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
					
					<% while (iNCount <= iNEnd and not oRS.EOF)
						
						itemID=oRS("ItemID")
		   				dbtype=oRS("dbType")
		   				get_detail itemID, dbtype
		   				%>
					    <table border="1" width="100%" bordercolordark="#FFFFFF" bgcolor="#EAF8DE" cellspacing="0" cellpadding="0"
					    	bordercolorlight="#000000">
					   		<tr>
						        <td width="75%">
									<% response.write(strType) %><br>
									<% response.write(strDetail) %><br>
									<% response.write(strPrice) %>
						        </td>
				      		</tr>
			    		</table>
			    		<table border="0" width="100%" cellspacing="0" cellpadding="0">
					      <tr>
					        <td width="100%" height="10"></td>
					      </tr>
					    </table>
						
						<% oRS.MoveNext
						iNCount = iNCount + 1
					wend %>
					</td>
				</tr>
			<% end if %>
		</table>
		Copy Right 2002, All rights reserved by MyMarket.Com.
		</center>
		<%
		oRS.Close
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing %>
	</body>
</html>
