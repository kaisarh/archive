<!--#include file="check_ban.asp"-->
<!--#include file='inc_acc_login.asp'-->
	<%	
		const MAX_SELL = 12
		const MAX_BUY = 2

		Dim oRS2, str, i

		Dim itemID, iCode, iDB, iUserID, bFound
		Dim strType, strPrice, strDetail
		Dim iSellCount, iBuyCount

		iUserID = Session("iUserID")
		sSQL = "SELECT * FROM Users WHERE UserID=" & iUserID
		oRS.Open sSQL, oConn, 1

		iSellCount = oRS("SellCount")
		iBuyCount = oRS("BuyCount")
	%>

	<%	
		'Dim oConn, oRS, oRS2, sSQL
		''//////////////////////////////// Sub //////////////////////////////////////
		sub get_detail(itemID, iCode)
	
			''if bFound=1 then oRS2.Close
			strType = "ERROR"
			strDetail = "Error"
			strPrice = "Error"
			
			iDB = int(iCode/10)
			
			select case iDB
				case 1:
					sSQL = "SELECT * FROM Alls WHERE ItemID=" & itemID
					oRS2.Open sSQL, oConn, 1
	
					if not oRS2.EOF then
						select case iCode
							case 11:
								str = "PC Accessory"

							case 12:
								str = "Electronic"

							case 13:
								str = "Vehicle"

							case 14:
								str = "Other"
						end select
		
						strType = "<b>Type: </b>" & str
						strDetail = oRS2("Brand") & " " & oRS2("Spec") & " " & oRS2("Name")
					end if
					
				case 2:
					sSQL = "SELECT * FROM PCs WHERE ItemID=" & itemID
					oRS2.Open sSQL, oConn, 1
	
					if not oRS2.EOF then
						strType = "<b>Type: </b>Computer"
						
						''TO DO: correct this details
						
						strDetail = oRS2("Processor") & ", " & oRS2("RAM") & " MB RAM, " & oRS2("HDD") & " GB HDD, " & _
									oRS2("Display") & ", " & oRS2("Monitor") & " Monitor "
					end if
					
				case 3:
					sSQL = "SELECT * FROM CDsNBooks WHERE ItemID=" & itemID
					oRS2.Open sSQL, oConn, 1
					
					if not oRS2.EOF then
						select case iCode
							case 31:
								str = "Audio/Video/MP3 CD"

							case 32:
								str = "Book"

							case 33:
								str = "DVD"

							case 34:
								str = "Software/Game CD"
						end select
						
						strType = "<b>Type: </b>" & str
						strDetail = oRS2("Author") & " - " & oRS2("Title")
					end if
			end select

			''response.write(iDB & "<br>")
			''response.write(iCode & "<br>")
			''response.write(sSql & "<br>")
			''response.write(strType & "<br>")
			''response.end

			oRS2.Close
			sSQL = "SELECT * FROM Items WHERE ItemID=" & itemID	 & " AND iCode=" & iCode
			oRS2.Open sSQL, oConn, 1
	
			if not oRS2.EOF then
				strPrice = "<b>Price: </b>" & oRS2("Price")
				if oRS2("PriceUnit")=1 then
					strPrice = strPrice & " Taka"
				else
					strPrice = strPrice & " US$"
				end if
				strDetail = "<b>Details: </b>" & strDetail
			end if

			oRS2.Close
			bFound = 1
		end sub
		''//////////////////////////////// End Sub //////////////////////////////////////
	%>

<html>
	<head>
		<title>User Management Page</title>

		<meta http-equiv="Content-Language" content="en-us" />
		<link rel="stylesheet" href="../style.css" />

	</head>
	<body>
		<table width="640">
			<tr>
			    <td width="50%"><p align="left">Login ID: <% =oRS("Login") %></td>
			    <td width="50%"><p align="right"><a href="change_pass.asp">
                <font size="2">Change Password</font></a><font size="2"> |
                <a href="add_edit/edit_user.asp">Edit Profile</a> | </font> <a href="log_out.asp">
                <font size="2">Log Out</font></a></td>
		  	</tr>
		  	<tr>
		    	<td width="100%" colspan="2"><hr size="1" color="#FFFFFF"></td>
		    </tr>
		    <tr>
		    	<td width="100%" colspan="2">Welcome <% =oRS("Name") %>,<p><b>
                System Messages:</b>
            <% 
            oRS.Close

			sSQL = "SELECT * FROM Notices WHERE UserID=" & iUserID & " ORDER BY InsertedDate DESC"
			oRS.Open sSQL, oConn, 1

			 if Not oRS.EOF then %>
				&nbsp;<a href="clear_msg.asp">Clear</a><br>
                <font size="1">&nbsp;</font><br>
		    	<table class=list width="100%">
            	<% i=0
            	while Not oRS.EOF AND i<10 %>
					   	<tr>
						    <td class=list width="100%"><font size="2">
								<%  Response.Write("<b>" & oRS("Title") & "</b> [" & oRS("InsertedDate") & "]&nbsp;" & _
									oRS("Msg"))
								 %></font>
					        </td>
			      		</tr>
		    		<% i = i + 1
		    		oRS.MoveNext
		    	Wend %>
    			</table>
			<% else %>
                <br><font size="1">&nbsp;</font><br>
		    	<table class=list width="100%">
				   		<tr>
					        <td class=list width="100%" align="middle"><font size="2">
					        	<b>No Messages!</b></font>
					        </td>
			      		</tr>
    			</table>
			<% end if 
			oRS.Close
            %>
                <p>
					<b>Your Buying Status:</b> [Items you have reserved]<font size="2">&nbsp;</font><b><font size="2"><% =iBuyCount & "/" & MAX_BUY %>&nbsp;
			        	<% if iBuyCount = 0 then %><a href="../main.asp">Buy An Item</a>
			        	<% elseif iBuyCount < MAX_BUY then %> </font>
			        		<a href="../main.asp"><font size="2">Buy More Items</font></a>
			        	<% end if %></b><br><font size="1">&nbsp;</font><br>

	   		<% ''//////////////////////////////// Buying Section //////////////////////////////////////

			sSQL = "SELECT * FROM Items WHERE BuyerID=" & iUserID & " AND iStatus<=1"
			''response.write(sSQL)
			oRS.Open sSQL, oConn, 1

			if not oRS.EOF then %>
			   	<table class=list width="100%">

				<% do while not oRS.EOF
	   				itemID=oRS("ItemID")
	   				iCode=oRS("iCode")
	   				get_detail itemID, iCode
	   				%>
						<tr>
								    <td class=list width="75%">
											<% response.write(strType) %>
									</td>
									<td class=list width="25%" align="middle">
											<a class=table href="trade/show_user.asp?bBuy=Yes&iUserID=<% Response.Write(oRS("SellerID")) %>">
                                            <font size="2">Show Seller</font></a><font size="2">&nbsp;
											<a class=table href="trade/cancel.asp?iItemID=<% =(itemID & "&iCode=" & iCode) %>">Cancel</a>
                                            </font>
									</td>
						</tr>
						<tr>
						    <td class=list width="100%" colspan="2">
								<% =strDetail %><br>&nbsp;
								<% =strPrice %>
					        </td>
			      		</tr>
				<% oRS.MoveNext
				if Not oRS.EOF then %>
						<tr>
						    <td width="100%" colspan="2" height="8">
					        </td>
			      		</tr>
				<% end if
				loop %>
		    		</table>
			<% else %>
				    <table class=acc width="100%">
				   		<tr>
					        <td width="100%" align="middle"><font size="2">&nbsp;
					        <b>No items!</b> </font>
					        </td>
			      		</tr>
		    		</table>
			<% end if ''//////////////////////////////// End Buying Section //////////////////////////////////////
			
			oRS.Close %>
				<br><font size="3"><b>
		        		Your Selling Status:</b> [Items you have added]&nbsp;</font><b><% =iSellCount & "/" & MAX_SELL %>
			        	<% if iSellCount = 0 then %>
			        		<a href="add_edit\add_menu.asp">Sell An Item</a>
			        	<% elseif iSellCount < MAX_SELL then %> 
			        		<a href="add_edit\add_menu.asp"><font size="2">Sell More Items</font></a>
			        	<% end if %></b><br><font size="1">&nbsp; </font><br>

	   		<% ''//////////////////////////////// Selling Section //////////////////////////////////////
			
			sSQL = "SELECT * FROM Items WHERE SellerID=" & iUserID & " AND iStatus<=1"
			oRS.Open sSQL, oConn, 1

			if not oRS.EOF then %>
			   	<table class=list width="100%">

				<% do while not oRS.EOF
	   				itemID=oRS("ItemID")
	   				iCode=oRS("iCode")
	   				get_detail itemID, iCode
	   				%>
	   				
								<% if oRS("iStatus")=1 then %>
									<tr>
										<td colspan="2" width="100%">
											<table class=acc width="100%">
								   				<tr>
									        		<td width="50%" colspan="2">
														<b><font size="2">Status: <font color=#FF0000>This item is reserved!</font></font></b><font size="2">
                                                		</font>
													</td>
													<td width="50%" align="right">
														<a class=table href="trade/show_user.asp?iUserID=<% Response.Write(oRS("BuyerID")) %>">
                                                		<font size="2">Show Buyer</font></a><font size="2">&nbsp;&nbsp;
															<a class=table href="trade/confirm.asp?iItemID=<% Response.Write(itemID & "&iCode=" & iCode) %>">Confirm Sell</a>&nbsp;&nbsp;
															<a class=table href="trade/cancel.asp?bSell=Yes&iItemID=<% Response.Write(itemID & "&iCode=" & iCode) %>">Cancel Sell</a>&nbsp;
                                                		</font>
													</td>
												</tr>
											</table>

										</td>
									</tr>
								<% end if %>	   				
							   		<tr>
								        <td class=list width="75%">
											<%=strType%>
										</td>
										<% if oRS("iStatus")=1 then %>
											<td class=list></td>
										<% else %>
											<td class=list align="middle">
												<a class=table href="add_edit/edit.asp?iItemID=<% =(itemID & "&iCode=" & iCode) %>">
                                                	<font size="2">Edit Item</font></a>&nbsp;
												<a class=table href="add_edit/del.asp?iItemID=<% =(itemID & "&iCode=" & iCode) %>">
                                                	<font size="2">Delete</font></a></td>
                                          	</td>
										<% end if %>
						      		</tr>
						      		<tr>
						      			<td class=list colspan="2" width="492">
											<% response.write(strDetail & "<br>&nbsp;" & strPrice ) %>
										</td>
									</tr>
				<% oRS.MoveNext
				if Not oRS.EOF then %>
						<tr>
						    <td width="100%" colspan="2" height="8">
					        </td>
			      		</tr>
				<% end if
				loop %>
		    		</table>
			<% else %>
				    <table class=acc width="100%">
				   		<tr>
					        <td width="100%" align="middle">&nbsp;<font size="2">
					        <b>No items!</b> </font>
					        </td>
			      		</tr>
		    		</table>

		  <% end if
	   		''//////////////////////////////// End Selling Section //////////////////////////////////////

			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			Set oRS2 = Nothing
			'Set oConn = Nothing
			%>

			<p align="center"><font size="2">Copyright © 2002, All rights reserved by
  <a href="../about.htm" style="color: white">K. 
  Haque</a></font></p>
		    </td>
		  </tr>
		</table>
		
    </body>
</html>