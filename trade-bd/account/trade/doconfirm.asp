<%
if not Session("bLogin") then
	Session("sLogin_Redirect")="account.asp"
	Session("sLogin_Reason")="You are not logged in!"
	Response.Redirect "../login.asp?bRedirect=yes"
end if
%>
<!--#include file="../../inc_db.asp"-->
<%
		Dim itemID, iUserID, i, bError, iBuyCount, iSellCount, str
		Dim iCode, iSellerID, iBuyerID, iDB
		Dim sMsg, sUrl

		itemID = Request.Querystring("iItemID")
		iCode = Request.Querystring("iCode")
		iUserID = Session("iUserID")
		
		sSql = "SELECT * FROM Items WHERE ItemID=" & itemID & " AND iCode=" & iCode
		oRS.Open sSql, oConn, 1

		iBuyerID = -1
		iSellerID=-1

		if not oRS.EOF then 
			iSellerID = oRS("SellerID")
			iBuyerID = oRS("BuyerID")
		end if

		oRS.Close			
		
		bError = true
		
		if iSellerID =-1 then
			sMsg = "Unexpected Error: Item not found!"

		elseif iBuyerID=-1 then
			sMsg = "Unexpected Error: Buyer ID is invalid!"
		
		elseif iSellerID <> iUserID then
			sMsg = "Error: You are not the seller of this item!"

		else
		
			''update buyer's record
			sSql = "SELECT * FROM Users WHERE UserID=" & iBuyerID
			oRS.Open sSql, oConn, 1
	
			iBuyCount = -1
			if not oRS.EOF then iBuyCount = oRS("BuyCount")
			oRS.Close

			'' update seller's record
			sSql = "SELECT * FROM Users WHERE UserID=" & iUserID
			oRS.Open sSql, oConn, 1

			iSellCount = -1
			if not oRS.EOF then iSellCount = oRS("SellCount")
			oRS.Close

			if iBuyCount=-1 then
				sMsg = "Unexpected Error: Buyer was not found!"
			elseif iSellCount=-1 then
				sMsg = "Unexpected Error: Your record was not found!"
			elseif iBuyCount=0 then
				sMsg = "Unexpected Error: BuyCount is invalid!"
			elseif iSellCount=0 then
				sMsg = "Unexpected Error: SellCount is invalid!"
			else
				''update in items
				sSql = "UPDATE Items SET iStatus=2, SoldDate='" & FormatDateTime(Now,2) & "' WHERE ItemID=" & itemID & " AND iCode=" & iCode
				Set oRS = oConn.Execute(sSql)

				''update in correcponding database
				iDB = int(iCode/10)

				if iDB = 1 then
					sSql = "Alls"
				elseif iDB = 2 then
					sSql = "PCs"
				elseif iDB = 3 then
					sSql = "CDsNBooks"
				end if

				sSql = "UPDATE " & sSql & " SET iStatus=2 WHERE ItemID=" & itemID
				Set oRS = oConn.Execute(sSql)

				''update buyer record
				sSql = "UPDATE Users SET BuyCount=" & iBuyCount-1 & " WHERE UserID=" & iBuyerID
				Set oRS = oConn.Execute(sSql)

				''update seller record
				sSql = "UPDATE Users SET SellCount=" & iSellCount-1 & " WHERE UserID=" & iUserID
				Set oRS = oConn.Execute(sSql)

				''update count record				
				sSql = "SELECT * FROM Count"
				oRS.Open sSql, oConn, 1

				sSql = "UPDATE Count SET iCode" & iCode & "='" & oRS("iCode" & iCode).value-1 & "'"
				oRS.Close
				Set oRS = oConn.Execute(sSql)

				sMsg = "The sell was confirmed successfully!"

				if iCode>=10 and iCode<20 then
						sSql = "SELECT * FROM Alls WHERE ItemID=" & itemID & " AND iCode=" & iCode
						oRS.Open sSql, oConn, 1
						if Not oRS.EOF then str = oRS("Name")
				elseif iCode>=20 and iCode<30 then
						sSql = "SELECT * FROM PCs WHERE ItemID=" & itemID & " AND iCode=" & iCode
						oRS.Open sSql, oConn, 1
						if Not oRS.EOF then str = oRS("Processor")
				elseif iCode>=30 and iCode<40 then
						sSql = "SELECT * FROM CDsNBooks WHERE ItemID=" & itemID & " AND iCode=" & iCode
						oRS.Open sSql, oConn, 1
						if Not oRS.EOF then str = oRS("Title")
				end if

				if Not oRS.EOF then oRS.Close

				sSql = "Insert Into Notices(UserID, Title, Msg, InsertedDate) VALUES(" & iBuyerID & ", 'Trade confirmed!', 'The trade for the " & str & " was confirmed by the seller!' ,'" & Now & "')"
				Set oRS = oConn.Execute(sSql)

				bError = false
			end if
		end if
		sURL = "../account.asp"

		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
	%><!--#include file="../../info.asp"-->