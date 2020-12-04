<%
	if not Session("bLogin") then
		Session("sLogin_Redirect")="account.asp"
		Session("sLogin_Reason")="You are not logged in!"
		Response.Redirect "../login.asp?bRedirect=yes"
	end if
%>
<!--#include file="../../inc_db.asp"-->
<%
		Dim itemID, iUserID, i, itemCount, bError
		Dim iCode, iSellerID, iBuyerID, bSell, str
		Dim sMsg, sUrl
	
		itemID = Request.Querystring("iItemID")
		iCode = Request.Querystring("iCode")
		
		bSell = false
		if Request.Querystring("bSell")="Yes" then bSell = true
		
		
		if bSell then	''if this is a cancel request by seller
			sSql = "SELECT * FROM Items WHERE ItemID=" & itemID & " AND iCode=" & iCode
			oRS.Open sSql, oConn, 1
	
			iUserID = oRS("BuyerID")
			oRS.Close			
		else			''if this is a cancel request by buyer
			iUserID = Session("iUserID")
		end if
		
		bError = true
		
		''update user's sells record
		sSql = "SELECT * FROM Users WHERE UserID=" & iUserID
		oRS.Open sSql, oConn, 1

		itemCount = -1
		if not oRS.EOF then itemCount = oRS("BuyCount")
		oRS.Close

		if itemCount=-1 then
			if bSell then
				sMsg = "Unexpected Error: Buyer not found!"
			else
				sMsg = "Unexpected Error: User not found!"
			end if
		else
			sSql = "SELECT * FROM Items WHERE ItemID=" & itemID & " AND iCode=" & iCode
			oRS.Open sSql, oConn, 1
	
			iSellerID=-1

			if not oRS.EOF then 
				iSellerID = oRS("SellerID")
				iBuyerID = oRS("BuyerID")
			end if
			
			oRS.Close
			
			if (iSellerID=-1) then
				sMsg = "Unexpected Error: Item not found!"
				
			elseif bSell AND iSellerID<>Session("iUserID") then
				sMsg = "Error: you are not the seller of this item!"
			
			elseif NOT bsell AND iBuyerID <> iUserID then
				sMsg = "Error: you did not reserve the item!"
			
			else
				''update in items
				sSql = "UPDATE Items SET BuyerID=0,iStatus=0 WHERE ItemID=" & itemID & " AND iCode=" & iCode
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

				sSql = "UPDATE " & sSql & " SET iStatus=0 WHERE ItemID=" & itemID
				Set oRS = oConn.Execute(sSql)
	
				''update user record
				sSql = "UPDATE Users SET BuyCount=" & itemCount-1 & " WHERE UserID=" & iUserID
				Set oRS = oConn.Execute(sSql)

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

				sSql = "Insert Into Notices(UserID, Title, Msg, InsertedDate) VALUES("

				if bSell then	''request by seller
					sMsg = "The sell was cancelled successfully!"
					sSql = sSql & iBuyerID
				else	''request by buyer
					sMsg = "Item reservation was cancelled successfully!"
					sSql = sSql & iSellerID
				end if

				sSql = sSql & ", 'Trade Cancelled!', 'The trade of your " & str & " was cancelled by the user!' ,'" & Now & "')"
				Set oRS = oConn.Execute(sSql)

				bError = false
			end if
		end if
		sURL = "../account.asp"

		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
%><!--#include file="../../info.asp"-->