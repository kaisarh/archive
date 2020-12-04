<% if NOT Session("bLogin") then
	Dim sMsg, sUrl
	sMsg = "You need to login in order to reserve this item!"
	sUrl = "../login.asp?bRedirect=yes"
	session("sLogin_Redirect")="trade/doreserve.asp?" & request.querystring
%>

	<!--#include file="../../info.asp"-->
 
<% response.end
end if
%>
<!--#include file="../../inc_db.asp"-->
<%
		Response.Expires = -1000
		Dim itemID, iUserID, i, itemCount, bError, str
		Dim iCode, iSellerID, iBuyerID

		itemID = Request.Querystring("iItemID")
		iCode = Request.Querystring("iCode")
		
		iUserID = Session("iUserID")
		bError = true
		
		''get buycount
		sSql = "SELECT * FROM Users WHERE UserID=" & iUserID
		oRS.Open sSql, oConn, 1

		itemCount = -1
		if not oRS.EOF then itemCount = oRS("BuyCount")
		oRS.Close
		''
		
		sUrl = "../../main.asp"

		if itemCount=-1 then
			sMsg = "Unexpected Error: User not found!"
		
		elseif itemCount>=2 then
			sMsg = "Sorry, you already have reserved the max number of items! " & _
				"You cannot reserve any more item at the moment, please read the FAQs."
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
			
			elseif iSellerID = iUserID then
				sMsg = "Error: You cannot reserve your own item!"
			
			elseif iBuyerID = iUserID then
				sMsg = "The item is already reserved by you!"
			
			elseif iBuyerID <> 0 then	''iStatus = 1
				sMsg = "Error: The item is already reserved by someone else!"

			else
				''update in items
				sSql = "UPDATE Items SET BuyerID=" & iUserID & ", iStatus=1 WHERE ItemID=" & itemID & " AND iCode=" & iCode
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

				sSql = "UPDATE " & sSql & " SET iStatus=1 WHERE ItemID=" & itemID
				Set oRS = oConn.Execute(sSql)

				''update buyer record
				sSql = "UPDATE Users SET BuyCount=" & itemCount+1 & " WHERE UserID=" & iUserID
				Set oRS = oConn.Execute(sSql)

				if iCode>=10 and iCode<20 then
						sSql = "SELECT * FROM Alls WHERE ItemID=" & itemID & " AND iCode=" & iCode
						oRS.Open sSql, oConn, 1
						if Not oRS.EOF then str = oRS("Name")
				elseif iCode>=20 and iCode<30 then
						sSql = "SELECT * FROM PCs WHERE ItemID=" & itemID & " AND iCode=" & iCode
						oRS.Open sSql, oConn, 1
						if Not oRS.EOF then str = oRS("Processor") & " PC "
				elseif iCode>=30 and iCode<40 then
						sSql = "SELECT * FROM CDsNBooks WHERE ItemID=" & itemID & " AND iCode=" & iCode
						oRS.Open sSql, oConn, 1
						if Not oRS.EOF then str = oRS("Title")
				end if

				if Not oRS.EOF then oRS.Close

				sSql = "Insert Into Notices(UserID, Title, Msg, InsertedDate) VALUES(" & iSellerID & ", 'Item reserved!', 'Your " & str & " was reserved!' ,'" & Now & "')"
				Set oRS = oConn.Execute(sSql)

				sMsg = "The item was reserved successfully!"
				sUrl = "show_user.asp?bBuy=Yes&iUserID=" & iSellerID
				bError = false

			end if
		end if
		
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
	%><!--#include file="../../info.asp"-->