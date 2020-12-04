<%
if Request.querystring("iCode")<>40 and not Session("bLogin") then
	session("login_reason")="You are not logged in!"
	response.redirect "login.asp"
end if
%>
<!--#include file="../../inc_db.asp"-->
<%
		
		Dim sDB
		Dim itemID, iUserID, sType, bFound, itemCount, iBuyerID
		Dim sMsg, sUrl
	
		itemID = int(Request.querystring("iItemID"))
		iCode = Request.querystring("iCode")
	
		iUserID = Session("iUserID")
	
		Select Case int(iCode/10)
			Case 1
				sDB = "Alls"
			Case 2
				sDB = "PCs"
			Case 3
				sDB = "CDsNBooks"
		End select
	
		sSql = "SELECT * FROM " & sDB & " WHERE ItemID=" & itemID
		oRS.Open sSql, oConn, 1
	
		bFound = false
		if NOT oRS.EOF then	bFound = true
		oRS.Close

		if bFound then	

			sSQL = "SELECT * FROM Items WHERE ItemID=" & itemID & " AND iCode=" & iCode
			oRS.Open sSQL, oConn, 1

			iBuyerID=0

			if not oRS.EOF then iBuyerID = oRS("BuyerID")
			oRS.Close
			
			if iBuyerID<>0 then
				sMsg = "The item is reserved by a member, you cannot deleted it now!"
			
			else
				sSql = "DELETE * FROM " & sDB & " WHERE ItemID=" & itemID
				Set oRS = oConn.Execute(sSql)
		
				sSql = "DELETE * FROM Items WHERE ItemID=" & itemID & " AND iCode=" & iCode
				Set oRS = oConn.Execute(sSql)
		
				''update user's sells record
				sSql = "SELECT * FROM Users WHERE UserID=" & iUserID
				oRS.Open sSql, oConn, 1
		
				itemCount = oRS("SellCount")
				oRS.Close
								
				sSql = "UPDATE Users SET SellCount=" & itemCount-1 & " WHERE UserID=" & iUserID
				Set oRS = oConn.Execute(sSql)
				
				''update count record				
				sSql = "SELECT * FROM Count"
				oRS.Open sSql, oConn, 1
				
				sSql = "UPDATE Count SET iCode" & iCode & "='" & oRS("iCode" & iCode).value-1 & "'"
				oRS.Close
				Set oRS = oConn.Execute(sSql)
				
				sMsg = "The item was deleted successfully!"
			end if
		else
			sMsg = "Error: Item was not found!"
		end if
	
		sUrl = "../account.asp"
		%>
		<!--#include file="../../info.asp"-->
		<%

		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
%>