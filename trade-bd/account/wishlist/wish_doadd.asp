<!--#include file="../../inc_db.asp"-->
<%
		sSQL = "INSERT INTO WishList(UserID, Wish, Price, InsertedDate) VALUES('" & _
					Session("iUserID") & "', '" & Request.form("Wish") & "', '" & _
					Request.form("Price") & " " & Request.form("PriceUnit") & "', '" & _
					now & "')"
					
					''goto account page
		Set oRS = oConn.Execute(sSQL)
					
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
		
		sMsg = "Data has been added to the Wish List!"
		sUrl = "wish_view.asp"
%><!--#include file='..\..\info.asp'-->