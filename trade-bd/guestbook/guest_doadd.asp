<!--#include file="../inc_db.asp"-->
<%
		sSQL = "INSERT INTO GuestBook(Name, Email, Comment, InsertedDate) VALUES('" & _
					Request.form("Name") & "', '" & Request.form("Email") & "', '" & _
					Request.form("Comment") & "', '" & now & "')"
					
					''goto account page
		Set oRS = oConn.Execute(sSQL)
					
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
		
		Dim sMsg, sUrl
		sMsg = "Thank you for your comments!"
		sUrl = "guest_view.asp"
%><!--#include file='../info.asp'-->