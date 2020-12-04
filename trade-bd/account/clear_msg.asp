<%
		if not Session("bLogin") then
			session("login_reason")="You are not logged in!"
			response.redirect "login.asp"
		end if
%>
<!--#include file="../inc_db.asp"-->
<%
		Dim iUserID

		iUserID = Session("iUserID")
		sSql = "DELETE * FROM Notices WHERE UserID=" & iUserID
		Set oRS = oConn.Execute(sSql)
		
		Response.Redirect("account.asp")
		Response.end
%>