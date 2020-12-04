<%
	if not session("bLogin") then
		session("sLogin_Redirect")="account.asp"
		session("sLogin_Reason")="Please enter you login and password!"
		response.redirect "login.asp?bRedirect=yes"
	end if
%>
