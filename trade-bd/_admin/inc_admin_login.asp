<%
if not Session("bAdmin") then
	Session("sLogin_Reason")="You are not logged in as Administrator!"
	Response.Redirect "admin_login.asp"
end if
%>
