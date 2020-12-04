<%
if not Session("bLogin") then
	Session("sLogin_Redirect")="add_edit/add_menu.asp"		'check for error on edit items
	Session("sLogin_Reason")="You are not logged in! You need to login first<br>in order to add an item to our database!"
	Response.Redirect "../login.asp?bRedirect=yes"
end if
%>