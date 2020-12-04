<%
'Dim sMsg, sYesUrl, sNoUrl, sCSSLoc

'sYesUrl = "dolog_out.asp"
'sNoUrl = "account.asp"
'sMsg = "Are you sure you want to log out?"
'sCSSLoc = "../style.css"

'<!--#include file="../ask.asp"-->
%>

<%
	Dim sMsg, sUrl, sCSSLoc
			
	if Session("bLogin") then
		Session("bLogin")=false
		Session("iUserID")=0
		sMsg = "You have been logged out!"
	else
		sMsg = "You are already not logged in!"
	end if
	
	sUrl = "../main.asp"
	sCSSLoc = "../style.css"
%><!--#include file="../info.asp"-->