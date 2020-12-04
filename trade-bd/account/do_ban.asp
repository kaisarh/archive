<!--#include file="../inc_db.asp"-->
<%
	Dim sR, IP
	IP = Request.ServerVariables("remote_addr")

	sSQL = "INSERT INTO BannedIP(IP, InDate) VALUES('" & IP & "','" & now & "')"

	Set oRS = oConn.Execute(sSQL)
	Session("bBanned") = true

	Dim sMsg, sUrl, sCSSLoc
	sCSSLoc = "../style.css"

	sMsg = "<font color=#FF0000>Sorry, your IP (" & IP & ") has been Banned for 10 mins for security reasons. To know more please read the FAQs.</font>"
	sUrl = "../main.asp"
%><!--#include file="../info.asp"-->