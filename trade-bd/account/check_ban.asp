<!--#include file="../inc_db.asp"-->
<%
	Dim sR
	Dim sMsg, sUrl, sCSSLoc
	Dim t

	sSQL = "SELECT * FROM BannedIP WHERE IP='" & Request.ServerVariables("remote_addr") & "'"
	oRS.Open sSQL, oConn, 1

	if Not oRS.EOF then
		t = DateDiff("n", oRS("InDate"), now)

		if t<10 then	''not over 10 minutes
			Session("bLogin")=false

			sCSSLoc = "../style.css"
			sMsg = "<font color=#FF0000>Sorry, your IP was Banned for 10 mins for security reasons. To know more please read the FAQs.</font>"
			sUrl = "../main.asp"
			%>
			<!--#include file="../info.asp"-->
			<%
			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing

			response.end
		else
			''10 min is over now remove the data

			sSQL = "DELETE * FROM BannedIP WHERE IP='" & Request.ServerVariables("remote_addr") & "'"
			Set oRS = oConn.Execute(sSQL)
			Session("iTry") = 0
			Session("change_iTry") = 0
		end if
	end if
	oRS.Close
%>