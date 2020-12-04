<!--#include file="check_ban.asp"--><%
	Function getv(x)
		getv = request.form(x)
	End Function

	Dim sUser, sPass, iTry, bBlank, do_redirect
	sCSSLoc = "../style.css"

	if Request.querystring("bRedirect")="yes" and Session("sLogin_Redirect") <> "" then
		do_redirect = true
	else
		do_redirect = false
	end if

	const MAX_TRY = 3

	if Session("bLogin") then
		sMsg = "You are already logged in!"
		sUrl = "account.asp"
		%><!--#include file="../info.asp"--><%
		response.end
	end if

	sUser = request.form("Login")
	sPass = request.form("Pass")

	iTry = Session("iTry")
	if iTry = "" then 
		iTry = 0
		Session("iTry") = 0
	end if

	if sUser = "" and sPass = "" then
		bBlank = true
	end if

	if (not bBlank) then
		sSQL = "SELECT * FROM Users WHERE Login='" & sUser & "'"
		oRS.Open sSQL, oConn, 1

		if oRS.EOF then 
			sR = "<font color=#FF0000>Error: Ivalid Login!</font>"
		elseif oRS("Pass") <> sPass then 
			sR = "<font color=#FF0000>Error: Ivalid Password!</font>"
		else
			Session("bLogin")=true
			Session("iUserID")=oRS("UserID")
			Session("iTry") = 0

			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing

			if do_redirect then
				sUrl = Session("sLogin_Redirect")
				Session("sLogin_Redirect") = ""
				Response.redirect sUrl
			else
				Response.redirect "account.asp"
			end if
			response.end
		end if

		iTry = iTry + 1
		Session("iTry") = iTry

		oRS.Close
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing

		if iTry >= MAX_TRY then
			response.redirect("do_ban.asp")
			response.end
		end if
	end if
%>
<html>

<head>
<title>Login</title>
<meta http-equiv="Content-Language" content="en-us" />
<link rel="stylesheet" href="../style.css" />
<script language="JavaScript">

function check(frm) {
	if (frm.Login.value.length < 1)
		window.alert("Please enter the Login correctly.");
	else if ( frm.Pass.value.length < 4)
		window.alert("Please enter your password correctly.");
	else
		frm.submit();
}

//-->
</script>
</head>

<body>

<table width="640">
  <tr>
    <td width="100%" align="middle"><% if sR <> "" then
			Response.Write(sR & "<br><br>")
		end if
		if session("sLogin_Reason") <> "" then
			sR = Session("sLogin_Reason") & "<p>"
			session("sLogin_Reason") = ""
			Response.Write(sR)
		end if %> <b style="font-weight: 400">Member Login:</b><p></p>
    <table width="300" border="1" bordercolor="#FFFFFF" style="border-collapse: collapse">
      <tr>
        <td width="100%"><font size="1">&nbsp; </font>
        <form method="POST" name="form1" action="login.asp<%if do_redirect then Response.write("?bRedirect=yes") end if%>">
          <table width="100%">
            <tr>
              <td width="12%" rowspan="4"></td>
              <td>Login:</td>
              <td><input type="text" name="Login" size="20" tabindex="1" /></td>
            </tr>
            <tr>
              <td>Password:</td>
              <td><input type="password" name="Pass" size="20" tabindex="2" /></td>
            </tr>
            <tr>
              <td height="10" colspan="2"></td>
            </tr>
            <tr>
              <td colspan="2">
              <input type="button" value="&lt;&lt; Back" name="b3" onClick="javascript:history.back(1)" tabindex="4" />&nbsp;
              <input type="reset" value="Reset" name="B2" tabindex="5" />&nbsp;
              <input type="button" value="Login &gt;&gt;" name="B1" onClick="javascript:check(this.form)" tabindex="3" />&nbsp;
              </td>
            </tr>
          </table>
        </form>
        <p></p>
        </td>
      </tr>
    </table>
    <p align="center"><font size="2">Copyright © 2002, All rights reserved by
    <a href="../about.htm" style="color: white">K. Haque</a></font></p>
    </td>
  </tr>
</table>
</body>
</html>