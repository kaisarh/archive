<!--#include file="check_ban.asp"-->
<%
		if NOT Session("bLogin") then
			session("sLogin_Redirect")="change_pass.asp"
			session("sLogin_Reason")="You are not logged in! You need to login first in order to change password!"
			response.redirect "login.asp?redirect=yes"
		end if

		Function getv(x)
			getv = request.form(x)
		End Function

		Dim sID
		Dim sUser, sPass, sNewPass, iTry, bDone, bBlank
		const MAX_TRY = 3

		sUser = request.form("Login")
		sPass = request.form("Pass")
		sNewPass = request.form("NewPass1")

		if sUser = "" and sPass = "" then 
			bBlank = true
		else
			bBlank = false
		end if
		bDone = 0

		iTry = Session("change_iTry")
		if iTry = "" then iTry = 0

		if iTry < MAX_TRY and not bBlank then

			sSQL = "SELECT * FROM Users WHERE Login='" & sUser & "'"
			oRS.Open sSQL, oConn, 1

			if oRS.EOF then 
				sR = "<font color=#FF0000>Error: Ivalid Login!</font><br>"
			elseif oRS("Pass") <> sPass then 
				sR = "<font color=#FF0000>Error: Ivalid Password!</font><br>"
			else
				oRS.Close
				sSQL = "UPDATE Users SET Pass='" & sNewPass & "' WHERE Login='" & sUser & "'"
				oRS.Open sSQL, oConn, 1

				bDone = 1
			end if

			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing
		end if

		if bDone = 1 then
			Session("change_iTry") = 0
			sMsg = "Your password has been changed successfully!"
			sUrl = "account.asp"
			sCSSLoc = "../style.css"
			%><!--#include file="../info.asp"--><%
			response.end

		elseif (not bBlank) then 
			iTry = iTry + 1
			Session("change_iTry") = iTry

			if iTry => MAX_TRY then
				response.redirect("do_ban.asp")
				response.end
			end if
		end if
%>
<html>

<head>
<title>Change Password</title>
<meta http-equiv="Content-Language" content="en-us" />
<link rel="stylesheet" href="../style.css" />
<script language="JavaScript">

function check(frm) {
	if (frm.Login.value.length < 1)
		window.alert("Please enter the Login correctly.");
	else if ( frm.Pass.value.length < 4)
		window.alert("Please enter your password correctly.");
	else if (frm.NewPass1.value.length < 4)
		window.alert("New password must be at least 4 chars.");
	else if (frm.NewPass1.value != frm.NewPass2.value)
		window.alert("The two passwords are not identical.");
	else
		frm.submit();
}

//-->
        </script>
</head>

<body>

<table width="640">
  <tr>
    <td width="100%"><center>
    <table width="640">
      <tr>
        <td width="100%" align="middle">
        <p><strong style="font-weight: 400">Change Password:</strong></p>
        		<% if sR <> "" then 
					Response.Write(sR)
				end if %><br>
        <table border="1" width="400" cellspacing="0" cellpadding="0" bordercolorlight="#B7B7B7" bordercolordark="#FFFFFF">
          <tr>
            <td width="100%"><br>
            <form method="POST" name="form1" action="change_pass.asp">
              <table border="0" width="100%" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="10%" rowspan="10"></td>
                  <td width="45%">Login:</td>
                  <td>
                  <input type="text" name="Login" size="17" </td tabindex="1"> </td>
                </tr>
                <tr>
                  <td>Password:</td>
                  <td><input type="password" name="Pass" size="17" tabindex="2"></td>
                </tr>
                <tr>
                  <td height="10" colspan="2"></td>
                </tr>
                <tr>
                  <td>New Password:</td>
                  <td>
                  <input type="password" name="NewPass1" size="17" tabindex="3"></td>
                </tr>
                <tr>
                  <td>Re-enter New Password:</td>
                  <td>
                  <input type="password" name="NewPass2" size="17" tabindex="4"></td>
                </tr>
                <tr>
                  <td height="10" colspan="2"></td>
                </tr>
                <tr>
                  <td>
                  <input type="button" value="&lt;&lt; Back" name="b3" onClick="javascript:location.href='account.asp'" tabindex="7">&nbsp;
                  <input type="reset" value="Reset" name="B2" tabindex="6"> </td>
                  <td>
                  <input type="button" value="Sumbit  &gt;&gt;" name="B1" onClick="javascript:check(this.form)" tabindex="5">&nbsp;
                  </td>
                </tr>
              </table>
            </form>
            </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    <p><font size="2">Copyright © 2002, All rights reserved by
    <a href="../about.htm" style="color: white">K. Haque</a></font></p>
    </center></td>
  </tr>
</table>

</body>

</html>