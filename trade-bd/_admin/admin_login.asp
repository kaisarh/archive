<html>
	<head>
		<meta http-equiv="Content-Language" content="en-us" />
		<link rel="stylesheet" href="../style.css" />
		<title>Admin Login</title>
	</head>
	<body>
		<center>
			<table border=0 width="640" cellspacing=0 cellpadding=0>
				<tr>
			    	<td width="100%" align="middle">
					<%
						Function getv(x)
							getv = request.form(x)
						End Function
						
						Dim sR
						Dim sUser, sPass, iTry, bBlank
						
						const MAX_TRY = 6
						const BAN_MSG = "<Br><strong><font color=#FF0000>Your IP has been Banned!</font></strong><Br>"
				
						if Session("bLogin") then
							Session("bLogin")=false
							Session("iUserID")=0
							Response.Write("You have been logged out as from General User.<br>")
						end if

						if Session("bAdmin") then
							Response.Write("<strong>You are already logged in!</strong>")
						else
							if Session("bBanned") then
								Response.write(BAN_MSG)
							else
								if Session("sLogin_Reason") <> "" then
									Response.Write(Session("sLogin_Reason"))
									session("sLogin_Reason") = ""
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
									Session("bAdmin")=false
									
									if sUser <> "admin" then 
										sR = "<font color=#FF0000>Error: Ivalid Login!</font>"
									elseif sPass <> "secret" then 
										sR = "<font color=#FF0000>Error: Ivalid Password!</font>"
									else
										sR = "Login Successful!"
										Session("bAdmin")=true
										Session("iTry") = 0

										Response.Redirect "admin.asp"
									end if
			
									Response.Write(sR)
								
									iTry = iTry + 1
									Session("iTry") = iTry
									
									if iTry >= MAX_TRY and not Session("bAdmin") then
										Response.Write(BAN_MSG)
										Session("banned")=true
									end if
								end if
								if iTry < MAX_TRY and not Session("bAdmin") then  %>
									<p><u><b>Administrator log in:</b></u></p>
									<table border="1" width="314" cellspacing="0" cellpadding="0" bordercolorlight="#000000"
									bordercolordark="#FFFFFF">
									  <tr>
									    <td width="311">&nbsp;
									    <form method="POST" name="form1" action="admin_login.asp">

										</form>

										<p align="left">
										<table border="0" width="100%" cellspacing="0" cellpadding="0">
									        <tr>
									          <td width="51"></td>
									          <td width="84">Login:</td>
									          <td width="174"><input type="text" name="Login" size="17"></td>
									        </tr>
									        <tr>
									          <td width="51"></td>
									          <td width="84">Password:</td>
									          <td width="174"><input type="password" name="Pass" size="17"></td>
									        </tr>
									        <tr>
									          <td height="8" width="51"></td>
									          <td height="8" width="84"></td>
									          <td height="8" width="174"></td>
									        </tr>
									        <tr>
									          <td width="51"></td>
									          <td width="84"><div align="right"><p><input type="submit" value="Login" name="B1">&nbsp; </td>
									          <td width="174">&nbsp;<input type="reset" value="Reset" name="B2"></td>
									        </tr>
									      </table>
									    </form>
									    </td>
									  </tr>
									</table>
								<% end if
							end if
						end if %>
					</td>
				</tr>
			</table>
			<p>Copy Right 2002, All rights reserved by MyMarket.Com.
		</center>
	</body>
</html>
