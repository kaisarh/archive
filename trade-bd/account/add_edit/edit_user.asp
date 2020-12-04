<!-- #include file='inc_add_login.asp' -->
<!--#include file="../../inc_db.asp"-->
<%
		iUserID = Session("iUserID")

		sSQL = "SELECT * FROM Users WHERE UserID=" & iUserID
		oRS.Open sSQL, oConn, 1

		if oRS.EOF then
			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing

			dim sMsg, sUrl
			sMsg = "Error: Data not found!"
			sUrl = "account.asp"
			%>
				<!--#include file="../../info.asp"-->
			<%
			response.end
		end if
%>

<html>

<head>
	<title>Sign Up</title>
	<meta http-equiv="Content-Language" content="en-us" />
	<link rel="stylesheet" href="../../style.css" />


<script language="JavaScript">
<!--
function isEmail(elem) {
	return !((elem.value.indexOf('@',0) < 1) || (elem.value.indexOf('.',0)<2) || (elem.value.length<6));
}

// bug: selection checking not here
function check(frm) {
	
	if (frm.Name.value.length < 2)
		window.alert("Please enter the Name correctly.");
	else if (!isEmail(frm.Email))
		window.alert("Please enter the E-mail address correctly.")
	else if (frm.Add1.value.length < 8 || frm.Add2.value.length < 4)
		window.alert("Please enter your complete address.");
	else if (frm.Phone.value.length < 6)
		window.alert("Please enter your phone number correctly.");
	else
		frm.submit();
}

function mark(){
		form1.Name.value="<%=oRS("Name")%>";
		form1.Add1.value="<%=oRS("Add1")%>";
		form1.Add2.value="<%=oRS("Add2")%>";
		form1.Phone.value="<%=oRS("Phone")%>";
		form1.Email.value="<%=oRS("Email")%>";
}

//-->
    </script>

</head>

<body onload="javascript:mark();">
  	<table width="640">
      <tr>
        <td width="100%">

<p align="justify">
<img border="0" src="../../images/arrow.gif"><font size="3">Edit your profile:</font><br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
First of all fill in the name, address &amp; phone number section. Then enter a 
valid e-mail address. This e-mail address will be used to confirm the account 
sign up and to contact with you. Login is the id you will be using to login to 
your account. You should remember your password. If you forget your password, 
your password will be sent to the e-mail address you are providing us with.</p>

<form method="POST" action="update.asp?iCode=40" name="form1">

<p align="center">
<strong>User Profile:</strong></p>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="39%" id="AutoNumber1">
    <tr>
      <td width="100%">
<table border="0" width="388" cellspacing="0" align="center">
    <tr>
    	<td width="796" rowspan="7">&nbsp;</td>
    	<td colspan="2" width="398">
        <p align="center"><br></td>
    	<td width="796" rowspan="7">&nbsp;</td>
    </tr>
    <tr>
      <td width="103" align="left">Name:</td>
      <td width="293"><input type="text" name="Name" size="20" tabindex="1"></td>
    </tr>
    <tr>
      <td valign="top" width="103">Address:</td>
      <td width="293"><input type="text" name="Add1" size="40" tabindex="2"><br>
      <input type="text" name="Add2" size="40" tabindex="3"></td>
    </tr>
    <tr>
      <td width="103">Phone:</td>
      <td width="293"><input type="text" name="Phone" size="20" tabindex="4"></td>
    </tr>
    <tr>
      <td width="103">Email Address:</td>
      <td width="293"><input type="text" name="Email" size="20" tabindex="5"></td>
    </tr>
    <tr>
      <td width="398" colspan="2" height="10"></td>
    </tr>
    <tr>
      <td width="103">
      <input type="button" value="&lt;&lt; Back" name="b3" onClick="javascript:history.back(1)"></td>
      <td align="center" width="293"><input type="reset" value="Clear All" name="b2" tabindex="10">&nbsp; 
      	<input type="button" value="Update &gt;&gt;" name="b1" onClick="javascript:check(this.form)"></td>
    </tr>
    <tr>
    	<td width="796">&nbsp;</td>
      <td width="103">
      &nbsp;</td>
      <td align="center" width="293">&nbsp;</td>
    	<td width="796">&nbsp;</td>
    </tr>
  </table>
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
  <p align="center"><font size="2">Copyright © 2002, All rights reserved by
  <a href="../../about.htm" style="color: white">K. Haque</a></font><br>
&nbsp;</p>
        </td>
      </tr>
	</table>
  	</body>
</html>

<%
			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing
%>