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
	else if (frm.Login.value.length < 1)
		window.alert("Please enter the Login correctly.");
	else if (frm.Add1.value.length < 8 || frm.Add2.value.length < 4)
		window.alert("Please enter your complete address.");
	else if (frm.Phone.value.length < 6)
		window.alert("Please enter your phone number correctly.");
	else if (frm.Pass1.value.length < 4)
		window.alert("Password must be at least 4 chars.");
	else if (frm.Pass1.value != frm.Pass2.value)
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
    <td width="100%">
    <p align="justify"><font size="3">
    <img border="0" src="../../images/step.gif">Step Two:<br>
    </font>&nbsp;&nbsp;&nbsp; Please fill the form below correctly. 
    It is important to give us your original contacting information. Whether you 
    buy or sell an item in this site, other users (i.e. your sellers or your buyers) 
    need to know your original address or contacting information in order to contact 
    with you.</p>
    <form method="POST" action="add.asp?iCode=40">
      <p align="center"><strong>User Sign Up Form:</strong></p>
      <div align="center">
        <center>
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="39%" id="AutoNumber1">
          <tr>
            <td width="100%">
            <table border="0" width="388" cellspacing="0" align="center">
              <tr>
                <td width="796" rowspan="11">&nbsp;</td>
                <td colspan="2" width="398">
                <p align="center"><br>
                </p>
                </td>
                <td width="796" rowspan="11">&nbsp;</td>
              </tr>
              <tr>
                <td width="103" align="left">Name:</td>
                <td width="293">
                <input type="text" name="Name" size="20" tabindex="1"><font size="2"> 
                [within 30 chars]</font></td>
              </tr>
              <tr>
                <td valign="top" width="103">Address:</td>
                <td width="293">
                <input type="text" name="Add1" size="40" tabindex="2"><br>
                <input type="text" name="Add2" size="40" tabindex="3"><br>
                <font size="2">&nbsp;[both within 50 chars, total within 100 chars]</font></td>
              </tr>
              <tr>
                <td width="103">Phone:</td>
                <td width="293">
                <input type="text" name="Phone" size="20" tabindex="4"><font size="2"> 
                [within 20 chars]</font></td>
              </tr>
              <tr>
                <td width="103">Email Address:</td>
                <td width="293">
                <input type="text" name="Email" size="20" tabindex="5"><font size="2"> 
                [within 30 chars]</font></td>
              </tr>
              <tr>
                <td width="398" colspan="2" height="10"></td>
              </tr>
              <tr>
                <td width="103">Login:</td>
                <td width="293"><input name="Login" size="10" tabindex="6"><font size="2"> 
                [within 10 chars]</font></td>
              </tr>
              <tr>
                <td width="103">Password:</td>
                <td width="293">
                <input type="password" name="Pass1" size="10" tabindex="7"><font size="2"> 
                [within 4 to 10 chars]</font></td>
              </tr>
              <tr>
                <td width="103">Confirm Password:</td>
                <td width="293">
                <input type="password" name="Pass2" size="10" tabindex="8"></td>
              </tr>
              <tr>
                <td width="398" colspan="2" height="10"></td>
              </tr>
              <tr>
                <td width="103">
                <input type="button" value="&lt;&lt; Back" name="b3" onClick="javascript:history.back(1)"></td>
                <td align="center" width="293">
                <input type="reset" value="Clear All" name="b2" tabindex="10">&nbsp;
                <input type="button" value="Submit &gt;&gt;" name="b1" onClick="javascript:check(this.form)"></td>
              </tr>
              <tr>
                <td width="796">&nbsp;</td>
                <td width="103">&nbsp;</td>
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
    <p align="justify"><img border="0" src="../../images/bulb.gif"><font size="3">Tips 
    for signing up:</font><br />
&nbsp;&nbsp;&nbsp; First of all fill in the name, address 
    &amp; phone number section. Then enter a valid e-mail address. This e-mail address 
    will be used to confirm the account sign up and to contact with you. Login is 
    the id you will be using to login to your account. You should remember your 
    password. If you forget your password, your password will be sent to the e-mail 
    address you are providing us with.</p>
    <p align="center"><font size="2">Copyright © 2002, All rights reserved by
    <a href="../../about.htm" style="color: white">K. Haque</a></font></p>
    </td>
    </tr>
  </table>

</body>

</html>