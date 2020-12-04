
<% if Session("bLogin") then
	Dim sMsg, sUrl
	sMsg = "You cannot sign up when you are logged in!"
	sUrl = "javascript:history.back(1);"
%>
	<!--#include file="../../info.asp"-->
<% 
response.end
end if %>


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

<p align="justify">
<font size="3"><img border="0" src="../../images/arrow.gif">Welcome to the Member Sign Up Page!</font><br />
Before you sign up please make sure you have read the 
following pages thoroughly:</p>
<blockquote>
  <ol type="a">
    <li>
  <p align="justify">Beginner's Page. (<a href="../../beginners.htm">Read Now</a>)</li>
    <li>
  <p align="justify">Rules &amp; Regulations. (<a href="../../rules.htm">Read 
  Now</a>)</li>
  </ol>
</blockquote>

<p align="justify">
It will take less than a minute to sign up in this site. There are 2 steps for 
signing up. The overview of the steps are given below. </p>
<blockquote>
<ol type="i">
  <li>
  <p align="justify">Step one, read and agree to terms and conditions.</li>
  <li>
  <p align="justify">Step two, fill up the form.</li>
</ol>
</blockquote>
<p align="justify">Click on the 'Proceed with Sign Up' button 
to sign up now, or 'Cancel' button to cancel the sign up.</p>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="50%" id="AutoNumber1">
    <tr>
      <td width="100%">
      <form method="POST" action="add_user_2.asp"> <p align="center">
        <input type="submit" value="Proceed with Sign Up" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="Cancel" name="B3" onclick="window.location='add_user_cancel.asp'"></p>
      </form>
      </td>
    </tr>
  </table>
  </center>
</div>
<p align="center"><font size="2">Copyright © 2002, All rights reserved by
<a href="../../about.htm" style="color: white">K. Haque</a></font></p>
        </td>
      </tr>
	</table>
      </body>
</html>