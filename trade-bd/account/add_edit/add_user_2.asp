
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
<font size="3"><img border="0" src="../../images/step.gif">Step One:<br>
</font>&nbsp;&nbsp;&nbsp; Please read the <b>Terms And Conditions</b> listed below very carefully. 
It is important for you to understand what you are agreeing with.</p>
<ol>
  <li>
  <p align="justify">You <b>must not</b> sell/buy (do trade of) any illegal or 
  harmful material/item here by any means.</li>
  <li>
  <p align="justify">You <b>must not</b> use bad/slang words in any place of 
  this site by any means. If you do, actions will be taken against you.</li>
  <li>
  <p align="justify">You <b>must not</b> harass other people/visitors of this 
  site by any means.&nbsp; </li>
  <li>
  <p align="justify"><b>Trade-bd.Com </b>is a free web based service, It <u><b>does not</b></u> and <b><u>
  will not</u> </b>take any responsibility for any possible harassment or fraud 
  activity done to you through this site.<b> </b></li>
  <li>
  <p align="justify">As <b>Trade-bd.Com </b>is a free web based service, It <u><b>does not</b></u> and <b><u>
  will not</u> </b>take any responsibility for any possible <u><b>damage</b></u> 
  to your property or your money you share through this site for any purpose.</li>
  <li>
  <p align="justify">As <b>Trade-bd.Com </b>is a free web based service, you 
  cannot have any objection about the quality, quantity of the services you get 
  from this site. You <b> <u>can suggest</u></b> for better options but you <u><b>
  cannot</b></u> claim those by any means.</li>
  <li>
  <p align="justify">Rules are subject to change with out any notice and the web 
  master of this site has the privilege to make any kind of modification to this site 
    or site content.</li>
</ol>
<p align="justify">&nbsp;&nbsp;&nbsp; You must agree to these terms and 
conditions for signing up as a member of this site . If you agree to then click the 'I Agree' button.</p>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="39%" id="AutoNumber1">
    <tr>
      <td width="100%">
      <form method="POST" action="add_user_3.asp"> <p align="center">
        <input type="submit" value="I agree" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="I do not agree" name="B3" onclick="window.location='add_user_cancel.asp'"></form>
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