
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
<font size="3"><img border="0" src="../../images/arrow.gif">Sign Up cancelled!</font><br />
&nbsp;&nbsp;&nbsp; You have cancelled the Sign Up process. You can always sign 
up again simply by clicking on the sign up button the menu (on the left side).</p>

<p align="center">
<font size="2">Copyright © 2002, All rights reserved by
<a href="../../about.htm" style="color: white">K. Haque</a></font><br>
&nbsp;</p>
        </td>
      </tr>
	</table>
      <p>&nbsp;</p>
</body>
</html>