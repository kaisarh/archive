<html>

<head>
<title>Sign Up</title>

<meta http-equiv="Content-Language" content="en-us" />
<link rel="stylesheet" href="../style.css" />

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
	
	else if (frm.Comment.value.length < 3)
		window.alert("Please enter the Comment correctly.");

	else
		frm.submit();
}

//-->
</script>
</head>

<body>
  	<table width="640" height="401">
      <tr>
        <td width="100%" height="397">
<p>
<font size="3"><img border="0" src="../images/arrow.gif">Add comments!</font><br />
&nbsp;&nbsp;&nbsp;
Please type comments below with your name and valid email address. Any 
suggestion about the site or site content, any comments or any request will 
carefully be reviewed.</p>
<div align="center">

<p align="center"><strong>Add comments to Guest Book:</strong></p>

<form method="POST" action="guest_doadd.asp">
  <table border="0" width="640" cellspacing="0">
    <tr>
      <td width="128" align="left"></td>
      <td width="105" align="left">Name:</td>
      <td width="412"><input type="text" name="Name" size="36" tabindex="1"></td>
    </tr>
    <tr>
      <td width="128" align="left"></td>
      <td width="105" align="left">Email Address:</td>
      <td width="412"><input type="text" name="Email" size="36" tabindex="2"></td>
    </tr>
    <tr>
      <td width="640" align="left" colspan="3" height="10"></td>
    </tr>
    <tr>
      <td width="128" align="left"></td>
      <td width="105" align="left" valign="top">Comments:</td>
      <td width="412"><textarea rows="5" name="Comment" cols="35" tabindex="3"></textarea></td>
    </tr>
    <tr>
      <td width="640" align="left" colspan="3" height="10"></td>
    </tr>
    <tr>
      <td width="128" align="left"></td>
      <td width="405" align="left" colspan="2">

      <input type="button" value="&lt;&lt; Back" name="b3" onClick="javascript:history.back(1)" tabindex="6">
      <input type="reset" value="Clear All" name="b2" tabindex="5">&nbsp; 
      <input type="button" value="Submit &gt;&gt;" name="b1" onClick="javascript:check(this.form)" tabindex="4"></td>
      
    </tr>
  </table>
  </form>
</div>
  <p align="center"><font size="2">Copyright © 2002, All rights reserved by
  <a href="../about.htm" style="color: white">K. 
  Haque</a></font>
  </p>
        </td>
      </tr>
	</table>
  </body>
</html>