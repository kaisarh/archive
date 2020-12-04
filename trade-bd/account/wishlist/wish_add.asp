<% if NOT Session("bLogin") then
	Dim sMsg, sUrl
	sMsg = "You need to login in order to add a wish to the Wish List!"
	sUrl = "../login.asp?bRedirect=yes"
	session("sLogin_Redirect")="wishlist/wish_add.asp"		'check for error on edit items
	%>
	<!--#include file="../../info.asp"-->
	<%
	response.end
end if %>
<html>

<head>
<title>Add a Wish</title>
<meta http-equiv="Content-Language" content="en-us" />
<link rel="stylesheet" href="../../style.css" />
<script language="JavaScript">
<!--
// bug: selection checking not here
function check(frm) {
	
	if (frm.Wish.value.length < 2)
		window.alert("Please enter the Wish correctly.");
	
	else if (!(frm.Price.value > 0))
		window.alert("Please enter the Price correctly.")

	else
		frm.submit();
}
//-->
    </script>
</head>

<body>

<table width="640">
  <tr>
    <td width="100%"><big><img border="0" src="../../images/arrow.gif">Add your 
    need to Wish List:</big><br>
&nbsp;&nbsp;&nbsp; If you are looking for something that is not on the list or already 
    sold, don&#39;t give up. Try this <strong>Wish List</strong>. Other members can 
    know about your need and may be able to offer you it by viewing this list.<form method="POST" action="wish_doadd.asp">
      <table width="640">
        <tr>
          <td width="115">&nbsp;</td>
          <td width="115">I need a</td>
          <td width="396"><input type="text" name="Wish" size="36" tabindex="1"><font size="2"> 
          [within 80 chars]</font></td>
        </tr>
        <tr>
          <td width="115">&nbsp;</td>
          <td width="115">Favoralbe Price<br>
&nbsp;</td>
          <td width="396"><input type="text" name="Price" size="6" tabindex="2">&nbsp;
          <select name="PriceUnit" size="1" tabindex="3">
          <option value="Taka">Taka</option>
          <option value="US$">US$</option>
          </select><br>
&nbsp;</td>
        </tr>
        <tr>
          <td width="115">
          <input type="button" value="&lt;&lt; Back" name="b3" onClick="javascript:history.back(1)" style="float: right" tabindex="6"></td>
          <td width="515" colspan="2">
          <input type="reset" value="Clear" name="b2" tabindex="5">&nbsp;&nbsp;
          <input type="button" value="Submit &gt;&gt;" name="b1" onClick="javascript:check(this.form);" tabindex="4"></td>
        </tr>
      </table>
    </form>
    <p align="center"><font size="2">Copyright © 2002, All rights reserved by
    <a href="../../about.htm" style="color: white">K. Haque</a></font></p>
    </td>
  </tr>
</table>

</body>

</html>