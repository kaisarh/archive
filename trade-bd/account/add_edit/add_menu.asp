<!-- #include file='inc_add_login.asp' -->

<html>

<head>
	<meta http-equiv="Content-Language" content="en-us" />
	<link rel="stylesheet" href="../../style.css" />

<script language="JavaScript">
<!--

function gourl(){
	var loc
	for (var i = 0; i < 9; i++) {
		if (this.document.form1.R1[i].checked) 
			loc = this.document.form1.R1[i].value;
	}

	window.location = loc;
}

//-->
    </script>

<title>Sell</title>
</head>

<body>
  	<table width="640">
      <tr>
        <td width="100%">
<p>
<font size="3"><img border="0" src="../../images/arrow.gif">Welcome to Add New Item Wizard!</font><br />
There are only two steps to add a new item. First, choose the item type you want 
to add, then you give us the details of about your item. The first step is in 
this page.</p>
<p align="center">
<strong>Currently available item types:</strong></p>
<div align="center"><center>

<table border="1" width="364" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" style="border-collapse: collapse">
  <tr>
    <td width="365">
    	<form name="form1">
          <table border="0" width="365" cellspacing="0" cellpadding="0">
            <tr>
              <td width="38">&nbsp;</td>
              <td width="327">
	              <font size="1">&nbsp;</font><br>
	              <input type="radio" value="add_all.asp?iCode=13" name="R1" checked>Vehicle 
                  (Bike, Car, Bus etc)<Br>
	              <input type="radio" value="add_pc.asp?" name="R1">Computer 
                  (Desktop, Laptop etc)<br>
	              <input type="radio" value="add_all.asp?iCode=11" name="R1">PC Accessory
	              (Processor, RAM, Printer etc)<br>
	              <input type="radio" value="add_all.asp?iCode=12" name="R1">Electronic Device (Mobile, 
                  Discman
	              etc)<br>
	              <input type="radio" value="add_cdsnbooks.asp?iCode=31" name="R1">Audio/Video/MP3 CD<br>
	              <input type="radio" value="add_cdsnbooks.asp?iCode=34" name="R1">Software/Game CD<br>
	              <input type="radio" value="add_cdsnbooks.asp?iCode=33" name="R1">DVD<br>
	              <input type="radio" value="add_cdsnbooks.asp?iCode=32" name="R1">Book<br>
	              <input type="radio" value="add_all.asp?iCode=14" name="R1">Others<br>
                  <font size="1">&nbsp;</font></td>
            </tr>
            <tr>
              <td width="365" colspan="2">
              <p align="center">
              <input type="button" value="&lt;&lt; Back" name="B4" OnClick="javascript:history.back(1)">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Next &gt;&gt;" name="B5" OnClick="javascript:gourl()">
              <font size="1">&nbsp; </font></td>
            </tr>
          </table>
        </form>
    </td>
  </tr>
</table>
</center></div>
<p align="center"><font size="2">Copyright © 2002, All rights reserved by
<a href="../../about.htm" style="color: white">K. Haque</a></font></p>
        </td>
      </tr>
	</table>
</body>
</html>