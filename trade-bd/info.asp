<%
''this is a sub page
'' pass the sMsg and sUrl variable

	if sIcon="" then
		''Dim sIcon
		sIcon = "i"
	end if
%>
<html>

<head>
	<title>Message</title>
	<meta http-equiv="Content-Language" content="en-us" />
	<link rel="stylesheet" 
	<% if sCSSLoc = "" then %>
		href="../../style.css"
	<% else %>
		href="<% =sCSSLoc %>"
	<% end if %>
	 />

</head>
<body>
  	<table width="640">
      <tr>
        <td width="100%">

&nbsp;<div align="center">
  <center>

<table class=msg width="365">
  <tr >
    <th class=msg width="359" colspan="3">&nbsp;Message</th>
  </tr>
  <tr>
    <td width="51">
    <font size="1">&nbsp;</font><div align="center">
      <center>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ffffff" width="60%" id="AutoNumber1">
        <tr>
          <td width="100%" valign="top">
          <p align="center"><b><font color="#ffffff" size="5"><% =sIcon %></font></b></td>
        </tr>
      </table>
      </center>
    </div>
    <p><font size="1">&nbsp;</font><br>
&nbsp;</td>
    <td width="282"><p align="justify"><font size="1"><br>
    </font><% =sMsg %>&nbsp;&nbsp; <br>
    <font size="1">&nbsp;</font><div align="center">
      <center>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ffffff" width="27%" id="AutoNumber2">
        <tr>
          <td width="100%">
          <p align="center">
		<a class=table href="<%=sUrl%>">Alright!</a></td>
        </tr>
      </table>
      </center>
    </div>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ffffff" width="100%" id="AutoNumber3">
      <tr>
        <td width="100%">&nbsp;</td>
      </tr>
    </table>
    </td>
    <td width="18">&nbsp;</td>
  </tr>
</table></center>
</div>

	<p align="center"><font size="2">Copyright � 2002, All rights reserved by
    <a href="http://www.trade-bd.com/about.htm">K. Haque</a></font></td>
      </tr>
	</table>

</body>
</html>