<!-- #include file='inc_add_login.asp' -->
<!-- #include file='inc_add_ado.asp' -->

<%
	if bLoad then
		sSQL = "SELECT * FROM PCs WHERE ItemID=" & iItemID
		oRS.Open sSQL, oConn, 1

		if oRS.EOF then
			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing
			
			dim sMsg
			sMsg = "Error: Data not found!"
			sUrl = "account.asp"
			%>
				<!--#include file="../../info.asp"-->
			<%
			response.end
		end if

		sSQL = "SELECT * FROM Items WHERE ItemID=" & iItemID & " AND iCode=" & iCode
		oRS2.Open sSQL, oConn, 1
	end if
%>

<html>

<head>
	<meta http-equiv="Content-Language" content="en-us" />
	<link rel="stylesheet" href="../../style.css" />
	<title>Add PC</title>

<script language="JavaScript">
<!--

function check(frm) {

	if (!(frm.Speed.value > 0))
		window.alert("Please enter the Speed correctly.");

	else if (!(frm.RAM.value > 0))
		window.alert("Please enter the RAM Size correctly.");

	else if (!(frm.HDD.value > 0))
		window.alert("Please enter the HDD Size correctly.");

	else if (frm.ModemC.value=="[None]" && frm.ModemS.value!="0")
		window.alert("Please select the Modem correctly.")
	
	else if (frm.ModemS.value=="0" && frm.ModemC.value!="[None]")
		window.alert("Please select the Modem correctly.")

	else if (frm.SoundC.value=="[None]" && frm.SoundT.value!="0")
		window.alert("Please select the Sound Card correctly.")
		
	else if (frm.SoundT.value=="0" && frm.SoundC.value!="[None]")
		window.alert("Please select the Sound Card correctly.")
		
	else if (!(frm.Used.value > 0 && frm.Used.value < 100))
		window.alert("Please enter the Used time correctly.")
	
	else if (!(frm.Price.value > 0))
		window.alert("Please enter the Price correctly.")
	
	else
		frm.submit();
}

function mark(){
	<%if bLoad then %>
		form1.Brand.value="<%=oRS("Brand")%>";
		form1.Processor.value="<%=oRS("Processor")%>";
		form1.Speed.value="<%=oRS("Speed")%>";
		form1.RAM.value="<%=oRS("RAM")%>";
		form1.HDD.value="<%=oRS("HDD")%>";
		form1.MB.value="<%=oRS("MB")%>";
		form1.DisT.value="<%=left(oRS("Display"),3)%>";
		form1.DisS.value="<%=mid(oRS("Display"),5)%>";
		
		form1.SoundC.value="<%=oRS("Sound")%>";
		form1.SoundT.value="<%=oRS("SoundBit")%>";
		form1.ModemC.value="<%=oRS("Modem")%>";
		form1.ModemS.value="<%=oRS("ModemSpeed")%>";
		form1.MonitorC.value="<%=oRS("Monitor")%>";
		form1.MonitorS.value="<%=oRS("MonitorSize")%>";
		form1.Acc.value="<%=oRS("Acc")%>";

		form1.C1.value="<%=mid(oRS("Options"),1,1)%>";
		form1.C2.value="<%=mid(oRS("Options"),2,1)%>";
		form1.C3.value="<%=mid(oRS("Options"),3,1)%>";
		form1.C4.value="<%=mid(oRS("Options"),4,1)%>";
		form1.C5.value="<%=mid(oRS("Options"),5,1)%>";
	
		form1.Used.value="<%=oRS2("Used")%>";
		form1.UsedUnit.value="<%=oRS2("UsedUnit")%>";
		form1.Price.value="<%=oRS2("Price")%>";
		form1.PriceUnit.value="<%=oRS2("PriceUnit")%>";
		form1.Condition.value="<%=oRS2("Condition")%>";
	<% end if %>
}

//-->
    </script>

</head>

<body onload="javascript:mark();">
  	<table width="640">
      <tr>
        <td width="100%">

<p>
    <font size="3"><img border="0" src="../../images/arrow.gif">Details of your item:</font><br />
Please enter the detailed information about the item. Please enter the most 
accurate information about 'Used Time' and present 'Condition'. Enter the price 
you would like your item to sell for.</p>

<table border="0" width="520" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="100%"><strong><% if bLoad then
  			Response.write("Update ")
		else
  			Response.write("Add a ")
      	end if %>Computer for sell:</strong>
      	
    <form method="POST" action="<% =sUrl %>iCode=20" name="form1">
      <table border="0" width="100%">
        <tr>
          <td width="15%" align="left">Brand:</td>
          <td><select name="Brand" size="1">
            <option selected value="- Local Clone -">[Clone PC]</option>
            <option value="IBM">IBM</option>
            <option value="Compaq">Compaq</option>
            <option value="Dell">Dell</option>
            <option value="HP">HP</option>
            <option value="Packer Bell">Packer Bell</option>
            <option value="Toshiba">Toshiba</option>
            <option value="Gateway">Gateway</option>
            <option value=" - Other Brand -"> - Other Brand -</option>
          </select></td>
        </tr>
        <tr>
          <td valign="top">Processor:</td>
          <td ><select name="Processor" size="1">
            <option value="Cyrix">Cyrix</option>
            <option value="Cyrix M2">Cyrix M2</option>
            <option value="AMD K6">AMD K6</option>
            <option value="AMD K62">AMD K62</option>
            <option value="AMD K7">AMD K7</option>
            <option value="AMD Duron">AMD Duron</option>
            <option value="AMD Athlon">AMD Athlon</option>
            <option value="Intel 386">Intel 386</option>
            <option value="Intel 486">Intel 486</option>
            <option selected value="Intel Pentium I">Intel Pentium I</option>
            <option value="Intel Pentium II">Intel Pentium II</option>
            <option value="Intel Pentium III">Intel Pentium III</option>
            <option value="Intel Pentium 4">Intel Pentium 4</option>
            <option value="Intel Celeron">Intel Celeron</option>
            <option value="- Other Type -">- Other Type -</option>
          </select> Speed: <input type="text" name="Speed" size="4" tabindex="2">MHz</td>
        </tr>
        <tr>
          <td>RAM:</td>
          <td ><input type="text" name="RAM" size="5" tabindex="1"> MB</td>
        </tr>
        <tr>
          <td>HDD:</td>
          <td ><input type="text" name="HDD" size="5" tabindex="2"> GB</td>
        </tr>
        <tr>
          <td>Motherborad Brand:</td>
          <td ><select name="MB" size="1">
            <option selected value="[Unknown]">[Unknown]</option>
            <option value="Intel">Intel</option>
            <option value="Asus">Asus</option>
            <option value="DFI">DFI</option>
            <option value="GigaByte">GigaByte</option>
            <option value="Octech">Octech</option>
            <option value="AOpen">AOpen</option>
            <option value="RedFox">RedFox</option>
            <option value="-Other Brand-">-Other Brand-</option>
          </select></td>
        </tr>
        <tr>
          <td valign="top">Display Card</td>
          <td ><select name="DisT" size="1">
            <option selected value="VGA">VGA</option>
            <option value="AGP">AGP</option>
          </select> VRAM: <select name="DisS" size="1">
            <option value="1 MB">1 MB</option>
            <option value="2 MB">2 MB</option>
            <option value="4 MB">4 MB</option>
            <option value="8 MB">8 MB</option>
            <option value="16 MB">16 MB</option>
            <option value="32 MB">32 MB</option>
            <option value="64 MB">64 MB</option>
            <option value="128 MB">128 MB</option>
          </select></td>
        </tr>
        <tr>
          <td valign="top">Sound Card</td>
          <td ><select name="SoundC" size="1">
            <option selected value="[None]">[None]</option>
            <option value="Creative">Creative</option>
            <option value="ESS">ESS</option>
            <option value="Yamaha">Yamaha</option>
            <option value="--Other--">--Other--</option>
          </select> <select name="SoundT" size="1">
            <option selected value="0">[None]</option>
            <option value="8">8</option>
            <option value="16">16</option>
            <option value="32">32</option>
            <option value="64">64</option>
            <option value="128">128</option>
            <option value="256">256</option>
            <option value="--Other--">--Other--</option>
          </select>Bit</td>
        </tr>
        <tr>
          <td valign="top">Modem:</td>
          <td ><select name="ModemC" size="1">
            <option selected value="[None]">[None]</option>
            <option value="US Robotics">US Robotics</option>
            <option value="Motrola">Motrola</option>
            <option value="Genius">Genius</option>
            <option value="Prolink">Prolink</option>
            <option value="--Other--">--Other--</option>
          </select> Speed: <select name="ModemS" size="1">
            <option selected value="0">[None]</option>
            <option value="9">9</option>
            <option value="33">33</option>
            <option value="56">56</option>
            <option value="128">128</option>
            <option value="256">256</option>
          </select>KBPS</td>
        </tr>
        <tr>
          <td valign="top">Monitor:</td>
          <td ><select name="MonitorC" size="1">
            <option value="[Same As CPU]">[Same As CPU]</option>
            <option value="Samsang">Samsang</option>
            <option value="LG">LG</option>
            <option value="Hyundai">Hyundai</option>
            <option value="ViewSonic">ViewSonic</option>
            <option value="Compro">Compro</option>
            <option value="--Other--">--Other--</option>
          </select> Size: <select name="MonitorS" size="1">
            <option value="14">14</option>
            <option value="15">15</option>
            <option value="17">17</option>
            <option value="21">21</option>
            <option value="--Other--">--Other--</option>
          </select>&quot;(Inch)</td>
        </tr>
        <tr>
          <td valign="top">Includes:</td>
          <td ><input type="checkbox" name="C1" value="ON" checked>Keyboard&nbsp; <input
          type="checkbox" name="C2" value="ON" checked>Mouse <input type="checkbox" name="C3"
          value="ON">Mini Speakers <input type="checkbox" name="C4" value="ON">Microphone <input
          type="checkbox" name="C5" value="ON">Dust Cover</td>
        </tr>
        <tr>
          <td valign="top">Additional Accessories:</td>
          <td ><input type="text" name="Acc" size="22" tabindex="2"> <small>[Type Within 30 Chars]</small></td>
        </tr>
        <!-- #include file='inc_price.asp' -->
      </table>
    </form>
    </td>
  </tr>
</table>
  <p align="center"><font size="2">Copyright © 2002, All rights reserved by
  <a href="../../about.htm" style="color: white">K. Haque</a></font></p>
        </td>
      </tr>
	</table>
  	</body>
</html>