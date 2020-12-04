<!-- #include file='inc_add_login.asp' -->
<!-- #include file='inc_add_ado.asp' -->

<%
	Dim sStr
	if iCode=11 then
  		sStr = "a PC Accessory"

  	elseif iCode=12 then
		sStr = "an Electronic Device"

  	elseif iCode=13 then
		sStr = "a Car"

  	elseif iCode=14 then
		sStr = "an Item"
  	end if

	if bLoad then
		sSQL = "SELECT * FROM Alls WHERE ItemID=" & iItemID
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

<script language="JavaScript">
<!--

// bug: selection checking not here

function check(frm) {

	if (frm.Brand.value=="")
		window.alert("Please enter the Brand correctly.");

	else if (frm.Spec.value=="")
		window.alert("Please enter the Specifications correctly.")

	else if (!(frm.Used.value > 0 && frm.Used.value < 100))
		window.alert("Please enter the Used time correctly.")

	else if (!(frm.Price.value > 0))
		window.alert("Please enter the Price correctly.")

	else
		frm.submit();
}

function mark(){
	<%if bLoad then %>
		form1.s<%=iCode%>.value="<%=oRS("Name")%>";
		form1.Brand.value="<%=oRS("Brand")%>";
		form1.Spec.value="<%=oRS("Spec")%>";
		form1.Acc.value="<%=oRS("Acc")%>";
	
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
	<table border="0" width="480" cellspacing="0" cellpadding="0" align="center">
	  <tr>
	    <td width="100%"><strong>
			<% if bLoad then
      			Response.write("Update ")
			else
      			Response.write("Add ")
	      	end if %>
	    	<% =sStr %> for sell:</strong>
	     <form method="POST" action="<% =sUrl %>iCode=<% =iCode %>" name="form1">
	      <table border="0" width="497">
	        <tr>
	          <td width="150" align="left" valign="top">Product Type:</td>
	          <td width="337"><table border="0" width="100%" cellspacing="0" cellpadding="0">
				<% if iCode=11 then %>
		            <tr>
		              <td width="100%"><select name="s11" size="1">
		                <option value="--Select--">--Select--</option>
		                <option value="Processor">Processor</option>
		                <option value="Mother Board">Mother Board</option>
		                <option value="RAM">RAM</option>
		                <option value="Hard Disk">Hard Disk</option>
		                <option value="Floppy Drive">Floppy Drive</option>
		                <option value="CD-Rom Drive">CD-Rom Drive</option>
		                <option value="DVD-Rom Drive">DVD-Rom Drive</option>
		                <option value="CD Writer">CD Writer</option>
		                <option value="Monitor">Monitor</option>
		                <option value="Keyboard">Keyboard</option>
		                <option value="Mouse">Mouse</option>
		                <option value="Printer">Printer</option>
		                <option value="Scanner">Scanner</option>
		                <option value="Speakers">Speakers</option>
		                <option value="VGA/AGP Card">VGA/AGP Card</option>
		                <option value="SoundCard">SoundCard</option>
		                <option value="Fax/Modem">Fax/Modem</option>
		                <option value="TV Tuner">TV Tuner</option>
		                <option value="MPEG Card">MPEG Card</option>
		              </select></td>
		              
		            </tr>
				<% elseif iCode=12 then %>
		            <tr>
		              <td><select name="s12" size="1">
		                <option value="--Select--">--Select--</option>
		                <option value="Mobile Set">Mobile Set</option>
		                <option value="Mobile &amp; SIM">Mobile &amp; SIM</option>
		                <option value="SIM Card">SIM Card</option>
		                <option value="TV">TV</option>
		                <option value="Freezer">Freezer</option>
		                <option value="Washing Machine">Washing Machine</option>
		                <option value="CD Player/Sound System">CD Player/Sound System</option>
		                <option value="Radio Cassate Recorder">Radio Cassate Recorder</option>
		                <option value="WalkMan [Cassate]">WalkMan [Cassate]</option>
		                <option value="DiskMan [CD]">DiskMan [CD]</option>
		                <option value="DiskMan [MD]">DiskMan [MD]</option>
		                <option value="MP3 Player">MP3 Player</option>
		                <option value="VCD Player">VCD Player</option>
		                <option value="DVD Player">DVD Player</option>
		                <option value="Still Camera">Still Camera</option>
		                <option value="Digital Camera">Digital Camera</option>
		                <option value="Video Camera/Camcorder">Video Camera/Camcorder</option>
		                <option value="PDA/PalmTop PCs">PDA/PalmTop PCs</option>
		              </select></td>
		            </tr>
				
				<% elseif iCode=13 then %>
		            <tr>
		              <td><select name="s13" size="1">
		                <option value="--Select--">--Select--</option>
		                <option value="Bike">Bike</option>
                        <option value="Motor Bike">Motor Bike</option>
		                <option value="Private Car">Private Car</option>
		                
		                <option value="Taxi CAB">Taxi CAB</option>
                        <option value="Jeep">Jeep</option>
                        <option value="CNG Autorickshaw">CNG Autorickshaw
                        </option>
                        <option value="Microbus">Microbus</option>
                        <option value="Mini Bus">Mini Bus</option>
                        <option value="Luxuary Bus">Luxuary Bus</option>
                        <option value="PickUp">PickUp</option>
                        <option value="Mini Truck">Mini Truck</option>
                        <option value="Heavy Truck">Heavy Truck</option>
		                
		              </select></td>
		            </tr>
				
				<% else%>
		            <tr>
		              <td><input type="text" name="s14" size="20" tabindex="1"><small>
                      <font size="1">[Type Within 30 Chars]</font></small></td>
		            </tr>
				<% end if %>
	            </table>
	          </td>
	        </tr>
	        <tr>
	          <td width="150">Brand/Company:</td>
	          <td width="337"><input type="text" name="Brand" size="30"><small><font
	          color="#008000"> </font><font size="1">[Type Within 30 Chars]</font></small></td>
	        </tr>
	        <tr>
	          <td width="150">Model/Specifications:</td>
	          <td width="337"><input type="text" name="Spec" size="30"><small><font
	          color="#008000"> </font><font size="1">[Type Within 50 Chars]</font></small></td>
	        </tr>
	        <tr>
	          <td width="150">Additional Accessories:</td>
	          <td width="337"><input type="text" name="Acc" size="30"><small><font color="#008000"> </font>
              <font size="1">[Type Within 50 Chars]</font></small></td>
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