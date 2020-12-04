<!-- #include file='inc_add_login.asp' -->
<!-- #include file='inc_add_ado.asp' -->

<%
	dim sStr1, sStr2
	
	''the following strings are needed whether or not we are bLoading
	if iCode=32 then
  		sStr1 = "Author Name"
  		sStr2 = "Book"
  	elseif iCode=31 then
  		sStr1 = "Artist Name"
  		sStr2 = "Audio/video/MP3 CD"
  	elseif iCode=33 then
  		sStr1 = "Description"
  		sStr2 = "DVD"
  	elseif iCode=34 then
  		sStr1 = "Company"
  		sStr2 = "Software/Game CD"
  	end if
	''
	
	if bLoad then
		sSQL = "SELECT * FROM CDsNBooks WHERE ItemID=" & iItemID
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

function check(frm) {
	
	if (frm.Author.value=="")
		window.alert("Please enter the <% =sStr1 %> correctly.");
	
	else if (frm.Title.value=="")
		window.alert("Please enter the Title correctly.")
	
	else if (!(frm.Used.value > 0 && frm.Used.value < 100))
		window.alert("Please enter the Used time correctly.")
	
	else if (!(frm.Price.value > 0))
		window.alert("Please enter the Price correctly.")
	else
		frm.submit();
}

function mark(){
	<% if bLoad then %>
		form1.Author.value="<%=oRS("Author")%>";
		form1.Title.value="<%=oRS("Title")%>";

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
	    <td width="100%">
	    	<strong>
	      		<% if bLoad then %>
          			Update <% =sStr2 %>:
    			<% else %>
    				Add a <% =sStr2 %> for sell:
		      	<%end if %>
	    	</strong>
	    	<form method="POST" action="<% =sUrl %>iCode=<% =iCode %>" name="form1">
		      
		      <table border="0" width="100%">
		        <tr>
		          <td width="20%"><%=sStr1%>:</td>
		          <td>
		          	<input type="text" name="Author" size="30">
		          		<small> <font size="1">[Type Within 30 Chars]</font></small>
		          </td>
		        </tr>
		        <tr>
		          <td valign="top">Title:</td>
		          <td>
		          	<input type="text" name="Title" size="30">
		          		<small> <font size="1">[Type Within 30 Chars]</font></small>
		          </td>
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