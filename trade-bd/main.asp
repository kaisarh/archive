<!--#include file="inc_db.asp"-->
<html>

<head>
<title>Main</title>
<meta http-equiv="Content-Language" content="en-us" />
<link rel="stylesheet" href="style.css" />
<base target="_self">
</head>

<body leftmargin="6">

<table width="640">
  <tr>
    <td width="100%">
    <p align="justify"><img border="0" src="images/arrow.gif"><b>Item List (currently available items):</b><br>
    The currently available items for sale are listed 
    below. Click on the categories to get a complete list of items of that category. 
    [Please note that categories with 0 items do not have a link].</p>
    <div align="center">
      <center>
      <table class="list" width="384">
        <tr>
          <th class="list" width="270">Category</th>
          <th class="list" width="98">No. of Items</th>
        </tr>
        <tr>
          <%
							Dim sColor
							Dim x, hit

							sSQL = "SELECT * FROM Count"
							Set oRS = oConn.Execute(sSQL)
							
							sColor = false
							
							sub add_tag(iCode, title)
								sColor = Not sColor
							   	%> 
        </tr>
        <tr>
          <td class="list" align="middle" width="270"><% if oRS("iCode" & iCode) > 0 then %>
          <a class="table" href="list.asp?iCode=<% =iCode %>"><% =title %></a> <% else %>
          <% =title %> <% end if %> </td>
          <td class="list" align="middle" width="98"><% =oRS("iCode" & iCode) %></td>
        </tr>
        <%	end sub
							hit = oRS("Hit")

							add_tag 20, "Computer (Desktop, Laptop etc)"
							add_tag 13, "Vehicle (Bike, Car etc)"
							add_tag 11, "PC Accessories (RAM, Printer etc)"
							add_tag 12, "Electronic Device (Mobile, Discman etc)"
							add_tag 31, "Audio/Video/MP3 CDs"
							add_tag 34, "Software/Games CDs"
							add_tag 33, "DVDs"
							add_tag 32, "Books"
							add_tag 14, "Others"

							oRS.Close
							Set oRS = oConn.Execute("UPDATE Count Set Hit=" & (hit+1))

							'Set oRS = Nothing
							'Set oConn = Nothing
							%> 
      </table>
      </center>
    </div>
    <p align="justify"><img border="0" src="images/arrow.gif"><b>Site News:<br>
&nbsp;&nbsp; November 23, 2002:</b> New look and feel!<b><br>
&nbsp;&nbsp; October 2, 2002:</b> Some bugs were fixed.<br>
    &nbsp;<br>
    <img border="0" src="images/arrow.gif"><b>Coming Soon:</b><br>
    There are a lot of new things that will be added to 
    this site very soon. For example:</p>
    <ul>
      <li>
      <p align="justify">Auction Functionality 
      </li>
      <li>
      <p align="justify">E-mail address validation 
      </li>
      <li>
      <p align="justify">Password Recovery 
      </li>
      <li>
      <p align="justify">Complete &#39;Site News&#39; section
      
      </li>
      <li>
      <p align="justify">And More... 
      </li>
    </ul>
    <p align="justify"><img border="0" src="images/arrow.gif"><b>Are you new here?</b><br>
    If you are new here then please visit our beginner&#39;s 
    page where you&#39;ll know about the basics of this site.
    <a href="beginners.htm">Click here</a> to go to the Beginner&#39;s Page.</p>
    <p align="center">This page has been visited <b><% =hit %></b> 
    times since 04-09-2002.</p>
    <p align="center">Copyright © 2002, All rights reserved by
    <a href="about.htm">K. Haque</a><br>
    This site is best viewed with medium font size at 800 x 600 resolution.</p>
    </td>
  </tr>
  <tr>
    <td width="100%">
    &nbsp;</td>
  </tr>
</table>
</body>

</html>