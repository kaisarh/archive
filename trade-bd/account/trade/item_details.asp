<!--#include file="../../inc_db.asp"-->
<%  		
  		Dim iItemID, iDB, bLoad, iCode

  		iCode = Request.querystring("iCode") ''type of item
		iItemID = Request.querystring("iItemID")

		const bgCol="#E3EFF2"
		iDB = int(iCode/10)
		bLoad = false

		select case iDB
			case 1 ''All
				select case iCode		
					case 11:
						sStr = "PC Accessory"

					case 12:
						sStr = "Electronics"

					case 13:
						sStr = "Car"
				end select

				sSql = "SELECT * FROM Alls WHERE ItemID=" & iItemID

			case 2	''PC

				sSql = "SELECT * FROM PCs WHERE ItemID=" & iItemID

			case 3 ''CD

				dim sStr2

				''the following strings are needed whether or not we are bLoading
				if iCode=32 then
			  		sStr = "Author"
			  		sStr2 = "Book"
			  	elseif iCode=31 then
			  		sStr = "Artist"
			  		sStr2 = "Audio/video/MP3 CD"
			  	elseif iCode=33 then
			  		sStr = "Description"
			  		sStr2 = "DVD"
			  	elseif iCode=34 then
			  		sStr = "Company"
			  		sStr2 = "Software/Game CD"
			  	end if
				''
				''sSql = "SELECT * FROM CDsNBooks WHERE iCode=" & iCode & " AND ItemID=" & iItemID
				sSql = "SELECT * FROM CDsNBooks WHERE ItemID=" & iItemID
		
		end select ''iDB
		
		oRS.Open sSql, oConn, 1

		if oRS.EOF then
			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing

			dim sMsg, sUrl
			sMsg = "Error: Data not found!"
			sUrl = "javascript:history.back()"
			%>
				<!--#include file="../../info.asp"-->
			<%
			response.end
		end if

		Dim UsedUnit, Condition, PriceUnit
		sSql = "SELECT * FROM Items WHERE ItemID=" & oRS("ItemID") & " AND iCode=" & oRS("iCode")
		oRS2.Open sSql, oConn, 1

		select case oRS2("UsedUnit")
			case 1
				UsedUnit = "Days"
			case 2
				UsedUnit = "Weeks"
			case 3
				UsedUnit = "Months"
			case 4
				UsedUnit = "Years"
		end select
			
		select case oRS2("Condition")
			case 1
				Condition = "100% New"
			case 2
				Condition = "Almost New"
			case 3
				Condition = "Not So Old"
			case 4
				Condition = "Old"
			case 5
				Condition = "Very Old"
		end select
		
		select case oRS2("PriceUnit")
			case 1
				PriceUnit = "Taka"
			case 2
				PriceUnit = "US$"
		end select

		if iDB=2 then ''PC
			Dim options
			if mid(oRS("Options"),1,1)=1 then options = "Keyboard, "
			if mid(oRS("Options"),2,1)=1 then options = options & "Mouse, "
			if mid(oRS("Options"),3,1)=1 then options = options & "Mini Speakers, "
			if mid(oRS("Options"),4,1)=1 then options = options & "Microphone, "
			if mid(oRS("Options"),5,1)=1 then options = options & "Dust Cover"
		end if
%>

	<html>

	<head>
	<title>Item Detials</title>
	<meta http-equiv="Content-Language" content="en-us" />
	<link rel="stylesheet" href="../../style.css" />

	</head>
	
	<body>
  	<table width="640">
      <tr>
        <td width="100%">
	<p align="justify">
<font size="3"><img border="0" src="../../images/arrow.gif">Item Details:</font><br />
Details about the selected item is shown below. You can reserve this item simply 
by clicking on the 'Reserve' button below. You'll not see the 'Reserve' button, 
if it is already reserved.<br>
		<br>
		<table width="633">
		  <tr>
		    <td width="627" align="middle">
			<% if iDB=1 then %>

       			  <table class=list width="393">
			   		<tr>
				    	<th class=list width="385" align="middle" colspan="2">
				 			<b>Item Details: <% =sStr %></b>
				    	 </th>
				  	</tr>
			        <tr>
				   	  <td class=list width="176" valign="top">Product Type:</td>
			          <td class=list width="203"><%=oRS("Name")%></td>
			        </tr>
			        <tr>
			          <td class=list width="176">Brand/Company:</td>
			          <td class=list width="203"><%=oRS("Brand")%>&nbsp;</td>
			        </tr>
			        <tr>
			          <td class=list width="176">Model/Specifications:</td>
			          <td class=list width="203"><%=oRS("Spec")%>&nbsp;</td>
			        </tr>
			        <tr>
			          <td class=list width="176">Additional Accessories:</td>
			          <td class=list width="203"><%=oRS("Acc")%>&nbsp;</td>
			        </tr>
			
			<% elseif iDB=2 then %>
			
       			  <table class=list width="460">
			   		<tr>
				    	<th class=list width="100%" align="middle" colspan="2">
				 			<b>Item Details: PC</b>
				    	 </th>
				  	</tr>
			        <tr>
			          <td class=list width="40%">Brand:</td>
			          <td class=list><%=oRS("Brand")%></td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Processor:</td>
			          <td class=list><%=oRS("Processor") & " " & oRS("Speed") & "MHz" %></td>
			        </tr>
			        <tr>
			          <td class=list>RAM:</td>
			          <td class=list><%=oRS("RAM") & " MB" %></td>
			        </tr>
			        <tr>
			          <td class=list>HDD:</td>
			          <td class=list><%=oRS("HDD") & " GB" %></td>
			        </tr>
			        <tr>
			          <td class=list>Motherborad Brand:</td>
			          <td class=list><%=oRS("MB")%></td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Display Card:</td>
			          <td class=list><%=oRS("Display")%></td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Sound Card:</td>
			          <td class=list><% =oRS("Sound") & "  " & oRS("SoundBit") & " Bit" %></td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Modem:</td>
			          <td class=list><% =oRS("Modem") & " "  & oRS("ModemSpeed") & " KBPs" %></td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Monitor:</td>
			          <td class=list><% =oRS("Monitor") & " " & oRS("MonitorSize") & " Inch" %></td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Includes:</td>
			          <td class=list><% =options %>&nbsp;</td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Additional Accessories:</td>
			          <td class=list><%=oRS("Acc")%>&nbsp;</td>
			        </tr>
 
			<% elseif iDB=3 then %>
       			  
       			  <table class=list width="380">
			   		<tr>
				    	<th class=list width="100%" colspan="2">
				 			<b>Item Details: <% =sStr2 %></b>
				    	 </th>
				  	</tr>
			        <tr>
			          <td class=list width="25%"><%=sStr%>:</td>
			          <td class=list><%=oRS("Author")%>&nbsp;</td>
			        </tr>
			        <tr>
			          <td class=list valign="top">Title:</td>
			          <td class=list><%=oRS("Title")%>&nbsp;</td>
			        </tr>
			        
			 <% end if %>
	
			        <tr>
			          <td class=list colspan="2">&nbsp;</td>
			        </tr>
	 		        <tr>
	 		          <td class=list>Inserted Date:</td>
			          <td class=list><%= oRS2("InsertedDate") %>&nbsp;</td>
		        	</tr>
	 		        <tr>
	 		          <td class=list>Used For:</td>
			          <td class=list><%= oRS2("Used") & " " & UsedUnit %>&nbsp;</td>
		        	</tr>
			        <tr>
			          <td class=list>Condition:</td>
			          <td class=list><% =Condition %>&nbsp;</td>
			        </tr>
			        <tr>
			          <td class=list>Price:</td>
			          <td class=list><b><%= oRS2("Price") & " " & PriceUnit %></b>&nbsp;</td>
			        </tr>
	      		</table>

				<table border="0" width="377" cellspacing="0" cellpadding="0">
			        <tr>
			          <td width="377" colspan="2">
			          	<font size="1">&nbsp;</font></td>
			        </tr>
			        <tr>
			          <td width="182">
			          	<p align="left"><a href="javascript:history.back();">
                        &lt;&lt;&nbsp;Back</a></td>
			          <td align="right" width="195">
				       	<p><% if oRS("iStatus")=0 then %>
				       		<a href="reserve.asp?iItemID=<% =iItemID %>&iCode=<% =iCode %>">
                        	Reserve Item &gt;&gt;</a>
                        <% else %>
                        	&nbsp;
                        <% end if %>
                        </td>
			        </tr>
	      	</table>
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
<%  if bLoad then
		oRS.Close
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
	end if
%>