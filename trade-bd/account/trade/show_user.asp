<!--#include file="../../inc_db.asp"-->
<%  		
  		Dim iUserID, iItemID, iCode
  		
  		iUserID = Request.querystring("iUserID")

		Dim sStr
		sSql = "SELECT * FROM Users WHERE UserID = " & iUserID
		oRS.Open sSql, oConn, 1

		if oRS.EOF then
			oRS.Close
			'oConn.Close
			'Set oRS = Nothing
			'Set oConn = Nothing
			
			dim sMsg, sUrl
			sMsg = "Error: User not found!"
			sUrl = "javascript:history.back()"
			%>
				<!--#include file="../../info.asp"-->
			<%
			response.end
		end if

%>

	<html>
	
	<head>
		<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
		<title>User Detials</title>
		<meta http-equiv="Content-Language" content="en-us" />
		<link rel="stylesheet" href="../../style.css" />

	</head>

	<body>
  	<table width="640">
      <tr>
        <td width="100%">
	<p>
    <font size="3"><img border="0" src="../../images/arrow.gif">Contacting Info:</font><br />
Here is the contacting information of the user of your interest. You are 
responsible for contacting with this person.</p>
	<table width="640">
	  	<tr>
	    	<td colspan="3" align="middle" width="634">
       			  <table class=list width="380">
			   		<tr>
				    	<th class=list width="100%" align="middle" colspan="2">
				 			<strong>
				 			<% if Request.querystring("bBuy")="Yes" then
							 	response.write("Seller ")
							 else
							 	response.write("Buyer ")
							 end if %>Details:
				 			</strong>
				    	 </th>
				  	</tr>
			        <tr>
				   	  <td class=list width="25%" align="middle">Name:</td>
			          <td class=list><%=oRS("Name")%>&nbsp;</td>
			        </tr>
			        <tr>
			          <td class=list rowspan="2" align="middle">Adress:</td>
			          <td class=list><%=oRS("Add1")%>&nbsp;</td>
			        </tr>
			        <tr>
			          <td class=list><%=oRS("Add2")%>&nbsp;</td>
			        </tr>
	
			        <tr>
			          <td class=list align="middle">Phone:</td>
			          <td class=list><%=oRS("Phone")%>&nbsp;</td>
			        </tr>
	
			        <tr>
			          <td class=list align="middle">Email:</td>
			          <td class=list><%=oRS("Email")%>&nbsp;</td>
			        </tr>
			     </table>
			</td>
	  	</tr>
	  	<tr>
			<td width="640" align="right" colspan="3">
	      		<font size="1">&nbsp;</font></td>
  	  	</tr>
	  	<tr>
			<td width="269" align="right">
	      		<a href="../account.asp">&lt; &lt; Account's Page</a>&nbsp;
	      	</td>
	      	<td width="112"></td>
	      	<td width="245">
	        	<a href="mailto:<% =oRS("Email") %>?Subject=Your Item at SellNBuy.Com">Send E-Mail 
                &gt;&gt;</a></td>
  	  	</tr>
	</table>
  <p align="center"><font size="2">Copyright © 2002, All rights reserved by
  <a href="../../about.htm" style="color: white">K. Haque</a></font></p>
        </td>
      </tr>
	</table>
  	</body>
</html>
<%
		oRS.Close
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
%>