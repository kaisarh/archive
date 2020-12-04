<!--#include file="../inc_db.asp"-->
<html>
	<head>
		<title>Showing guest book entries</title>

		<meta http-equiv="Content-Language" content="en-us" />
		<link rel="stylesheet" href="../style.css" />

	</head>
	<body>

		<% sub page_link(s, pg_num)
			response.write("<a href=" & chr(34) & "guest_view.asp?pg=" & pg_num & _
					chr (34) & ">" & s & "</a>")
		end sub

		Dim iNTotal, iNStart, iNEnd, iNCount, iPg
		const PAGE_SIZE = 10
		sSQL = "SELECT * FROM GuestBook ORDER BY InsertedDate DESC"
		
		oRS.Open sSQL, oConn, 1
		iNTotal = oRS.RecordCount
	%>
		<table width="640">
				<tr>
					<td width="100%" colspan="3">
			    		<table width="640">
							<tr>
							    <td width="50%" align="left"><strong>Guest Book</strong></td>
							    <td align="right">
                                <a href="../main.asp">Main</a> | <a
							    	href="guest_add.asp">Add Comments</a>
							    	</td>
							</tr>
						</table>
					</td>
				</tr>
			  	<tr>
			    	<td width="100%" colspan="3"><hr size="1" color="#FFFFFF"></td>
			    </tr>
			
	<%	if not oRS.EOF then %>
				<tr>
					<td width="84%">
				<%
				''this is the current page number
				iPg = Request.querystring("pg")
				if iPg = "" then iPg = 1
				
				''page size should be set first
				oRS.PageSize = PAGE_SIZE
				oRS.AbsolutePage = iPg
			
			
				''other variables to handle
				iNStart = (iPg-1)*PAGE_SIZE + 1
				iNEnd = iNStart + PAGE_SIZE - 1
				if iNEnd > iNTotal then iNEnd = iNTotal
				iNCount = iNStart
			
				
				''now we set the Prev and next button	
				
				Response.write ("Total <Strong>" & iNTotal & "</Strong> Entries! ") 
				Response.write ("Displaying <Strong>" & _
					iNStart & " to " & iNEnd & "</Strong>")
		
				''now we write the links
				%>
					</td>
					<td width="8%" align="right">
				<%		if iNStart <> 1 then page_link "Prev", iPg-1 %>
					</td>

					<td align="right">
				<%		if iNTotal > iNEnd then page_link "Next", iPg+1 %>
					</td>
				</tr>
				<tr>
					<td width="100%" colspan="3" height="6"></td>
				</tr>
				<tr>
					<td width="100%" colspan="3">
			    
			<% while (iNCount <= iNEnd and not oRS.EOF) %>
					
					<table class=list width="640">
						<tr>
							<td class=list><font size="2"><b>Name: </b> <% =oRS("Name") %>,&nbsp;&nbsp; <b>E-mail:</b>
								<a class=table href="emailto:<% =oRS("Email") %>"><% =oRS("Email") %></a>,&nbsp;&nbsp;
								<b>Date:</b> <% =FormatDateTime(oRS("InsertedDate"),1) %>
								<br>&nbsp;&nbsp;<b>Comment:</b>&nbsp;
								<% = Replace(oRS("Comment"), CStr(vbcrlf), "<br>", 1) %></font>
							</td>
						</tr>
					</table>
					<table width="640">
						<tr>
							<td width="100%" height=8></td>
						</tr>
					</table>
					<% oRS.MoveNext
					iNCount = iNCount + 1
				wend
		end if %>
				</td>
			</tr>
			<tr align="middle">
				<td width="100%" colspan="3">
  <p align="center"><font size="2">Copyright © 2002, All rights reserved by
  <a href="../about.htm" style="color: white">K. 
  Haque</a></font></p>

				</td>
			</tr>
		</table>
	<%
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing %>
	</body>
</html>