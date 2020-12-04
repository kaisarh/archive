		<%
			oRS.Open sSQL, oConn, 1
			iNTotal = oRS.RecordCount
		%>

		<center>
			<table border=0 width="640" cellspacing=1 cellpadding=0>
				<tr>
				    <td width="50%" colspan="3"><p align="left">Admininstrator's Page: <strong><% =sPageTitle %></strong></td>
			  	</tr>
			  	<tr>
			    	<td width="100%" colspan="3"><hr size="1"></td>
			    </tr>
   				<tr>
					<td width="100%" colspan="3" height="6"></td>
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
						<table border="1" width="100%" bordercolordark="#FFFFFF" bgcolor="#EAF8DE" cellspacing="0" cellpadding="0"
						    	bordercolorlight="#000000">
