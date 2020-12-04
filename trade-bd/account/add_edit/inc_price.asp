				<tr>
 		          <td width="100%" colspan="2" height="10"></td>
 		        </tr>
 		        <tr>
 		          <td>Used For:</td>
		          <td>
		          	<input type="text" name="Used" size="4">
		          	<select name="UsedUnit" size="1" tabindex="5">
		          		<option value="1">Days</option>
						<option value="2">Weeks</option>    
						<option value="3">Months</option>
						<option value="4">Years</option>
					</select>
				  </td>
	        	</tr>
		        <tr>
		          <td>Condition:</td>
		          <td>
			          <select name="Condition" size="1">
			            <option value="1">100% New</option>
			            <option value="2">Almost New</option>
			            <option selected value="3">Not So Old</option>
			            <option value="4">Old</option>
			            <option value="5">Very Old</option>
			          </select>
		          </td>
		        </tr>
		        <tr>
		          <td width="100%" colspan="2" height="10"></td>
		        </tr>
		        <tr>
		          <td>Price:</td>
		          <td>
			      	<input type="text" name="Price" size="5">
			        <select name="PriceUnit" size="1">
			        	<option value="1">Taka</option>
			            <option value="2">US$</option>
			        </select>
			      </td>
		        </tr>
		        <tr>
		          <td width="100%" colspan="2" height="10"></td>
		        </tr>
		        
		        <tr>
		          <td>
		          	<input type="button" value="&lt;&lt; Back" name="b3" onClick="javascript:history.back()">
		          </td>
		          <td>
			          <input type="reset" value="Reset" name="b2">&nbsp; 
			          <input type="button" value=
			          	<% if bLoad then
			          		response.write(chr(34) & "Update &gt;&gt;" & chr(34))
			          	else
			          		response.write(chr(34) & "Submit &gt;&gt;" & chr(34))
			          	end if %>
			          	name="b1" onClick="javascript:check(this.form)">
		          </td>
		        </tr>
   				<%  if bLoad then
						oRS.Close
						'oConn.Close
						'Set oRS = Nothing
						'Set oConn = Nothing
					end if
				%>
