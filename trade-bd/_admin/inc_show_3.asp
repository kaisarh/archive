							<% oRS.MoveNext
							iNCount = iNCount + 1
							wend %>
						</table>
					</td>
				</tr>
				<tr align="middle">
					<td width="100%" colspan="3">
						<br><strong>Search:</strong> Search the Data Base<p>
						<form method="POST" action="<%=sPageURL%>">
				  				<p>
				  					<input type="text" name="Search" size="34"> &nbsp;
				  					<input type="submit" value="Search" name="b1">
				  				</p>
						</form>
					</td>
				</tr>
		<% end if %>
		<tr align="middle">
			<td width="100%" colspan="3">
				Copy Right 2002, All rights reserved by MyMarket.Com.
			</td>
		</tr>
		</table>
		</center>
		<% oRS.Close
		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing %>
	</body>
</html>
