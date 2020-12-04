<html>
	<head>
		<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
		<title>Sign Up</title>
	</head>
	<body>
		<%
			Dim sMsg, sUrl
			if Session("bAdmin") then
				Session("bLogin")=false
				Session("iUserID")=0
				Session("bAdmin")=false
				sMsg = "You have been logged out from Administrator account!"
			else
				sMsg = "You are already not logged in as Administrator!"
			end if
		%>
		<center>
		<br>
		<strong><% =sMsg %></strong>
		<br><br>
		<a href="..\main.asp">Main Page</a>
		
		</center>
	</body>
</html>
