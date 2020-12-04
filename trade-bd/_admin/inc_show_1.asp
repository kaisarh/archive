<!--#include file='inc_admin_login.asp'-->

<html>
	<head>
		<title>
		</title>
	</head>
	<body>

		<%
		
		Dim oConn, oRS, sSQL
		Dim iNTotal, iNStart, iNEnd, iNCount, iPg
		Dim bFound, sPageURL, sPageTitle, sSrc, str
		
		''sPage is the page name, like "show_items.asp"

		sub page_link(s, pg_num)
			str = "<a href=" & chr(34) & sPageURL & "?pg=" & pg_num
			if sSrc <> "" then str = str & "&src=" & sSrc
			str = str & chr (34) & ">" & s & "</a>"

			response.write(str)
		
		end sub

		Response.Expires = -1000
		const PAGE_SIZE = 20
		
		sSrc = Request.form("Search")
		if sSrc = "" then sSrc = Request.querystring("src")
				
		''let's create our objects	
		Set oConn = Server.CreateObject("ADODB.Connection")
		Set oRS = Server.CreateObject( "ADODB.Recordset" )
		oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/data.mdb")
		
		%>
			