<%
Response.Expires = -1000

Dim sSql, sSql2

If IsObject(Session("db_conn")) Then
    Set oConn = Session("db_conn")
    Set oRs = Session("db_rs")
    Set oRs2 = Session("db_rs2")
	
	on error resume next
	oRS.Close
	oRS2.Close
	on error goto 0

Else
    Set oConn = Server.CreateObject("ADODB.Connection")
    Set oRs = Server.CreateObject("ADODB.Recordset")
    Set oRs2 = Server.CreateObject("ADODB.Recordset")

	'oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/trade-bd.mdb")
	'oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/data.mdb")
	oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=E:\Inetpub\wwwroot\trade-bd\db\data-2000.mdb"
    Set Session("db_conn") = oConn
    Set Session("db_rs") = oRs
    Set Session("db_rs2") = oRs2

End If

'<!--#include file="inc_db.asp"-->
%>