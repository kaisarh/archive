<%  		
  		Dim iUserID, bLoad, sUrl
  		Dim iItemID, iCode
  		
		iItemID = Request.querystring("iItemID")
		iCode = Request.Querystring("iCode")
		
  		bLoad = false
		sUrl = "add.asp?"
		
		if iItemID <> "" then
			iUserID = Session("iUserID")

			bLoad = true
			sUrl = "update.asp?iItemID=" & iItemID & "&"
%>
			<!--#include file="../../inc_db.asp"-->
<%					
		end if
%>