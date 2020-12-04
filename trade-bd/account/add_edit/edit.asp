<!--#include file='inc_acc_login.asp'-->

<%  		
  		Dim iDB, iCode
  		
		iCode = Request.Querystring("iCode")
		iDB = int(iCode/10)
		
		select case iDB
			case 1
				sUrl = "add_all.asp"
			case 2
				sUrl = "add_pc.asp"
			case 3
				sUrl = "add_cdsnbooks.asp"
		end select
		
		sUrl = sUrl & "?" & Request.QueryString()

		Response.Redirect sUrl
		Response.End
				
%>
