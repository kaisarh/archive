<%
if not Session("bLogin") then
	Session("sLogin_Redirect")="../account.asp"
	Session("sLogin_Reason")="You are not logged in!"
	Response.Redirect "login.asp?bRedirect=yes"
end if
%>
<!--#include file="../../inc_db.asp"-->
<%
		Function getv2(x)
			dim s
			s = request.form(x)
			if len(s)=0 then s=" "
			getv2 = "'" & s & "'"
		End Function
		
		Function geti2(x)
			dim s
			s = request.form(x)
			if len(s)=0 then s=0
			geti2 = s
		End Function

		Dim i, itemCount, bError
		Dim iItemID, iCode, iDB, iUserID
		Dim sMsg, sUrl
	
		iCode = Request.querystring("iCode")
		iItemID = Request.querystring("iItemID")
		iUserID = Session("iUserID")
		bError = true

		sMsg = "Database was updated successfully!"
		sURL = "../account.asp"

		''we check for all error using javascript
		if iCode = 40 then
			sSql = "Update Users Set Name="  & getv2("Name") & _
					",Add1=" & getv2("Add1") & _
					",Add2=" & getv2("Add2") & _
					",Phone=" & getv2("Phone") & _
					",Email=" & getv2("Email") & _
					" WHERE UserID=" & iUserID
			sStr = ""

		elseif iCode = 20 then
			sSql = "Update PCs Set Brand="  & getv2("Brand") & _
					",Processor=" & getv2("Processor") & _
					",Speed=" & geti2("Speed") & _
					",RAM=" & geti2("RAM") & _
					",HDD=" & geti2("HDD") & _
					",MB=" & getv2("MB") & _
					",Display='" & geti2("DisT") & " " & geti2("DisS") & "'" & _
					",Sound=" & getv2("SoundC") & _
					",SoundBit=" & geti2("SoundT") & _
					",Modem=" & getv2("ModemC") & _
					",ModemSpeed=" & geti2("ModemS") & _
					",Monitor=" & getv2("MonitorC") & _
					",MonitorSize=" & geti2("MonitorS") & _
					" WHERE ItemID=" & iItemID
			sStr = "PCs"

		elseif iCode>30 AND iCode<40 then
			sSql = "Update CDsNBooks Set Author="  & getv2("Author") & _
					",Title=" & getv2("Title") & _
					",CopyType=0" & _
					" WHERE ItemID=" & iItemID
			sStr = "CDsNBooks"

		else	
			sSql = "Update Alls Set Name="  & getv2("s" & iCode) & _
					",Brand=" & getv2("Brand") & _
					",Spec=" & getv2("Spec") & _
					",Acc=" & getv2("Acc") & _
					" WHERE ItemID=" & iItemID
			sStr = "Alls"

		end if

		Set oRS = oConn.Execute(sSql)

		if sStr <> "" then
			sSql = "Update Items Set Used=" & geti2("Used") & _
					",UsedUnit=" & geti2("UsedUnit") & _
					",Condition=" & geti2("Condition") & _
					",Price=" & geti2("Price") & _
					",PriceUnit=" & geti2("PriceUnit") & _
					" WHERE ItemID=" & iItemID & " AND iCode=" & iCode
			''response.write(sSql)
			''response.end
			Set oRS = oConn.Execute(sSql)
		end if

		bError = false

		'oConn.Close
		'Set oRS = Nothing
		'Set oConn = Nothing
	%><!--#include file="..\..\info.asp"-->