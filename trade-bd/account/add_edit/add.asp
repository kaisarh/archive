<%
if Request.Querystring("iCode")<>40 and not Session("bLogin") then
	Session("sLogin_Redirect")="../account.asp"
	Session("sLogin_Reason")="You are not logged in! You need to login first in order to add an item to our database!"
	Response.Redirect "../login.asp?bRedirect=yes"
end if

Function getv(x)
	getv = getv2(x) & ", "
End Function

Function getv2(x)
	dim s
	s = request.form(x)
	if len(s)=0 then s=" "
	getv2 = "'" & s & "'"

End Function

Function geti(x)
	geti = geti2(x) & ", "
End Function

Function geti2(x)
	dim s
	s = request.form(x)
	if len(s)=0 then s=0
	geti2 = s
End Function


function get_count(x)
	''assuming oRS were closed
	sSQL = "SELECT * FROM Count"
	
	oRS.Open sSQL, oConn, 1
	get_count = oRS(x).value+1
	oRS.CLose
	
	sSQL = "UPDATE Count SET " & x & "=" & get_count & ""
	Set oRS = oConn.Execute(sSQL)
	''oRS now also closed
end function
%>
<!--#include file="../../inc_db.asp"-->

<%		Dim i, itemCount, bExists, bError
		Dim itemID, iCode, iDB, iUserID
		
		Dim sMsg, sUrl
		
		bExists = false
		bError = true
		
		iCode = Request.querystring("iCode")


		''we check for all error using javascript
		if iCode = 40 then
			''check for existing login
			sSQL = "SELECT * FROM Users WHERE Login=" & getv2("Login")
			oRS.Open sSQL, oConn, 1

			if NOT oRS.EOF then 
				bExists = true
			end if
			oRS.Close
			
			sUrl = "javascript:history.back()"
			
			if bExists then
				sMsg = "Login ID already exists, please choose another one."
			else
				''check for existing email address...
				sSQL = "SELECT * FROM Users WHERE Email=" & getv2("Email")
				oRS.Open sSQL, oConn, 1
	
				if NOT oRS.EOF then 
					bExists = true
				end if
				
				oRS.Close
				
				if bExists then
					sMsg = "You already have an account with this email address!"
				else
					itemID = get_count("SNUser")

					sSQL = "INSERT INTO Users(UserID, Name, Add1, Add2, Phone, Email, Login, Pass) VALUES(" & _
						itemID & ", " & getv("Name") & getv("Add1") & getv("Add2") & getv("Phone") & _
						getv("Email") & getv("Login") & getv2("Pass1") & ")"

					''goto account page
					Set oRS = oConn.Execute(sSQL)

					Session("bLogin")=true
					Session("iUserID")=itemID
					Session("iTry") = 0

					sMsg = "User Signup Successful, You'll be taken to your account page now!"
					sUrl = "../account.asp"
					bError = false
				end if
			end if
		else
			iUserID = Session("iUserID")

			''update user's sells record
			sSQL = "SELECT * FROM Users WHERE UserID=" & iUserID

			oRS.Open sSQL, oConn, 1

			itemCount = -1
			if not oRS.EOF then itemCount = oRS("SellCount")
			oRS.Close

			if itemCount=-1 then
				sMsg = "Unexpected Error: User not found!"
			elseif itemCount=10 then
				sMsg = "Sorry, you already have added the max number of items!"
			else
				if iCode = 20 then
						itemID = get_count("SNPC")

						sSQL = "INSERT INTO PCs(ItemID, iCode, Brand, Processor, Speed, RAM, HDD, MB, Display, " & _
							"Sound, SoundBit, Modem, ModemSpeed, Monitor, MonitorSize) VALUES(" & _
							itemID & ", 20, " & getv("Brand") & getv("Processor") & geti("Speed") & _
							geti("RAM") & geti("HDD") & getv("MB") & _
							"'" & geti2("DisT") & " " & geti2("DisS") & "', " & _
							getv("SoundC") & geti("SoundT") & getv("ModemC") & geti("ModemS") & _
							getv("MonitorC") & geti2("MonitorS") & ")"
	
				elseif iCode>30 AND iCode<40 then
						itemID = get_count("SNCD")
	
						sSQL = "INSERT INTO CDsNBooks(ItemID, iCode, Author, Title, CopyType) VALUES(" & _
							itemID & ", " & iCode & ", " & getv("Author") & getv("Title") & "0)"
	
				else	
						itemID = get_count("SNAll")
	
						sSQL = "INSERT INTO Alls(ItemID, iCode, Name, Brand, Spec, Acc) VALUES(" & _
							itemID & ", " &  iCode & ", " & _
							getv("s" & iCode) & getv("Brand") & getv("Spec") & getv2("Acc") & ")"
				end if
				Set oRS = oConn.Execute(sSQL)

				''user not found already handled
				sSQL = "INSERT INTO Items(ItemID, iCode, SellerID, Used, UsedUnit, Condition, Price, PriceUnit, InsertedDate) VALUES(" & _
					itemID & ", " & iCode & ", " & iUserID & ", " & geti("Used") & geti("UsedUnit")& geti("Condition") & geti("Price") & geti("PriceUnit") & _
					"'" & FormatDateTime(Now,2) & "')"
				Set oRS = oConn.Execute(sSQL)

				sSQL = "UPDATE Users SET SellCount=" & itemCount+1 & " WHERE UserID=" & iUserID
				Set oRS = oConn.Execute(sSQL)

				sMsg = "Your item was added into our database successfully!"
				bError = false
			end if
			sURL = "../account.asp"
		end if

if not bError then
	sSQL = "SELECT * FROM Count"
	oRS.Open sSQL, oConn, 1

	sSQL = "UPDATE Count SET iCode" & iCode & "='" & oRS("iCode" & iCode).value+1 & "'"
	oRS.Close
	Set oRS = oConn.Execute(sSQL)
end if

'oConn.Close
'Set oRS = Nothing
'Set oConn = Nothing
%><!--#include file="../../info.asp"-->