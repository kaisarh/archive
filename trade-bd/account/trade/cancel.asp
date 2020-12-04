<%
Dim sMsg, sYesUrl, sNoUrl

sYesUrl = "docancel.asp?" & request.querystring()
sNoUrl = "../account.asp"
sMsg = "Are you sure you want to cancel this item?"

%><!--#include file="../../ask.asp"-->