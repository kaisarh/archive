<%
Dim sMsg, sYesUrl, sNoUrl

sYesUrl = "dodel.asp?" & request.querystring()
sNoUrl = "../account.asp"
sMsg = "Are you sure you want to delete this item?"

%><!--#include file="../../ask.asp"-->