<%
Dim sMsg, sYesUrl, sNoUrl

sYesUrl = "doconfirm.asp?" & request.querystring()
sNoUrl = "../account.asp"
sMsg = "Are you sure this item was sold?"

%><!--#include file="../../ask.asp"-->