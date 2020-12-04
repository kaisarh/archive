<%
Dim sMsg, sYesUrl, sNoUrl
sYesUrl = "doreserve.asp?" & request.querystring()
sNoUrl = "javascript:history.back();"
sMsg = "Are you sure you want to reserve and buy this item?"

%><!--#include file="../../ask.asp"-->