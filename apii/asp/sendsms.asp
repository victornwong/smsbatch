<%
Dim objWinHttp
Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
objWinHttp.Open "GET", "http://sample.onewaysms.com.au:xxxx/api.aspx?apiusername=123&apipassword=xyz&mobileno=61412345678&senderid=onewaysms&languagetype=1&message=test sms from api", false
objWinHttp.setrequestHeader "Content-Type", "application/x-www-formurlencoded"
objWinHttp.Send
Response.Write objWinHttp.ResponseText
Response.End
%>