<!DOCTYPE html>
<html>
<body>

<%
Dim httpRequest, httpResponse
  
Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
httpRequest.Open "POST", "https://node-fake-api-server.herokuapp.com/", False
httpRequest.SetRequestHeader "Content-Type", "application/json"
httpRequest.SetRequestHeader "X-FakeAPI-Action", "register"
httpRequest.Send
httpResponse = httpRequest.ResponseText

data=split(httpRequest.ResponseText,"apikey",-1,1)
    if ubound(data)>0 then
        kvpair=split(data(1),chr(34),-1,1)
		Response.Write ("<br/>" & kvpair(2))
%>

</body>
</html>