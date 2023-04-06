<%
Set objData = CreateDOM
objData.async = false

if (false) then
	Set objDataXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
	objDataXMLHTTP.open "GET", "", false 
	objDataXMLHTTP.setRequestHeader "Content-Type", "text/xml"
	objDataXMLHTTP.send
	objData.load(objDataXMLHTTP.responseBody)
else
	objData.load(Server.MapPath("Customers_Server.xml"))
end if

Set objStyle = CreateDOM
objStyle.async = false
objStyle.load(Server.MapPath("Customers_Server.xsl"))
Session.CodePage = 65001

Response.ContentType = "text/html"
'correction statement on the next line take care of the browser error
'objStyle.setProperty "AllowXsltScript", true
Response.Write objData.transformNode(objStyle)

Function CreateDOM()
	On Error Resume Next
	Dim tmpDOM

	Set tmpDOM = Nothing
	Set tmpDOM = Server.CreateObject("MSXML2.DOMDocument.6.0")
	
	Set CreateDOM = tmpDOM
End Function

%>
