Attribute VB_Name = "bas_SOAP"
Public Sub test_with_unsecure_server()

Dim sURL As String
Dim sEnv As String
Dim xmlhtp As New MSXML2.XMLHTTP
Dim xmlDoc As New DOMDocument
sURL = "http://www.webservicex.net/stockquote.asmx"

        sEnv = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:web=""http://www.webserviceX.NET/"">"
        sEnv = sEnv & "<soap:Header/>"
        sEnv = sEnv & "<soap:Body>"
        sEnv = sEnv & "<web:GetQuote>"
        sEnv = sEnv & "<web:symbol>IBM</web:symbol>"
        sEnv = sEnv & "</web:GetQuote>"
        sEnv = sEnv & "</soap:Body>"
        sEnv = sEnv & "</soap:Envelope>"

With xmlhtp
    
    .Open "post", sURL, False
    '.setRequestHeader "Host", "service.leads360.com"
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "soapAction", """http://www.webserviceX.NET/GetQuote"""
    '.setRequestHeader "Accept-encoding", "zip"
    .send sEnv
    xmlDoc.LoadXML .responseText
    
    MsgBox (.responseText)
    
    xmlDoc.Save ThisWorkbook.Path & "\test_soap_unsecure.xml"

End With




End Sub




Public Sub test_with_unsecure_server_serverxmlhttp()

Dim sURL As String
Dim sEnv As String
Dim xmlhtp As New MSXML2.ServerXMLHTTP
Dim xmlDoc As New DOMDocument
sURL = "http://www.webservicex.net/stockquote.asmx"

        sEnv = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:web=""http://www.webserviceX.NET/"">"
        sEnv = sEnv & "<soap:Header/>"
        sEnv = sEnv & "<soap:Body>"
        sEnv = sEnv & "<web:GetQuote>"
        sEnv = sEnv & "<web:symbol>IBM</web:symbol>"
        sEnv = sEnv & "</web:GetQuote>"
        sEnv = sEnv & "</soap:Body>"
        sEnv = sEnv & "</soap:Envelope>"

With xmlhtp
    
    .Open "post", sURL, False
    '.setRequestHeader "Host", "service.leads360.com"
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "soapAction", """http://www.webserviceX.NET/GetQuote"""
    '.setRequestHeader "Accept-encoding", "zip"
    .send sEnv
    xmlDoc.LoadXML .responseText
    
    MsgBox (.responseText)
    
    xmlDoc.Save ThisWorkbook.Path & "\test_soap_unsecure_server_xml.xml"

End With

End Sub



Public Sub test_with_Tmsg_bloomberg()

Dim sURL As String
Dim sEnv As String
Dim xmlhtp As New ServerXMLHTTP60

Dim xmlDoc As New DOMDocument
sURL = "https://bws.bloomberg.com/TmsgServiceSOAP"

'        sEnv = "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:web=""http://www.webserviceX.NET/"">"
'        sEnv = sEnv & "<soap:Header/>"
'        sEnv = sEnv & "<soap:Body>"
'        sEnv = sEnv & "<web:GetQuote>"
'        sEnv = sEnv & "<web:symbol>IBM</web:symbol>"
'        sEnv = sEnv & "</web:GetQuote>"
'        sEnv = sEnv & "</soap:Body>"
'        sEnv = sEnv & "</soap:Envelope>"



sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tmsg=""http://www.bloomberg.com/services/tmsg"">"
   sEnv = sEnv & "<soapenv:Header/>"
   sEnv = sEnv & "<soapenv:Body>"
      sEnv = sEnv & "<tmsg:TradeIdeaRead>"
         sEnv = sEnv & "<tmsg:Senders>"
            sEnv = sEnv & "<tmsg:Sender>"
               sEnv = sEnv & "<tmsg:Login>jstouff</tmsg:Login>"
            sEnv = sEnv & "</tmsg:Sender>"
         sEnv = sEnv & "</tmsg:Senders>"
      sEnv = sEnv & "</tmsg:TradeIdeaRead>"
   sEnv = sEnv & "</soapenv:Body>"
sEnv = sEnv & "</soapenv:Envelope>"



With xmlhtp
    
    '.Open "POST", sURL, False, "/1.3.6.1.4.1.1814.3.1.4=4485", "ct-84ESh"
    '.Open "POST", sURL, False, "D:\blp\data\tmsg-webservice-csharp\piccbws.p12", "ct-84ESh"
    .setRequestHeader "Accept-Encoding", "gzip,deflate"
    .setRequestHeader "Content-Type", "text/xml;charset=utf-8"
    .setRequestHeader "SOAPAction", """http://www.bloomberg.com/services/tmsg/TradeIdeaRead"""
    .setRequestHeader "Connection", "Keep-Alive"
    .setRequestHeader "Host", "bws.bloomberg.com"

    
    .send sEnv
    
    
    
    
    xmlDoc.LoadXML .responseText
    
    MsgBox (.responseText)
    
    xmlDoc.Save ThisWorkbook.Path & "\test_soap_secure_bbg_tmsg.xml"

End With

End Sub
