Attribute VB_Name = "bas_bws_TMSG"
Private Const bws_TMSG_sheet As String = "bws_tmsg"
    Private Const c_bws_TMSG_id As Integer = 1
    Private Const c_bws_TMSG_status As Integer = 2
    Private Const c_bws_TMSG_open_datetime As Integer = 3
    Private Const c_bws_TMSG_ticker As Integer = 4
    Private Const c_bws_TMSG_side As Integer = 5
    Private Const c_bws_TMSG_shares As Integer = 6
    Private Const c_bws_TMSG_cost As Integer = 7
    Private Const c_bws_TMSG_last_price As Integer = 8
    Private Const c_bws_TMSG_target_price As Integer = 9
    Private Const c_bws_TMSG_sender As Integer = 10
    Private Const c_bws_TMSG_close_datetime As Integer = 11


Private Const bws_TMSG_web_service_url As String = "https://bws.bloomberg.com/TmsgServiceSOAP"
Private Const bws_TMSG_web_service_https_login As String = "/1.3.6.1.4.1.1814.3.1.4=4485"
Private Const bws_TMSG_web_service_https_password As String = "ct-84ESh"

Private Const bws_TMSG_web_service_limit_record As Integer = 100



Public Function SOAP_bws_tmsg_READ(ByVal soap_query As String) As DOMDocument

Dim xmlDoc As DOMDocument, xmlDocTmpPartial As DOMDocument

Set xmlDoc = SOAP_bws_tmsg(soap_query, "http://www.bloomberg.com/services/tmsg/TradeIdeaRead")

Dim oRoot As IXMLDOMElement, oRootTmpPartial As IXMLDOMElement
Set oRoot = xmlDoc.DocumentElement

Dim oTotalNumberOfRecords As IXMLDOMElement
Set oTotalNumberOfRecords = oRoot.getElementsByTagName("TotalNumberOfRecords")(0)

If CDbl(oTotalNumberOfRecords.Text) > bws_TMSG_web_service_limit_record Then
    
    Dim oMainTradeIdeasList As IXMLDOMElement
    Set oMainTradeIdeasList = oRoot.getElementsByTagName("TradeIdeasList")(0)
    
    Dim oMainEndRecordNumber As IXMLDOMElement
    Set oMainEndRecordNumber = oRoot.getElementsByTagName("EndRecordNumber")(0)
    
    Dim oTmpTradeIdeaFromPartial As IXMLDOMElement
    
    
    'tranformation de la query text en xml
    Dim xmlDocInputSOAP As New DOMDocument, oRootInputSOAP As IXMLDOMElement
    xmlDocInputSOAP.LoadXML soap_query
    
    Set oRootInputSOAP = xmlDocInputSOAP.DocumentElement
    
    Dim oTmsgTradeIdeaRead As IXMLDOMElement
    Set oTmsgTradeIdeaRead = oRootInputSOAP.getElementsByTagName("tmsg:TradeIdeaRead")(0)
    
    Dim oTmsgRecoredRange As IXMLDOMElement
    Dim oTmsgStartRecordNumber As IXMLDOMElement
    Dim oTmsgEndRecordNumber As IXMLDOMElement
    
    'check si tmsg:RecordRange existe deja dans la structure
    Set oTmsgRecoredRange = oTmsgTradeIdeaRead.getElementsByTagName("tmsg:RecordRange")(0)
    
    If oTmsgRecoredRange Is Nothing Then
        
        'append la zone tmsg:RecordRange
        Set oTmsgRecoredRange = xmlDocInputSOAP.createElement("tmsg:RecordRange")
        oTmsgTradeIdeaRead.appendChild oTmsgRecoredRange
        
        
        Set oTmsgStartRecordNumber = xmlDocInputSOAP.createElement("tmsg:StartRecordNumber")
            oTmsgStartRecordNumber.Text = "?"
        
        Set oTmsgEndRecordNumber = xmlDocInputSOAP.createElement("tmsg:EndRecordNumber")
            oTmsgEndRecordNumber.Text = "?"
            
            oTmsgRecoredRange.appendChild oTmsgStartRecordNumber
            oTmsgRecoredRange.appendChild oTmsgEndRecordNumber
    
    Else
        
        Set oTmsgStartRecordNumber = oTmsgRecoredRange.getElementsByTagName("tmsg:StartRecordNumber")(0)
        Set oTmsgEndRecordNumber = oTmsgRecoredRange.getElementsByTagName("tmsg:EndRecordNumber")(0)
    
    End If
    
    
    Dim nbre_recall As Integer
    nbre_recall = Int(CDbl(oTotalNumberOfRecords.Text) / bws_TMSG_web_service_limit_record)
    
    For i = 1 To nbre_recall
        
        Dim start_record As Integer, end_record As Integer
        start_record = (i * bws_TMSG_web_service_limit_record) + 1
        end_record = start_record + (bws_TMSG_web_service_limit_record - 1)
        
        oTmsgStartRecordNumber.Text = start_record
        oTmsgEndRecordNumber.Text = end_record
        
        Set xmlDocTmpPartial = SOAP_bws_tmsg(CStr(xmlDocInputSOAP.XML), "http://www.bloomberg.com/services/tmsg/TradeIdeaRead")
        
        xmlDocTmpPartial.Save ThisWorkbook.Path & "\partial_intermed_soap_tmsg_read.xml"
        
        Set oRootTmpPartial = xmlDocTmpPartial.DocumentElement
        
        'append sur la structure principal
        For Each oTmpTradeIdeaFromPartial In oRootTmpPartial.getElementsByTagName("TradeIdea")
            oMainTradeIdeasList.appendChild oTmpTradeIdeaFromPartial
        Next
        
        'ajustement du nbre de record
        oMainEndRecordNumber.Text = oRootTmpPartial.getElementsByTagName("EndRecordNumber")(0).Text
        
        
        'tmp save to check content
        'xmlDoc.Save ThisWorkbook.Path & "\partial_soap_tmsg_read.xml"
    Next i
    
End If

Set SOAP_bws_tmsg_READ = xmlDoc

End Function


Private Function SOAP_bws_tmsg(ByVal soap_query As String, ByVal soap_action As String) As DOMDocument

Dim vec_header() As Variant

vec_header = Array(Array("Accept-Encoding", "gzip,deflate"), Array("Content-Type", "text/xml;charset=utf-8"), Array("Connection", "Keep-Alive"), Array("Host", "bws.bloomberg.com"))

    ReDim Preserve vec_header(UBound(vec_header, 1) + 1)
    vec_header(UBound(vec_header, 1)) = Array("SOAPAction", """" & soap_action & """")
    
    Set SOAP_bws_tmsg = SOAP_standardize_query(soap_query, bws_TMSG_web_service_url, vec_header, bws_TMSG_web_service_https_login, bws_TMSG_web_service_https_password)
    
End Function



Public Sub Tmsg_bloomberg_in_excel()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim sEnv As String
Dim xmlDoc As New DOMDocument


sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tmsg=""http://www.bloomberg.com/services/tmsg"">"
   sEnv = sEnv & "<soapenv:Header/>"
   sEnv = sEnv & "<soapenv:Body>"
      sEnv = sEnv & "<tmsg:TradeIdeaRead>"
      
'      sEnv = sEnv & "<tmsg:TimeRange>"
'            sEnv = sEnv & "<tmsg:StartDate>2012-11-15T00:00:00</tmsg:StartDate>"
'            sEnv = sEnv & "<tmsg:EndDate>2012-11-21T00:00:00</tmsg:EndDate>"
'         sEnv = sEnv & "</tmsg:TimeRange>"
         
         sEnv = sEnv & "<tmsg:Receivers>"
            sEnv = sEnv & "<tmsg:Receiver>"
               sEnv = sEnv & "<tmsg:PortfolioIdentifier>"
                  sEnv = sEnv & "<tmsg:PortfolioName>MS/Pictet Ideas</tmsg:PortfolioName>"
                  sEnv = sEnv & "<tmsg:FirmId>3063</tmsg:FirmId>"
               sEnv = sEnv & "</tmsg:PortfolioIdentifier>"
            sEnv = sEnv & "</tmsg:Receiver>"
         sEnv = sEnv & "</tmsg:Receivers>"
         
      sEnv = sEnv & "</tmsg:TradeIdeaRead>"
   sEnv = sEnv & "</soapenv:Body>"
sEnv = sEnv & "</soapenv:Envelope>"


Set xmlDoc = SOAP_bws_tmsg_READ(sEnv)
Dim oRoot As IXMLDOMElement
Set oRoot = xmlDoc.DocumentElement

' parsing
Dim oTradeIdea As IXMLDOMElement
    Dim oIdeaId As IXMLDOMElement
        Dim oIdeaIdBloombergId As IXMLDOMElement
    Dim oReceiver As IXMLDOMElement
        Dim oReceiverPortfolioIdentifier As IXMLDOMElement
            Dim oReceiverPortfolioIdentifierPortfolioName As IXMLDOMElement
            Dim oReceiverPortfolioIdentifierFirmId As IXMLDOMElement
    Dim oStatus As IXMLDOMElement
    Dim oDirection As IXMLDOMElement
    Dim oConviction As IXMLDOMElement
    Dim oTargetPrice As IXMLDOMElement
    Dim oSender As IXMLDOMElement
        Dim oSenderLogin As IXMLDOMElement
        Dim oSenderSenderName As IXMLDOMElement
        Dim oSenderFirmName As IXMLDOMElement
    Dim oInstrument As IXMLDOMElement
        Dim oInstrumentSecurity As IXMLDOMElement
            Dim oInstrumentSecurityIdentifier As IXMLDOMElement
                Dim oInstrumentSecurityIdentifierBloombergId As IXMLDOMElement
                Dim oInstrumentSecurityIdentifierParseKey As IXMLDOMElement
                Dim oInstrumentSecurityIdentifierCompagnyName As IXMLDOMElement
                Dim oInstrumentSecurityIdentifierAssetClass As IXMLDOMElement
    Dim oOpenTimestamp As IXMLDOMElement
    Dim oCloseTimestamp As IXMLDOMElement
    Dim oMsgSubject As IXMLDOMElement
    Dim oMsgBody As IXMLDOMElement
        Dim oMsgBodyBodyContent As IXMLDOMElement
    Dim oInvestment As IXMLDOMElement
    
                
        

Worksheets(bws_TMSG_sheet).Cells.Clear


'header



k = 2
For Each oTradeIdea In oRoot.getElementsByTagName("TradeIdea")
    
    Set oIdeaId = oTradeIdea.getElementsByTagName("IdeaId")(0)
        Set oBloombergId = oIdeaId.getElementsByTagName("BloombergId")(0)
    Set oReceiver = oTradeIdea.getElementsByTagName("Receiver")(0)
        Set oReceiverPortfolioIdentifier = oReceiver.getElementsByTagName("PortfolioIdentifier")(0)
            Set oReceiverPortfolioIdentifierPortfolioName = oReceiverPortfolioIdentifier.getElementsByTagName("PortfolioName")(0)
            Set oReceiverPortfolioIdentifierFirmId = oReceiverPortfolioIdentifier.getElementsByTagName("FirmId")(0)
    Set oStatus = oTradeIdea.getElementsByTagName("Status")(0)
    Set oDirection = oTradeIdea.getElementsByTagName("Direction")(0)
    Set oConviction = oTradeIdea.getElementsByTagName("Conviction")(0)
    Set oTargetPrice = oTradeIdea.getElementsByTagName("TargetPrice")(0) 'optional
    Set oSender = oTradeIdea.getElementsByTagName("Sender")(0)
        Set oSenderLogin = oSender.getElementsByTagName("Login")(0)
        Set oSenderSenderName = oSender.getElementsByTagName("SenderName")(0)
        Set oSenderFirmName = oSender.getElementsByTagName("FirmName")(0)
    Set oInstrument = oTradeIdea.getElementsByTagName("Instrument")(0)
        Set oInstrumentSecurity = oInstrument.getElementsByTagName("Security")(0)
            Set oInstrumentSecurityIdentifier = oInstrumentSecurity.getElementsByTagName("Identifier")(0)
                Set oInstrumentSecurityIdentifierBloombergId = oInstrumentSecurityIdentifier.getElementsByTagName("BloombergId")(0)
                Set oInstrumentSecurityIdentifierParseKey = oInstrumentSecurityIdentifier.getElementsByTagName("ParseKey")(0)
                Set oInstrumentSecurityIdentifierCompagnyName = oInstrumentSecurityIdentifier.getElementsByTagName("CompagnyName")(0)
                Set oInstrumentSecurityIdentifierAssetClass = oInstrumentSecurityIdentifier.getElementsByTagName("AssetClass")(0)
    Set oOpenTimestamp = oTradeIdea.getElementsByTagName("OpenTimestamp")(0)
    Set oCloseTimestamp = oTradeIdea.getElementsByTagName("CloseTimestamp")(0)
    Set oMsgSubject = oTradeIdea.getElementsByTagName("MsgSubject")(0)
    Set oMsgBody = oTradeIdea.getElementsByTagName("MsgBody")(0)
        Set oMsgBodyBodyContent = oMsgBody.getElementsByTagName("BodyContent")(0)
    Set oInvestment = oTradeIdea.getElementsByTagName("Investment")(0)
    
    
    'impression du tableau
    If oBloombergId Is Nothing Then
    Else
        Worksheets(bws_TMSG_sheet).Cells(k, c_bws_TMSG_id) = CDbl(oBloombergId.Text)
    End If
    
    If oStatus Is Nothing Then
    Else
        Worksheets(bws_TMSG_sheet).Cells(k, c_bws_TMSG_status) = oStatus.Text
    End If
    
    Dim datetime_open As Date
    If oOpenTimestamp Is Nothing Then
    Else
        date_txt = oOpenTimestamp.Text
        
            date_txt_year = CDbl(Left(date_txt, 4))
            date_txt_month = CDbl(Mid(date_txt, 6, 2))
            date_txt_day = CDbl(Mid(date_txt, 9, 2))
            date_txt_hour = CDbl(Mid(date_txt, 12, 2))
            date_txt_minute = CDbl(Mid(date_txt, 15, 2))
            date_txt_second = CDbl(Mid(date_txt, 18, 2))
        
        datetime_open = DateSerial(date_txt_year, date_txt_month, date_txt_day) + TimeSerial(date_txt_hour, date_txt_minute, date_txt_second)
        
        Worksheets(bws_TMSG_sheet).Cells(k, c_bws_TMSG_open_datetime) = datetime_open
        
    End If
    
    
    If oInstrumentSecurityIdentifierParseKey Is Nothing Then
    Else
        Worksheets(bws_TMSG_sheet).Cells(k, c_bws_TMSG_ticker) = oInstrumentSecurityIdentifierParseKey.Text
    End If
    
    
    If oDirection Is Nothing Then
    Else
        Worksheets(bws_TMSG_sheet).Cells(k, c_bws_TMSG_side) = oDirection.Text
    End If
    
    
    If oTargetPrice Is Nothing Then
    Else
        If IsNumeric(oTargetPrice.Text) Then
            Worksheets(bws_TMSG_sheet).Cells(k, c_bws_TMSG_target_price) = CDbl(oTargetPrice.Text)
        End If
    End If
    
    
    If oSenderFirmName Is Nothing And oSenderSenderName Is Nothing Then
    Else
        Worksheets(bws_TMSG_sheet).Cells(k, c_bws_TMSG_sender) = oSenderFirmName.Text & " / " & oSenderSenderName.Text
    End If
    
    
    
    
    k = k + 1
    
Next
    

End Sub
