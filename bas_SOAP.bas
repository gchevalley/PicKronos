Attribute VB_Name = "bas_SOAP"


Public Function SOAP_standardize_query(ByVal soap_query As String, ByVal web_service_url As String, ByVal vec_http_request_header As Variant, Optional ByVal login As Variant, Optional ByVal password As Variant, Optional ByVal output_xml_file As Variant) As DOMDocument

Dim i As Integer, j As Integer, k As Integer

Dim xmlHttp As New xmlHttp
Dim secureXmlHttp As New ServerXMLHTTP60

Dim xmlDoc As New DOMDocument

If UCase(Left(web_service_url, Len("HTTPS"))) = "HTTPS" And IsMissing(login) = False And IsMissing(password) = False Then
    
    'secure connection
    With secureXmlHttp
        
        .Open "POST", web_service_url, False, login, password
        
        'inject header parameters
        For i = 0 To UBound(vec_http_request_header, 1)
            .setRequestHeader vec_http_request_header(i)(0), vec_http_request_header(i)(1)
        Next i
        
        .send soap_query
    
        xmlDoc.LoadXML .responseText
        
        
    End With
Else
    
    With xmlHttp
        
        .Open "POST", web_service_url, False
        
        'inject header parameters
        For i = 0 To UBound(vec_http_request_header, 1)
            .setRequestHeader vec_http_request_header(i)(0), vec_http_request_header(i)(1)
        Next i
        
        .send soap_query
    
        xmlDoc.LoadXML .responseText
        
    End With
    
End If


If IsMissing(output_xml_file) = False Then
    xmlDoc.Save output_xml_file
End If

Set SOAP_standardize_query = xmlDoc

End Function




