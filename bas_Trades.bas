Attribute VB_Name = "bas_Trades"
Private Const key_r_plus_user As String = "user"
Private Const key_r_plus_password As String = "password"
Private Const key_r_plus_cash_account As String = "acct_cash"
Private Const key_r_plus_derivatives_account As String = "acct_deriv"

Public Const dim_vec_ticker_details_ticker As Integer = 0
Public Const dim_vec_ticker_details_crncy As Integer = 1
Public Const dim_vec_ticker_details_line As Integer = 2
Public Const dim_vec_ticker_details_delta As Integer = 3
Public Const dim_vec_ticker_details_region As Integer = 4
Public Const dim_vec_ticker_details_theta As Integer = 5
Public Const dim_vec_ticker_details_net_pos As Integer = 6
Public Const dim_vec_ticker_details_src As Integer = 7
Public Const dim_vec_ticker_details_broker As Integer = 8

Public Const dim_vec_trade_ticker As Integer = 0
Public Const dim_vec_trade_qty As Integer = 1
Public Const dim_vec_trade_price As Integer = 2
Public Const dim_vec_trade_crncy As Integer = 3 'txt
Public Const dim_vec_trade_region As Integer = 4
Public Const dim_vec_trade_line As Integer = 5
Public Const dim_vec_trade_s3 As Integer = 6
Public Const dim_vec_trade_s2 As Integer = 7
Public Const dim_vec_trade_s1 As Integer = 8
Public Const dim_vec_trade_p As Integer = 9
Public Const dim_vec_trade_r1 As Integer = 10
Public Const dim_vec_trade_r2 As Integer = 11
Public Const dim_vec_trade_r3 As Integer = 12
Public Const dim_vec_trade_last_price As Integer = 13
Public Const dim_vec_trade_position_in_equity_db As Integer = 14
Public Const dim_vec_trade_source As Integer = 15
Public Const dim_vec_trade_order_type As Integer = 16
Public Const dim_vec_trade_broker As Integer = 17

Public Const dim_vec_change_rate_txt  As Integer = 0
Public Const dim_vec_change_rate_code As Integer = 1
Public Const dim_vec_change_rate_rate As Integer = 2

Public Const dim_vec_simple_trade_ticker As Integer = 0
Public Const dim_vec_simple_trade_qty As Integer = 1
Public Const dim_vec_simple_trade_price As Integer = 2
Public Const dim_vec_simple_trade_order_type As Integer = 3
Public Const dim_vec_simple_trade_src As Integer = 4


Public Const l_format2_header As Integer = 99

Public Const c_format2_ticker  As Integer = 1
Public Const c_format2_aim_account As Integer = 2
Public Const c_format2_strategy  As Integer = 3
Public Const c_format2_limit_type  As Integer = 4
Public Const c_format2_side  As Integer = 5
Public Const c_format2_qty  As Integer = 6
Public Const c_format2_price  As Integer = 7
Public Const c_format2_time_limit  As Integer = 8
Public Const c_format2_broker  As Integer = 9
Public Const c_format2_last_price  As Integer = 10

Public Const c_format2_s3  As Integer = 11
Public Const c_format2_s2  As Integer = 12
Public Const c_format2_s1  As Integer = 13
Public Const c_format2_p  As Integer = 14
Public Const c_format2_r1  As Integer = 15
Public Const c_format2_r2  As Integer = 16
Public Const c_format2_r3  As Integer = 17

Public Const c_format2_pre_market_start_column  As Integer = 18

Public Const c_format2_valeur_eur  As Integer = 20
Public Const c_format2_delta  As Integer = 21
Public Const c_format2_theta  As Integer = 22
Public Const c_format2_eps  As Integer = 23
Public Const c_format2_perso_rel_1d  As Integer = 24

Public Const c_format2_dmi  As Integer = 25
Public Const c_format2_pct_bol  As Integer = 26

Public Const c_format2_source  As Integer = 27

Public Const l_format2_get_out_pct_last_price  As Integer = 86
Public Const c_format2_get_out_pct_last_price  As Integer = 21


Public Const c_format2_eqs_name As Integer = 105
Public Const c_format2_eqs_type As Integer = 106
Public Const c_format2_eqs_folder As Integer = 107
Public Const c_format2_eqs_order_type As Integer = 108
Public Const c_format2_eqs_insert_username As Integer = 109
Public Const c_format2_eqs_stop_formula As Integer = 110
Public Const c_format2_eqs_target_formula As Integer = 111
Public Const c_format2_eqs_custom_formula_start As Integer = 112

Public check_formula_syntax_vec_equity_db_header As Variant
Public check_formula_syntax_vec_equity_db_securities As Variant


Private Const form_height_without_custom_area As Integer = 387
Private Const form_height_with_custom_area As Integer = 500


Public Enum ElocateType
    REQUEST = 0
    Notify = 1
    query = 2
End Enum


Public Sub deploy_update_bas_trades()

Dim i As Integer, j As Integer, k As Integer

'mise en place des champs customfield1 pour format2
Dim blub As OLEObject


'checkbox qui control si le titre est dans ces heures traitable
Set blub = Worksheets("FORMAT2").OLEObjects.Add(ClassType:="Forms.CheckBox.1", link:=False, DisplayAsIcon:=False, Left:=818, Top:=945, Width:=120, Height:=17)
blub.name = "CB_check_tradeable_hour"
blub.Object.Caption = "Check if tradeable hours"


For i = 1 To 3
    
    'label
    Set blub = Worksheets("FORMAT2").OLEObjects.Add(ClassType:="Forms.Label.1", link:=False, DisplayAsIcon:=False, Left:=550, Top:=1125 + (i * 25), Width:=40, Height:=17)
    blub.name = "L_format2_customfield" & i
    blub.Object.Caption = "Custom " & i
    
    
    'cb type
    Set blub = Worksheets("FORMAT2").OLEObjects.Add(ClassType:="Forms.ComboBox.1", link:=False, DisplayAsIcon:=False, Left:=600, Top:=1125 + (i * 25), Width:=72, Height:=17)
    blub.name = "CB_format2_customfield" & i & "_type"
    
    
    'cb name
    Set blub = Worksheets("FORMAT2").OLEObjects.Add(ClassType:="Forms.ComboBox.1", link:=False, DisplayAsIcon:=False, Left:=675, Top:=1125 + (i * 25), Width:=120, Height:=17)
    blub.name = "CB_format2_customfield" & i & "_name"
    
    'cb sens
    Set blub = Worksheets("FORMAT2").OLEObjects.Add(ClassType:="Forms.ComboBox.1", link:=False, DisplayAsIcon:=False, Left:=800, Top:=1125 + (i * 25), Width:=33, Height:=17)
    blub.name = "CB_format2_customfield" & i & "_sens"
    
    'tb value
    Set blub = Worksheets("FORMAT2").OLEObjects.Add(ClassType:="Forms.TextBox.1", link:=False, DisplayAsIcon:=False, Left:=835, Top:=1125 + (i * 25), Width:=120, Height:=17)
    blub.name = "TB_format2_customfield" & i & "_value"
    

Next i



End Sub


Public Sub print_orders_queue_redi_plus()

Dim debug_test As Variant
debug_test = get_redi_orders

Dim tmp_wrbk As String, tmp_wrksht As String

tmp_wrbk = "Book2"
tmp_wrksht = "Sheet1"

Workbooks(tmp_wrbk).Worksheets(tmp_wrksht).Cells.Clear

For i = 0 To UBound(debug_test, 1)
    For j = 0 To UBound(debug_test(i), 1)
        Workbooks(tmp_wrbk).Worksheets(tmp_wrksht).Cells(i + 1, j + 1) = debug_test(i)(j)
    Next j
Next i


End Sub


Public Function get_redi_amount_in_queue(ByVal queue As String, ByVal ticker As String, ByVal side As String) As Long

get_redi_amount_in_queue = 0

ticker = get_symbol_redi_plus(ticker)

Dim i As Long, j As Long, k As Long, m As Long, n  As Long

Dim tmp_qty As Double


Dim list_orders As Variant

If InStr(UCase(queue), UCase("exe")) <> 0 Then
    list_orders = get_redi_exec
ElseIf InStr(UCase(queue), UCase("ord")) <> 0 Then
    list_orders = get_redi_orders
End If



If IsEmpty(list_orders) = False Then
    
    'repere la colonne symbol + qty
    
    For i = 0 To UBound(list_orders(0), 1)
        
        If list_orders(0)(i) = "Symbol" Then
            dim_symbol = i
        ElseIf list_orders(0)(i) = "Side" Then
            dim_side = i
        ElseIf list_orders(0)(i) = "OrderQty" Then
            dim_order_qty = i
        ElseIf list_orders(0)(i) = "ExecQty" Then
            dim_exec_qty = i
        End If
        
    Next i
    
    
    For i = 1 To UBound(list_orders, 1)
        
        If Replace(UCase(list_orders(i)(dim_symbol)), ".", " ") = UCase(ticker) Then
                
            If UCase(Left(list_orders(i)(dim_side), 1)) = "B" Then
                tmp_qty = Abs(list_orders(i)(dim_order_qty) - list_orders(i)(dim_exec_qty))
            ElseIf UCase(Left(list_orders(i)(dim_side), 1)) = "S" Then
                tmp_qty = -Abs(list_orders(i)(dim_order_qty) - list_orders(i)(dim_exec_qty))
            End If
            
                
            If UCase(Left(side, 1)) = "A" Then
                get_redi_amount_in_queue = get_redi_amount_in_queue + tmp_qty
            ElseIf UCase(Left(side, 1)) = "B" Then
                
                If tmp_qty > 0 Then
                    get_redi_amount_in_queue = get_redi_amount_in_queue + tmp_qty
                End If
                
            ElseIf UCase(Left(side, 1)) = "S" Then
                
                If tmp_qty < 0 Then
                    get_redi_amount_in_queue = get_redi_amount_in_queue + tmp_qty
                End If
                
            End If
            
        End If
        
    Next i
    
    
Else
    get_redi_amount_in_queue = 0
    Exit Function
End If

End Function


Public Function get_redi_orders() As Variant

Dim vtable As Variant
Dim vwhere As Variant
Dim verr As Variant


If IsRediReady Then
    
    If IsEmpty(ThisWorkbook.RediOrders) = True Then
        If ThisWorkbook.OrderQuery Is Nothing Then
            Set ThisWorkbook.OrderQuery = New RediLib.CacheControl
        End If
        
        ThisWorkbook.OrderQuery.UserID = ""
        ThisWorkbook.OrderQuery.Password = ""
        vtable = "Message"
        vwhere = "true"
        
        MessageQuery = ThisWorkbook.OrderQuery.Submit(vtable, vwhere, verr)
        
        ThisWorkbook.OrderQuery.Revoke verr
    End If
    
    get_redi_orders = ThisWorkbook.RediOrders
    'get_redi_exec = ThisWorkbook.RediExec
End If

End Function


Public Function get_redi_exec() As Variant

Dim vtable As Variant
Dim vwhere As Variant
Dim verr As Variant

If IsRediReady Then
    
    If IsEmpty(ThisWorkbook.RediExec) = True Then
        If ThisWorkbook.OrderQuery Is Nothing Then
            Set ThisWorkbook.OrderQuery = New RediLib.CacheControl
        End If
        
        ThisWorkbook.OrderQuery.UserID = ""
        ThisWorkbook.OrderQuery.Password = ""
        vtable = "Message"
        vwhere = "true"
        
        MessageQuery = ThisWorkbook.OrderQuery.Submit(vtable, vwhere, verr)
        
        ThisWorkbook.OrderQuery.Revoke verr
    End If
    
    'get_redi_orders = ThisWorkbook.RediOrders
    get_redi_exec = ThisWorkbook.RediExec
End If

End Function


Public Function IsRediReady() As Boolean

Dim RediApp As RediLib.Application
Set RediApp = New RediLib.Application
Dim strUserID As String

strUserID = RediApp.UserID

If strUserID = "" Then
    MsgBox "REDIPlus is not currently running.  Please login to REDIPlus first.", vbInformation, Title
    DoEvents
    RediApp.Quit (False)
    IsRediReady = False
Else
    IsRediReady = True
End If

Set RediApp = Nothing

End Function


Public Function redi_elocate(ByVal ticker As String, ByVal qty As Long)

Dim debug_test As Variant

Dim redi_userid As String, redi_password As String
redi_userid = Worksheets("FORMAT2").Cells(7, 21)
redi_password = Decrypter(Worksheets("FORMAT2").Cells(8, 21))

Dim objRedi As New RediLib.ELOCATE

If IsRediReady Then
    
    objRedi.UserID = ""
    objRedi.Password = ""
    
    objRedi.UserID = redi_userid
    objRedi.Password = redi_password
    
    debug_test = ElocateType.REQUEST
    objRedi.ElocateType = ElocateType.REQUEST
    objRedi.AddInfo get_symbol_redi_plus(ticker), Abs(qty), "GLOC"
    objRedi.Submit
    
    
End If

End Function


Sub update_bridge_with_GVA(ByVal vec_product_id As Variant, ByVal vec_underlying_id As Variant, ByVal vec_description As Variant, ByVal vec_currency As Variant)

Dim debug_test As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

'mount les bridges deja présent
Dim bridge_product_id() As Variant

Dim sql_query As String
sql_query = "SELECT * FROM t_bridge"
Dim extract_bridge As Variant
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_bridge, 1) <> 0 Then
    ReDim bridge_product_id(UBound(extract_bridge, 1) - 1)
Else
    ReDim bridge_product_id(0)
End If


j = 0
For i = 1 To UBound(extract_bridge, 1)
    bridge_product_id(j) = extract_bridge(i, 0)
    
    j = j + 1
Next i


'mount les exceptions
Dim extract_exception As Variant
sql_query = "SELECT * FROM t_exception"
extract_exception = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'mount les instruments
Dim extract_instrument As Variant
sql_query = "SELECT * FROM t_instrument"
extract_instrument = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With


Dim vec_new_equity() As Variant
Dim vec_new_future() As Variant
Dim vec_new_option() As Variant

Dim count_new_equity As Integer, count_new_future As Integer, count_new_option As Integer
    count_new_equity = 0
    count_new_future = 0
    count_new_option = 0

Dim count_new_line_in_bridge As Integer
count_new_line_in_bridge = 0

With rst
    
    .ActiveConnection = conn
    .Open "t_bridge", LockType:=adLockOptimistic


        'insere les nouvelles lignes
        For i = 0 To UBound(vec_product_id, 1)
            For j = 0 To UBound(bridge_product_id, 1)
                If vec_product_id(i) = bridge_product_id(j) Then
                    'deja présent dans bridge
                    Exit For
                Else
                    If j = UBound(bridge_product_id, 1) Then
                        'insertion de la ligne
                        .AddNew
                            
                            count_new_line_in_bridge = count_new_line_in_bridge + 1
                            
                            
                            'PRODUCT ID
                            .fields("gs_id") = vec_product_id(i)
                            
                            
                            'UNDERLYING ID
                            For k = 1 To UBound(extract_exception, 1)
                                If extract_exception(k, 0) = vec_product_id(i) Then
                                    'charge l'underlying de l'exception
                                    .fields("gs_underlying_id") = extract_exception(k, 1)
                                    Exit For
                                Else
                                    If k = UBound(extract_exception, 1) Then
                                        .fields("gs_underlying_id") = vec_underlying_id(i)
                                    End If
                                End If
                            Next k
                            
                            
                            'DESCRIPTION
                            .fields("gs_description") = vec_description(i)
                            
                            
                            'INSTRUMENT ID
                            If .fields("gs_id") = .fields("gs_underlying_id") Then
                                
                                'equity
                                For k = 1 To UBound(extract_instrument, 1)
                                    If extract_instrument(k, 1) = "equity" Then
                                        .fields("system_instrument_id") = extract_instrument(k, 0)
                                        Exit For
                                    End If
                                Next k
                                
                                
                                ReDim Preserve vec_new_equity(count_new_equity)
                                vec_new_equity(count_new_equity) = vec_product_id(i)
                                count_new_equity = count_new_equity + 1
                                
                            Else
                                If InStr(UCase(Left(vec_description(i), 4)), "CALL") <> 0 Or InStr(UCase(Left(vec_description(i), 3)), "PUT") <> 0 Then
                                    
                                    'option
                                    For k = 1 To UBound(extract_instrument, 1)
                                        If extract_instrument(k, 1) = "option" Then
                                            .fields("system_instrument_id") = extract_instrument(k, 0)
                                            Exit For
                                        End If
                                    Next k
                                    
                                    ReDim Preserve vec_new_option(count_new_option)
                                    vec_new_option(count_new_option) = vec_product_id(i)
                                    count_new_option = count_new_option + 1
                                    
                                Else
                                    
                                    'future
                                    For k = 1 To UBound(extract_instrument, 1)
                                        If extract_instrument(k, 1) = "future" Then
                                            .fields("system_instrument_id") = extract_instrument(k, 0)
                                            Exit For
                                        End If
                                    Next k
                                    
                                    ReDim Preserve vec_new_future(count_new_future)
                                    vec_new_future(count_new_future) = vec_product_id(i)
                                    count_new_future = count_new_future + 1
                                    
                                End If
                            End If
                        
                        
                        .Update
                        
                    End If
                End If
            Next j
        Next i
        
End With


'update des tables
If count_new_equity > 0 Then
    Call insert_new_equity_gva(vec_new_equity)
End If

If count_new_future > 0 Then
    Call insert_new_future_gva(vec_new_future)
End If

If count_new_option > 0 Then
    Call insert_new_option_gva(vec_new_option)
End If


MsgBox ("New lines in bridge : " & count_new_line_in_bridge)

End Sub


Sub insert_new_equity_gva(ByVal id As Variant) 'reception de vecteurs

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim matrix_new_equity() As Variant
ReDim matrix_new_equity(UBound(id, 1), 20)

    Dim dim_gs_id As Integer, dim_gs_isin As Integer, dim_gs_sedol As Integer, dim_gs_name As Integer, dim_gs_ric As Integer, dim_gs_crncy As Integer, _
        dim_gs_longshot As Integer, dim_bbg_NAME As Integer, dim_bbg_sector As Integer, dim_bbg_industry As Integer, _
        dim_system_ticker As Integer, dim_system_crncy_code As Integer, dim_system_sector_code As Integer, dim_system_industry_code As Integer
    Dim dim_already_found_in_db As Integer
    
    dim_gs_id = 0
    dim_gs_isin = 1
    dim_gs_sedol = 2
    dim_gs_name = 3
    dim_gs_ric = 4
    dim_gs_crncy = 5
    dim_gs_longshot = 6
    dim_bbg_NAME = 7
    dim_bbg_sector = 8
    dim_bbg_industry = 9
    dim_system_ticker = 10
    dim_system_crncy_code = 11
    dim_system_sector_code = 12
    dim_system_industry_code = 13
    
    dim_already_found_in_db = 14



For i = 0 To UBound(id, 1)
    matrix_new_equity(i, dim_gs_id) = id(i)
Next i


'mount les id deja présent
Dim extract_equity As Variant
sql_query = "SELECT gs_id FROM t_equity"
extract_equity = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'mount les currency code
Dim extract_crncy As Variant
sql_query = "SELECT system_code, system_name FROM t_currency"
extract_crncy = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'mount database folio
Dim vec_db_folio_id() As Variant
Dim vec_db_folio_isin() As Variant
Dim vec_db_folio_sedol() As Variant
Dim vec_db_folio_descritpion() As Variant
Dim vec_db_folio_ric() As Variant
Dim vec_db_folio_crncy() As Variant
Dim vec_db_folio_ticker() As Variant

Dim l_sht_db_folio_header As Integer, c_sht_db_folio_id As Integer, c_sht_db_folio_u_id As Integer, c_sht_db_folio_description As Integer, _
    c_sht_db_folio_isin As Integer, c_sht_db_folio_sedol As Integer, c_sht_db_folio_ticker As Integer, c_sht_db_folio_ric As Integer, _
    c_sht_db_folio_crncy As Integer


l_sht_db_folio_header = 12

c_sht_db_folio_description = 6
c_sht_db_folio_id = 10
c_sht_db_folio_u_id = 11
c_sht_db_folio_isin = 2
c_sht_db_folio_sedol = 3
c_sht_db_folio_ticker = 4
c_sht_db_folio_ric = 5
c_sht_db_folio_crncy = 9

k = 0
For i = l_sht_db_folio_header + 1 To 32000
    If Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id) = "" And Worksheets("Database_Folio").Cells(i + 1, c_sht_db_folio_id) = "" And Worksheets("Database_Folio").Cells(i + 2, c_sht_db_folio_id) = "" Then
        Exit For
    Else
        If Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id) <> "" And Worksheets("Database_Folio").Cells(i, c_sht_db_folio_description) <> "" Then
            ReDim Preserve vec_db_folio_id(k)
            ReDim Preserve vec_db_folio_isin(k)
            ReDim Preserve vec_db_folio_sedol(k)
            ReDim Preserve vec_db_folio_descritpion(k)
            ReDim Preserve vec_db_folio_ric(k)
            ReDim Preserve vec_db_folio_crncy(k)
            ReDim Preserve vec_db_folio_ticker(k)
            
            vec_db_folio_id(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id)
            vec_db_folio_isin(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_isin)
            vec_db_folio_sedol(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_sedol)
            vec_db_folio_descritpion(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_description)
            vec_db_folio_ric(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_ric)
            vec_db_folio_crncy(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_crncy)
            vec_db_folio_ticker(k) = Replace(Worksheets("Database_Folio").Cells(i, c_sht_db_folio_ticker), " EQUITY", " Equity")
            
            k = k + 1
            
        End If
    End If
Next i


Dim check_wrkb As Workbook
Dim FoundWrbk As Boolean

FoundWrbk = False
For Each check_wrkb In Workbooks
    If check_wrkb.name = db_folio Then
        FoundWrbk = True
        Exit For
    End If
Next

If FoundWrbk = False Then
    Workbooks.Open filename:=base_path & db_folio, readOnly:=True
End If

Dim found_data As Boolean

Dim vec_ticker() As Variant
k = 0
For i = 0 To UBound(matrix_new_equity, 1)
    
    For n = 0 To UBound(extract_equity, 1)
    
        If extract_equity(n, 0) = matrix_new_equity(i, dim_gs_id) Then
            'l entrée est déjà présente
            matrix_new_equity(i, dim_already_found_in_db) = True
            Exit For
        Else
            If n = UBound(extract_equity, 1) Then
                matrix_new_equity(i, dim_already_found_in_db) = False
                found_data = False
                
                For j = 2 To 32000
                    If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 1) = "" And Workbooks(db_folio).Worksheets("Sheet1").Cells(j + 2, 1) = "" And Workbooks(db_folio).Worksheets("Sheet1").Cells(j + 3, 1) = "" Then
                        Exit For
                    Else
                        If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 11) = matrix_new_equity(i, dim_gs_id) Then
                            
                            found_data = True
                            
                            'isin
                            If matrix_new_equity(i, dim_gs_isin) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 2) <> "" Then
                                    matrix_new_equity(i, dim_gs_isin) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 2)
                                ElseIf Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 17) <> "" Then
                                    matrix_new_equity(i, dim_gs_isin) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 17)
                                End If
                            End If
                            
                            
                            'sedol
                            If matrix_new_equity(i, dim_gs_sedol) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 3) <> "" Then
                                    matrix_new_equity(i, dim_gs_sedol) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 3)
                                End If
                            End If
                            
                            
                            'longshot
                            If matrix_new_equity(i, dim_gs_longshot) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 12) <> "" Then
                                    matrix_new_equity(i, dim_gs_longshot) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 12)
                                End If
                            End If
                            
                            
                            
                            'ric
                            If matrix_new_equity(i, dim_gs_ric) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 5) <> "" Then
                                    matrix_new_equity(i, dim_gs_ric) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 5)
                                End If
                            End If
                            
                            
                            
                            'ticker
                            If matrix_new_equity(i, dim_system_ticker) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 4) <> "" Then
                                    matrix_new_equity(i, dim_system_ticker) = Replace(Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 4), "EQUITY", "Equity")
                                    
                                    ReDim Preserve vec_ticker(k)
                                    vec_ticker(k) = matrix_new_equity(i, dim_system_ticker)
                                    k = k + 1
                                    
                                ElseIf Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 15) <> "" Then
                                    matrix_new_equity(i, dim_system_ticker) = Replace(Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 15), "EQUITY", "Equity")
                                    
                                    ReDim Preserve vec_ticker(k)
                                    vec_ticker(k) = matrix_new_equity(i, dim_system_ticker)
                                    k = k + 1
                                End If
                            End If
                            
                            
                            'currency
                            If matrix_new_equity(i, dim_gs_crncy) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 9) <> "" Then
                                    matrix_new_equity(i, dim_gs_crncy) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 9)
                                    
                                    For m = 0 To UBound(extract_crncy, 1)
                                        If UCase(extract_crncy(m, 1)) = UCase(matrix_new_equity(i, dim_gs_crncy)) Then
                                            matrix_new_equity(i, dim_system_crncy_code) = extract_crncy(m, 0)
                                            Exit For
                                        End If
                                    Next m
                                    
                                End If
                            End If
                            
                            
                            'description
                            If matrix_new_equity(i, dim_gs_name) = "" Then
                                
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6) <> "" And InStr(Left(UCase(Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6)), 4), "CALL") = 0 And InStr(Left(UCase(Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6)), 3), "PUT") = 0 Then
                                    matrix_new_equity(i, dim_gs_name) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6)
                                End If
                            End If
                            
                        End If
                    End If
                Next j
                
                
                'si rien dans folio, complete ce qui est possible avec la sheet database_folio
                If found_data = False Then
                    For j = 0 To UBound(vec_db_folio_id, 1)
                        If vec_db_folio_id(j) = matrix_new_equity(i, dim_gs_id) Then
                            matrix_new_equity(i, dim_gs_isin) = vec_db_folio_isin(j)
                            matrix_new_equity(i, dim_gs_sedol) = vec_db_folio_sedol(j)
                            matrix_new_equity(i, dim_gs_ric) = vec_db_folio_ric(j)
                            matrix_new_equity(i, dim_system_ticker) = vec_db_folio_ticker(j)
                            
                            matrix_new_equity(i, dim_gs_crncy) = vec_db_folio_crncy(j)
                            
                            For m = 0 To UBound(extract_crncy, 1)
                                If UCase(extract_crncy(m, 1)) = UCase(matrix_new_equity(i, dim_gs_crncy)) Then
                                    matrix_new_equity(i, dim_system_crncy_code) = extract_crncy(m, 0)
                                    Exit For
                                End If
                            Next m
                            
                            matrix_new_equity(i, dim_gs_name) = vec_db_folio_descritpion(j)
                            
                            Exit For
                        End If
                    Next j
                End If
                
                
            End If
        End If
    Next n
Next i

Workbooks(db_folio).Close False


'envoi du resultat dans la table t_equity
'envoie des positions dans la base de données
'mount la connexion
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With



With rst
    
    .ActiveConnection = conn
    .Open "t_equity", LockType:=adLockOptimistic
    
    
    For i = 0 To UBound(matrix_new_equity, 1)
        If matrix_new_equity(i, dim_already_found_in_db) = False Then
            
            .AddNew
            
                .fields("gs_id") = matrix_new_equity(i, dim_gs_id)
                .fields("gs_isin") = matrix_new_equity(i, dim_gs_isin)
                .fields("gs_sedol") = matrix_new_equity(i, dim_gs_sedol)
                .fields("gs_name") = matrix_new_equity(i, dim_gs_name)
                .fields("gs_ric") = matrix_new_equity(i, dim_gs_ric)
                .fields("gs_currency") = matrix_new_equity(i, dim_gs_crncy)
                .fields("longshot_id") = matrix_new_equity(i, dim_gs_longshot)
                .fields("bbg_name") = matrix_new_equity(i, dim_bbg_NAME)
                .fields("bbg_sector") = matrix_new_equity(i, dim_bbg_sector)
                .fields("bbg_industry") = matrix_new_equity(i, dim_bbg_industry)
                .fields("system_ticker") = matrix_new_equity(i, dim_system_ticker)
                .fields("system_currency_code") = matrix_new_equity(i, dim_system_crncy_code)
                .fields("system_sector_code") = matrix_new_equity(i, dim_system_sector_code)
                .fields("system_industry_code") = matrix_new_equity(i, dim_system_industry_code)
            
            .Update
            
        End If
        
    Next i
End With


rst.Close
conn.Close


End Sub


Sub insert_new_future_pict_exec(ByVal vec_future As Variant)

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim fileFolio As File_Folio
Set fileFolio = New File_Folio
fileFolio.set_file_path = base_path & "\GS_Folio\" & folio_all_view

Dim matrix_folio As Variant
matrix_folio = fileFolio.get_content_as_a_matrix()


Dim dim_folio_id As Integer, dim_folio_ticker As Integer, dim_folio_crncy As Integer, dim_folio_description As Integer, _
    dim_folio_qty_yesterday_close As Integer, dim_folio_underlying_id As Integer, dim_folio_yesterday_close_price As Integer, _
    dim_folio_product_type As Integer, dim_folio_option_strike As Integer, dim_folio_option_put_call As Integer, _
    dim_folio_option_contract_size As Integer, dim_folio_option_expiration_date As Integer, dim_folio_option_exercise_style As Integer, _
    dim_folio_isin As Integer



'detect les dimensions
For i = 1 To UBound(matrix_folio, 2)
    If matrix_folio(0, i) = "Identifier" Then
        dim_folio_id = i
    ElseIf matrix_folio(0, i) = "CCY" Then
        dim_folio_crncy = i
    ElseIf matrix_folio(0, i) = "Description" Then
        dim_folio_description = i
    End If
Next i


'database_folio
Dim vec_db_folio_id() As Variant
Dim vec_db_folio_isin() As Variant
Dim vec_db_folio_sedol() As Variant
Dim vec_db_folio_descritpion() As Variant
Dim vec_db_folio_ric() As Variant
Dim vec_db_folio_crncy() As Variant
Dim vec_db_folio_ticker() As Variant

Dim l_sht_db_folio_header As Integer, c_sht_db_folio_id As Integer, c_sht_db_folio_u_id As Integer, c_sht_db_folio_description As Integer, _
    c_sht_db_folio_isin As Integer, c_sht_db_folio_sedol As Integer, c_sht_db_folio_ticker As Integer, c_sht_db_folio_ric As Integer, _
    c_sht_db_folio_crncy As Integer


l_sht_db_folio_header = 12

c_sht_db_folio_description = 6
c_sht_db_folio_id = 10
c_sht_db_folio_u_id = 11
c_sht_db_folio_isin = 2
c_sht_db_folio_sedol = 3
c_sht_db_folio_ticker = 4
c_sht_db_folio_ric = 5
c_sht_db_folio_crncy = 9

k = 0
For i = l_sht_db_folio_header + 1 To 32000
    If Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id) = "" And Worksheets("Database_Folio").Cells(i + 1, c_sht_db_folio_id) = "" And Worksheets("Database_Folio").Cells(i + 2, c_sht_db_folio_id) = "" Then
        Exit For
    Else
        If Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id) <> "" And Worksheets("Database_Folio").Cells(i, c_sht_db_folio_description) <> "" Then
            ReDim Preserve vec_db_folio_id(k)
            ReDim Preserve vec_db_folio_descritpion(k)
            ReDim Preserve vec_db_folio_crncy(k)
            
            vec_db_folio_id(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id)
            vec_db_folio_descritpion(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_description)
            vec_db_folio_crncy(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_crncy)

            k = k + 1
            
        End If
    End If
Next i








'remonte l'état actuel de la table equity
Dim extract_future As Variant
sql_query = "SELECT gs_id FROM t_future"
extract_future = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

'crncy
Dim extract_crncy As Variant
sql_query = "SELECT system_code, system_name FROM t_currency"
extract_crncy = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With

Dim found_data As Boolean

With rst
    
    .ActiveConnection = conn
    .Open "t_future", LockType:=adLockOptimistic

    For i = 0 To UBound(vec_future, 1)
        
        found_data = False
        
        For j = 0 To UBound(extract_future, 1)
            If vec_future(i) = extract_future(j, 0) Then
                Exit For
            Else
                If j = UBound(extract_future, 1) Then
                    
                    For k = 1 To UBound(matrix_folio, 1)
                        If vec_future(i) = matrix_folio(k, dim_folio_id) Then
                            
                            found_data = True
                            
                            .AddNew
                            
                                .fields("gs_id") = vec_future(i)
                                .fields("gs_name") = matrix_folio(k, dim_folio_description)
                                .fields("gs_currency") = UCase(matrix_folio(k, dim_folio_crncy))
                                
                                
                                For m = 1 To UBound(extract_crncy, 1)
                                    If UCase(matrix_folio(k, dim_folio_crncy)) = extract_crncy(m, 1) Then
                                        .fields("system_currency_code") = extract_crncy(m, 0)
                                        Exit For
                                    End If
                                Next m
                                
                            .Update
                            
                            Exit For
                        End If
                    Next k
                    
                    If found_data = False Then
                        
                        'tentative avec database_folio
                        For k = 0 To UBound(vec_db_folio_id, 1)
                            If vec_db_folio_id(k) = vec_future(i) Then
                                
                                found_data = True
                                
                                .AddNew
                                
                                    .fields("gs_id") = vec_future(i)
                                    .fields("gs_name") = vec_db_folio_descritpion(k)
                                    .fields("gs_currency") = UCase(vec_db_folio_crncy(k))
                                    .fields("system_ticker") = Replace(vec_db_folio_ticker(k), " EQUITY", " Equity")
                                    
                                
                                .Update
                                
                                Exit For
                            End If
                        Next k
                    End If
                    
                End If
            End If
        Next j
    Next i
    
    .Close
    
End With

conn.Close

End Sub


Sub insert_new_future_gva(ByVal id As Variant) 'reception de vecteurs

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim matrix_new_future() As Variant
ReDim matrix_new_future(UBound(id, 1), 20)

    Dim dim_gs_id As Integer, dim_gs_name As Integer, dim_gs_crncy As Integer, dim_bbg_NAME As Integer, _
        dim_bbg_settlement_date As Integer, dim_bbg_future_contract_size As Integer, dim_system_settlement_date As Integer, _
        dim_system_crncy_code As Integer
    
    
    dim_gs_id = 0
    dim_gs_name = 1
    dim_gs_crncy = 2
    dim_bbg_NAME = 3
    dim_bbg_settlement_date = 4
    dim_bbg_future_contract_size = 5
    dim_system_settlement_date = 6
    dim_system_crncy_code = 7
    
    dim_already_found_in_db = 8



For i = 0 To UBound(id, 1)
    matrix_new_future(i, dim_gs_id) = id(i)
Next i

'mount les id deja présent
Dim extract_future As Variant
sql_query = "SELECT gs_id FROM t_future"
extract_future = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'mount les currency code
Dim extract_crncy As Variant
sql_query = "SELECT system_code, system_name FROM t_currency"
extract_crncy = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


Dim check_wrkb As Workbook
Dim FoundWrbk As Boolean

FoundWrbk = False
For Each check_wrkb In Workbooks
    If check_wrkb.name = db_folio Then
        FoundWrbk = True
        Exit For
    End If
Next

If FoundWrbk = False Then
    Workbooks.Open filename:=base_path & db_folio, readOnly:=True
End If

Dim vec_ticker() As Variant
k = 0
For i = 0 To UBound(matrix_new_future, 1)
    For n = 0 To UBound(extract_future, 1)
        If extract_future(n, 0) = matrix_new_future(i, dim_gs_id) Then
            matrix_new_future(i, dim_already_found_in_db) = True
            Exit For
        Else
            If n = UBound(extract_future, 1) Then
                
                matrix_new_future(i, dim_already_found_in_db) = False
                
                For j = 2 To 32000
                    If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 1) = "" And Workbooks(db_folio).Worksheets("Sheet1").Cells(j + 2, 1) = "" And Workbooks(db_folio).Worksheets("Sheet1").Cells(j + 3, 1) = "" Then
                        Exit For
                    Else
                        If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 10) = matrix_new_future(i, dim_gs_id) Then
                            
                            'name
                            If matrix_new_future(i, dim_gs_name) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6) <> "" Then
                                    matrix_new_future(i, dim_gs_name) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6)
                                End If
                            End If
                            
                            
                            'currency
                            If matrix_new_future(i, dim_gs_crncy) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 9) <> "" Then
                                    matrix_new_future(i, dim_gs_crncy) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 9)
                                    
                                    For m = 0 To UBound(extract_crncy, 1)
                                        If UCase(extract_crncy(m, 1)) = UCase(matrix_new_future(i, dim_gs_crncy)) Then
                                            matrix_new_future(i, dim_system_crncy_code) = extract_crncy(m, 0)
                                            Exit For
                                        End If
                                    Next m
                                    
                                End If
                            End If
                            
                        End If
                    End If
                Next j
            End If
        End If
    Next n
Next i

Workbooks(db_folio).Close False

'envoi du resultat dans la table t_equity
'envoie des positions dans la base de données
'mount la connexion
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With



With rst
    
    .ActiveConnection = conn
    .Open "t_future", LockType:=adLockOptimistic
    
    
    For i = 0 To UBound(matrix_new_future, 1)
        
        If matrix_new_future(i, dim_already_found_in_db) = False Then
        
            .AddNew
            
                .fields("gs_id") = matrix_new_future(i, dim_gs_id)
                .fields("gs_name") = matrix_new_future(i, dim_gs_name)
                .fields("gs_currency") = matrix_new_future(i, dim_gs_crncy)
                .fields("bbg_name") = matrix_new_future(i, dim_bbg_NAME)
                .fields("bbg_settlement_date") = matrix_new_future(i, dim_bbg_settlement_date)
                .fields("bbg_future_contract_size") = matrix_new_future(i, dim_bbg_future_contract_size)
                .fields("system_settlement_date") = matrix_new_future(i, dim_system_settlement_date)
                .fields("system_currency_code") = matrix_new_future(i, dim_system_crncy_code)
    
            .Update
        End If
        
    Next i
End With


rst.Close
conn.Close

End Sub


Sub insert_new_option_gva(ByVal id As Variant) 'reception de vecteurs

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim matrix_new_option() As Variant
ReDim matrix_new_option(UBound(id, 1), 20)

    Dim dim_gs_id As Integer, dim_gs_name As Integer, dim_bbg_TICKER As Integer, dim_bbg_strike As Integer, _
        dim_bbg_expiry As Integer, dim_bbg_exercise_type As Integer, dim_bbg_put_call As Integer, dim_bbg_option_contract_size As Integer
        
    
    
    dim_gs_id = 0
    dim_gs_name = 1
    dim_bbg_TICKER = 2
    dim_bbg_strike = 3
    dim_bbg_expiry = 4
    dim_bbg_exercise_type = 5
    dim_bbg_put_call = 6
    dim_bbg_option_contract_size = 7
    
    
    dim_already_found_in_db = 8



For i = 0 To UBound(id, 1)
    matrix_new_option(i, dim_gs_id) = id(i)
Next i

'mount les id deja présent
Dim extract_option As Variant
sql_query = "SELECT gs_id FROM t_option"
extract_option = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'mount les currency code
Dim extract_crncy As Variant
sql_query = "SELECT system_code, system_name FROM t_currency"
extract_crncy = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


Dim check_wrkb As Workbook
Dim FoundWrbk As Boolean

FoundWrbk = False
For Each check_wrkb In Workbooks
    If check_wrkb.name = db_folio Then
        FoundWrbk = True
        Exit For
    End If
Next

If FoundWrbk = False Then
    Workbooks.Open filename:=base_path & db_folio, readOnly:=True
End If

Dim vec_ticker() As Variant

Dim underlying_ticker As String
Dim underlying_ticker_market As String
Dim date_txt As String
Dim strike As Double
Dim pos_space As Integer
Dim pos_start_date As Integer

k = 0
For i = 0 To UBound(matrix_new_option, 1)
    For n = 0 To UBound(extract_option, 1)
        If extract_option(n, 0) = matrix_new_option(i, dim_gs_id) Then
            matrix_new_option(i, dim_already_found_in_db) = True
            Exit For
        Else
            If n = UBound(extract_option, 1) Then
                
                matrix_new_option(i, dim_already_found_in_db) = False
                
                For j = 2 To 32000
                    If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 1) = "" And Workbooks(db_folio).Worksheets("Sheet1").Cells(j + 2, 1) = "" And Workbooks(db_folio).Worksheets("Sheet1").Cells(j + 3, 1) = "" Then
                        Exit For
                    Else
                        If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 10) = matrix_new_option(i, dim_gs_id) Then
                            
                            'name
                            If matrix_new_option(i, dim_gs_name) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6) <> "" Then
                                    matrix_new_option(i, dim_gs_name) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 6)
                                End If
                            End If
                            
                            'ticker
                            If matrix_new_option(i, dim_bbg_TICKER) = "" Then
                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 14) <> "" Then
                                    matrix_new_option(i, dim_bbg_TICKER) = Replace(Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 14), "Equity EQUITY", "Equity")
                                    matrix_new_option(i, dim_bbg_TICKER) = Replace(Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 14), "Index EQUITY", "Index")
                                    
                                    
'                                    'tentative d'extract
'                                    If InStr(UCase(matrix_new_option(i, dim_bbg_ticker)), "EQUITY") <> 0 Then
'                                        underlying_ticker = Left(matrix_new_option(i, dim_bbg_ticker), InStr(matrix_new_option(i, dim_bbg_ticker), " ") - 1)
'                                        underlying_ticker_market = Mid(matrix_new_option(i, dim_bbg_ticker), Len(underlying_ticker) + 1, 2)
'                                        pos_start_date = Len(underlying_ticker) + 1 + Len(underlying_ticker_market) + 1
'                                        date_txt = Mid(pos_start_date, matrix_new_option(i, dim_bbg_ticker), InStr(pos_start_date + 1, matrix_new_option(i, dim_bbg_ticker), " ") - pos_start_date)
'
'                                        If Len(date_txt) = 8 Then
'                                            'weekly
'
'                                        ElseIf Len(date_txt) = 5 Then
'                                            'monthly
'
'                                        End If
'
'                                    ElseIf InStr(UCase(matrix_new_option(i, dim_bbg_ticker)), "INDEX") <> 0 Then
'
'                                    End If
                                End If
                            End If
                            
                            
'                            'currency
'                            If matrix_new_option(i, dim_gs_crncy) = "" Then
'                                If Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 9) <> "" Then
'                                    matrix_new_option(i, dim_gs_crncy) = Workbooks(db_folio).Worksheets("Sheet1").Cells(j, 9)
'
'                                    For m = 0 To UBound(extract_crncy, 1)
'                                        If UCase(extract_crncy(m, 1)) = UCase(matrix_new_option(i, dim_gs_crncy)) Then
'                                            matrix_new_option(i, dim_system_crncy_code) = extract_crncy(m, 0)
'                                            Exit For
'                                        End If
'                                    Next m
'
'                                End If
'                            End If
                        Exit For

                        End If
                    End If
                Next j
            End If
        End If
    Next n
Next i

Workbooks(db_folio).Close False


'envoi du resultat dans la table t_equity
'envoie des positions dans la base de données
'mount la connexion
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With



With rst
    
    .ActiveConnection = conn
    .Open "t_option", LockType:=adLockOptimistic
    
    
    For i = 0 To UBound(matrix_new_option, 1)
        
        If matrix_new_option(i, dim_already_found_in_db) = False Then
        
            .AddNew
            
                .fields("gs_id") = matrix_new_option(i, dim_gs_id)
                .fields("gs_name") = matrix_new_option(i, dim_gs_name)
                
                .fields("bbg_ticker") = matrix_new_option(i, dim_bbg_TICKER)
                .fields("bbg_strike") = matrix_new_option(i, dim_bbg_strike)
                .fields("bbg_expiry") = matrix_new_option(i, dim_bbg_expiry)
                .fields("bbg_exercise_type") = matrix_new_option(i, dim_bbg_exercise_type)
                .fields("bbg_put_call") = matrix_new_option(i, dim_bbg_put_call)
                .fields("bbg_option_contract_size") = matrix_new_option(i, dim_bbg_option_contract_size)
    
            .Update
            
        End If
        
    Next i
End With


rst.Close
conn.Close

End Sub



'standardiser la reception de la matrix
Sub update_bridge_pict_exec(ByVal vec_product_id As Variant, Optional ByVal matrix_data As Variant)

Dim debug_test As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

'mount les bridges deja présent
Dim bridge_product_id() As Variant

Dim sql_query As String
sql_query = "SELECT * FROM t_bridge"
Dim extract_bridge As Variant
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_bridge, 1) <> 0 Then
    ReDim bridge_product_id(UBound(extract_bridge, 1) - 1)
Else
    ReDim bridge_product_id(0)
End If


j = 0
For i = 1 To UBound(extract_bridge, 1)
    bridge_product_id(j) = extract_bridge(i, 0)
    
    j = j + 1
Next i


'mount les exceptions
Dim extract_exception As Variant
sql_query = "SELECT * FROM t_exception"
extract_exception = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'mount les instruments
Dim extract_instrument As Variant
sql_query = "SELECT * FROM t_instrument"
extract_instrument = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'remonte les entrée de database_folio
Dim l_db_folio_header As Integer, c_db_folio_id As Integer, c_db_folio_u_id As Integer, c_db_folio_description As Integer
l_db_folio_header = 12
c_db_folio_id = 10
c_db_folio_u_id = 11
c_db_folio_description = 6

Dim vec_db_folio_id() As Variant
Dim vec_db_folio_u_id() As Variant
Dim vec_db_folio_description() As Variant

k = 0
ReDim vec_db_folio_id(k)
For i = l_db_folio_header To 32000
    If Worksheets("Database_Folio").Cells(i, 1) = "" And Worksheets("Database_Folio").Cells(i + 1, 1) = "" And Worksheets("Database_Folio").Cells(i + 2, 1) = "" Then
        Exit For
    Else
        If Worksheets("Database_Folio").Cells(i, c_db_folio_id) <> "" And Worksheets("Database_Folio").Cells(i, c_db_folio_u_id) <> "" Then
            ReDim Preserve vec_db_folio_id(k)
            ReDim Preserve vec_db_folio_u_id(k)
            ReDim Preserve vec_db_folio_description(k)
            
            vec_db_folio_id(k) = Worksheets("Database_Folio").Cells(i, c_db_folio_id)
            vec_db_folio_u_id(k) = Worksheets("Database_Folio").Cells(i, c_db_folio_u_id)
            vec_db_folio_description(k) = Worksheets("Database_Folio").Cells(i, c_db_folio_description)
            
            k = k + 1
        End If
    End If
Next i

'remonte la view all de folio
Dim fileFolio As File_Folio
Set fileFolio = New File_Folio
fileFolio.set_file_path = base_path & "\GS_Folio\" & folio_all_view

Dim matrix_folio As Variant
matrix_folio = fileFolio.get_content_as_a_matrix()


Dim dim_folio_id As Integer, dim_folio_ticker As Integer, dim_folio_crncy As Integer, dim_folio_description As Integer, _
    dim_folio_qty_yesterday_close As Integer, dim_folio_underlying_id As Integer, dim_folio_yesterday_close_price As Integer, _
    dim_folio_product_type As Integer

'detect les dimensions
For i = 1 To UBound(matrix_folio, 2)
    If matrix_folio(0, i) = "Identifier" Then
        dim_folio_id = i
    ElseIf matrix_folio(0, i) = "Market Data Symbol" Then
        dim_folio_ticker = i
    ElseIf matrix_folio(0, i) = "CCY" Then
        dim_folio_crncy = i
    ElseIf matrix_folio(0, i) = "Description" Then
        dim_folio_description = i
    ElseIf matrix_folio(0, i) = "Qty - Yesterday's Close" Then
        dim_folio_qty_yesterday_close = i
    ElseIf matrix_folio(0, i) = "Underlyer Product ID" Then
        dim_folio_underlying_id = i
    ElseIf matrix_folio(0, i) = "Yesterday's Close (Local)" Then
        dim_folio_yesterday_close_price = i
    ElseIf matrix_folio(0, i) = "Product Type" Then
        dim_folio_product_type = i
    End If
Next i


Dim vec_underlying_id() As Variant
Dim vec_description() As Variant

    ReDim vec_underlying_id(UBound(vec_product_id, 1))
    ReDim vec_description(UBound(vec_product_id, 1))

For i = 0 To UBound(vec_product_id, 1)
    For j = 0 To UBound(matrix_folio, 1)
        If vec_product_id(i) = matrix_folio(j, dim_folio_id) Then
            
            'If matrix_folio(j, dim_folio_underlying_id) <> "" Then
                vec_underlying_id(i) = matrix_folio(j, dim_folio_underlying_id)
            'End If
            
            'If matrix_folio(j, dim_folio_description) <> "" Then
                vec_description(i) = matrix_folio(j, dim_folio_description)
            'End If
            
            Exit For
        End If
    Next j
    
    If IsEmpty(vec_underlying_id(i)) = True Then
        For j = 0 To UBound(vec_db_folio_id, 1)
            If vec_product_id(i) = vec_db_folio_id(j) Then
                
                vec_underlying_id(i) = vec_db_folio_u_id(j)
                vec_description(i) = vec_db_folio_description(j)
                
                Exit For
            End If
        Next j
    End If
Next i


Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With

Dim vec_bug_no_enough_info() As Variant
Dim count_bug_no_enough_info As Integer
    count_bug_no_enough_info = 0

Dim vec_new_equity() As Variant
Dim vec_new_future() As Variant
Dim vec_new_option() As Variant

Dim vec_new_exception() As Variant


Dim count_new_equity As Integer, count_new_future As Integer, count_new_option As Integer, count_new_exception As Integer
    count_new_equity = 0
    count_new_future = 0
    count_new_option = 0
    count_new_exception = 0
    
    ReDim vec_new_exception(0)

Dim count_new_line_in_bridge As Integer
count_new_line_in_bridge = 0

Dim reply As Variant

With rst
    
    .ActiveConnection = conn
    .Open "t_bridge", LockType:=adLockOptimistic


            'insere les nouvelles lignes
            For i = 0 To UBound(vec_product_id, 1)
                If vec_product_id(i) <> "" And IsEmpty(vec_underlying_id(i)) = False And IsEmpty(vec_description(i)) = False Then
                    For j = 0 To UBound(bridge_product_id, 1)
                        If vec_product_id(i) = bridge_product_id(j) Then
                            'deja présent dans bridge
                            Exit For
                        Else
                            If j = UBound(bridge_product_id, 1) Then
                                'insertion de la ligne
                                .AddNew
                                    
                                    count_new_line_in_bridge = count_new_line_in_bridge + 1
                                    
                                    
                                    'PRODUCT ID
                                    .fields("gs_id") = vec_product_id(i)
                                    
                                    
                                    'UNDERLYING ID
                                    For k = 1 To UBound(extract_exception, 1)
                                        If extract_exception(k, 0) = vec_product_id(i) Then
                                            'charge l'underlying de l'exception
                                            .fields("gs_underlying_id") = extract_exception(k, 1)
                                            Exit For
                                        Else
                                            If k = UBound(extract_exception, 1) Then
                                                .fields("gs_underlying_id") = vec_underlying_id(i)
                                            End If
                                        End If
                                    Next k
                                    
                                    
                                    'DESCRIPTION
                                    .fields("gs_description") = vec_description(i)
                                    
                                    
                                    'INSTRUMENT ID
                                    If .fields("gs_id") = .fields("gs_underlying_id") Then
                                        
                                        'equity
                                        For k = 1 To UBound(extract_instrument, 1)
                                            If extract_instrument(k, 1) = "equity" Then
                                                .fields("system_instrument_id") = extract_instrument(k, 0)
                                                Exit For
                                            End If
                                        Next k
                                        
                                        
                                        ReDim Preserve vec_new_equity(count_new_equity)
                                        vec_new_equity(count_new_equity) = vec_product_id(i)
                                        count_new_equity = count_new_equity + 1
                                        
                                    Else
                                        If InStr(UCase(Left(vec_description(i), 4)), "CALL") <> 0 Or InStr(UCase(Left(vec_description(i), 3)), "PUT") <> 0 Then
                                            
                                            'option
                                            For k = 1 To UBound(extract_instrument, 1)
                                                If extract_instrument(k, 1) = "option" Then
                                                    .fields("system_instrument_id") = extract_instrument(k, 0)
                                                    Exit For
                                                End If
                                            Next k
                                            
                                            ReDim Preserve vec_new_option(count_new_option)
                                            vec_new_option(count_new_option) = vec_product_id(i)
                                            count_new_option = count_new_option + 1
                                            
                                        Else
                                            
                                            'exception adr mise en place de l'exception + instru equity
                                            reply = MsgBox("Is this product a future or an ADR ? - " & vec_description(i), vbYesNo, "Future / ADR")
                                            
                                            If reply = vbYes Then
                                                'future
                                                
                                                For k = 1 To UBound(extract_instrument, 1)
                                                    If extract_instrument(k, 1) = "future" Then
                                                        .fields("system_instrument_id") = extract_instrument(k, 0)
                                                        Exit For
                                                    End If
                                                Next k
                                                
                                                ReDim Preserve vec_new_future(count_new_future)
                                                vec_new_future(count_new_future) = vec_product_id(i)
                                                count_new_future = count_new_future + 1
                                            Else
                                                'adr
                                                 .fields("gs_underlying_id") = vec_product_id(i)
                                                
                                                For k = 1 To UBound(extract_instrument, 1)
                                                    If extract_instrument(k, 1) = "equity" Then
                                                        .fields("system_instrument_id") = extract_instrument(k, 0)
                                                        Exit For
                                                    End If
                                                Next k
                                                
                                                'insertion de l'exception
                                                For k = 0 To UBound(vec_new_exception, 1)
                                                    If vec_product_id(i) = vec_new_exception(k) Then
                                                        Exit For
                                                    Else
                                                        If k = UBound(vec_new_exception, 1) Then
                                                            
                                                            ReDim Preserve vec_new_exception(count_new_exception)
                                                            vec_new_exception(count_new_exception) = vec_product_id(i)
                                                            count_new_exception = count_new_exception + 1
                                                            
                                                            ReDim Preserve vec_new_equity(count_new_equity)
                                                            vec_new_equity(count_new_equity) = vec_product_id(i)
                                                            count_new_equity = count_new_equity + 1
                                                            
                                                        End If
                                                    End If
                                                Next k
                                                
                                            End If
                                            
                                            
                                        End If
                                    End If
                                
                                
                                .Update
                                
                            End If
                        End If
                    Next j
                Else
                    ReDim Preserve vec_bug_no_enough_info(count_bug_no_enough_info)
                    vec_bug_no_enough_info(count_bug_no_enough_info) = vec_product_id(i)
                    count_bug_no_enough_info = count_bug_no_enough_info + 1
                End If
            Next i
    
    .Close
        
End With


'creation des entrées pour les exceptions
If count_new_exception > 0 Then
    With rst
    
        .ActiveConnection = conn
        .Open "t_exception", LockType:=adLockOptimistic
        
        For i = 0 To UBound(vec_new_exception, 1)
            .AddNew
            
                .fields("gs_id") = vec_new_exception(i)
                .fields("gs_underlying_id") = vec_new_exception(i)
            
            .Update
        Next i
        
        .Close
    
    End With
End If


conn.Close

If count_bug_no_enough_info > 0 Then
    MsgBox ("impossible d'insérer les données dans bridge pour " & count_bug_no_enough_info & " car les informations sont incomplètes")
    
    For i = 0 To UBound(vec_bug_no_enough_info, 1)
        Debug.Print vec_bug_no_enough_info(i)
    Next i
End If



'update des tables
If count_new_equity > 0 Then
   Call insert_new_equity_pict_exec(vec_new_equity)
End If

'If count_new_future > 0 Then
'    Call insert_new_future_pict_exec(vec_new_future)
'End If
'
If count_new_option > 0 Then
    Call insert_new_option_pict_exec(vec_new_option)
End If


MsgBox ("New lines in bridge : " & count_new_line_in_bridge)

End Sub


Sub insert_new_equity_pict_exec(ByVal vec_equity)

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim fileFolio As File_Folio
Set fileFolio = New File_Folio
fileFolio.set_file_path = base_path & "\GS_Folio\" & folio_all_view

Dim matrix_folio As Variant
matrix_folio = fileFolio.get_content_as_a_matrix()


Dim dim_folio_id As Integer, dim_folio_ticker As Integer, dim_folio_crncy As Integer, dim_folio_description As Integer, _
    dim_folio_qty_yesterday_close As Integer, dim_folio_underlying_id As Integer, dim_folio_yesterday_close_price As Integer, _
    dim_folio_product_type As Integer, dim_folio_option_strike As Integer, dim_folio_option_put_call As Integer, _
    dim_folio_option_contract_size As Integer, dim_folio_option_expiration_date As Integer, dim_folio_option_exercise_style As Integer, _
    dim_folio_isin As Integer



'detect les dimensions
For i = 1 To UBound(matrix_folio, 2)
    If matrix_folio(0, i) = "Identifier" Then
        dim_folio_id = i
    ElseIf matrix_folio(0, i) = "Market Data Symbol" Then
        dim_folio_ticker = i
    ElseIf matrix_folio(0, i) = "CCY" Then
        dim_folio_crncy = i
    ElseIf matrix_folio(0, i) = "Description" Then
        dim_folio_description = i
    ElseIf matrix_folio(0, i) = "Qty - Yesterday's Close" Then
        dim_folio_qty_yesterday_close = i
    ElseIf matrix_folio(0, i) = "Underlyer Product ID" Then
        dim_folio_underlying_id = i
    ElseIf matrix_folio(0, i) = "Yesterday's Close (Local)" Then
        dim_folio_yesterday_close_price = i
    ElseIf matrix_folio(0, i) = "Product Type" Then
        dim_folio_product_type = i
    ElseIf matrix_folio(0, i) = "Strike Price" Then
        dim_folio_option_strike = i
    ElseIf matrix_folio(0, i) = "Put/Call" Then
        dim_folio_option_put_call = i
    ElseIf matrix_folio(0, i) = "Contract Size" Then
        dim_folio_option_contract_size = i
    ElseIf matrix_folio(0, i) = "Expiration Date" Then
        dim_folio_option_expiration_date = i
    ElseIf matrix_folio(0, i) = "Option Exercise Style" Then
        dim_folio_option_exercise_style = i
    ElseIf matrix_folio(0, i) = "ISIN" Then
        dim_folio_isin = i
    End If
Next i


'database_folio
Dim vec_db_folio_id() As Variant
Dim vec_db_folio_isin() As Variant
Dim vec_db_folio_sedol() As Variant
Dim vec_db_folio_descritpion() As Variant
Dim vec_db_folio_ric() As Variant
Dim vec_db_folio_crncy() As Variant
Dim vec_db_folio_ticker() As Variant

Dim l_sht_db_folio_header As Integer, c_sht_db_folio_id As Integer, c_sht_db_folio_u_id As Integer, c_sht_db_folio_description As Integer, _
    c_sht_db_folio_isin As Integer, c_sht_db_folio_sedol As Integer, c_sht_db_folio_ticker As Integer, c_sht_db_folio_ric As Integer, _
    c_sht_db_folio_crncy As Integer


l_sht_db_folio_header = 12

c_sht_db_folio_description = 6
c_sht_db_folio_id = 10
c_sht_db_folio_u_id = 11
c_sht_db_folio_isin = 2
c_sht_db_folio_sedol = 3
c_sht_db_folio_ticker = 4
c_sht_db_folio_ric = 5
c_sht_db_folio_crncy = 9

k = 0
For i = l_sht_db_folio_header + 1 To 32000
    If Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id) = "" And Worksheets("Database_Folio").Cells(i + 1, c_sht_db_folio_id) = "" And Worksheets("Database_Folio").Cells(i + 2, c_sht_db_folio_id) = "" Then
        Exit For
    Else
        If Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id) <> "" And Worksheets("Database_Folio").Cells(i, c_sht_db_folio_description) <> "" Then
            ReDim Preserve vec_db_folio_id(k)
            ReDim Preserve vec_db_folio_isin(k)
            ReDim Preserve vec_db_folio_sedol(k)
            ReDim Preserve vec_db_folio_descritpion(k)
            ReDim Preserve vec_db_folio_ric(k)
            ReDim Preserve vec_db_folio_crncy(k)
            ReDim Preserve vec_db_folio_ticker(k)
            
            vec_db_folio_id(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_id)
            vec_db_folio_isin(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_isin)
            vec_db_folio_sedol(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_sedol)
            vec_db_folio_descritpion(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_description)
            vec_db_folio_ric(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_ric)
            vec_db_folio_crncy(k) = Worksheets("Database_Folio").Cells(i, c_sht_db_folio_crncy)
            vec_db_folio_ticker(k) = Replace(Worksheets("Database_Folio").Cells(i, c_sht_db_folio_ticker), " EQUITY", " Equity")
            
            k = k + 1
            
        End If
    End If
Next i








'remonte l'état actuel de la table equity
Dim extract_equity As Variant
sql_query = "SELECT gs_id FROM t_equity"
extract_equity = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

'crncy
Dim extract_crncy As Variant
sql_query = "SELECT system_code, system_name FROM t_currency"
extract_crncy = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With

Dim found_data As Boolean

With rst
    
    .ActiveConnection = conn
    .Open "t_equity", LockType:=adLockOptimistic

    For i = 0 To UBound(vec_equity, 1)
        
        found_data = False
        
        For j = 0 To UBound(extract_equity, 1)
            If vec_equity(i) = extract_equity(j, 0) Then
                Exit For
            Else
                If j = UBound(extract_equity, 1) Then
                    
                    For k = 1 To UBound(matrix_folio, 1)
                        If vec_equity(i) = matrix_folio(k, dim_folio_id) Then
                            
                            found_data = True
                            
                            .AddNew
                            
                                .fields("gs_id") = vec_equity(i)
                                .fields("gs_isin") = matrix_folio(k, dim_folio_isin)
                                .fields("gs_name") = matrix_folio(k, dim_folio_description)
                                .fields("gs_currency") = UCase(matrix_folio(k, dim_folio_crncy))
                                .fields("system_ticker") = Replace(matrix_folio(k, dim_folio_ticker), " EQUITY", " Equity")
                                
                                
                                For m = 1 To UBound(extract_crncy, 1)
                                    If UCase(matrix_folio(k, dim_folio_crncy)) = extract_crncy(m, 1) Then
                                        .fields("system_currency_code") = extract_crncy(m, 0)
                                        Exit For
                                    End If
                                Next m
                                
                            .Update
                            
                            Exit For
                        End If
                    Next k
                    
                    If found_data = False Then
                        
                        'tentative avec database_folio
                        For k = 0 To UBound(vec_db_folio_id, 1)
                            If vec_db_folio_id(k) = vec_equity(i) Then
                                
                                found_data = True
                                
                                .AddNew
                                
                                    .fields("gs_id") = vec_equity(i)
                                    .fields("gs_isin") = vec_db_folio_isin(k)
                                    .fields("gs_name") = vec_db_folio_descritpion(k)
                                    .fields("gs_currency") = UCase(vec_db_folio_crncy(k))
                                    .fields("system_ticker") = Replace(vec_db_folio_ticker(k), " EQUITY", " Equity")
                                    
                                
                                .Update
                                
                                Exit For
                            End If
                        Next k
                    End If
                    
                End If
            End If
        Next j
    Next i
    
    .Close
    
End With

conn.Close


End Sub




Sub insert_new_option_pict_exec(ByVal vec_option)

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim fileFolio As File_Folio
Set fileFolio = New File_Folio
fileFolio.set_file_path = base_path & "\GS_Folio\" & folio_all_view

Dim matrix_folio As Variant
matrix_folio = fileFolio.get_content_as_a_matrix()


Dim dim_folio_id As Integer, dim_folio_ticker As Integer, dim_folio_crncy As Integer, dim_folio_description As Integer, _
    dim_folio_qty_yesterday_close As Integer, dim_folio_underlying_id As Integer, dim_folio_yesterday_close_price As Integer, _
    dim_folio_product_type As Integer, dim_folio_option_strike As Integer, dim_folio_option_put_call As Integer, _
    dim_folio_option_contract_size As Integer, dim_folio_option_expiration_date As Integer, dim_folio_option_exercise_style As Integer, _
    dim_folio_isin As Integer



'detect les dimensions
For i = 1 To UBound(matrix_folio, 2)
    If matrix_folio(0, i) = "Identifier" Then
        dim_folio_id = i
    ElseIf matrix_folio(0, i) = "Market Data Symbol" Then
        dim_folio_ticker = i
    ElseIf matrix_folio(0, i) = "CCY" Then
        dim_folio_crncy = i
    ElseIf matrix_folio(0, i) = "Description" Then
        dim_folio_description = i
    ElseIf matrix_folio(0, i) = "Qty - Yesterday's Close" Then
        dim_folio_qty_yesterday_close = i
    ElseIf matrix_folio(0, i) = "Underlyer Product ID" Then
        dim_folio_underlying_id = i
    ElseIf matrix_folio(0, i) = "Yesterday's Close (Local)" Then
        dim_folio_yesterday_close_price = i
    ElseIf matrix_folio(0, i) = "Product Type" Then
        dim_folio_product_type = i
    ElseIf matrix_folio(0, i) = "Strike Price" Then
        dim_folio_option_strike = i
    ElseIf matrix_folio(0, i) = "Put/Call" Then
        dim_folio_option_put_call = i
    ElseIf matrix_folio(0, i) = "Contract Size" Then
        dim_folio_option_contract_size = i
    ElseIf matrix_folio(0, i) = "Expiration Date" Then
        dim_folio_option_expiration_date = i
    ElseIf matrix_folio(0, i) = "Option Exercise Style" Then
        dim_folio_option_exercise_style = i
    ElseIf matrix_folio(0, i) = "ISIN" Then
        dim_folio_isin = i
    End If
Next i


'database_folio

'remonte l'état actuel de la table equity
Dim extract_equity As Variant
sql_query = "SELECT gs_id FROM t_option"
extract_equity = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

'crncy
Dim extract_crncy As Variant
sql_query = "SELECT system_code, system_name FROM t_currency"
extract_crncy = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With

Dim found_data As Boolean
Dim date_tmp As Date

With rst
    
    .ActiveConnection = conn
    .Open "t_option", LockType:=adLockOptimistic

    For i = 0 To UBound(vec_option, 1)
        
        found_data = False
        
        For j = 0 To UBound(extract_equity, 1)
            If vec_option(i) = extract_equity(j, 0) Then
                Exit For
            Else
                If j = UBound(extract_equity, 1) Then
                    
                    For k = 1 To UBound(matrix_folio, 1)
                        If vec_option(i) = matrix_folio(k, dim_folio_id) Then
                            
                            found_data = True
                            
                            .AddNew
                            
                                .fields("gs_id") = vec_option(i)
                                .fields("gs_name") = matrix_folio(k, dim_folio_description)
                                .fields("bbg_ticker") = Replace(matrix_folio(k, dim_folio_ticker), " EQUITY", " Equity")
                                .fields("bbg_strike") = matrix_folio(k, dim_folio_option_strike)
                                
                                .fields("bbg_expiry") = matrix_folio(k, dim_folio_option_expiration_date)
                                    .fields("bbg_expiry_day") = day(matrix_folio(k, dim_folio_option_expiration_date))
                                    .fields("bbg_expiry_month") = Month(matrix_folio(k, dim_folio_option_expiration_date))
                                    .fields("bbg_expiry_year") = year(matrix_folio(k, dim_folio_option_expiration_date))
                                
                                .fields("bbg_exercise_type") = matrix_folio(k, dim_folio_option_exercise_style)
                                .fields("bbg_put_call") = matrix_folio(k, dim_folio_option_put_call)
                                .fields("bbg_option_contract_size") = matrix_folio(k, dim_folio_option_contract_size)
                                
                                
                                
                            .Update
                            
                            Exit For
                        End If
                    Next k
                    
                    If found_data = False Then
                        'tentative avec database_folio
                        debug_test = "test"
                    End If
                    
                End If
            End If
        Next j
    Next i
    
    .Close
    
End With

conn.Close


End Sub




Sub sync_and_update_db()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim sql_query As String

Dim extract_instrument As Variant
sql_query = "SELECT system_id, system_table FROM t_instrument"
extract_instrument = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_bridge As Variant
sql_query = "SELECT gs_id, system_instrument_id FROM t_bridge"
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_instrument_entry As Variant
ReDim extract_instrument_entry(UBound(extract_instrument, 1))

For i = 1 To UBound(extract_instrument, 1)
    sql_query = "SELECT gs_id FROM " & extract_instrument(i, 1)
    extract_instrument_entry(i) = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
Next i


Dim vec_product_no_found_id() As Variant
Dim vec_product_no_found_instru_code() As Variant

m = 0
ReDim vec_product_no_found_id(m)
Dim found_entry As Boolean
For i = 1 To UBound(extract_bridge, 1)
    
    found_entry = False
    
    For j = 1 To UBound(extract_instrument, 1)
        If extract_bridge(i, 1) = extract_instrument(j, 0) Then
            
            For k = 1 To UBound(extract_instrument_entry(j))
                If extract_instrument_entry(j)(k, 0) = extract_bridge(i, 0) Then
                    found_entry = True
                    Exit For
                End If
            Next k
            
            If found_entry = False Then
                For k = 0 To UBound(vec_product_no_found_id, 1)
                    If vec_product_no_found_id(k) = extract_bridge(i, 0) Then
                        Exit For
                    Else
                        If k = UBound(vec_product_no_found_id, 1) Then
                            ReDim Preserve vec_product_no_found_id(m)
                            ReDim Preserve vec_product_no_found_instru_code(m)
                            
                            vec_product_no_found_id(m) = extract_bridge(i, 0)
                            vec_product_no_found_instru_code(m) = extract_bridge(i, 1)
                            
                            m = m + 1
                        End If
                    End If
                Next k
            End If
            
            Exit For
        End If
    Next j
Next i




Dim vec_new_entry() As Variant

For i = 1 To UBound(extract_instrument, 1)

    k = 0
    ReDim vec_new_entry(k)
    
    For j = 0 To UBound(vec_product_no_found_id, 1)
        If extract_instrument(i, 0) = vec_product_no_found_instru_code(j) Then
            ReDim Preserve vec_new_entry(k)
            vec_new_entry(k) = vec_product_no_found_id(j)
            k = k + 1
        End If
    Next j
    
    
    'insertion dans les tables
    If k > 0 Then
        
        Debug.Print "Instruement_id=" & i
        For j = 0 To UBound(vec_new_entry, 1)
            
            Debug.Print vec_new_entry(j)
        Next j
        
        If InStr(UCase(extract_instrument(i, 1)), "EQUITY") <> 0 Then
            Call insert_new_equity_pict_exec(vec_new_entry)
        ElseIf InStr(UCase(extract_instrument(i, 1)), "FUTURE") <> 0 Then
            Call insert_new_future_pict_exec(vec_new_entry)
        ElseIf InStr(UCase(extract_instrument(i, 1)), "OPTION") <> 0 Then
            Call insert_new_option_pict_exec(vec_new_entry)
        End If
        
    End If
Next i


MsgBox ("terminé")

End Sub


Sub mgmt_expiry_from_trades(ByVal vec_trades_lines)

Application.Calculation = xlCalculationManual

Dim sql_query As String

Dim debug_test As Variant
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, u As Integer

Dim l_kronos_trades_header As Integer

l_kronos_trades_header = 16


'remonte les colonnes
Dim c_kronos_trades_product_id As Integer, c_kronos_trades_underlying_id As Integer, c_kronos_trades_qty As Integer, _
    c_kronos_trades_is_exec As Integer, c_kronos_trades_settlement_price As Integer, c_kronos_trades_crncy As Integer, _
    c_kronos_trades_derivative_id As Integer, c_kronos_trades_derivative_qty_to_close As Integer
    


c_kronos_trades_product_id = 0
c_kronos_trades_underlying_id = 0
c_kronos_trades_qty = 0
c_kronos_trades_is_exec = 0
c_kronos_trades_settlement_price = 0
c_kronos_trades_crncy = 0
c_kronos_trades_derivative_id = 0 '32
c_kronos_trades_derivative_qty_to_close = 0 '33

For i = 1 To 250
    If Worksheets("Trades").Cells(l_kronos_trades_header, i) = "" And Worksheets("Trades").Cells(l_kronos_trades_header, i + 1) = "" And Worksheets("Trades").Cells(l_kronos_trades_header, i + 2) = "" And Worksheets("Trades").Cells(l_kronos_trades_header, i + 3) = "" And Worksheets("Trades").Cells(l_kronos_trades_header, i + 4) = "" And Worksheets("Trades").Cells(l_kronos_trades_header, i + 5) = "" Then
        Exit For
    Else
        If Worksheets("Trades").Cells(l_kronos_trades_header, i) = "Identifier" And c_kronos_trades_underlying_id = 0 Then
            c_kronos_trades_underlying_id = i
        ElseIf Worksheets("Trades").Cells(l_kronos_trades_header, i) = "Identifier" And c_kronos_trades_underlying_id <> 0 Then
            c_kronos_trades_product_id = i
        ElseIf Worksheets("Trades").Cells(l_kronos_trades_header, i) = "#" And c_kronos_trades_qty = 0 Then
            c_kronos_trades_qty = i
        ElseIf Worksheets("Trades").Cells(l_kronos_trades_header, i) = "Exercice" Then
            c_kronos_trades_is_exec = i
        ElseIf Worksheets("Trades").Cells(l_kronos_trades_header, i) = "Price" Then
            c_kronos_trades_settlement_price = i
        ElseIf Worksheets("Trades").Cells(l_kronos_trades_header, i) = "Currency" Then
            c_kronos_trades_crncy = i
        ElseIf Worksheets("Trades").Cells(l_kronos_trades_header, i) = "Option_Id" Then
            c_kronos_trades_derivative_id = i
        ElseIf Worksheets("Trades").Cells(l_kronos_trades_header, i) = "Qty Option" Then
            c_kronos_trades_derivative_qty_to_close = i
        End If
    End If
Next i



'mount les details de la ligne courante
Dim product_id As String
Dim underlying_id As String
Dim new_postion_to_add As Double
Dim settlement_price As Double
Dim currency_code As Integer

Dim derivative_id As String
Dim derivative_qty_to_close As Double


Dim matrix_trades() As Variant
ReDim matrix_trades(UBound(vec_trades_lines), 10)
    Dim dim_matrix_trades_id As Integer, dim_matrix_trades_u_id As Integer, dim_matrix_trades_new_pos_size As Integer, _
        dim_matrix_trades_new_pos_price As Integer, dim_matrix_trades_derivatives_id As Integer, _
        dim_matrix_trades_derivatives_qty As Integer, dim_matrix_trades_crncy_code As Integer
        
    dim_matrix_trades_id = 0
    dim_matrix_trades_u_id = 1
    dim_matrix_trades_crncy_code = 2
    dim_matrix_trades_new_pos_size = 3
    dim_matrix_trades_new_pos_price = 4
    dim_matrix_trades_derivatives_id = 5
    dim_matrix_trades_derivatives_qty = 6



For i = 0 To UBound(vec_trades_lines, 1)
    matrix_trades(i, dim_matrix_trades_id) = Worksheets("Trades").Cells(vec_trades_lines(i), c_kronos_trades_product_id)
    matrix_trades(i, dim_matrix_trades_u_id) = Worksheets("Trades").Cells(vec_trades_lines(i), c_kronos_trades_underlying_id)
    matrix_trades(i, dim_matrix_trades_crncy_code) = Worksheets("Trades").Cells(vec_trades_lines(i), c_kronos_trades_crncy)
    
    matrix_trades(i, dim_matrix_trades_new_pos_size) = Worksheets("Trades").Cells(vec_trades_lines(i), c_kronos_trades_qty)
    matrix_trades(i, dim_matrix_trades_new_pos_price) = Worksheets("Trades").Cells(vec_trades_lines(i), c_kronos_trades_settlement_price)
    
    matrix_trades(i, dim_matrix_trades_derivatives_id) = Worksheets("Trades").Cells(vec_trades_lines(i), c_kronos_trades_derivative_id)
    matrix_trades(i, dim_matrix_trades_derivatives_qty) = -Worksheets("Trades").Cells(vec_trades_lines(i), c_kronos_trades_derivative_qty_to_close)
    
Next i


'charge les données du dérivé de la base de données
Dim extract_instrument As Variant
sql_query = "SELECT system_id, system_name  FROM t_instrument"
extract_instrument = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


Dim extract_bridge As Variant
sql_query = "SELECT gs_id, gs_underlying_id, system_instrument_id FROM t_bridge"
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_trader As Variant
sql_query = "SELECT system_code, system_first_name, system_surname, gs_UserID FROM t_trader"
extract_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_trading_account As Variant
sql_query = "SELECT gs_account_number, system_trader_code, system_broker_code FROM t_trading_account WHERE gs_main_account = TRUE"
extract_trading_account = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset

With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With


Dim initial_trader As String, trader_txt As String, trader_code As Integer, trader_gs_user_id As String, trading_account As String, _
    broker_code As Integer

Dim is_code_12 As Boolean, not_found_in_bridge As Boolean

Dim vec_equity_product_not_found() As Variant
Dim count_no_entry_in_bridge As Integer
count_no_entry_in_bridge = 0
ReDim Preserve vec_equity_product_not_found(count_no_entry_in_bridge)

With rst
    
    .ActiveConnection = conn
    .Open "t_trade", LockType:=adLockOptimistic
    
    For u = 0 To UBound(matrix_trades, 1)
        
        is_code_12 = False
        not_found_in_bridge = False
        
        product_id = matrix_trades(u, dim_matrix_trades_id)
        underlying_id = matrix_trades(u, dim_matrix_trades_u_id)
        currency_code = matrix_trades(u, dim_matrix_trades_crncy_code)
        
        new_postion_to_add = matrix_trades(u, dim_matrix_trades_new_pos_size)
        settlement_price = matrix_trades(u, dim_matrix_trades_new_pos_price)
        
        derivative_id = matrix_trades(u, dim_matrix_trades_derivatives_id)
        derivative_qty_to_close = matrix_trades(u, dim_matrix_trades_derivatives_qty)
        
        
            If derivative_id = "" Then 'FUT dans le fixing
                
                For i = 1 To UBound(extract_bridge, 1)
                    If extract_bridge(i, 0) = product_id Then
                        
                        not_found_in_bridge = False
                        
                        For j = 1 To UBound(extract_instrument, 1)
                            If extract_bridge(i, 2) = extract_instrument(j, 0) Then
                                
                                'traitement si FUT
                                If extract_instrument(j, 1) = "future" Then
                                    
                                    'insertion d'une pos de cloture dans trades, la netpos->0 et le pnl prend en compte le fixing
                                    'Select Case MsgBox("Close the actual postion with " & new_postion_to_add & "@" & settlement_price & " Cash Flow=" & -new_postion_to_add * settlement_price, vbYesNo, "Close position")
                                        
                                        'Case vbYes
                                            
                                            'essai d'attraper le nom du trader dans GVA report
                                            trader_txt = Worksheets("Cointrin").Cells(5, 2)
                                            
                                            If trader_txt = "" Then
                                                initial_trader = UCase(InputBox("Trader : LA/AM/JS"))
                                            Else
                                                initial_trader = Left(trader_txt, 1)
                                                initial_trader = initial_trader & Mid(trader_txt, InStr(trader_txt, " ") + 1, 1)
                                            End If
                                            
                                            For k = 1 To UBound(extract_trader, 1)
                                                If Left(initial_trader, 1) = Left(extract_trader(k, 1), 1) And Right(initial_trader, 1) = Left(extract_trader(k, 2), 1) Then
                                                    
                                                    trader_code = extract_trader(k, 0)
                                                    trader_gs_user_id = extract_trader(k, 3)
                                                    
                                                    'remonte main account
                                                    For m = 1 To UBound(extract_trading_account, 1)
                                                        If extract_trading_account(m, 1) = trader_code Then
                                                            trading_account = extract_trading_account(m, 0)
                                                            broker_code = extract_trading_account(m, 2)
                                                            Exit For
                                                        End If
                                                    Next m
                                                    
                                                    Exit For
                                                End If
                                            Next k
                                            
                                            
                                            .AddNew
                                                
                                                .fields("gs_date") = Date
                                                .fields("gs_time") = Time
                                                
                                                .fields("gs_unique_id") = "internal_exec_fixing_fut" & product_id & "_" & year(Date) & Month(Date) & day(Date) & Hour(Time) & Minute(Time) & Second(Time)
                                                '.fields("gs_unique_id") = "internal_exec_fixing_fut_" & product_id
                                                
                                                .fields("gs_security_id") = product_id
                                                
                                                .fields("gs_exec_qty") = new_postion_to_add
                                                .fields("gs_exec_price") = settlement_price
                                                .fields("gs_order_qty") = new_postion_to_add
                                                
                                                If new_postion_to_add > 0 Then
                                                    .fields("gs_side") = "B"
                                                    .fields("gs_side_detailed") = "B"
                                                Else
                                                    .fields("gs_side") = "S"
                                                    .fields("gs_side_detailed") = "S"
                                                End If
                                                
                                                .fields("gs_trading_account") = trading_account
                                                .fields("gs_user_id") = trader_gs_user_id
                                                
                                                .fields("gs_close_price") = settlement_price
                                                
                                                .fields("system_ytd_pnl_reversal") = 0
                                                .fields("system_position_reversal") = 0
                                                .fields("system_comm_reversal") = 0
                                                
                                                .fields("system_comm_reversal") = 0
                                                
                                                .fields("system_currency_code") = currency_code
                                                
                                                .fields("system_broker_id") = broker_code
                                                
                                                .fields("system_commission_local_currency") = 0
                                                
                                                .fields("system_trader_code") = trader_code
                                                
                                                .fields("system_exercise") = True
                                                
                                            
                                            .Update
                                            
                                            Exit For
                                            
                                        'Case vbNo
                                            'Exit Sub
                                        
                                    'End Select
                                    
                                End If
                            End If
                        Next j
                        
                        Exit For
                    End If
                Next i
            
            Else 'derivatives
                
                For i = 1 To UBound(extract_bridge, 1)
                    If extract_bridge(i, 0) = product_id Then
                        For j = 1 To UBound(extract_instrument, 1)
                            If extract_bridge(i, 2) = extract_instrument(j, 0) Then
                                
processing_expiry_deriv_equity:
                                
                                'Select Case MsgBox("Close the derivative position ?", vbYesNo, "Close position")
                                        
                                    'Case vbYes
                                        
                                        'essai d'attraper le nom du trader dans GVA report
                                        trader_txt = Worksheets("Cointrin").Cells(5, 2)
                                        
                                        If trader_txt = "" Then
                                            initial_trader = UCase(InputBox("Trader : LA/AM/JS"))
                                        Else
                                            initial_trader = Left(trader_txt, 1)
                                            initial_trader = initial_trader & Mid(trader_txt, InStr(trader_txt, " ") + 1, 1)
                                        End If
                                        
                                        For k = 1 To UBound(extract_trader, 1)
                                            If Left(initial_trader, 1) = Left(extract_trader(k, 1), 1) And Right(initial_trader, 1) = Left(extract_trader(k, 2), 1) Then
                                                
                                                trader_code = extract_trader(k, 0)
                                                trader_gs_user_id = extract_trader(k, 3)
                                                
                                                'remonte main account
                                                For m = 1 To UBound(extract_trading_account, 1)
                                                    If extract_trading_account(m, 1) = trader_code Then
                                                        trading_account = extract_trading_account(m, 0)
                                                        broker_code = extract_trading_account(m, 2)
                                                        Exit For
                                                    End If
                                                Next m
                                                
                                                Exit For
                                            End If
                                        Next k
                                    'Case vbNo
                                        'Exit Sub
                                'End Select
                                
                                'ajuste la pos du dérivé
                                If extract_instrument(j, 1) = "equity" Or is_code_12 = True Or not_found_in_bridge = True Then 'option sur equity
                                    If new_postion_to_add <> 0 Then
                                        
                                        .AddNew
                                        
                                            .fields("gs_date") = Date
                                            .fields("gs_time") = Time
                                            
                                            '.fields("gs_unique_id") = "internal_exec_derivative_" & product_id
                                            .fields("gs_unique_id") = "internal_exec_derivative_" & product_id & "_" & derivative_id & "_" & year(Date) & Month(Date) & day(Date) & Hour(Time) & Minute(Time) & Second(Time)
                                            .fields("gs_security_id") = product_id
                                            
                                            .fields("gs_exec_qty") = new_postion_to_add
                                            .fields("gs_exec_price") = settlement_price
                                            .fields("gs_order_qty") = new_postion_to_add
                                            
                                            If new_postion_to_add > 0 Then
                                                .fields("gs_side") = "B"
                                                .fields("gs_side_detailed") = "B"
                                            Else
                                                .fields("gs_side") = "S"
                                                .fields("gs_side_detailed") = "S"
                                            End If
                                            
                                            .fields("gs_trading_account") = trading_account
                                            .fields("gs_user_id") = trader_gs_user_id
                                            
                                            .fields("gs_close_price") = settlement_price
                                            
                                            .fields("system_ytd_pnl_reversal") = 0
                                            .fields("system_position_reversal") = 0
                                            .fields("system_comm_reversal") = 0
                                            
                                            .fields("system_comm_reversal") = 0
                                            
                                            .fields("system_currency_code") = currency_code
                                            
                                            .fields("system_broker_id") = broker_code
                                            
                                            .fields("system_commission_local_currency") = 0
                                            
                                            .fields("system_trader_code") = trader_code
                                            
                                            .fields("system_exercise") = True
                                        
                                        .Update
                                    End If
                                ElseIf extract_instrument(j, 1) = "index" Then 'option sur index
                                    'cash settlement... creation d'un pnl sur l'index
                                    
                                End If
                                    
                                    
                                'cloture de la pos d'option (partial or complete)
                                .AddNew
                                
                                    .fields("gs_date") = Date
                                    .fields("gs_time") = Time
                                    
                                    '.fields("gs_unique_id") = "internal_exec_derivative_" & derivative_id
                                    .fields("gs_unique_id") = "internal_exec_derivative_" & derivative_id & product_id & "_" & year(Date) & Month(Date) & day(Date) & Hour(Time) & Minute(Time) & Second(Time) & "_" & Rnd()
                                    
                                    .fields("gs_security_id") = derivative_id
                                    
                                    .fields("gs_exec_qty") = derivative_qty_to_close
                                    .fields("gs_exec_price") = 0
                                    .fields("gs_order_qty") = derivative_qty_to_close
                                    
                                    If derivative_qty_to_close > 0 Then
                                        .fields("gs_side") = "B"
                                        .fields("gs_side_detailed") = "B"
                                    Else
                                        .fields("gs_side") = "S"
                                        .fields("gs_side_detailed") = "S"
                                    End If
                                    
                                    .fields("gs_trading_account") = trading_account
                                    .fields("gs_user_id") = trader_gs_user_id
                                    
                                    .fields("gs_close_price") = 0
                                    
                                    .fields("system_ytd_pnl_reversal") = 0
                                    .fields("system_position_reversal") = 0
                                    .fields("system_comm_reversal") = 0
                                    
                                    .fields("system_comm_reversal") = 0
                                    
                                    .fields("system_currency_code") = currency_code
                                    
                                    .fields("system_broker_id") = broker_code
                                    
                                    .fields("system_commission_local_currency") = 0
                                    
                                    .fields("system_trader_code") = trader_code
                                    
                                    .fields("system_exercise") = True
                                
                                .Update
                                    
                                Exit For
                            End If
                            
                        Next j
                        
                        Exit For
                    Else
                        If i = UBound(extract_bridge, 1) Then
                            
                            'retrouve le main pnl product id
                            For j = 27 To 32000 Step 2
                                If Worksheets("Equity_Database").Cells(j, 1) = "" And Worksheets("Equity_Database").Cells(j + 2, 1) = "" And Worksheets("Equity_Database").Cells(j + 4, 1) = "" Then
                                    Exit For
                                Else
                                    If Worksheets("Equity_Database").Cells(j, 1) = product_id Then
                                    
                                        If Worksheets("Equity_Database").Cells(j, 4) = 12 Or Worksheets("Equity_Database").Cells(j, 4) = 13 Then
                                            'code 12
                                            is_code_12 = True
                                            
                                            product_id = Worksheets("Equity_Database").Cells(j, 97)
                                            j = 0
                                            GoTo processing_expiry_deriv_equity
                                            
                                        Else
                                            'la valeur n'a pas été ouverte dans le pont
                                            j = 0
                                            not_found_in_bridge = True
                                            
                                            If new_postion_to_add <> 0 Then
                                                For k = 0 To UBound(vec_equity_product_not_found, 1)
                                                    If vec_equity_product_not_found(k) = product_id Then
                                                        Exit For
                                                    Else
                                                        If k = UBound(vec_equity_product_not_found, 1) Then
                                                            ReDim Preserve vec_equity_product_not_found(count_no_entry_in_bridge)
                                                            vec_equity_product_not_found(count_no_entry_in_bridge) = product_id
                                                            
                                                            count_no_entry_in_bridge = count_no_entry_in_bridge + 1
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                            
                                            
                                            GoTo processing_expiry_deriv_equity
                                        End If
                                        
                                        Exit For
                                    End If
                                End If
                            Next j
                        End If
                    End If
                Next i
            End If
        Next u
    .Close
End With

'update du bridge si necessaire
If count_no_entry_in_bridge > 0 Then
    Call update_bridge_pict_exec(vec_equity_product_not_found)
End If

MsgBox ("terminé")

End Sub


Sub mgmt_expiry_from_exe(ByVal vec_exe_lines)

Application.Calculation = xlCalculationManual

Dim sql_query As String

Dim debug_test As Variant
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, u As Integer

Dim l_kronos_exe_header As Integer

l_kronos_exe_header = 16


'remonte les colonnes
Dim c_kronos_exe_product_id As Integer, c_kronos_exe_underlying_id As Integer, c_kronos_exe_qty As Integer, _
    c_kronos_exe_is_exec As Integer, c_kronos_exe_settlement_price As Integer, c_kronos_exe_crncy As Integer, _
    c_kronos_exe_derivative_id As Integer, c_kronos_exe_derivative_qty_to_close As Integer, c_kronos_exe_price_factor As Integer, _
    c_kronos_exe_option_strike As Integer, c_kronos_exe_price_underlying As Integer, c_kronos_exe_money_underlying As Integer
    


c_kronos_exe_product_id = 0
c_kronos_exe_underlying_id = 0
c_kronos_exe_qty = 0
c_kronos_exe_is_exec = 0
c_kronos_exe_settlement_price = 0
c_kronos_exe_crncy = 0
c_kronos_exe_derivative_id = 0 '32
c_kronos_exe_derivative_qty_to_close = 0 '33
c_kronos_exe_price_factor = 0
c_kronos_exe_option_strike = 0
c_kronos_exe_price_underlying = 0
c_kronos_exe_money_underlying = 0

For i = 1 To 250
    If Worksheets("Exe").Cells(l_kronos_exe_header, i) = "" And Worksheets("Exe").Cells(l_kronos_exe_header, i + 1) = "" And Worksheets("Exe").Cells(l_kronos_exe_header, i + 2) = "" And Worksheets("Exe").Cells(l_kronos_exe_header, i + 3) = "" And Worksheets("Exe").Cells(l_kronos_exe_header, i + 4) = "" And Worksheets("Exe").Cells(l_kronos_exe_header, i + 5) = "" Then
        Exit For
    Else
        If Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Identifier" And c_kronos_exe_underlying_id = 0 Then
            c_kronos_exe_underlying_id = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Identifier" And c_kronos_exe_underlying_id <> 0 Then
            c_kronos_exe_product_id = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Currency" Then
            c_kronos_exe_crncy = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Option_Id" Then
            c_kronos_exe_derivative_id = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Qty Option" Then
            c_kronos_exe_derivative_qty_to_close = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Price_underlying" Then
            c_kronos_exe_price_underlying = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Price Factor" Then
            c_kronos_exe_price_factor = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Strike" Then
            c_kronos_exe_option_strike = i
        ElseIf Worksheets("Exe").Cells(l_kronos_exe_header, i) = "Money" Then
            c_kronos_exe_money_underlying = i
        End If
    End If
Next i



'mount les details de la ligne courante
Dim product_id As String
Dim underlying_id As String
Dim new_postion_to_add As Double
Dim settlement_price As Double
Dim currency_code As Integer

Dim derivative_id As String
Dim derivative_qty_to_close As Double


Dim matrix_trades() As Variant
ReDim matrix_trades(UBound(vec_exe_lines), 15)
    Dim dim_matrix_trades_id As Integer, dim_matrix_trades_u_id As Integer, dim_matrix_trades_new_pos_size As Integer, _
        dim_matrix_trades_new_pos_price As Integer, dim_matrix_trades_derivatives_id As Integer, _
        dim_matrix_trades_derivatives_qty As Integer, dim_matrix_trades_crncy_code As Integer, dim_matrix_trades_derivatives_cash_settlement As Integer
        
    dim_matrix_trades_id = 0
    dim_matrix_trades_u_id = 1
    dim_matrix_trades_crncy_code = 2
    dim_matrix_trades_new_pos_size = 3
    dim_matrix_trades_new_pos_price = 4
    dim_matrix_trades_derivatives_id = 5
    dim_matrix_trades_derivatives_qty = 6
    dim_matrix_trades_derivatives_cash_settlement = 7
    dim_matrix_trades_derivatives_price_factor = 8
    dim_matrix_trades_derivatives_strike = 9
    dim_matrix_trades_derivatives_money = 10
    



For i = 0 To UBound(vec_exe_lines, 1)
    matrix_trades(i, dim_matrix_trades_id) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_product_id)
    matrix_trades(i, dim_matrix_trades_u_id) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_underlying_id)
    matrix_trades(i, dim_matrix_trades_crncy_code) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_crncy)
    
    matrix_trades(i, dim_matrix_trades_derivatives_id) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_derivative_id)
    matrix_trades(i, dim_matrix_trades_derivatives_qty) = -Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_derivative_qty_to_close)
    matrix_trades(i, dim_matrix_trades_derivatives_cash_settlement) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_price_underlying)
    matrix_trades(i, dim_matrix_trades_derivatives_price_factor) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_price_factor)
    matrix_trades(i, dim_matrix_trades_derivatives_strike) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_option_strike)
    matrix_trades(i, dim_matrix_trades_derivatives_money) = Worksheets("Exe").Cells(vec_exe_lines(i), c_kronos_exe_money_underlying)
Next i


'charge les données du dérivé de la base de données
Dim extract_instrument As Variant
sql_query = "SELECT system_id, system_name  FROM t_instrument"
extract_instrument = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


Dim extract_bridge As Variant
sql_query = "SELECT gs_id, gs_underlying_id, system_instrument_id FROM t_bridge"
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_trader As Variant
sql_query = "SELECT system_code, system_first_name, system_surname, gs_UserID FROM t_trader"
extract_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_trading_account As Variant
sql_query = "SELECT gs_account_number, system_trader_code, system_broker_code FROM t_trading_account WHERE gs_main_account = TRUE"
extract_trading_account = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset

With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With


Dim initial_trader As String, trader_txt As String, trader_code As Integer, trader_gs_user_id As String, trading_account As String, _
    broker_code As Integer, derivative_price_factor As Double, derivative_strike As Double, derivative_money As String, _
    derivative_price_settlement As Double



With rst
    
    .ActiveConnection = conn
    .Open "t_trade", LockType:=adLockOptimistic
    
    For u = 0 To UBound(matrix_trades, 1)
        
        product_id = matrix_trades(u, dim_matrix_trades_id)
        underlying_id = matrix_trades(u, dim_matrix_trades_u_id)
        currency_code = matrix_trades(u, dim_matrix_trades_crncy_code)
        
        
        derivative_id = matrix_trades(u, dim_matrix_trades_derivatives_id)
        derivative_qty_to_close = matrix_trades(u, dim_matrix_trades_derivatives_qty)
        derivative_price_settlement = matrix_trades(u, dim_matrix_trades_derivatives_cash_settlement)
        derivative_price_factor = matrix_trades(u, dim_matrix_trades_derivatives_price_factor)
        derivative_strike = matrix_trades(u, dim_matrix_trades_derivatives_strike)
        derivative_money = matrix_trades(u, dim_matrix_trades_derivatives_money)
        
            If derivative_id = "" Then 'FUT dans le fixing
            
            Else 'derivatives
                
                For i = 1 To UBound(extract_bridge, 1)
                    If extract_bridge(i, 0) = underlying_id Then
                        For j = 1 To UBound(extract_instrument, 1)
                            If extract_bridge(i, 2) = extract_instrument(j, 0) Then
                                
                                
                                'Select Case MsgBox("Close the derivative position ?", vbYesNo, "Close position")
                                        
                                    'Case vbYes
                                        
                                        'essai d'attraper le nom du trader dans GVA report
                                        trader_txt = Worksheets("Cointrin").Cells(5, 2)
                                        
                                        If trader_txt = "" Then
                                            initial_trader = UCase(InputBox("Trader : LA/AM/JS"))
                                        Else
                                            initial_trader = Left(trader_txt, 1)
                                            initial_trader = initial_trader & Mid(trader_txt, InStr(trader_txt, " ") + 1, 1)
                                        End If
                                        
                                        For k = 1 To UBound(extract_trader, 1)
                                            If Left(initial_trader, 1) = Left(extract_trader(k, 1), 1) And Right(initial_trader, 1) = Left(extract_trader(k, 2), 1) Then
                                                
                                                trader_code = extract_trader(k, 0)
                                                trader_gs_user_id = extract_trader(k, 3)
                                                
                                                'remonte main account
                                                For m = 1 To UBound(extract_trading_account, 1)
                                                    If extract_trading_account(m, 1) = trader_code Then
                                                        trading_account = extract_trading_account(m, 0)
                                                        broker_code = extract_trading_account(m, 2)
                                                        Exit For
                                                    End If
                                                Next m
                                                
                                                Exit For
                                            End If
                                        Next k
                                    'Case vbNo
                                        'Exit Sub
                                'End Select
                                
                                'ajuste la pos du dérivé
                                If extract_instrument(j, 1) = "equity" Then 'option sur equity

                                ElseIf extract_instrument(j, 1) = "index" Then 'option sur index
                                    'cash settlement... creation d'un pnl sur l'index
'                                    If derivative_money = "ATM" Or derivative_money = "ITM" Then
'                                        .AddNew
'
'                                            .fields("gs_date") = Date
'                                            .fields("gs_time") = Time
'
'                                            .fields("gs_unique_id") = "internal_exec_derivative_" & underlying_id & "_" & derivative_id & "_" & Year(Date) & Month(Date) & Day(Date) & Hour(Time) & Minute(Time) & Second(Time)
'                                            .fields("gs_security_id") = underlying_id
'
'                                            .fields("gs_exec_qty") = 0
'                                            .fields("gs_exec_price") = 0
'                                            .fields("gs_order_qty") = 0
'
'                                            .fields("gs_side") = ""
'                                            .fields("gs_side_detailed") = ""
'
'                                            .fields("gs_trading_account") = trading_account
'                                            .fields("gs_user_id") = trader_gs_user_id
'
'                                            .fields("gs_close_price") = derivative_price_settlement
'
'                                            .fields("system_ytd_pnl_reversal") = -derivative_qty_to_close * derivative_price_factor * (derivative_strike - derivative_price_settlement)
'                                            .fields("system_position_reversal") = 0
'                                            .fields("system_comm_reversal") = 0
'
'                                            .fields("system_comm_reversal") = 0
'
'                                            .fields("system_currency_code") = currency_code
'
'                                            .fields("system_broker_id") = broker_code
'
'                                            .fields("system_commission_local_currency") = 0
'
'                                            .fields("system_trader_code") = trader_code
'
'                                            .fields("system_exercise") = True
'
'                                        .Update
'                                    End If
                                End If
                                    
                                    
                                'cloture de la pos d'option (partial or complete)
                                .AddNew
                                
                                    .fields("gs_date") = Date
                                    .fields("gs_time") = Time
                                    
                                    .fields("gs_unique_id") = "internal_exec_derivative_" & derivative_id & "_" & year(Date) & Month(Date) & day(Date) & Hour(Time) & Minute(Time) & Second(Time)
                                    .fields("gs_security_id") = derivative_id
                                    
                                    .fields("gs_exec_qty") = derivative_qty_to_close
                                    .fields("gs_exec_price") = 0
                                    .fields("gs_order_qty") = derivative_qty_to_close
                                    
                                    If derivative_qty_to_close > 0 Then
                                        .fields("gs_side") = "B"
                                        .fields("gs_side_detailed") = "B"
                                    Else
                                        .fields("gs_side") = "S"
                                        .fields("gs_side_detailed") = "S"
                                    End If
                                    
                                    .fields("gs_trading_account") = trading_account
                                    .fields("gs_user_id") = trader_gs_user_id
                                    
                                    .fields("gs_close_price") = 0
                                    
                                    If derivative_money = "ATM" Or derivative_money = "ITM" Then
                                        .fields("system_ytd_pnl_reversal") = -derivative_qty_to_close * derivative_price_factor * Abs(derivative_strike - derivative_price_settlement)
                                    Else
                                        .fields("system_ytd_pnl_reversal") = 0
                                    End If
                                    
                                    .fields("system_position_reversal") = 0
                                    .fields("system_comm_reversal") = 0
                                    
                                    .fields("system_comm_reversal") = 0
                                    
                                    .fields("system_currency_code") = currency_code
                                    
                                    .fields("system_broker_id") = broker_code
                                    
                                    .fields("system_commission_local_currency") = 0
                                    
                                    .fields("system_trader_code") = trader_code
                                    
                                    .fields("system_exercise") = True
                                
                                .Update
                                    
                                Exit For
                            End If
                            
                        Next j
                        
                        Exit For
                    End If
                Next i
            End If
        Next u
    .Close
End With

MsgBox ("terminé")

End Sub

Sub import_trades_from_redi_plus_intraday_extract(ByVal file_extract_rplus As String, Optional ByVal last_pos_only As Integer, Optional ByVal limit_time As Date, Optional ByVal only_code_20 As Boolean)

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim date_tmp As Date, date_min As Date, date_tmp_1 As Date, date_tmp_2 As Date

Dim debug_test As Variant

Dim file_book As String

Dim l_extract_rplus_header As Integer, l_book_equity_db_header As Integer

Dim l_extract_rplus_last_line As Integer

Dim c_extract_rplus_Time As Integer, c_extract_rplus_Side As Integer, c_extract_rplus_ExeQty As Integer, _
c_extract_rplus_Symbol As Integer, c_extract_rplus_ExecPr As Integer, c_extract_rplus_Orig_Qt As Integer, _
c_extract_rplus_OrdrPr As Integer, c_extract_rplus_OrdExecQty As Integer, c_extract_rplus_Exch As Integer, _
c_extract_rplus_UserSubId As Integer, c_extract_rplus_date As Integer, c_extract_rplus_ProductType As Integer, _
c_extract_rplus_LastMarket As Integer, c_extract_rplus_BrSeq As Integer, c_extract_rplus_Stat As Integer, _
c_extract_rplus_MatchOmsKeyLineSeq As Integer, c_extract_rplus_BID As Integer, c_extract_account_number As Integer

Dim step_next_line As Integer


Dim c_extract_rplus_last_column As Integer


Dim c_book_equity_db_uid As Integer, c_book_equity_db_name As Integer, c_book_equity_db_isin As Integer, _
c_book_equity_db_ticker As Integer, c_book_equity_db_crncy_code As Integer

Dim l_book_equity_db_last_line As Integer



Dim l_book_index_db_header As Integer

Dim c_book_index_db_uid As Integer, c_book_equity_db_status As Integer, c_book_index_db_name As Integer, c_book_index_db_ticker As Integer, _
c_book_index_db_FuturesMaturities_ID_1 As Integer, c_book_index_db_quotite As Integer, c_book_index_db_crncy As Integer, _
c_book_index_db_settlement As Integer, c_book_index_db_status As Integer

Dim l_book_index_db_last_line As Integer


Dim l_book_parameters_crncy_header As Integer

Dim c_book_parameters_crncy_crncy As Integer, c_book_parameters_crncy_code As Integer

Dim error_msg_if_uid_not_found As String, error_color_line As Integer

Dim account_equity As Variant, account_future As Variant

'global cst
l_extract_rplus_header = 1

file_book = ThisWorkbook.name
account_equity = Array(Array("GOLDMAN", "G4364339"), Array("jpm", "G4365014"), Array("jpm", "JPMS ALGO"), Array("DBK", "DB EU ALGO"), Array("jpm", "JPME ALGO"))
account_future = Array(Array("GOLDMAN", "50043176-T"))
error_msg_if_uid_not_found = "NO ENTRY IN DB"
error_color_line = 6
l_book_equity_db_header = 25
    c_book_equity_db_uid = 0
    c_book_equity_db_name = 0
    c_book_equity_db_status = 0
    c_book_equity_db_isin = 0
    c_book_equity_db_ticker = 0
    c_book_equity_db_crncy_code = 0

l_book_index_db_header = 25
    c_book_index_db_uid = 0 '1
    c_book_index_db_name = 0 '2
    c_book_index_db_status = 0
    c_book_index_db_ticker = 0 '34
    c_book_index_db_FuturesMaturities_ID_1 = 0 '31
    c_book_index_db_quotite = 0 '113
    c_book_index_db_crncy = 0 '107
    c_book_index_db_settlement = 0 '33

l_book_parameters_crncy_header = 13
    c_book_parameters_crncy_crncy = 1
    c_book_parameters_crncy_code = 5
    


'list exceptions
Dim dim_exception_ticker As Integer, dim_line_in_equity_db As Integer
Dim matrix_exceptions(25, 10) As Variant
    dim_exception_ticker = 1
    dim_line_in_equity_db = 2
    
    
'    i = 1
'    matrix_exceptions(i, dim_exception_ticker) = "APC GY EQUITY"
'    matrix_exceptions(i, dim_line_in_equity_db) = 12000
'    i = i + 1
'
'    matrix_exceptions(i, dim_exception_ticker) = "AAPL UQ EQUITY"
'    matrix_exceptions(i, dim_line_in_equity_db) = 12500
'    i = i + 1
    
    matrix_exceptions(i, dim_exception_ticker) = "BRK.B (*)"
    matrix_exceptions(i, dim_line_in_equity_db) = 12532
    i = i + 1
    



Application.Calculation = xlCalculationManual


'ouvre la book d'extract r+ si necessaire
Dim check_wrkb As Workbook
Dim FoundWrbk As Boolean
Dim src_path As String

src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))


FoundWrbk = False
For Each check_wrkb In Workbooks
    If check_wrkb.name = file_extract_rplus Then
        FoundWrbk = True
        Exit For
    End If
Next

If FoundWrbk = False Then
    Workbooks.Open filename:=src_path & file_extract_rplus, readOnly:=True
    'Workbooks.Open FileName:="Q:\front\LONGSHOT\Sheet1.xls", readOnly:=True
End If


'last line files
For i = 1 To 32000
    If Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, 1) = "" And Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i + 1, 1) = "" And Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i + 2, 1) = "" Then
        l_extract_rplus_last_line = i - 1
        Exit For
    End If
Next i


'recup des colonnes
For i = 1 To 250
    If Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "" And Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i + 1) = "" And Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i + 2) = "" Then
        c_extract_rplus_last_column = i - 1
        Exit For
    End If
    
    
    If Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "Time" Then
        c_extract_rplus_Time = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "Side" Then
        c_extract_rplus_Side = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "ExeQty" Then
        c_extract_rplus_ExeQty = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "Symbol" Then
        c_extract_rplus_Symbol = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "ExecPr" Then
        c_extract_rplus_ExecPr = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "Orig Qty" Then
        c_extract_rplus_Orig_Qt = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "OrdrPr" Then
        c_extract_rplus_OrdrPr = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "OrdExecQty" Then
        c_extract_rplus_OrdExecQty = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "BID" Then
        c_extract_rplus_BID = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "Exch" Then
        c_extract_rplus_Exch = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "UserSubId" Then
        c_extract_rplus_UserSubId = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "Date" Then
        c_extract_rplus_date = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "ProductType" Then
        c_extract_rplus_ProductType = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "LastMarket" Then
        c_extract_rplus_LastMarket = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "BrSeq" Then
        c_extract_rplus_BrSeq = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "Stat" Then
        c_extract_rplus_Stat = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "MatchOmsKeyLineSeq" Then
        c_extract_rplus_MatchOmsKeyLineSeq = i
    ElseIf Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(l_extract_rplus_header, i) = "AcctNum" Then
        c_extract_account_number = i
    End If
    
Next i


'nettoie la sheet des virgules
For i = 1 To l_extract_rplus_last_line
    For j = 1 To c_extract_rplus_last_column
        If InStr(1, Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j), ",") <> 0 Then
            Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j) = Replace(Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j), ",", "")
        End If
    Next j
Next i


'remonte les trades dans une matrix brute
Dim matrix_brut_extract_trades() As Variant

Dim dim_matrix_brut_traitement As Integer, dim_matrix_brut_nbre_trades As Integer, dim_matrix_brut_cost_total As Integer, _
dim_matrix_brut_cost_mid As Integer, dim_matrix_brut_order_status As Integer, dim_matrix_brut_construct_ticker As Integer, _
dim_matrix_brut_construct_market As Integer, dim_matrix_brut_isin As Integer, dim_matrix_brut_ticker As Integer, _
dim_matrix_brut_uid As Integer, dim_matrix_brut_name As Integer, dim_matrix_brut_crncy_code As Integer, dim_matrix_brut_maturity_id As Integer, _
dim_matrix_brut_quotite As Integer, dim_matrix_brut_settlement As Integer, dim_matrix_brut_construct_found_ticker_in_db As Integer


'column supp pour simplifier le traitement futur
i = 1
dim_matrix_brut_traitement = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_nbre_trades = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_cost_total = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_cost_mid = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_order_status = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_construct_found_ticker_in_db = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_construct_ticker = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_construct_market = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_isin = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_ticker = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_uid = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_name = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_crncy_code = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_maturity_id = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_quotite = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_settlement = c_extract_rplus_last_column + i
i = i + 1

dim_matrix_brut_need_import_in_excel = c_extract_rplus_last_column + i
i = i + 1



Dim date_txt_monthEU As Variant, date_txt_yearEU As Variant, date_txt_dayEU As Variant


ReDim matrix_brut_extract_trades(l_extract_rplus_last_line, c_extract_rplus_last_column + 20)

For i = l_extract_rplus_header To l_extract_rplus_last_line
    
    
    For j = 1 To c_extract_rplus_last_column
        
        'colonne de date
        If j = c_extract_rplus_date And i > l_extract_rplus_header Then
            
            date_txt_monthEU = Left(Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j), 2)

            'année à 4 ou 2 chiffres
            If Len(Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j)) = 10 Then
                date_txt_yearEU = Right(Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j), 4)
            ElseIf Len(Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j)) = 8 Then
                date_txt_yearEU = "20" & Right(Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j), 2)
            End If
            
            date_txt_dayEU = Mid(Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j), 4, 2)
            
            date_tmp = date_txt_dayEU & "." & date_txt_monthEU & "." & date_txt_yearEU
            
            matrix_brut_extract_trades(i, j) = date_tmp
        Else
            matrix_brut_extract_trades(i, j) = Workbooks(file_extract_rplus).Worksheets(Left(file_extract_rplus, InStr(file_extract_rplus, ".") - 1)).Cells(i, j)
        End If
    Next j
    
    
    matrix_brut_extract_trades(i, dim_matrix_brut_traitement) = False
    matrix_brut_extract_trades(i, dim_matrix_brut_nbre_trades) = 1
    matrix_brut_extract_trades(i, dim_matrix_brut_cost_total) = 0
    matrix_brut_extract_trades(i, dim_matrix_brut_cost_mid) = 0
    matrix_brut_extract_trades(i, dim_matrix_brut_construct_found_ticker_in_db) = False
    matrix_brut_extract_trades(i, dim_matrix_brut_need_import_in_excel) = True
    
Next i

'fermeture du fichier d'extract de trades
Workbooks(file_extract_rplus).Close (False)


'count les trades uniques
k = 1
Dim vec_trades() As String
Dim found_id As Boolean

ReDim Preserve vec_trades(k)
For i = 2 To UBound(matrix_brut_extract_trades, 1)
    found_id = False

    For j = 0 To UBound(vec_trades, 1)
        If vec_trades(j) = matrix_brut_extract_trades(i, c_extract_rplus_ProductType) & "_" & matrix_brut_extract_trades(i, c_extract_rplus_MatchOmsKeyLineSeq) & "_" & matrix_brut_extract_trades(i, c_extract_rplus_BID) Then
           found_id = True
           Exit For
        End If
    Next j
    
    
    If found_id = False Then
        ReDim Preserve vec_trades(k)
        vec_trades(k) = matrix_brut_extract_trades(i, c_extract_rplus_ProductType) & "_" & matrix_brut_extract_trades(i, c_extract_rplus_MatchOmsKeyLineSeq) & "_" & matrix_brut_extract_trades(i, c_extract_rplus_BID)
        k = k + 1
    End If

Next i

Dim nbre_unique_trades As Integer
nbre_unique_trades = k - 1 'ou ubound(vec_trades,1)



'construction de la matrix des trades unique
Dim matrix_trades() As Variant
ReDim matrix_trades(nbre_unique_trades, UBound(matrix_brut_extract_trades, 2) + 20)


'regroupe et aggrege les trades similaires
k = 1
For i = 2 To UBound(matrix_brut_extract_trades, 1)

    If matrix_brut_extract_trades(i, dim_matrix_brut_traitement) = False Then
        'insertion de la ligne
        For j = 1 To UBound(matrix_brut_extract_trades, 2)
            matrix_trades(k, j) = matrix_brut_extract_trades(i, j)
        Next j
        
        
        'calcul du cout d'achat / de la vente
        matrix_trades(k, dim_matrix_brut_cost_total) = matrix_trades(k, c_extract_rplus_ExeQty) * matrix_trades(k, c_extract_rplus_ExecPr)
        
        
        'ticker
        If UCase(matrix_trades(k, c_extract_rplus_ProductType)) = "STOCK" Then
            
            If InStr(matrix_trades(k, c_extract_rplus_BID), ".") <> 0 Then
                matrix_trades(k, dim_matrix_brut_construct_ticker) = Replace(matrix_trades(k, c_extract_rplus_BID), ".", " ") & " Equity"
                matrix_trades(k, dim_matrix_brut_construct_market) = Mid(matrix_trades(k, c_extract_rplus_BID), (InStr(matrix_trades(k, c_extract_rplus_BID), ".")) + 1, 2)
            Else
                matrix_trades(k, dim_matrix_brut_construct_ticker) = matrix_trades(k, c_extract_rplus_BID) & " US Equity"
                matrix_trades(k, dim_matrix_brut_construct_market) = "US"
            End If
        End If
        
        
        matrix_brut_extract_trades(i, dim_matrix_brut_traitement) = True
        
        
        'repere toutes les autres lignes qui appartiennent a ce trade
        For j = i + 1 To UBound(matrix_brut_extract_trades, 1)
            
            'If matrix_trades(k, c_extract_rplus_Symbol) = matrix_brut_extract_trades(j, c_extract_rplus_Symbol) And matrix_trades(k, c_extract_rplus_ProductType) = matrix_brut_extract_trades(j, c_extract_rplus_ProductType) And matrix_trades(k, c_extract_rplus_MatchOmsKeyLineSeq) = matrix_brut_extract_trades(j, c_extract_rplus_MatchOmsKeyLineSeq) Then
            If matrix_trades(k, c_extract_rplus_BID) = matrix_brut_extract_trades(j, c_extract_rplus_BID) And matrix_trades(k, c_extract_rplus_ProductType) = matrix_brut_extract_trades(j, c_extract_rplus_ProductType) And matrix_trades(k, c_extract_rplus_MatchOmsKeyLineSeq) = matrix_brut_extract_trades(j, c_extract_rplus_MatchOmsKeyLineSeq) Then
                
                'sum exeQty
                matrix_trades(k, c_extract_rplus_ExeQty) = matrix_trades(k, c_extract_rplus_ExeQty) + matrix_brut_extract_trades(j, c_extract_rplus_ExeQty)
                
                'count nbre trades
                matrix_trades(k, dim_matrix_brut_nbre_trades) = matrix_trades(k, dim_matrix_brut_nbre_trades) + 1
                
                'marque la ligne comme traitee
                matrix_brut_extract_trades(j, dim_matrix_brut_traitement) = True
                
                'total cost
                matrix_trades(k, dim_matrix_brut_cost_total) = matrix_trades(k, dim_matrix_brut_cost_total) + (matrix_brut_extract_trades(j, c_extract_rplus_ExeQty) * matrix_brut_extract_trades(j, c_extract_rplus_ExecPr))
                
                
            End If
            
        Next j
        
        
        'mid cost
        If matrix_trades(k, c_extract_rplus_ExeQty) <> 0 Then
            matrix_trades(k, dim_matrix_brut_cost_mid) = matrix_trades(k, dim_matrix_brut_cost_total) / matrix_trades(k, c_extract_rplus_ExeQty)
        Else
            matrix_trades(k, dim_matrix_brut_cost_mid) = 0
        End If
        
        'partial / complete ?
        If matrix_trades(k, c_extract_rplus_ExeQty) = matrix_trades(k, c_extract_rplus_Orig_Qt) Then
            matrix_trades(k, dim_matrix_brut_order_status) = "Total"
        Else
            matrix_trades(k, dim_matrix_brut_order_status) = "Partial"
        End If
        
        k = k + 1
        
    Else
        'la ligne est deja etait sommee avec les trades avec le meme id
    End If
    
Next i


'complete avec les donnnées d'equityDB & DBIndex / Bloomberg
'repere les colonnes dans equity_db
For i = 1 To 250
    
    If c_book_equity_db_uid <> 0 And c_book_equity_db_name <> 0 And c_book_equity_db_isin <> 0 And c_book_equity_db_ticker <> 0 Then
        
        Exit For
        
    Else
        If Workbooks(file_book).Worksheets("Equity_Database").Cells(l_book_equity_db_header, i) = "Underlying_Id" Or Workbooks(file_book).Worksheets("Equity_Database").Cells(l_book_equity_db_header, i) = "Identifier" Then
            c_book_equity_db_uid = i
        ElseIf Workbooks(file_book).Worksheets("Equity_Database").Cells(l_book_equity_db_header, i) = "Equities_Name" Then
            c_book_equity_db_name = i
        ElseIf Workbooks(file_book).Worksheets("Equity_Database").Cells(l_book_equity_db_header, i) = "Position Statut" Then
            c_book_equity_db_status = i
        ElseIf Workbooks(file_book).Worksheets("Equity_Database").Cells(l_book_equity_db_header, i) = "ISIN" Then
            c_book_equity_db_isin = i
        ElseIf Workbooks(file_book).Worksheets("Equity_Database").Cells(l_book_equity_db_header, i) = "BLOOMBERG" Then
            c_book_equity_db_ticker = i
        ElseIf Workbooks(file_book).Worksheets("Equity_Database").Cells(l_book_equity_db_header, i) = "Currency" Then
            c_book_equity_db_crncy_code = i
        End If
    End If
Next i

'attrape la derniere ligne d'equity db
step_next_line = 2
For i = l_book_equity_db_header + 2 To 32000 Step step_next_line
    If Workbooks(file_book).Worksheets("Equity_Database").Cells(i, c_book_equity_db_uid) = "" And Workbooks(file_book).Worksheets("Equity_Database").Cells(i + 1 * step_next_line, c_book_equity_db_uid) = "" And Workbooks(file_book).Worksheets("Equity_Database").Cells(i + 2 * step_next_line, c_book_equity_db_uid) = "" Then
        l_book_equity_db_last_line = i - step_next_line
        Exit For
    End If
Next i

Dim matrix_equity_db() As Variant
ReDim matrix_equity_db(((l_book_equity_db_last_line - (l_book_equity_db_header + step_next_line)) / step_next_line) + 1, 10)


'remonte la liste des id, ticker et isin deja present
Dim dim_equity_db_uid As Integer, dim_equity_db_name As Integer, dim_equity_db_ticker As Integer, dim_equity_db_isin As Integer, _
dim_equity_db_crncy_code As Integer, dim_equity_db_status As Integer, dim_equity_db_line As Integer

dim_equity_db_uid = 1
dim_equity_db_name = 2
dim_equity_db_ticker = 3
dim_equity_db_isin = 4
dim_equity_db_crncy_code = 5
dim_equity_db_status = 6
dim_equity_db_line = 7


k = 1
For i = l_book_equity_db_header + 2 To l_book_equity_db_last_line Step step_next_line
    matrix_equity_db(k, dim_equity_db_uid) = Workbooks(file_book).Worksheets("Equity_Database").Cells(i, c_book_equity_db_uid)
    matrix_equity_db(k, dim_equity_db_name) = Workbooks(file_book).Worksheets("Equity_Database").Cells(i, c_book_equity_db_name)
    matrix_equity_db(k, dim_equity_db_ticker) = Workbooks(file_book).Worksheets("Equity_Database").Cells(i, c_book_equity_db_ticker)
    matrix_equity_db(k, dim_equity_db_isin) = Workbooks(file_book).Worksheets("Equity_Database").Cells(i, c_book_equity_db_isin)
    matrix_equity_db(k, dim_equity_db_crncy_code) = Workbooks(file_book).Worksheets("Equity_Database").Cells(i, c_book_equity_db_crncy_code)
    matrix_equity_db(k, dim_equity_db_status) = Workbooks(file_book).Worksheets("Equity_Database").Cells(i, c_book_equity_db_status)
    matrix_equity_db(k, dim_equity_db_line) = i
    
    k = k + 1
Next i



'repere les colonnes d'index_db
For i = 1 To 250
    If Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Underlying_Id" Or Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Identifier" Then
        c_book_index_db_uid = i
    ElseIf Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Futures_Name" Then
        c_book_index_db_name = i
    ElseIf Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Position Statut" Then
        c_book_index_db_status = i
    ElseIf Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Bloomberg" Then
        c_book_index_db_ticker = i
    ElseIf Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "FuturesMaturities_ID_1" Then
        c_book_index_db_FuturesMaturities_ID_1 = i
    ElseIf Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Quotite" Then
        c_book_index_db_quotite = i
    ElseIf Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Currency" Then
        c_book_index_db_crncy = i
    ElseIf Workbooks(file_book).Worksheets("Index_Database").Cells(l_book_index_db_header, i) = "Settlement" Then
        'c_book_index_db_settlement = i
        c_book_index_db_settlement = 33
    End If
Next i


step_next_line = 3
For i = l_book_index_db_header + 2 To 32000 Step step_next_line
    If Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_uid) = "" And Workbooks(file_book).Worksheets("Index_Database").Cells(i + 1 * step_next_line, c_book_index_db_uid) = "" And Workbooks(file_book).Worksheets("Index_Database").Cells(i + 2 * step_next_line, c_book_index_db_uid) = "" Then
        l_book_index_db_last_line = i - step_next_line
        Exit For
    End If
Next i


'remonte la matrix index_database
Dim matrix_index_db() As Variant
'ReDim matrix_index_db(((l_book_index_db_last_line - l_book_index_db_header + 2) / step_next_line) + 1, 10)
ReDim matrix_index_db(((l_book_index_db_last_line - l_book_index_db_header + 2)) + 1, 10)

Dim dim_index_db_uid As Integer, dim_index_db_maturity_id As Integer, dim_index_db_name As Integer, _
    dim_index_db_ticker As Integer, dim_index_db_crncy_code As Integer, dim_index_db_line As Integer, _
    dim_index_db_settlement As Integer, dim_index_db_quotite As Integer, dim_index_db_status As Integer

dim_index_db_uid = 1
dim_index_db_maturity_id = 2
dim_index_db_name = 3
dim_index_db_ticker = 4
dim_index_db_crncy_code = 5
dim_index_db_quotite = 6
dim_index_db_settlement = 7
dim_index_db_status = 8
dim_index_db_line = 9


k = 1
For i = l_book_index_db_header + 2 To l_book_index_db_last_line Step step_next_line
    matrix_index_db(k, dim_index_db_uid) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_uid)
    matrix_index_db(k, dim_index_db_maturity_id) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_FuturesMaturities_ID_1)
    matrix_index_db(k, dim_index_db_ticker) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_ticker)
    matrix_index_db(k, dim_index_db_name) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_name)
    matrix_index_db(k, dim_index_db_quotite) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_quotite)
    matrix_index_db(k, dim_index_db_settlement) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_settlement)
    matrix_index_db(k, dim_index_db_crncy_code) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_crncy)
    matrix_index_db(k, dim_index_db_status) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_status)
    
    matrix_index_db(k, dim_index_db_line) = i
    
    k = k + 1
    
    If Workbooks(file_book).Worksheets("Index_Database").Cells(i + 1, c_book_index_db_ticker) <> "" Then
        
        matrix_index_db(k, dim_index_db_uid) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_uid)
        matrix_index_db(k, dim_index_db_maturity_id) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_FuturesMaturities_ID_1 + 1)
        matrix_index_db(k, dim_index_db_ticker) = Workbooks(file_book).Worksheets("Index_Database").Cells(i + 1, c_book_index_db_ticker)
        matrix_index_db(k, dim_index_db_name) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_name)
        matrix_index_db(k, dim_index_db_quotite) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_quotite)
        matrix_index_db(k, dim_index_db_settlement) = Workbooks(file_book).Worksheets("Index_Database").Cells(i + 1, c_book_index_db_settlement)
        matrix_index_db(k, dim_index_db_crncy_code) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_crncy)
        matrix_index_db(k, dim_index_db_status) = Workbooks(file_book).Worksheets("Index_Database").Cells(i, c_book_index_db_status)
        
        matrix_index_db(k, dim_index_db_line) = i
    
        k = k + 1
        
        
    End If
Next i





'rec des isin grace au construct ticker et au ticker d'equity_db
date_min = "01.01.2050"

Dim found_ticker_in_equity_db As Boolean
Dim found_ticker_in_ticker_not_found As Boolean
Dim vec_ticker_not_found() As String

Dim nbre_ticker_not_found As Integer
nbre_ticker_not_found = 0


Dim vec_trades_lines_ticker() As Integer
Dim nbre_bug_lines_ticker As Integer
nbre_bug_lines_ticker = 0


ReDim Preserve vec_ticker_not_found(nbre_ticker_not_found)

For i = 1 To UBound(matrix_trades, 1)
    
    found_ticker_in_equity_db = False
    
    date_tmp = matrix_trades(i, c_extract_rplus_date)
    
    If date_tmp < date_min Then
        date_min = date_tmp
    End If
    
    
    If matrix_trades(i, c_extract_rplus_Time) <= limit_time Then
          matrix_trades(i, dim_matrix_brut_need_import_in_excel) = False
          
          GoTo check_and_update_with_local_db_next_trade
    End If
    
    If matrix_trades(i, c_extract_rplus_ProductType) = "STOCK" And matrix_trades(i, dim_matrix_brut_isin) = "" Then
        
        For j = 1 To UBound(matrix_equity_db, 1)
            If UCase(matrix_trades(i, dim_matrix_brut_construct_ticker)) = UCase(matrix_equity_db(j, dim_equity_db_ticker)) Then
                
                matrix_trades(i, dim_matrix_brut_ticker) = matrix_equity_db(j, dim_equity_db_ticker)
                matrix_trades(i, dim_matrix_brut_isin) = matrix_equity_db(j, dim_equity_db_isin)
                matrix_trades(i, dim_matrix_brut_uid) = matrix_equity_db(j, dim_equity_db_uid)
                matrix_trades(i, dim_matrix_brut_name) = matrix_equity_db(j, dim_equity_db_name)
                matrix_trades(i, dim_matrix_brut_crncy_code) = matrix_equity_db(j, dim_equity_db_crncy_code)
                matrix_trades(i, dim_matrix_brut_construct_found_ticker_in_db) = True
                
                If only_code_20 = True And (matrix_equity_db(j, dim_equity_db_status) <> 11 And matrix_equity_db(j, dim_equity_db_status) <> 21 And matrix_equity_db(j, dim_equity_db_status) <> 22 And matrix_equity_db(j, dim_equity_db_status) <> 23 And matrix_equity_db(j, dim_equity_db_status) <> 32 And matrix_equity_db(j, dim_equity_db_status) <> 33) Then
                    matrix_trades(i, dim_matrix_brut_need_import_in_excel) = False
                End If
                
                found_ticker_in_equity_db = True
                
                Exit For
            End If
        Next j
        
        
        If found_ticker_in_equity_db = False Then
            
            
            For k = 1 To UBound(matrix_exceptions, 1)
                If UCase(Left(matrix_trades(i, dim_matrix_brut_construct_ticker), Len(matrix_exceptions(k, dim_exception_ticker)))) = UCase(matrix_exceptions(k, dim_exception_ticker)) And matrix_exceptions(k, dim_exception_ticker) <> "" Then
                    
                    matrix_trades(i, dim_matrix_brut_ticker) = Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_ticker)
                    matrix_trades(i, dim_matrix_brut_isin) = Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_isin)
                    matrix_trades(i, dim_matrix_brut_uid) = Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_uid)
                    matrix_trades(i, dim_matrix_brut_name) = Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_name)
                    matrix_trades(i, dim_matrix_brut_crncy_code) = Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_crncy_code)
                    matrix_trades(i, dim_matrix_brut_construct_found_ticker_in_db) = True
                    
                    If only_code_20 = True And (Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_status) <> 11 And Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_status) <> 21 And Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_status) <> 22 And Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_status) <> 23 And Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_status) <> 32 And Workbooks(file_book).Worksheets("Equity_Database").Cells(matrix_exceptions(k, dim_line_in_equity_db), c_book_equity_db_status) <> 33) Then
                        matrix_trades(i, dim_matrix_brut_need_import_in_excel) = False
                    End If
                    
                    found_ticker_in_equity_db = True
                    
                    Exit For
                End If
            Next k
            
            If found_ticker_in_equity_db = False Then 'meme apres avoir passé dans la liste d'exception
                found_ticker_in_ticker_not_found = False
                
                ReDim Preserve vec_trades_lines_ticker(nbre_bug_lines_ticker)
                vec_trades_lines_ticker(nbre_bug_lines_ticker) = i
                
                For j = 0 To UBound(vec_ticker_not_found, 1)
                    If vec_ticker_not_found(j) = matrix_trades(i, dim_matrix_brut_construct_ticker) Then
                        found_ticker_in_ticker_not_found = True
                        Exit For
                    End If
                Next j
                
                If found_ticker_in_ticker_not_found = False Then
                    ReDim Preserve vec_ticker_not_found(nbre_ticker_not_found)
                    vec_ticker_not_found(nbre_ticker_not_found) = matrix_trades(i, dim_matrix_brut_construct_ticker)
                    nbre_ticker_not_found = nbre_ticker_not_found + 1
                End If
            End If
        Else
            'profite de completer les trades du meme symbol
            For k = i + 1 To UBound(matrix_trades, 1)
                If matrix_trades(k, dim_matrix_brut_construct_ticker) = matrix_trades(i, dim_matrix_brut_construct_ticker) Then
                    
                    matrix_trades(k, dim_matrix_brut_ticker) = matrix_equity_db(j, dim_equity_db_ticker)
                    matrix_trades(k, dim_matrix_brut_isin) = matrix_equity_db(j, dim_equity_db_isin)
                    matrix_trades(k, dim_matrix_brut_uid) = matrix_equity_db(j, dim_equity_db_uid)
                    matrix_trades(k, dim_matrix_brut_name) = matrix_equity_db(j, dim_equity_db_name)
                    matrix_trades(k, dim_matrix_brut_crncy_code) = matrix_equity_db(j, dim_equity_db_crncy_code)
                    matrix_trades(k, dim_equity_db_status) = matrix_equity_db(j, dim_equity_db_status)
                    
                    If only_code_20 = True And (matrix_equity_db(j, dim_equity_db_status) <> 11 & matrix_equity_db(j, dim_equity_db_status) <> 21 And matrix_equity_db(j, dim_equity_db_status) <> 22 And matrix_equity_db(j, dim_equity_db_status) <> 23 And matrix_equity_db(j, dim_equity_db_status) <> 32 And matrix_equity_db(j, dim_equity_db_status) <> 33) Then
                        matrix_trades(k, dim_matrix_brut_need_import_in_excel) = False
                    End If
                    
                    matrix_trades(k, dim_matrix_brut_construct_found_ticker_in_db) = True
                    
                    
                End If
            Next k
        End If
    
    ElseIf matrix_trades(i, c_extract_rplus_ProductType) = "FUTURE" And matrix_trades(i, dim_matrix_brut_uid) = "" Then
        
        For j = 1 To UBound(matrix_index_db, 1)
            If UCase(matrix_trades(i, c_extract_rplus_BID)) = UCase(matrix_index_db(j, dim_index_db_ticker)) Then
                matrix_trades(i, dim_matrix_brut_ticker) = matrix_index_db(j, dim_index_db_ticker)
                matrix_trades(i, dim_matrix_brut_uid) = matrix_index_db(j, dim_index_db_uid)
                matrix_trades(i, dim_matrix_brut_name) = matrix_index_db(j, dim_index_db_name)
                matrix_trades(i, dim_matrix_brut_crncy_code) = matrix_index_db(j, dim_index_db_crncy_code)
                matrix_trades(i, dim_matrix_brut_maturity_id) = matrix_index_db(j, dim_index_db_maturity_id)
                matrix_trades(i, dim_matrix_brut_quotite) = matrix_index_db(j, dim_index_db_quotite)
                matrix_trades(i, dim_matrix_brut_settlement) = matrix_index_db(j, dim_index_db_settlement)
                
                If only_code_20 = True And (matrix_index_db(j, dim_index_db_status) <> 2) Then
                    matrix_trades(i, dim_matrix_brut_need_import_in_excel) = False
                End If
                
                matrix_trades(i, dim_matrix_brut_construct_found_ticker_in_db) = True
                
                
                
                'complete les fut avec meme ticker
                For k = i + 1 To UBound(matrix_trades, 1)
                    If matrix_trades(k, c_extract_rplus_BID) = matrix_trades(i, c_extract_rplus_BID) Then
                        matrix_trades(k, dim_matrix_brut_ticker) = matrix_index_db(j, dim_index_db_ticker)
                        matrix_trades(k, dim_matrix_brut_uid) = matrix_index_db(j, dim_index_db_uid)
                        matrix_trades(k, dim_matrix_brut_name) = matrix_index_db(j, dim_index_db_name)
                        matrix_trades(k, dim_matrix_brut_crncy_code) = matrix_index_db(j, dim_index_db_crncy_code)
                        matrix_trades(k, dim_matrix_brut_maturity_id) = matrix_index_db(j, dim_index_db_maturity_id)
                        matrix_trades(k, dim_matrix_brut_quotite) = matrix_index_db(j, dim_index_db_quotite)
                        matrix_trades(k, dim_matrix_brut_settlement) = matrix_index_db(j, dim_index_db_settlement)
                        
                        If only_code_20 = True And (matrix_index_db(j, dim_index_db_status) <> 2) Then
                            matrix_trades(k, dim_matrix_brut_need_import_in_excel) = False
                        End If
                        
                        matrix_trades(k, dim_matrix_brut_construct_found_ticker_in_db) = True
                    End If
                Next k
                
            End If
        Next j
    
    End If
check_and_update_with_local_db_next_trade:
Next i



If nbre_ticker_not_found > 0 Then

    Dim l_book_parameters_crncy_last_line As Integer
    
    For i = l_book_parameters_crncy_header To 32
        If Workbooks(file_book).Worksheets("Parametres").Cells(i, c_book_parameters_crncy_crncy) = "" Then
            l_book_parameters_crncy_last_line = i - 1
            Exit For
        End If
    Next i
    
    Dim parameters_crncy() As Variant
    Dim dim_crncy_txt As Integer, dim_crncy_code As Integer
    ReDim parameters_crncy(l_book_parameters_crncy_last_line - l_book_parameters_crncy_header, 2)
        dim_crncy_txt = 1
        dim_crncy_code = 2
    
    k = 1
    For i = l_book_parameters_crncy_header + 1 To l_book_parameters_crncy_last_line
        parameters_crncy(k, dim_crncy_txt) = Workbooks(file_book).Worksheets("Parametres").Cells(i, c_book_parameters_crncy_crncy)
        parameters_crncy(k, dim_crncy_code) = Workbooks(file_book).Worksheets("Parametres").Cells(i, c_book_parameters_crncy_code)
        
        k = k + 1
    Next i
    
    'descend les isin des tickers construct introuvable
    Dim l_blp_bb As BlpData
    Set l_blp_bb = New BlpData
    Dim return_bbg_data As Variant
    return_bbg_data = l_blp_bb.BLPSubscribe(vec_ticker_not_found, Array("ID_ISIN", "NAME", "CRNCY"))
        'complete la matrix de trades
        For i = 1 To UBound(matrix_trades, 1)
            For j = 0 To UBound(return_bbg_data, 1)
                If UCase(matrix_trades(i, dim_matrix_brut_construct_ticker)) = UCase(vec_ticker_not_found(j)) Then
                    
                    If Left(return_bbg_data(j, 0), 1) <> "#" Then
                        matrix_trades(i, dim_matrix_brut_isin) = return_bbg_data(j, 0)
                        matrix_trades(i, dim_matrix_brut_name) = return_bbg_data(j, 1)
                        
                        For k = 1 To UBound(parameters_crncy, 1)
                            If UCase(return_bbg_data(j, 2)) = UCase(parameters_crncy(k, dim_crncy_txt)) Then
                                matrix_trades(i, dim_matrix_brut_crncy_code) = parameters_crncy(k, dim_crncy_code)
                            End If
                        Next k
                        
                        'maintenant que l'isin est connu, il est existe pe dans equity db
                        For k = 1 To UBound(matrix_equity_db, 1)
                            If matrix_trades(i, dim_matrix_brut_isin) = matrix_equity_db(k, dim_equity_db_isin) Then
                                
                                matrix_trades(i, dim_matrix_brut_name) = matrix_equity_db(k, dim_equity_db_name) 'prend le nom d'equity DB pour eviter les conflits
                                matrix_trades(i, dim_matrix_brut_ticker) = matrix_equity_db(k, dim_equity_db_ticker)
                                matrix_trades(i, dim_matrix_brut_uid) = matrix_equity_db(k, dim_equity_db_uid)
                                
                                If only_code_20 = True And (matrix_equity_db(k, dim_equity_db_status) <> 21 And matrix_equity_db(k, dim_equity_db_status) <> 22 And matrix_equity_db(k, dim_equity_db_status) <> 23) Then
                                    matrix_trades(i, dim_matrix_brut_need_import_in_excel) = False
                                End If
                                
                                matrix_trades(i, dim_matrix_brut_construct_found_ticker_in_db) = True
                            End If
                        Next k
                    
                    End If
                    
                End If
            Next j
        Next i
        
End If



'envoi des trades retraités dans la sheet trades du book
Dim l_book_trades_header As Integer, l_book_trades_last_line As Integer, l_book_trades_start_line_search As Integer

Dim c_book_trades_uid As Integer, c_book_trades_instrument_id As Integer, c_book_trades_ticket As Integer, _
c_book_trades_exec_qty As Integer, c_book_trades_date As Integer, c_book_trades_quantity_used As Integer, c_book_trades_qty_open As Integer, c_book_trades_name As Integer, _
c_book_trades_price As Integer, c_book_trades_value As Integer, c_book_trades_status As Integer, c_book_trades_value_date As Integer, _
c_book_trades_characteristics As Integer, c_book_trades_option_quotity As Integer, c_book_trades_currency As Integer, _
c_book_trades_commission As Integer, c_book_trades_isin As Integer, c_book_trades_broker As Integer, c_book_trades_strategy As Integer, _
c_book_trades_statut As Integer, c_book_trades_futuresmaturities_id As Integer, c_book_trades_maturity As Integer, _
c_book_trades_cover As Integer, c_book_trades_bq_seq As Integer, c_book_trades_id_oms_gs As Integer

l_book_trades_header = 16
    c_book_trades_uid = 0
    c_book_trades_instrument_id = 0
    c_book_trades_ticket = 0
    c_book_trades_exec_qty = 0
    c_book_trades_date = 0
    c_book_trades_quantity_used = 0
    c_book_trades_qty_open = 0
    c_book_trades_name = 0
    c_book_trades_price = 0
    c_book_trades_value = 0
    c_book_trades_status = 0
    c_book_trades_value_date = 0
    c_book_trades_characteristics = 0
    c_book_trades_option_quotity = 0
    c_book_trades_currency = 0
    c_book_trades_commission = 0
    c_book_trades_isin = 0
    c_book_trades_broker = 0
    c_book_trades_strategy = 0
    c_book_trades_statut = 0
    c_book_trades_futuresmaturities_id = 0
    c_book_trades_maturity = 0
    c_book_trades_cover = 28 'rajouter header colonne au plus vite
    c_book_trades_bq_seq = 29 'rajouter header colonne au plus vite
    c_book_trades_id_oms_gs = 30
    
    

For i = 1 To 250
    If Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Underlying_ID" Or (Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Identifier" And c_book_trades_uid = 0) Then
        c_book_trades_uid = i '1
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Instrument_ID" Or (Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Identifier" And c_book_trades_uid <> 0) Then
        c_book_trades_instrument_id = i '2
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Ticket" Then
        c_book_trades_ticket = i '3
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "#" Then
        c_book_trades_exec_qty = i '4
        c_book_trades_exec_qty = 4
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Date" Then
        c_book_trades_date = i '5
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Quantity Used" Then
        c_book_trades_quantity_used = i '6
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Quantity Open" Then
        c_book_trades_qty_open = i '7
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Equities_Name" Then
        c_book_trades_name = i '8
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Price" Then
        c_book_trades_price = i '9
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Value" Then
        c_book_trades_value = i '10
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Status" Then
        c_book_trades_status = i '11
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Value Date" Then
        c_book_trades_value_date = i '15
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Characteristics" Then
        c_book_trades_characteristics = i '16
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Option Quotity" Then
        c_book_trades_option_quotity = i '17
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Currency" Then
        c_book_trades_currency = i '18
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Commission" Then
        c_book_trades_commission = i '19
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "ISIN" Then
        c_book_trades_isin = i '20
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Broker" Then
        c_book_trades_broker = i '22
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Strategy" Then
        c_book_trades_strategy = i '23
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Statut" Then
        c_book_trades_statut = i '24
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "FuturesMaturities_Id" Then
        c_book_trades_futuresmaturities_id = i '25
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "Maturity" Then
        c_book_trades_maturity = i '26
    ElseIf Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_header, i) = "MatchOmsKeyLineSeq" Then
        c_book_trades_id_oms_gs = i '30
    End If
Next i


Dim nbre_ticket_trades As Double
nbre_ticket_trades = 0

'determine debut/fin/start search
l_book_trades_start_line_search = l_book_trades_header + 1
For i = l_book_trades_header + 1 To 32000
    
    date_tmp_1 = Workbooks(file_book).Worksheets("Trades").Cells(i, c_book_trades_value_date)
    date_tmp_2 = Workbooks(file_book).Worksheets("Trades").Cells(i + 1, c_book_trades_value_date)
    
    If date_tmp_1 < date_min And date_tmp_2 >= date_min And l_book_trades_start_line_search = (l_book_trades_header + 1) Then
        l_book_trades_start_line_search = i + 1
    End If
    
    If Workbooks(file_book).Worksheets("Trades").Cells(i, c_book_trades_ticket) = "" And Workbooks(file_book).Worksheets("Trades").Cells(i + 1, c_book_trades_ticket) = "" And Workbooks(file_book).Worksheets("Trades").Cells(i + 2, c_book_trades_ticket) = "" And Workbooks(file_book).Worksheets("Trades").Cells(i + 3, c_book_trades_ticket) = "" Then
        l_book_trades_last_line = i - 1
        
        
        If l_book_trades_last_line = l_book_trades_header Then
            nbre_ticket_trades = 0
            Exit For
        Else
            nbre_ticket_trades = Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_ticket)
            Exit For
        End If
    End If
Next i



'remonte les codes de brokers
Dim l_parameters_broker_first_line As Integer
    l_parameters_broker_first_line = 14

Dim c_parameters_broker_name As Integer, c_parameters_broker_code
    c_parameters_broker_name = 239
    c_parameters_broker_code = 238

Dim vec_broker_name() As Variant
Dim vec_broker_code() As Variant

k = 0
For i = l_parameters_broker_first_line To 150
    If Workbooks(file_book).Worksheets("Parametres").Cells(i, c_parameters_broker_name) = "" Then
        Exit For
    Else
        ReDim Preserve vec_broker_name(k)
        ReDim Preserve vec_broker_code(k)
        
        vec_broker_name(k) = Workbooks(file_book).Worksheets("Parametres").Cells(i, c_parameters_broker_name)
        vec_broker_code(k) = Workbooks(file_book).Worksheets("Parametres").Cells(i, c_parameters_broker_code)
        
        k = k + 1
    End If
Next i

Dim broker_name As String
Dim found_entry_in_trades As Boolean
Dim l_com_formula As Variant, l_trade_broker_code As Integer, l_trade_type As String, l_underlying_ccy As Integer, l_underlying_name As String, l_trade_size As Double, l_trade_rate As Double, L_value As Variant

Dim found_account As Boolean

Dim start_trades_import As Integer

If IsEmpty(last_pos_only) = True Or last_pos_only < 0 Then
    start_trades_import = 1
Else
    If UBound(matrix_trades, 1) > last_pos_only Then
        start_trades_import = UBound(matrix_trades, 1) - last_pos_only
    Else
        start_trades_import = 1
    End If
End If


'ENVOI dans trades dans la sheet excel
For i = start_trades_import To UBound(matrix_trades, 1)
    
    If matrix_trades(i, dim_matrix_brut_need_import_in_excel) = False Then
        GoTo next_trade
    End If
    
    found_entry_in_trades = False
    
    date_tmp = matrix_trades(i, c_extract_rplus_date)
    
    If matrix_trades(i, c_extract_rplus_ProductType) = "STOCK" Then
        
        found_account = False
        
        For j = 0 To UBound(account_equity, 1)
            If account_equity(j)(1) = matrix_trades(i, c_extract_rplus_Exch) Then
                found_account = True
                broker_name = account_equity(j)(0)
                Exit For
            End If
        Next j
        
        If found_account = False Then
            For j = 0 To UBound(account_equity, 1)
                If account_equity(j)(1) = matrix_trades(i, c_extract_account_number) Then
                    found_account = True
                    broker_name = account_equity(j)(0)
                    Exit For
                End If
            Next j
        End If
        
        
        If found_account = False Then
            GoTo next_trade
        Else
        
            'l'entrée existe-t-elle deja dans la sheet trades ?
            For j = l_book_trades_start_line_search To l_book_trades_last_line
                If matrix_trades(i, dim_matrix_brut_uid) = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_uid) And matrix_trades(i, c_extract_rplus_MatchOmsKeyLineSeq) = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_id_oms_gs) And date_tmp = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_value_date) Then
                    found_entry_in_trades = True 'edit equity line
                    
                    'ajuste les colonnes
                        
                        'execQty
                        If Left(matrix_trades(i, c_extract_rplus_Side), 1) = "B" Then
                            'achat de titre
                            'remarque : directement possible de prendre execQty car a été sommé sur tous les trades avec le même OMS_id
                            Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_exec_qty) = matrix_trades(i, c_extract_rplus_ExeQty)
                        ElseIf Left(matrix_trades(i, c_extract_rplus_Side), 1) = "S" Then
                            'vente
                            Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_exec_qty) = -matrix_trades(i, c_extract_rplus_ExeQty)
                        End If
                        
                        'price
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_price) = matrix_trades(i, c_extract_rplus_OrdrPr)
                        
                        
                        'commission
                        
                        For k = 0 To UBound(vec_broker_name, 1)
                            If UCase(vec_broker_name(k)) = UCase(broker_name) Then
                                l_trade_broker_code = vec_broker_code(k)
                                Exit For
                            End If
                        Next k
                        
                        l_trade_type = "E"
                        l_underlying_ccy = matrix_trades(i, dim_matrix_brut_crncy_code)
                        l_underlying_name = matrix_trades(i, dim_matrix_brut_name)
                        l_trade_size = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_exec_qty)
                        l_trade_rate = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_price)
                        
                        l_com_formula = set_Commission(l_trade_broker_code, l_trade_type, l_underlying_ccy, l_underlying_name, l_trade_size, l_trade_rate)
                        L_value = "=" & Replace(l_com_formula, ";", ",")
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_commission).Value = L_value
                        
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_broker) = broker_name
                        
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_cover) = matrix_trades(i, dim_matrix_brut_order_status)
                        
                        
                        
                        'ecrase certaines données statiques (important si equity_db était incomplet lors du dernier passage)
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_name) = matrix_trades(i, dim_matrix_brut_name)
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_currency) = matrix_trades(i, dim_matrix_brut_crncy_code)
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_isin) = matrix_trades(i, dim_matrix_brut_isin)
                        
                    
                    Exit For
                End If
            Next j
            
            If found_entry_in_trades = False Then 'nouvelle entrée equity
                
                'nouveau ticket
                nbre_ticket_trades = nbre_ticket_trades + 1
                l_book_trades_last_line = l_book_trades_last_line + 1
                
                
                If matrix_trades(i, dim_matrix_brut_construct_found_ticker_in_db) = False Then
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_uid) = error_msg_if_uid_not_found
                    
                    'passage en couleur de la ligne
                    For j = 1 To c_book_trades_id_oms_gs
                        Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, j).Interior.ColorIndex = error_color_line
                        Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, j).Interior.Pattern = xlSolid
                    Next j
                Else
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_uid) = matrix_trades(i, dim_matrix_brut_uid)
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_instrument_id) = matrix_trades(i, dim_matrix_brut_uid)
                End If
                
                
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_ticket) = nbre_ticket_trades
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_characteristics) = LCase(matrix_trades(i, c_extract_rplus_ProductType))
                
                If Left(UCase(matrix_trades(i, c_extract_rplus_Side)), 1) = "B" Then
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_exec_qty) = matrix_trades(i, c_extract_rplus_ExeQty)
                ElseIf Left(UCase(matrix_trades(i, c_extract_rplus_Side)), 1) = "S" Then
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_exec_qty) = -matrix_trades(i, c_extract_rplus_ExeQty)
                End If
                
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_date) = Format(matrix_trades(i, c_extract_rplus_Time), "hh:mm:ss")
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_date).NumberFormat = "hh:mm"
                
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_quantity_used) = 0
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_qty_open).FormulaR1C1 = "=(RC4-RC6)"
                
                If matrix_trades(i, dim_matrix_brut_construct_found_ticker_in_db) = False Then
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_name) = matrix_trades(i, c_extract_rplus_Symbol)
                Else
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_name) = matrix_trades(i, dim_matrix_brut_name)
                End If
                
                
                
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_price) = matrix_trades(i, c_extract_rplus_OrdrPr)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_value).FormulaR1C1 = "=IF(RC11=""C"",0,RC7*RC9)*RC17"
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_status).FormulaR1C1 = "=IF(RC7=0,""C"",""O"")"
                
                date_tmp = matrix_trades(i, c_extract_rplus_date)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_value_date) = date_tmp
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_value_date).NumberFormat = "dd.mm.yyyy"
                
                
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_characteristics) = LCase(matrix_trades(i, c_extract_rplus_ProductType))
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_option_quotity) = 1
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_currency) = matrix_trades(i, dim_matrix_brut_crncy_code)
                
                
                'commission
                
                For k = 0 To UBound(vec_broker_name, 1)
                    If UCase(vec_broker_name(k)) = UCase(broker_name) Then
                        l_trade_broker_code = vec_broker_code(k)
                        Exit For
                    End If
                Next k
                
                l_trade_type = "E"
                l_underlying_ccy = matrix_trades(i, dim_matrix_brut_crncy_code)
                l_underlying_name = matrix_trades(i, dim_matrix_brut_name)
                l_trade_size = Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_exec_qty)
                l_trade_rate = Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_price)
                
                l_com_formula = set_Commission(l_trade_broker_code, l_trade_type, l_underlying_ccy, l_underlying_name, l_trade_size, l_trade_rate)
                L_value = "=" & Replace(l_com_formula, ";", ",")
                
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_commission).Value = L_value
                
                
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_isin) = matrix_trades(i, dim_matrix_brut_isin)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_broker) = broker_name
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_statut) = "OPEN"
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_bq_seq) = matrix_trades(i, c_extract_rplus_BrSeq)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_cover) = matrix_trades(i, dim_matrix_brut_order_status)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_id_oms_gs) = matrix_trades(i, c_extract_rplus_MatchOmsKeyLineSeq)
                
            End If
        End If
    
    ElseIf matrix_trades(i, c_extract_rplus_ProductType) = "FUTURE" Then
        
        
        found_account = False
        For j = 0 To UBound(account_future, 1)
            If account_future(j)(1) = matrix_trades(i, c_extract_account_number) Then
                found_account = True
                broker_name = account_future(j)(0)
                Exit For
            End If
        Next j

        If found_account = False Then
            GoTo next_trade
        Else
        
            'l'entrée existe-t-elle deja dans la sheet trades ?
            For j = l_book_trades_start_line_search To l_book_trades_last_line
                If matrix_trades(i, dim_matrix_brut_uid) = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_uid) And matrix_trades(i, c_extract_rplus_MatchOmsKeyLineSeq) = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_id_oms_gs) And date_tmp = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_value_date) Then
                    found_entry_in_trades = True 'ajuste la ligne de FUT
    
                    'ajuste les colonnes
    
                        'execQty
                        If Left(matrix_trades(i, c_extract_rplus_Side), 1) = "B" Then
                            'achat de titre
                            'remarque : directement possible de prendre execQty car a été sommé sur tous les trades avec le même OMS_id
                            Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_exec_qty) = matrix_trades(i, c_extract_rplus_ExeQty)
                        ElseIf Left(matrix_trades(i, c_extract_rplus_Side), 1) = "S" Then
                            'vente
                            Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_exec_qty) = -matrix_trades(i, c_extract_rplus_ExeQty)
                        End If
    
                        'price
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_price) = matrix_trades(i, c_extract_rplus_OrdrPr)
    
    
                        'commission
                        
                        For k = 0 To UBound(vec_broker_name, 1)
                            If UCase(vec_broker_name(k)) = UCase(broker_name) Then
                                l_trade_broker_code = vec_broker_code(k)
                                Exit For
                            End If
                        Next k
                        
                        l_trade_type = "I"
                        l_underlying_ccy = matrix_trades(i, dim_matrix_brut_crncy_code)
                        l_underlying_name = matrix_trades(i, dim_matrix_brut_name)
                        l_trade_size = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_exec_qty)
                        l_trade_rate = Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_price)
    
                        l_com_formula = set_Commission(l_trade_broker_code, l_trade_type, l_underlying_ccy, l_underlying_name, l_trade_size, l_trade_rate)
                        L_value = "=" & Replace(l_com_formula, ";", ",")
                        
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_broker) = broker_name
    
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_commission).Value = L_value
    
                        Workbooks(file_book).Worksheets("Trades").Cells(j, c_book_trades_cover) = matrix_trades(i, dim_matrix_brut_order_status)
    
    
                    Exit For
                End If
            Next j
    
    
            If found_entry_in_trades = False Then 'entree de la ligne de FUT
                'nouveau ticket
                nbre_ticket_trades = nbre_ticket_trades + 1
                l_book_trades_last_line = l_book_trades_last_line + 1
    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_uid) = matrix_trades(i, dim_matrix_brut_uid)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_instrument_id) = matrix_trades(i, dim_matrix_brut_maturity_id)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_ticket) = nbre_ticket_trades
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_characteristics) = LCase(matrix_trades(i, c_extract_rplus_ProductType))
    
                If Left(UCase(matrix_trades(i, c_extract_rplus_Side)), 1) = "B" Then
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_exec_qty) = matrix_trades(i, c_extract_rplus_ExeQty)
                ElseIf Left(UCase(matrix_trades(i, c_extract_rplus_Side)), 1) = "S" Then
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_exec_qty) = -matrix_trades(i, c_extract_rplus_ExeQty)
                End If
    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_date) = Format(matrix_trades(i, c_extract_rplus_Time), "hh:mm:ss")
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_date).NumberFormat = "hh:mm"
                    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_quantity_used) = 0
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_qty_open).FormulaR1C1 = "=(RC4-RC6)"
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_name) = matrix_trades(i, dim_matrix_brut_name)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_price) = matrix_trades(i, c_extract_rplus_OrdrPr)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_value).FormulaR1C1 = "=IF(RC11=""C"",0,RC7*RC9)*RC17"
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_status).FormulaR1C1 = "=IF(RC7=0,""C"",""O"")"
    
                date_tmp = matrix_trades(i, c_extract_rplus_date)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_value_date) = date_tmp
                    Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_value_date).NumberFormat = "dd.mm.yyyy"
    
    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_characteristics) = LCase(matrix_trades(i, c_extract_rplus_ProductType))
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_option_quotity) = matrix_trades(i, dim_matrix_brut_quotite)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_currency) = matrix_trades(i, dim_matrix_brut_crncy_code)
    
    
                'commission
                
                For k = 0 To UBound(vec_broker_name, 1)
                    If UCase(vec_broker_name(k)) = UCase(broker_name) Then
                        l_trade_broker_code = vec_broker_code(k)
                        Exit For
                    End If
                Next k
                
                l_trade_type = "I"
                l_underlying_ccy = matrix_trades(i, dim_matrix_brut_crncy_code)
                l_underlying_name = matrix_trades(i, dim_matrix_brut_name)
                l_trade_size = Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_exec_qty)
                l_trade_rate = Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_price)
    
                l_com_formula = set_Commission(l_trade_broker_code, l_trade_type, l_underlying_ccy, l_underlying_name, l_trade_size, l_trade_rate)
                L_value = "=" & Replace(l_com_formula, ";", ",")
    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_commission).Value = L_value
    
    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_isin) = 0
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_broker) = broker_name
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_statut) = "OPEN"
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_currency) = matrix_trades(i, dim_matrix_brut_crncy_code)
    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_bq_seq) = matrix_trades(i, c_extract_rplus_BrSeq)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_cover) = matrix_trades(i, dim_matrix_brut_order_status)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_id_oms_gs) = matrix_trades(i, c_extract_rplus_MatchOmsKeyLineSeq)
    
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_futuresmaturities_id) = matrix_trades(i, dim_matrix_brut_maturity_id)
                Workbooks(file_book).Worksheets("Trades").Cells(l_book_trades_last_line, c_book_trades_maturity) = matrix_trades(i, dim_matrix_brut_settlement)
    
            End If
        
        End If
        
    End If
next_trade:
Next i

Application.Calculation = xlCalculationAutomatic
frm_Import_Trades.Hide

End Sub



'standardiser la reception de la matrix
Sub update_bridge_universal(ByVal vec_product_id As Variant, ByVal vec_description As Variant)

Dim debug_test As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

'mount les bridges deja présent
Dim bridge_product_id() As Variant

Dim sql_query As String
sql_query = "SELECT * FROM t_bridge"
Dim extract_bridge As Variant
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_bridge, 1) <> 0 Then
    ReDim bridge_product_id(UBound(extract_bridge, 1) - 1)
Else
    ReDim bridge_product_id(0)
End If


j = 0
For i = 1 To UBound(extract_bridge, 1)
    bridge_product_id(j) = extract_bridge(i, 0)
    
    j = j + 1
Next i

'isole uniquement les produits manquants pour éviter des rec inutiles
Dim vec_not_found_product_id() As Variant
Dim vec_not_found_product_description() As Variant

ReDim vec_not_found_product_id(0)

k = 0
For i = 0 To UBound(vec_product_id, 1)
    For j = 0 To UBound(extract_bridge, 1)
        If vec_product_id(i) = extract_bridge(j, 0) Then
            Exit For
        Else
            If j = UBound(extract_bridge, 1) Then
                ReDim Preserve vec_not_found_product_id(k)
                ReDim Preserve vec_not_found_product_description(k)
                
                vec_not_found_product_id(k) = vec_product_id(i)
                vec_not_found_product_description(k) = vec_description(i)
                
                k = k + 1
            End If
        End If
    Next j
Next i


If k > 0 Then
    vec_product_id = vec_not_found_product_id
    vec_description = vec_not_found_product_description
Else
    MsgBox ("Toutes les positions traitées sont déjà dans le bridge")
    Exit Sub
End If


'mount les exceptions
Dim extract_exception As Variant
sql_query = "SELECT * FROM t_exception"
extract_exception = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'mount les instruments
Dim extract_instrument As Variant
sql_query = "SELECT * FROM t_instrument"
extract_instrument = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'remonte les pos d'option d'open
Dim l_open_header As Integer, c_open_product_id As Integer, c_open_underlying_id As Integer, c_open_product_type As Integer, _
        l_open_last_line As Integer
    l_open_header = 25
    c_open_product_id = 2
    c_open_underlying_id = 1
    c_open_product_type = 6
    
Dim vec_open_option() As Variant, vec_open_future() As Variant
ReDim vec_open_option(0)
    vec_open_option(0) = Array("", "")
ReDim vec_open_future(0)
    vec_open_future(0) = Array("", "")

m = 0
n = 0
For i = l_open_header + 1 To 32000
    If Worksheets("Open").Cells(i, c_open_underlying_id) = "" And Worksheets("Open").Cells(i + 1, c_open_underlying_id) = "" And Worksheets("Open").Cells(i + 2, c_open_underlying_id) = "" And Worksheets("Open").Cells(i + 3, c_open_underlying_id) = "" Then
        l_open_last_line = i - 1
        Exit For
    Else
        If Worksheets("Open").Cells(i, c_open_product_type) = "C" Or Worksheets("Open").Cells(i, c_open_product_type) = "P" Then
            For j = 0 To UBound(vec_open_option, 1)
                If Worksheets("Open").Cells(i, c_open_product_id) = vec_open_option(j)(0) Then
                    Exit For
                Else
                    If j = UBound(vec_open_option, 1) Then
                        ReDim Preserve vec_open_option(m)
                        vec_open_option(m) = Array(Worksheets("Open").Cells(i, c_open_product_id).Value, Worksheets("Open").Cells(i, c_open_underlying_id).Value, "")
                        m = m + 1
                    End If
                End If
            Next j
        ElseIf Worksheets("Open").Cells(i, c_open_product_type) = "F" Then
            For j = 0 To UBound(vec_open_future, 1)
                If Worksheets("Open").Cells(i, c_open_product_id) = vec_open_future(j)(0) Then
                    Exit For
                Else
                    If j = UBound(vec_open_future, 1) Then
                        ReDim Preserve vec_open_future(n)
                        vec_open_future(n) = Array(Worksheets("Open").Cells(i, c_open_product_id).Value, Worksheets("Open").Cells(i, c_open_underlying_id).Value, "")
                        n = n + 1
                    End If
                End If
            Next j
        Else
            
        End If
    End If
Next i


'recupere les descriptions pour les options et fut
Dim l_options_folio_header As Integer, c_options_folio_product_id As Integer, c_options_folio_description As Integer
l_options_folio_header = 10
c_options_folio_product_id = 1
c_options_folio_description = 14

For i = 0 To UBound(vec_open_option, 1)
    For j = l_options_folio_header + 2 To 32000
        If Worksheets("Options_Folio").Cells(j, c_options_folio_product_id) = "" And Worksheets("Options_Folio").Cells(j + 1, c_options_folio_product_id) = "" Then
            Exit For
        Else
            If vec_open_option(i)(0) = Worksheets("Options_Folio").Cells(j, c_options_folio_product_id) Then
                vec_open_option(i)(2) = Worksheets("Options_Folio").Cells(j, c_options_folio_description)
                Exit For
            End If
        End If
    Next j
Next i


Dim l_futures_folio_header As Integer, c_futures_folio_product_id As Integer, c_futures_folio_description As Integer
l_futures_folio_header = 10
c_futures_folio_product_id = 1
c_futures_folio_description = 16

For i = 0 To UBound(vec_open_future, 1)
    For j = l_futures_folio_header + 2 To 32000
        If Worksheets("Futures_Folio").Cells(j, c_futures_folio_product_id) = "" And Worksheets("Futures_Folio").Cells(j + 1, c_futures_folio_product_id) = "" Then
            Exit For
        Else
            If vec_open_future(i)(0) = Worksheets("Futures_Folio").Cells(j, c_futures_folio_product_id) Then
                vec_open_future(i)(2) = Worksheets("Futures_Folio").Cells(j, c_futures_folio_description)
                Exit For
            End If
        End If
    Next j
Next i


'remonte les pos d'equity d'equity_database
Dim l_equity_db_header As Integer, c_equity_db_product_id As Integer, c_equity_db_compagny_name As Integer, _
    c_equity_db_last_line As Integer

l_equity_db_header = 25
c_equity_db_product_id = 1
c_equity_db_compagny_name = 2

Dim vec_equity_db_equity() As Variant

k = 0
For i = l_equity_db_header + 2 To 3200 Step 2
    If Worksheets("Equity_Database").Cells(i, c_equity_db_product_id) = "" And Worksheets("Equity_Database").Cells(i + 2, c_equity_db_product_id) = "" And Worksheets("Equity_Database").Cells(i + 4, c_equity_db_product_id) = "" Then
        Exit For
    Else
        ReDim Preserve vec_equity_db_equity(k)
        vec_equity_db_equity(k) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_product_id).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_product_id).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_compagny_name).Value)
        k = k + 1
    End If
Next i



'remonte les entrée de database_folio
Dim l_db_folio_header As Integer, c_db_folio_id As Integer, c_db_folio_u_id As Integer, c_db_folio_description As Integer
l_db_folio_header = 12
c_db_folio_id = 10
c_db_folio_u_id = 11
c_db_folio_description = 6

Dim vec_db_folio_product() As Variant

k = 0
ReDim vec_db_folio_id(k)
For i = l_db_folio_header To 32000
    If Worksheets("Database_Folio").Cells(i, 1) = "" And Worksheets("Database_Folio").Cells(i + 1, 1) = "" And Worksheets("Database_Folio").Cells(i + 2, 1) = "" Then
        Exit For
    Else
        If Worksheets("Database_Folio").Cells(i, c_db_folio_id) <> "" And Worksheets("Database_Folio").Cells(i, c_db_folio_u_id) <> "" Then
            ReDim Preserve vec_db_folio_product(k)
            vec_db_folio_product(k) = Array(Worksheets("Database_Folio").Cells(i, c_db_folio_id).Value, Worksheets("Database_Folio").Cells(i, c_db_folio_u_id).Value, Worksheets("Database_Folio").Cells(i, c_db_folio_description).Value)
            k = k + 1
        End If
    End If
Next i

'remonte la view all de folio
Dim fileFolio As File_Folio
Set fileFolio = New File_Folio
fileFolio.set_file_path = base_path & "\GS_Folio\" & folio_all_view

Dim matrix_folio As Variant
matrix_folio = fileFolio.get_content_as_a_matrix()


Dim dim_folio_id As Integer, dim_folio_ticker As Integer, dim_folio_crncy As Integer, dim_folio_description As Integer, _
    dim_folio_qty_yesterday_close As Integer, dim_folio_underlying_id As Integer, dim_folio_yesterday_close_price As Integer, _
    dim_folio_product_type As Integer

'detect les dimensions
For i = 1 To UBound(matrix_folio, 2)
    If matrix_folio(0, i) = "Identifier" Then
        dim_folio_id = i
    ElseIf matrix_folio(0, i) = "Market Data Symbol" Then
        dim_folio_ticker = i
    ElseIf matrix_folio(0, i) = "CCY" Then
        dim_folio_crncy = i
    ElseIf matrix_folio(0, i) = "Description" Then
        dim_folio_description = i
    ElseIf matrix_folio(0, i) = "Qty - Yesterday's Close" Then
        dim_folio_qty_yesterday_close = i
    ElseIf matrix_folio(0, i) = "Underlyer Product ID" Then
        dim_folio_underlying_id = i
    ElseIf matrix_folio(0, i) = "Yesterday's Close (Local)" Then
        dim_folio_yesterday_close_price = i
    ElseIf matrix_folio(0, i) = "Product Type" Then
        dim_folio_product_type = i
    End If
Next i



''remonte GVA
'Dim l_gva_header As Integer, l_gva_last_line As Integer, c_gva_product_id As Integer, c_gva_underlying_id As Integer, _
'    c_gva_l_s As Integer, c_gva_description As Integer
'
'    l_gva_header = 8
'    c_gva_product_id = 1
'    c_gva_underlying_id = 2
'    c_gva_l_s = 3
'    c_gva_description = 4
'
'Dim vec_gva_product() As Variant
'ReDim vec_gva_product(0)
'vec_gva_product(0) = Array("", "")
'
'k = 0
'For i = l_gva_header + 1 To 32000
'    If Worksheets("GS_Folio_Geneva").Cells(i, c_gva_product_id) = "" And Worksheets("GS_Folio_Geneva").Cells(i + 1, c_gva_product_id) = "" Then
'        l_gva_last_line = i - 1
'        Exit For
'    Else
'        If Worksheets("GS_Folio_Geneva").Cells(i, c_gva_l_s) <> "Cash" Then
'            For j = 0 To UBound(vec_gva_product, 1)
'                If Worksheets("GS_Folio_Geneva").Cells(i, c_gva_product_id) = vec_gva_product(j)(0) Then
'                    Exit For
'                Else
'                    If j = UBound(vec_gva_product, 1) Then
'                        ReDim Preserve vec_gva_product(k)
'                        vec_gva_product(k) = Array(Worksheets("GS_Folio_Geneva").Cells(i, c_gva_product_id).value, Worksheets("GS_Folio_Geneva").Cells(i, c_gva_underlying_id).value, Worksheets("GS_Folio_Geneva").Cells(i, c_gva_description).value)
'                        k = k + 1
'                    End If
'                End If
'            Next j
'        End If
'    End If
'Next i



'RECONCILLATION
Dim vec_issue_not_enough_data() As Variant
Dim count_issue_not_enough_data As Integer
    count_issue_not_enough_data = 0

Dim vec_issue_kronos() As Variant
Dim count_issue_kronos As Integer
    count_issue_kronos = 0

Dim data_to_import() As Variant
ReDim data_to_import(UBound(vec_product_id, 1), 10)

Dim dim_data_product_id As Integer, dim_data_underlying_id As Integer, dim_data_description As Integer, _
    dim_data_description_jeffery As Integer, dim_data_instrument_id As Integer, dim_data_need_import As Integer
    
    dim_data_product_id = 0
    dim_data_underlying_id = 1
    dim_data_description = 2
    dim_data_description_jeffery = 3
    dim_data_instrument_id = 4
    dim_data_need_import = 5


Dim reply As Variant
Dim vec_adr() As Variant
Dim count_adr As Integer
    count_adr = 0




Dim vec_bug_no_enough_info() As Variant
Dim count_bug_no_enough_info As Integer
    count_bug_no_enough_info = 0

Dim vec_new_equity() As Variant
Dim vec_new_future() As Variant
Dim vec_new_option() As Variant

Dim vec_new_exception() As Variant


Dim count_new_equity As Integer, count_new_future As Integer, count_new_option As Integer, count_new_exception As Integer
    count_new_equity = 0
    count_new_future = 0
    count_new_option = 0
    count_new_exception = 0
    
    ReDim vec_new_exception(0)



For i = 0 To UBound(vec_product_id, 1)
    
    found_data = False
    
    data_to_import(i, dim_data_product_id) = vec_product_id(i)
    data_to_import(i, dim_data_description_jeffery) = vec_description(i)
    data_to_import(i, dim_data_need_import) = False
    
    'passe en revue les options d'open
    For j = 0 To UBound(vec_open_option, 1)
        If vec_product_id(i) = vec_open_option(j)(0) Then
            
            data_to_import(i, dim_data_underlying_id) = vec_open_option(j)(1)
            data_to_import(i, dim_data_description) = vec_open_option(j)(2)
            data_to_import(i, dim_data_instrument_id) = 3
            
            data_to_import(i, dim_data_need_import) = True
            
            ReDim Preserve vec_new_option(count_new_option)
            vec_new_option(count_new_option) = data_to_import(i, dim_data_product_id)
            count_new_option = count_new_option + 1
            
            GoTo rec_next_product
            Exit For
        End If
    Next j
    
    If data_to_import(i, dim_data_need_import) = False Then
        'passe en revue les FUT
        For j = 0 To UBound(vec_open_future, 1)
            If vec_product_id(i) = vec_open_future(j)(0) Then
                
                data_to_import(i, dim_data_underlying_id) = vec_open_future(j)(1)
                data_to_import(i, dim_data_description) = vec_open_future(j)(2)
                data_to_import(i, dim_data_instrument_id) = 2
            
                data_to_import(i, dim_data_need_import) = True
                
                ReDim Preserve vec_new_future(count_new_future)
                vec_new_future(count_new_future) = data_to_import(i, dim_data_product_id)
                count_new_future = count_new_future + 1
                
                GoTo rec_next_product
                Exit For
            End If
        Next j
    End If
    
    
    If data_to_import(i, dim_data_need_import) = False Then
        'il doit donc s'agir d'une equity
        For j = 0 To UBound(vec_equity_db_equity, 1)
            If vec_product_id(i) = vec_equity_db_equity(j)(0) Then
                
                data_to_import(i, dim_data_underlying_id) = vec_equity_db_equity(j)(1)
                data_to_import(i, dim_data_description) = vec_equity_db_equity(j)(2)
                data_to_import(i, dim_data_instrument_id) = 1
                
                ReDim Preserve vec_new_equity(count_new_equity)
                vec_new_equity(count_new_equity) = data_to_import(i, dim_data_product_id)
                count_new_equity = count_new_equity + 1
                
                data_to_import(i, dim_data_need_import) = True
                GoTo rec_next_product
                Exit For
            End If
        Next j
    End If
    
    
    If data_to_import(i, dim_data_need_import) = False Then
        'rien dans le système Kronos, il a probleme avec la position
        'premiere tentative avec database folio
        For j = 0 To UBound(vec_db_folio_product, 1)
            If vec_product_id(i) = vec_db_folio_product(j)(0) Then
                
                data_to_import(i, dim_data_underlying_id) = vec_db_folio_product(j)(1)
                data_to_import(i, dim_data_description) = vec_db_folio_product(j)(2)
                
                'determination du produit
                If vec_db_folio_product(j)(0) = vec_db_folio_product(j)(1) Then
                    data_to_import(i, dim_data_instrument_id) = 1
                    
                    ReDim Preserve vec_new_equity(count_new_equity)
                    vec_new_equity(count_new_equity) = data_to_import(i, dim_data_product_id)
                    count_new_equity = count_new_equity + 1
                    
                Else
                    If Left(UCase(vec_db_folio_product(j)(2)), 4) = "CALL" Or Left(UCase(vec_db_folio_product(j)(2)), 3) = "PUT" Then
                        data_to_import(i, dim_data_instrument_id) = 3
                        
                        ReDim Preserve vec_new_option(count_new_option)
                        vec_new_option(count_new_option) = data_to_import(i, dim_data_product_id)
                        count_new_option = count_new_option + 1
                    Else
                        'Future ou ADR mais impossible de déterminer
                        reply = MsgBox("If " & data_to_import(i, dim_data_description) & " is a future then press Yes, else if it's an ADR press no", vbYesNo, "Future / ADR")
                        
                        If reply = vbYes Then
                            data_to_import(i, dim_data_instrument_id) = 2
                            ReDim Preserve vec_new_future(count_new_future)
                            vec_new_future(count_new_future) = data_to_import(i, dim_data_product_id)
                            count_new_future = count_new_future + 1
                        ElseIf reply = vbNo Then
                            data_to_import(i, dim_data_instrument_id) = 1
                            data_to_import(i, dim_data_underlying_id) = data_to_import(i, dim_data_product_id)
                            
                            ReDim Preserve vec_new_equity(count_new_equity)
                            vec_new_equity(count_new_equity) = data_to_import(i, dim_data_product_id)
                            count_new_equity = count_new_equity + 1
                            
                            'ajoute exception
                            ReDim Preserve vec_adr(count_adr)
                            vec_adr(count_adr) = data_to_import(i, dim_data_product_id)
                            count_adr = count_adr + 1
                        End If
                    End If
                End If
                
                
                'construction du msg pour system kronos
                ReDim Preserve vec_issue_kronos(count_issue_kronos)
                vec_issue_kronos(count_issue_kronos) = Array(data_to_import(i, dim_data_product_id), data_to_import(i, dim_data_underlying_id), data_to_import(i, dim_data_description), data_to_import(i, dim_data_instrument_id), "database_folio")
                count_issue_kronos = count_issue_kronos + 1
                
                data_to_import(i, dim_data_need_import) = True
                GoTo rec_next_product
                Exit For
                
            End If
            
        Next j
    End If
    
    
    
'    If data_to_import(i, dim_data_need_import) = False Then
'        'derniere tentative avec GVA
'        For j = 0 To UBound(vec_gva_product, 1)
'            If vec_product_id(i) = vec_gva_product(j)(0) Then
'
'                data_to_import(i, dim_data_underlying_id) = vec_gva_product(j)(1)
'                data_to_import(i, dim_data_description) = vec_gva_product(j)(2)
'
'
'                'determination du produit
'                If vec_gva_product(j)(0) = vec_gva_product(j)(1) Then
'                    data_to_import(i, dim_data_instrument_id) = 1
'                Else
'                    If Left(UCase(vec_gva_product(j)(2)), 4) = "CALL" Or Left(UCase(vec_gva_product(j)(2)), 3) = "PUT" Then
'                        data_to_import(i, dim_data_instrument_id) = 3
'
'                        ReDim Preserve vec_new_option(count_new_option)
'                        vec_new_option(count_new_option) = data_to_import(i, dim_data_product_id)
'                        count_new_option = count_new_option + 1
'                    Else
'                        'Future ou ADR mais impossible de déterminer
'                        reply = MsgBox("If " & data_to_import(i, dim_data_description) & " is a future then press Yes, else if it's an ADR press no", vbYesNo, "Future / ADR")
'
'                        If reply = vbYes Then
'                            data_to_import(i, dim_data_instrument_id) = 2
'
'                            ReDim Preserve vec_new_future(count_new_future)
'                            vec_new_future(count_new_future) = data_to_import(i, dim_data_product_id)
'                            count_new_future = count_new_future + 1
'                        ElseIf reply = vbNo Then
'                            data_to_import(i, dim_data_instrument_id) = 1
'                            data_to_import(i, dim_data_underlying_id) = data_to_import(i, dim_data_product_id)
'
'                            ReDim Preserve vec_new_equity(count_new_equity)
'                            vec_new_equity(count_new_equity) = data_to_import(i, dim_data_product_id)
'                            count_new_equity = count_new_equity + 1
'
'
'                            'ajoute exception
'                            ReDim Preserve vec_adr(count_adr)
'                            vec_adr(count_adr) = data_to_import(i, dim_data_product_id)
'                            count_adr = count_adr + 1
'                        End If
'                    End If
'                End If
'
'
'                'construction du msg pour system kronos
'                ReDim Preserve vec_issue_kronos(count_issue_kronos)
'                vec_issue_kronos(count_issue_kronos) = Array(data_to_import(i, dim_data_product_id), data_to_import(i, dim_data_underlying_id), data_to_import(i, dim_data_description), data_to_import(i, dim_data_instrument_id), "geneva")
'                count_issue_kronos = count_issue_kronos + 1
'
'                data_to_import(i, dim_data_need_import) = True
'                GoTo rec_next_product
'                Exit For
'
'            End If
'
'        Next j
        
        
        If data_to_import(i, dim_data_need_import) = False Then
            'impossible d'effectuer l'entrée dans bridge
            ReDim Preserve vec_issue_not_enough_data(count_issue_not_enough_data)
            vec_issue_not_enough_data(count_issue_not_enough_data) = Array(vec_product_id(i), "", vec_description(i))
            count_issue_not_enough_data = count_issue_not_enough_data + 1
        End If
        
    'End If
    
    
rec_next_product:
Next i





Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With



Dim count_new_line_in_bridge As Integer
count_new_line_in_bridge = 0


With rst
    
    .ActiveConnection = conn
    .Open "t_bridge", LockType:=adLockOptimistic


            'insere les nouvelles lignes
            For i = 0 To UBound(data_to_import, 1)
                If data_to_import(i, dim_data_need_import) = True Then
                    
                    .AddNew
                    
                        .fields("gs_id") = data_to_import(i, dim_data_product_id)
                        .fields("gs_underlying_id") = data_to_import(i, dim_data_underlying_id)
                        .fields("gs_description") = data_to_import(i, dim_data_description)
                        .fields("gs_pict_exec_description") = data_to_import(i, dim_data_description_jeffery)
                        .fields("system_instrument_id") = data_to_import(i, dim_data_instrument_id)
                        
                        count_new_line_in_bridge = count_new_line_in_bridge + 1
                    
                    .Update
                    
                    
                End If
            Next i
    
    .Close
        
End With


'creation des entrées pour les exceptions
If count_adr > 0 Then
    With rst
    
        .ActiveConnection = conn
        .Open "t_exception", LockType:=adLockOptimistic
        
        For i = 0 To UBound(vec_new_exception, 1)
            .AddNew
            
                .fields("gs_id") = vec_adr(i)
                .fields("gs_underlying_id") = vec_adr(i)
            
            .Update
        Next i
        
        .Close
    
    End With
End If

conn.Close

Dim c_cointrin_product_id As Integer, c_cointrin_underlying_id As Integer, c_cointrin_description As Integer, _
    c_cointrin_product As Integer, c_cointrin_source_info As Integer

Dim l_cointrin_header As Integer
l_cointrin_header = 8

c_cointrin_product_id = 100
c_cointrin_underlying_id = 101
c_cointrin_description = 102
c_cointrin_product = 103
c_cointrin_source_info = 104


Worksheets("Cointrin").Columns(c_cointrin_product_id).Clear
Worksheets("Cointrin").Columns(c_cointrin_underlying_id).Clear
Worksheets("Cointrin").Columns(c_cointrin_description).Clear
Worksheets("Cointrin").Columns(c_cointrin_product).Clear
Worksheets("Cointrin").Columns(c_cointrin_source_info).Clear

Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_product_id) = "product_id"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_underlying_id) = "underlying_id"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_description) = "description"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_product) = "product"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_source_info) = "source"



k = l_cointrin_header + 1
If count_issue_not_enough_data > 0 Then
    MsgBox ("impossible d'insérer les données dans bridge pour " & count_issue_not_enough_data & " car les informations sont incomplètes")
    
    For i = 0 To UBound(vec_issue_not_enough_data, 1)
        Debug.Print vec_issue_not_enough_data(i)(0)
        
        Worksheets("Cointrin").Cells(k, c_cointrin_product_id) = vec_issue_not_enough_data(i)(0)
        Worksheets("Cointrin").Cells(k, c_cointrin_underlying_id) = "NO DATA"
        Worksheets("Cointrin").Cells(k, c_cointrin_description) = vec_issue_not_enough_data(i)(1)
        
    Next i
    
    k = k + 2
    
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_product_id).Activate
End If


If count_issue_kronos > 0 Then
    'impression du rapport / déclenchement d'opérations
    MsgBox ("erreur au niveau du système Kronos, consultez le rapport")
    
    For i = 0 To UBound(vec_issue_kronos, 1)
        Worksheets("Cointrin").Cells(k, c_cointrin_product_id) = vec_issue_kronos(i)(0)
        Worksheets("Cointrin").Cells(k, c_cointrin_underlying_id) = vec_issue_kronos(i)(1)
        Worksheets("Cointrin").Cells(k, c_cointrin_description) = vec_issue_kronos(i)(2)
        
        If vec_issue_kronos(i)(3) = 1 Then
            Worksheets("Cointrin").Cells(k, c_cointrin_product) = "equity"
        ElseIf vec_issue_kronos(i)(3) = 2 Then
            Worksheets("Cointrin").Cells(k, c_cointrin_product) = "future"
        ElseIf vec_issue_kronos(i)(3) = 3 Then
            Worksheets("Cointrin").Cells(k, c_cointrin_product) = "option"
        End If
        
        Worksheets("Cointrin").Cells(k, c_cointrin_source_info) = vec_issue_kronos(i)(4)
        
        k = k + 1
    Next i
    
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_product_id).Activate
    
End If



'update des tables
If count_new_equity > 0 Then
   Call insert_new_equity_pict_exec(vec_new_equity)
End If

If count_new_future > 0 Then
    Call insert_new_future_pict_exec(vec_new_future)
End If

If count_new_option > 0 Then
    Call insert_new_option_pict_exec(vec_new_option)
End If


MsgBox ("New lines in bridge : " & count_new_line_in_bridge)

End Sub


Sub load_trades_for_one_security(Optional product_id As String, Optional underlying_id As String)

Dim debug_test As Variant
Dim sql_query As String

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim trader_txt As String, trader_code As Integer
trader_code = 0
trader_txt = Worksheets("Cointrin").Cells(5, 2)

Application.Calculation = xlManual

If trader_txt = "" Then
    Exit Sub
Else
    Dim extract_trader As Variant
    sql_query = "SELECT system_code, system_first_name, system_surname FROM t_trader"
    extract_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    For i = 1 To UBound(extract_trader, 1)
        If trader_txt = extract_trader(i, 1) & " " & extract_trader(i, 2) Then
            trader_code = extract_trader(i, 0)
            Exit For
        End If
    Next i
End If

If trader_code = 0 Then
    Exit Sub
End If


Dim l_cointrin_header As Integer
l_cointrin_header = 8


Dim c_cointrin_product_id As Integer, c_cointrin_underlying_id As Integer, c_cointrin_close_price As Integer, _
    c_cointrin_price_factor As Integer
    
    c_cointrin_product_id = 0
    c_cointrin_underlying_id = 0
    c_cointrin_close_price = 0
    c_cointrin_price_factor = 0

For i = 1 To 250
    If c_cointrin_product_id <> 0 And c_cointrin_underlying_id <> 0 And c_cointrin_close_price <> 0 And c_cointrin_price_factor <> 0 Then
        Exit For
    Else
        If Worksheets("Cointrin").Cells(l_cointrin_header, i) = "Identifier" And c_cointrin_product_id = 0 Then
            c_cointrin_product_id = i
        ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "Identifier" And c_cointrin_product_id <> 0 Then
            c_cointrin_underlying_id = i
        ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "gva_close_price" Then
            c_cointrin_close_price = i
        ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "factor" Then
            c_cointrin_price_factor = i
        End If
    End If
Next i


Dim c_cointrin_extract_trade_product_id As Integer, c_cointrin_extract_trade_date As Integer, c_cointrin_extract_trade_time As Integer, _
    c_cointrin_extract_trade_qty As Integer, c_cointrin_extract_trade_price As Integer, c_cointrin_extract_trade_mv As Integer, _
    c_cointrin_extract_trade_comm As Integer, c_cointrin_extract_trade_broker As Integer, c_cointrin_extract_trade_reason As Integer, _
    c_cointrin_extract_trade_unique_id As Integer
    
    c_cointrin_extract_trade_product_id = 50
    c_cointrin_extract_trade_date = 51
    c_cointrin_extract_trade_time = 52
    c_cointrin_extract_trade_qty = 53
    c_cointrin_extract_trade_price = 54
    c_cointrin_extract_trade_mv = 55
    c_cointrin_extract_trade_comm = 56
    c_cointrin_extract_trade_broker = 57
    c_cointrin_extract_trade_reason = 58
    c_cointrin_extract_trade_unique_id = 59



Dim is_product_id As Boolean, is_underlying_id As Boolean
    is_product_id = False
    is_underlying_id = False
    
Dim id As String


If (IsMissing(product_id) = True Or product_id = "") And IsMissing(underlying_id) = False Then
    is_underlying_id = True
    id = underlying_id
ElseIf IsMissing(product_id) = False And (IsMissing(underlying_id) = True Or underlying_id = "") Then
    is_product_id = True
    id = product_id
Else
    Exit Sub
End If


'remonte la liste des product id qu'il va falloir afficher
sql_query = "SELECT gs_id, gs_description, gs_pict_exec_description FROM t_bridge "

If is_product_id = True Then
    sql_query = sql_query & " WHERE gs_id=""" & id & """"
ElseIf is_underlying_id = True Then
    sql_query = sql_query & " WHERE gs_underlying_id=""" & id & """"
End If

Dim extract_bridge As Variant
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_trading_account As Variant
sql_query = "SELECT gs_account_number FROM t_trading_account WHERE system_trader_code=" & trader_code
extract_trading_account = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)



'clean area trades
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_product_id).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_date).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_time).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_qty).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_price).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_mv).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_comm).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_broker).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_reason).Clear
    Worksheets("Cointrin").Columns(c_cointrin_extract_trade_unique_id).Clear
    
    
    'header
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_product_id) = "product_id"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_product_id).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_date) = "date"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_date).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_time) = "time"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_time).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_qty) = "exec_qty"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_qty).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_price) = "exec_price"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_price).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_mv) = "market_value"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_mv).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_comm) = "comm"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_comm).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_broker) = "broker"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_broker).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_reason) = "comments"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_reason).Font.Bold = True
    Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_unique_id) = "unique_id"
        Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_extract_trade_unique_id).Font.Bold = True



Dim dim_trades_date As Integer, dim_trades_time As Integer, dim_trades_unique_id As Integer, dim_trades_security_id As Integer, _
    dim_trades_exec_qty As Integer, dim_trades_exec_price As Integer, dim_trades_account As Integer, _
    dim_trades_broker As Integer, dim_trades_comm As Integer, dim_trades_pnl_reversal As Integer


'pour chaque product id, remonter les trades
k = 0
Dim new_product As Boolean

Dim comments As String

Dim total_qty As Double
Dim sum_market_value As Double
Dim close_value As Double
Dim price_factor As Double


Dim extract_trades As Variant
For i = 1 To UBound(extract_bridge, 1)
    
    total_qty = 0
    sum_market_value = 0
    close_value = 0
    price_factor = 1
    
    'repere la valeur en close + price factor
    For j = l_cointrin_header To 32000
        If Worksheets("Cointrin").Cells(j, c_cointrin_product_id) = "" Then
            Exit For
        Else
            If Worksheets("Cointrin").Cells(j, c_cointrin_product_id) = extract_bridge(i, 0) Then
                close_value = Worksheets("Cointrin").Cells(j, c_cointrin_close_price)
                price_factor = Worksheets("Cointrin").Cells(j, c_cointrin_price_factor)
                Exit For
            End If
        End If
    Next j
    
    sql_query = "SELECT * FROM t_trade WHERE gs_security_id=""" & extract_bridge(i, 0) & """ ORDER BY gs_security_id, gs_date, gs_time"
    
    extract_trades = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    'dimension
    For j = 0 To UBound(extract_trades, 2)
        If extract_trades(0, j) = "gs_date" Then
            dim_trades_date = j
        ElseIf extract_trades(0, j) = "gs_time" Then
            dim_trades_time = j
        ElseIf extract_trades(0, j) = "gs_unique_id" Then
            dim_trades_unique_id = j
        ElseIf extract_trades(0, j) = "gs_security_id" Then
            dim_trades_security_id = j
        ElseIf extract_trades(0, j) = "gs_exec_qty" Then
            dim_trades_exec_qty = j
        ElseIf extract_trades(0, j) = "gs_exec_price" Then
            dim_trades_exec_price = j
        ElseIf extract_trades(0, j) = "gs_trading_account" Then
            dim_trades_account = j
        ElseIf extract_trades(0, j) = "gs_exec_broker" Then
            dim_trades_broker = j
        ElseIf extract_trades(0, j) = "system_commission_local_currency" Then
            dim_trades_comm = j
        ElseIf extract_trades(0, j) = "system_ytd_pnl_reversal" Then
            dim_trades_pnl_reversal = j
        End If
    Next j
    
    
    If UBound(extract_trades, 1) > 0 Then
        
        If IsNull(extract_bridge(i, 1)) = False Or IsNull(extract_bridge(i, 2)) = False Then
            
            If IsNull(extract_bridge(i, 1)) = False Then
                
                If k = 0 Then
                Else
                    k = k + 1
                End If
                
                Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_product_id) = extract_bridge(i, 1)
                
                If price_factor <> 1 Then
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_product_id) = Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_product_id) & " (price factor=" & price_factor & ")"
                End If
                
                k = k + 1
            ElseIf IsNull(extract_bridge(i, 2)) = False Then
                
                If k = 0 Then
                Else
                    k = k + 1
                End If
                
                Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_product_id) = extract_bridge(i, 2)
                
                If price_factor <> 1 Then
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_product_id) = Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_product_id) & " (price factor=" & price_factor & ")"
                End If
                
                k = k + 1
            End If
            
            
        End If
        
        n = 0
        For j = 1 To UBound(extract_trades, 1)
            For m = 0 To UBound(extract_trading_account, 1)
                If extract_trades(j, dim_trades_account) = extract_trading_account(m, 0) Then
                    'le trades apparatient bien au trader - impression de la ligne
                    
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_product_id) = extract_trades(j, dim_trades_security_id)
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_date) = extract_trades(j, dim_trades_date)
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_date).NumberFormat = "dd.mm.yyyy"
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_time) = extract_trades(j, dim_trades_time)
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_time).NumberFormat = "hh:mm"
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_qty) = extract_trades(j, dim_trades_exec_qty)
                        total_qty = total_qty + extract_trades(j, dim_trades_exec_qty)
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_qty).NumberFormat = "#,##0"
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_price) = extract_trades(j, dim_trades_exec_price)
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_price).NumberFormat = "#,##0.00"
                        
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_mv) = price_factor * (extract_trades(j, dim_trades_exec_qty) * extract_trades(j, dim_trades_exec_price)) + extract_trades(j, dim_trades_pnl_reversal)
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_mv).NumberFormat = "#,##0.00"
                        
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_comm) = extract_trades(j, dim_trades_comm)
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_comm).NumberFormat = "#,##0.00"
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_broker) = extract_trades(j, dim_trades_broker)
                    
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = extract_trades(j, dim_trades_unique_id)
                    
                    Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_unique_id) = extract_trades(j, dim_trades_unique_id)
                    
                    If InStr(UCase(extract_trades(j, dim_trades_unique_id)), "GVA") <> 0 Then
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = "GVA reversal"
                    ElseIf InStr(UCase(extract_trades(j, dim_trades_unique_id)), "PICT_EXEC") <> 0 Then
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = ""
                    ElseIf InStr(UCase(extract_trades(j, dim_trades_unique_id)), "FIXING_FUT") <> 0 Then
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = "Fixing"
                    ElseIf InStr(UCase(extract_trades(j, dim_trades_unique_id)), "EXEC_DERIVATIVE") <> 0 Then
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = "Expiry/Exe/Assign"
                        
                        If extract_trades(j, dim_trades_pnl_reversal) <> 0 Then
                            Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) & " " & Round(extract_trades(j, dim_trades_pnl_reversal), 0)
                        End If
                        
                    ElseIf InStr(UCase(extract_trades(j, dim_trades_unique_id)), "MANUAL") <> 0 Then
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = "Manual"
                    ElseIf InStr(UCase(extract_trades(j, dim_trades_unique_id)), "STOCKS#DIV") <> 0 Then
                        comments = Mid(extract_trades(j, dim_trades_unique_id), InStr(UCase(extract_trades(j, dim_trades_unique_id)), "STOCKS#DIV"))
                        comments = Left(comments, InStr(comments, "_") - 1)
                        comments = Replace(comments, "#", " ")
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = comments
                    
                    ElseIf InStr(UCase(extract_trades(j, dim_trades_unique_id)), "CASH#DIV") <> 0 Then
                        comments = Mid(extract_trades(j, dim_trades_unique_id), InStr(UCase(extract_trades(j, dim_trades_unique_id)), "CASH#DIV"))
                        comments = Left(comments, InStr(comments, "_") - 1)
                        comments = Replace(comments, "#", " ")
                        Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_reason) = comments
                        
                    End If
                    
                    k = k + 1
                    n = n + 1
                    
                    Exit For
                End If
            Next m
        Next j
        
        'calcul de la valeur en close
        If n > 0 And total_qty <> 0 Then
            Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_time) = "CLOSE VALUE"
            Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_qty) = -total_qty
                Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_qty).NumberFormat = "#,##0"
            
            Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_price) = close_value
                Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_price).NumberFormat = "#,##0.00"
            
            Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_mv) = -total_qty * close_value * price_factor
                Worksheets("Cointrin").Cells(l_cointrin_header + 1 + k, c_cointrin_extract_trade_mv).NumberFormat = "#,##0.00"
            
            k = k + 1
        End If
        
    End If
    
Next i

Sheets("Cointrin").Activate
Worksheets("Cointrin").Cells(l_cointrin_header + 1, c_cointrin_extract_trade_product_id).Activate

Application.Calculation = xlAutomatic

End Sub




Sub export_div_stock_in_db()

Application.Calculation = xlCalculationManual

Dim sql_query As String

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim l_trades_header As Integer, c_trades_product_id As Integer, c_trades_underlying_id As Integer, c_trades_ticket As Integer, _
    c_trades_qty As Integer, c_trades_time As Integer, c_trades_qty_open As Integer, c_trades_sec_name As Integer, _
    c_trades_price As Integer, c_trades_value As Integer, c_trades_date As Integer, c_trades_comments As Integer, _
    c_trades_com As Integer, c_trades_crncy As Integer, l_trades_last_line As Integer
    
    l_trades_header = 16
    c_trades_product_id = 2
    c_trades_underlying_id = 1
    c_trades_ticket = 3
    c_trades_qty = 4
    c_trades_time = 5
    c_trades_qty_open = 7
    c_trades_sec_name = 8
    c_trades_price = 9
    c_trades_value = 10
    c_trades_date = 15
    c_trades_comments = 16
    c_trades_crncy = 18
    c_trades_com = 19


'repere les lignes
Dim dim_product As Integer, dim_date As Integer, dim_time As Integer, dim_qty As Integer, dim_comment As Integer, _
    dim_currency As Integer


dim_product = 0
dim_date = 1
dim_time = 2
dim_qty = 3
dim_comment = 4
dim_currency = 5

Dim vec_dvd() As Variant
Dim vec_dvd_line() As Variant
k = 0
For i = l_trades_header To 32000
    If Worksheets("Trades").Cells(i, c_trades_product_id) = "" And Worksheets("Trades").Cells(i + 1, c_trades_product_id) = "" And Worksheets("Trades").Cells(i + 2, c_trades_product_id) = "" Then
        Exit For
    Else
        If Left(Worksheets("Trades").Cells(i, c_trades_comments), 16) = "Stocks dividends" Then
            ReDim Preserve vec_dvd(k)
            ReDim Preserve vec_dvd_line(k)
            
            vec_dvd(k) = Array(Worksheets("Trades").Cells(i, c_trades_product_id).Value, Worksheets("Trades").Cells(i, c_trades_date).Value, Worksheets("Trades").Cells(i, c_trades_time).Value, Worksheets("Trades").Cells(i, c_trades_qty).Value, Worksheets("Trades").Cells(i, c_trades_comments).Value, Worksheets("Trades").Cells(i, c_trades_crncy).Value)
            vec_dvd_line(k) = i
            
            k = k + 1
        End If
    End If
Next i

If k > 0 Then
    
    Dim conn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'remonte le trader
    Dim txt_trader As String
    txt_trader = Worksheets("Cointrin").Cells(5, 2)
    
    If txt_trader = "" Then
        Exit Sub
    End If
    
    Dim trader_code As Integer, gs_under_id As String
        trader_code = 0
    
    sql_query = "SELECT system_code, system_first_name, system_surname, gs_UserID FROM t_trader"
    Dim extract_trader As Variant
    extract_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    
    For i = 0 To UBound(extract_trader, 1)
        If txt_trader = extract_trader(i, 1) & " " & extract_trader(i, 2) Then
            trader_code = extract_trader(i, 0)
            gs_under_id = extract_trader(i, 3)
            Exit For
        End If
    Next i
    
    If trader_code = 0 Then
        Exit Sub
    End If
    
    'remonte le main account
    Dim trader_main_account As String
    
    Dim extract_main_account As Variant
    sql_query = "SELECT gs_account_number FROM t_trading_account WHERE system_trader_code=" & trader_code & " AND gs_main_account=TRUE"
    extract_main_account = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    trader_main_account = extract_main_account(1, 0)
    
    
    'creation des lignes dans la table de trade
    With conn
        .Provider = "Microsoft.JET.OLEDB.4.0"
        .Open db_cointrin_trades_path
    End With
    
    With rst
    
    .ActiveConnection = conn
    .Open "t_trade", LockType:=adLockOptimistic
        
        For i = 0 To UBound(vec_dvd, 1)
            
            .AddNew
            
                .fields("gs_date") = vec_dvd(i)(dim_date)
                .fields("gs_time") = vec_dvd(i)(dim_time)
                
                '.fields("gs_unique_id") = "dvd_stock_" & vec_dvd(i)(dim_product) & "_" & Year(vec_dvd(i)(dim_date)) & Month(vec_dvd(i)(dim_date)) & Day(vec_dvd(i)(dim_date)) & "_" & Hour(vec_dvd(i)(dim_time)) & Minute(vec_dvd(i)(dim_time))
                '.fields("gs_unique_id") = "dvd_stock_" & vec_dvd(i)(dim_product) & "_" & Replace(vec_dvd(i)(dim_comment), " ", "#") & "_" & Year(vec_dvd(i)(dim_date)) & Month(vec_dvd(i)(dim_date)) & Day(vec_dvd(i)(dim_date)) & "_" & Hour(vec_dvd(i)(dim_time)) & Minute(vec_dvd(i)(dim_time))
                .fields("gs_unique_id") = "dvd_stock_" & vec_dvd(i)(dim_product) & "_" & Replace(vec_dvd(i)(dim_comment), " ", "#") & "_" & year(Date) & Month(Date) & day(Date) & "_" & Hour(Time) & Minute(Time) & Second(Time)
                
                .fields("gs_security_id") = vec_dvd(i)(dim_product)
                .fields("gs_exec_qty") = vec_dvd(i)(dim_qty)
                .fields("gs_exec_price") = 0
                .fields("gs_order_qty") = vec_dvd(i)(dim_qty)
                
                If vec_dvd(i)(dim_qty) >= 0 Then
                    .fields("gs_side") = "B"
                    .fields("gs_side_detailed") = "B"
                Else
                    .fields("gs_side") = "S"
                    .fields("gs_side_detailed") = "S"
                End If
                
                .fields("gs_trading_account") = trader_main_account
                .fields("gs_user_id") = gs_under_id
                
                .fields("system_currency_code") = vec_dvd(i)(dim_currency)
                
                .fields("system_commission_local_currency") = 0
                
                .fields("system_trader_code") = trader_code
                
            .Update
            
        Next i
        
    End With
    
    
    'mise des qty a 0 dans la sheet Trades
    For i = 0 To UBound(vec_dvd_line, 1)
        Worksheets("Trades").Cells(vec_dvd_line(i), c_trades_qty) = 0
        Worksheets("Trades").Cells(vec_dvd_line(i), c_trades_qty_open) = 0
    Next i
    
End If

Application.Calculation = xlCalculationAutomatic

End Sub


Sub export_div_cash_in_db()

Application.Calculation = xlCalculationManual

Dim debug_test As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim c_exe_product_id As Integer, c_exe_underlying_id As Integer, c_exe_nombre_share As Integer, c_exe_net_amount As Integer, _
    c_exe_exec_result As Integer, c_exe_comments As Integer, c_exe_vba_code As Integer, c_exe_currency As Integer, _
    l_exe_header As Integer, l_exe_last_line As Integer, c_exe_now As Integer
    
    l_exe_header = 16
    c_exe_product_id = 1
    c_exe_underlying_id = 2
    c_exe_nombre_share = 3
    c_exe_now = 4
    c_exe_net_amount = 5
    c_exe_exec_result = 8
    c_exe_comments = 14
    c_exe_currency = 15
    c_exe_vba_code = 18


'remonte tous les lignes de type "Cash div"
Dim date_tmp As Date, time_tmp As Date

    Dim dim_product_id As Integer, dim_cf As Integer, dim_comments As Integer, dim_crncy As Integer, dim_date As Integer, _
        dim_time As Integer
    
    dim_product_id = 0
    dim_date = 1
    dim_time = 2
    dim_cf = 3
    dim_comments = 4
    dim_crncy = 5
    

Dim vec_cash_div_exe() As Variant
Dim vec_cas_div_exe_line() As Variant

k = 0
For i = l_exe_header To 32000
    If Worksheets("Exe").Cells(i, c_exe_product_id) = "" And Worksheets("Exe").Cells(i + 1, c_exe_product_id) = "" Then
        l_exe_last_line = i - 1
        Exit For
    Else
        If Left(Worksheets("Exe").Cells(i, c_exe_comments), 14) = "Cash dividends" Then
            ReDim Preserve vec_cash_div_exe(k)
            ReDim Preserve vec_cas_div_exe_line(k)
            
            date_tmp = day(Worksheets("Exe").Cells(i, c_exe_now)) & "." & Month(Worksheets("Exe").Cells(i, c_exe_now)) & "." & year(Worksheets("Exe").Cells(i, c_exe_now))
            time_tmp = Hour(Worksheets("Exe").Cells(i, c_exe_now)) & ":" & Minute(Worksheets("Exe").Cells(i, c_exe_now)) & ":" & Second(Worksheets("Exe").Cells(i, c_exe_now))
            
            vec_cash_div_exe(k) = Array(Worksheets("Exe").Cells(i, c_exe_product_id).Value, date_tmp, time_tmp, Worksheets("Exe").Cells(i, c_exe_exec_result).Value, Worksheets("Exe").Cells(i, c_exe_comments).Value, Worksheets("Exe").Cells(i, c_exe_currency).Value)
            vec_cas_div_exe_line(k) = i
            
            k = k + 1
            
        End If
    End If
Next i


'insertion des Reversal PnL dans la table de trades
If k > 0 Then
    
    Dim conn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'remonte le trader
    Dim txt_trader As String
    txt_trader = Worksheets("Cointrin").Cells(5, 2)
    
    If txt_trader = "" Then
        Exit Sub
    End If
    
    Dim trader_code As Integer, gs_under_id As String
        trader_code = 0
    
    sql_query = "SELECT system_code, system_first_name, system_surname, gs_UserID FROM t_trader"
    Dim extract_trader As Variant
    extract_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    
    For i = 0 To UBound(extract_trader, 1)
        If txt_trader = extract_trader(i, 1) & " " & extract_trader(i, 2) Then
            trader_code = extract_trader(i, 0)
            gs_under_id = extract_trader(i, 3)
            Exit For
        End If
    Next i
    
    If trader_code = 0 Then
        Exit Sub
    End If
    
    'remonte le main account
    Dim trader_main_account As String
    
    Dim extract_main_account As Variant
    sql_query = "SELECT gs_account_number FROM t_trading_account WHERE system_trader_code=" & trader_code & " AND gs_main_account=TRUE"
    extract_main_account = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    trader_main_account = extract_main_account(1, 0)
    
    
    'creation des lignes dans la table de trade
    With conn
        .Provider = "Microsoft.JET.OLEDB.4.0"
        .Open db_cointrin_trades_path
    End With
    
    With rst
    
    .ActiveConnection = conn
    .Open "t_trade", LockType:=adLockOptimistic
        
        For i = 0 To UBound(vec_cash_div_exe, 1)
            .AddNew
                
                .fields("gs_date") = vec_cash_div_exe(i)(dim_date)
                .fields("gs_time") = vec_cash_div_exe(i)(dim_time)
                
                .fields("gs_unique_id") = "cash_dvd_" & vec_cash_div_exe(i)(dim_product_id) & "_" & Replace(vec_cash_div_exe(i)(dim_comments), " ", "#") & "_" & year(Date) & Month(Date) & day(Date) & "_" & Hour(Time) & Minute(Time) & Second(Time)
                
                .fields("gs_security_id") = vec_cash_div_exe(i)(dim_product_id)
                
                .fields("gs_exec_qty") = 0
                .fields("gs_exec_price") = 0
                .fields("gs_order_qty") = 0
                .fields("gs_side") = "B"
                .fields("gs_side_detailed") = "B"
                .fields("gs_trading_account") = trader_main_account
                .fields("gs_user_id") = gs_under_id
                .fields("gs_close_price") = 0
                .fields("system_ytd_pnl_reversal") = vec_cash_div_exe(i)(dim_cf)
                .fields("system_currency_code") = vec_cash_div_exe(i)(dim_crncy)
                .fields("system_commission_local_currency") = 0
                .fields("system_trader_code") = trader_code
                
            .Update
        Next i
    
    End With
    
    
    'mise a 0 des nbre, price, cf dans la sheet exe
    For i = 0 To UBound(vec_cas_div_exe_line, 1)
        Worksheets("Exe").Cells(vec_cas_div_exe_line(i), c_exe_nombre_share) = 0
        Worksheets("Exe").Cells(vec_cas_div_exe_line(i), c_exe_net_amount) = 0
        Worksheets("Exe").Cells(vec_cas_div_exe_line(i), c_exe_exec_result) = 0
    Next i
    
    
End If

Application.Calculation = xlCalculationAutomatic

End Sub


Sub load_form_morning_procedure()

frm_Morning_Procedure.Show

End Sub

Sub load_trades_from_open_id()

Dim l_open_header As Integer, c_open_product_id As Integer, c_open_underlying_id As Integer
    l_open_header = 25
    c_open_product_id = 2
    c_open_underlying_id = 1

Dim id As String
If ActiveSheet.name = "Open" And ActiveCell.row > l_open_header Then
    id = Worksheets("Open").Cells(ActiveCell.row, c_open_underlying_id)
    Call load_trades_for_one_security(, id)
End If


End Sub


Sub export_dvd_cash_and_stock_to_db()

Call export_div_cash_in_db
Call export_div_stock_in_db

End Sub


Sub export_close_price_gva(ByVal path_to_gva As String)

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim l_gva_header As Integer, c_gva_product_id As Integer, c_gva_underlying_id As Integer, c_gva_close_price As Integer
l_gva_header = 8
c_gva_product_id = 4
c_gva_underlying_id = 5
c_gva_close_price = 10


Dim check_wrkb As Workbook
Dim FoundWrbk As Boolean
Dim file_gva
file_gva = Right(path_to_gva, InStr(StrReverse(path_to_gva), "\") - 1)

FoundWrbk = False
For Each check_wrkb In Workbooks
    If check_wrkb.name = file_gva Then
        FoundWrbk = True
        Exit For
    End If
Next

If FoundWrbk = False Then
    Workbooks.Open filename:=path_to_gva, readOnly:=True
End If



For i = 1 To 250
    If Workbooks(file_gva).Worksheets("Sheet1").Cells(l_gva_header, i) = "Product ID" Then
        c_gva_product_id = i
    ElseIf Workbooks(file_gva).Worksheets("Sheet1").Cells(l_gva_header, i) = "Underlying Product Id" Then
        c_gva_underlying_id = i
    ElseIf Workbooks(file_gva).Worksheets("Sheet1").Cells(l_gva_header, i) = "Price Close (Local)" Then
        c_gva_close_price = i
    End If
Next i


'deduction de la date
Dim date_txt As String, date_txt_month As String, date_txt_year As String, date_txt_day As String

For i = 1 To l_gva_header - 1
    If Workbooks(file_gva).Worksheets("Sheet1").Cells(i, 1) = "Business Date:" Then
        date_txt = Workbooks(file_gva).Worksheets("Sheet1").Cells(i, 2)
        Exit For
    End If
Next i

date_txt_year = Right(date_txt, 4)
date_txt_month = Left(date_txt, InStr(date_txt, " ") - 1)
date_txt_day = Mid(date_txt, InStr(date_txt, " ") + 1, (InStr(date_txt, ",") - (InStr(date_txt, " ") + 1)))


Select Case date_txt_month
    Case "Jan"
        date_txt_month = "01"
    Case "Feb"
        date_txt_month = "02"
    Case "Mar"
        date_txt_month = "03"
    Case "Apr"
        date_txt_month = "04"
    Case "May"
        date_txt_month = "05"
    Case "Jun"
        date_txt_month = "06"
    Case "Jul"
        date_txt_month = "07"
    Case "Aug"
        date_txt_month = "08"
    Case "Sep"
        date_txt_month = "09"
    Case "Oct"
        date_txt_month = "10"
    Case "Nov"
        date_txt_month = "11"
    Case "Dec"
        date_txt_month = "12"
End Select

Dim business_date As Date
business_date = date_txt_day & "." & date_txt_month & "." & date_txt_year



'extraction des prix
k = 0
Dim gva_entries() As Variant
Dim dim_product_id As Integer, dim_price_close As Integer

ReDim gva_entries(0)
gva_entries(0) = Array("", "")
For i = l_gva_header + 1 To 32000
    If Workbooks(file_gva).Worksheets("Sheet1").Cells(i, c_gva_product_id) = "" Then
        Exit For
    Else
        For j = 0 To UBound(gva_entries, 1)
            If gva_entries(j)(dim_product_id) = Workbooks(file_gva).Worksheets("Sheet1").Cells(i, c_gva_product_id) Then
                Exit For
            Else
                If j = UBound(gva_entries, 1) Then
                    ReDim Preserve gva_entries(k)
                    gva_entries(k) = Array(Workbooks(file_gva).Worksheets("Sheet1").Cells(i, c_gva_product_id).Value, Workbooks(file_gva).Worksheets("Sheet1").Cells(i, c_gva_close_price).Value)
                    k = k + 1
                End If
            End If
        Next j
    End If
Next i

Workbooks(file_gva).Close False


'insertion des données
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


'creation des lignes dans la table de trade
With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With


'destruction des doublons
For i = 0 To UBound(gva_entries, 1)
    sql_query = "DELETE FROM t_close_price WHERE gs_date=" & FormatDateSQL(business_date) & " AND gs_product_id=""" & gva_entries(i)(0) & """"
    conn.Execute sql_query
Next i



With rst
    
    .ActiveConnection = conn
    .Open "t_close_price", LockType:=adLockOptimistic
    
    
    For i = 0 To UBound(gva_entries, 1)
        
        .AddNew
            
            .fields("gs_date") = business_date
            .fields("gs_product_id") = gva_entries(i)(0)
            .fields("price_close_gva") = gva_entries(i)(1)
            
        .Update
        
    Next i


End With

End Sub


Sub report_yest_daily_pnl()

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer


Dim base_path
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

'remonte le bridge
Dim extract_bridge As Variant
sql_query = "SELECT * FROM t_bridge"
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim dim_bridge_gs_id As Integer, dim_bridge_gs_underlying_id As Integer, dim_bridge_description As Integer, _
    dim_bridge_instu_id As Integer, dim_bridge_pict_exec_description As Integer

For i = 0 To UBound(extract_bridge, 2)
    If extract_bridge(0, i) = "gs_id" Then
        dim_bridge_gs_id = i
    ElseIf extract_bridge(0, i) = "gs_underlying_id" Then
        dim_bridge_gs_underlying_id = i
    ElseIf extract_bridge(0, i) = "gs_description" Then
        dim_bridge_description = i
    ElseIf extract_bridge(0, i) = "gs_pict_exec_description" Then
        dim_bridge_pict_exec_description = i
    ElseIf extract_bridge(0, i) = "system_instrument_id" Then
        dim_bridge_instu_id = i
    End If
Next i


'remonte le id avec la correction des prix necessaire
Dim extract_exception_gva_price As Variant
sql_query = "SELECT * FROM t_exception_gva_price"
extract_exception_gva_price = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'remonde les id avec correction sur les price factor
Dim extract_exception_gva_price_factor As Variant
sql_query = "SELECT * FROM t_exception_price_factor"
extract_exception_gva_price_factor = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)


'charge les exception (ADR etc.)
Dim extract_exception As Variant
sql_query = "SELECT gs_id, gs_underlying_id FROM t_exception"
extract_exception = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim dim_exception_product_id As Integer, dim_exception_underlying_id As Integer

For i = 0 To UBound(extract_exception, 2)
    If extract_exception(0, i) = "gs_id" Then
        dim_exception_product_id = i
    ElseIf extract_exception(0, i) = "gs_underlying_id" Then
        dim_exception_underlying_id = i
    End If
Next i


'remonte trader code
Dim trader_line As Integer
Dim trader_code As Integer
Dim extract_trader As Variant
sql_query = "SELECT system_code AS code, system_first_name as prenom, system_surname as nom FROM t_trader ORDER BY system_surname"
extract_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim dim_trader_code As Integer, dim_trader_prenom As Integer, dim_trader_nom As Integer

For i = 0 To UBound(extract_trader, 2)
    If extract_trader(0, i) = "code" Then
        dim_trader_code = i
    ElseIf extract_trader(0, i) = "prenom" Then
        dim_trader_prenom = i
    ElseIf extract_trader(0, i) = "nom" Then
        dim_trader_nom = i
    End If
Next i

For i = 0 To UBound(extract_trader, 1)
    If extract_trader(i, dim_trader_prenom) & " " & extract_trader(i, dim_trader_nom) = Worksheets("Cointrin").Cells(5, 2) Then
        trader_code = extract_trader(i, dim_trader_code)
        trader_line = i
        Exit For
    End If
Next i


'remonte instrument
Dim extract_instrument As Variant
sql_query = "SELECT * FROM t_instrument"
extract_instrument = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim dim_instrument_id As Integer, dim_instrument_name As Integer

For i = 0 To UBound(extract_instrument, 2)
    If extract_instrument(0, i) = "system_id" Then
        dim_instrument_id = i
    ElseIf extract_instrument(0, i) = "gs_name" Then
        dim_instrument_name = i
    End If
Next i



'remonte les distinct date < today
Dim extract_distinct_date As Variant
sql_query = "SELECT DISTINCT TOP 5 t_trade.gs_date, COUNT(t_trade.gs_unique_id) AS NbreTrades "
sql_query = sql_query & " FROM t_trade "
sql_query = sql_query & " WHERE t_trade.gs_date<" & FormatDateSQL(Date) & " "
sql_query = sql_query & " GROUP BY t_trade.gs_date "
sql_query = sql_query & " ORDER BY t_trade.gs_date DESC"
extract_distinct_date = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)



Dim date_t_1 As Date, date_t_2 As Date
date_t_1 = extract_distinct_date(1, 0)
date_t_2 = extract_distinct_date(2, 0)


'determination des pos agrégée en t-2
Dim extract_net_pos_t_2 As Variant
sql_query = "SELECT [t_trade].[gs_security_id] AS PB_ID, [t_trading_account].[system_trader_code] AS trader_code, [t_trade].[system_currency_code] AS crncy_code, SUM([t_trade].[gs_exec_qty]) AS NetPos, Sum([t_trade].[gs_exec_qty]*[t_trade].[gs_exec_price]) AS MarketValue, SUM([t_trade].[system_commission_local_currency]+[t_trade].[system_comm_reversal]) AS COMT, Sum([t_trade].[system_ytd_pnl_reversal]+[t_trade].[system_comm_reversal]) AS Reversal_YTD_PnL, SUM([t_trade].[system_custom_ytd_pnl_reversal]) AS CUSTOM_reversal_ytd_pnl"
sql_query = sql_query & " FROM t_trade INNER JOIN t_trading_account ON [t_trade].[gs_trading_account] = [t_trading_account].[gs_account_number]"
sql_query = sql_query & " WHERE [t_trade].[gs_date]<=" & FormatDateSQL(date_t_2) & " "
sql_query = sql_query & " GROUP BY [t_trade].[gs_security_id], [t_trading_account].[system_trader_code], [t_trade].[system_currency_code]"
sql_query = sql_query & " HAVING [t_trading_account].[system_trader_code]=" & trader_code
sql_query = sql_query & " ORDER BY [t_trade].[gs_security_id]"
extract_net_pos_t_2 = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)



'determination des pos agrégée en t-1
Dim extract_net_pos_t_1 As Variant
sql_query = "SELECT [t_trade].[gs_security_id] AS PB_ID, [t_trading_account].[system_trader_code] AS trader_code, [t_trade].[system_currency_code] AS crncy_code, SUM([t_trade].[gs_exec_qty]) AS NetPos, Sum([t_trade].[gs_exec_qty]*[t_trade].[gs_exec_price]) AS MarketValue, SUM([t_trade].[system_commission_local_currency]+[t_trade].[system_comm_reversal]) AS COMT, Sum([t_trade].[system_ytd_pnl_reversal]+[t_trade].[system_comm_reversal]) AS Reversal_YTD_PnL, SUM([t_trade].[system_custom_ytd_pnl_reversal]) AS CUSTOM_reversal_ytd_pnl"
sql_query = sql_query & " FROM t_trade INNER JOIN t_trading_account ON [t_trade].[gs_trading_account] = [t_trading_account].[gs_account_number]"
sql_query = sql_query & " WHERE [t_trade].[gs_date]<=" & FormatDateSQL(date_t_1) & " "
sql_query = sql_query & " GROUP BY [t_trade].[gs_security_id], [t_trading_account].[system_trader_code], [t_trade].[system_currency_code]"
sql_query = sql_query & " HAVING [t_trading_account].[system_trader_code]=" & trader_code
sql_query = sql_query & " ORDER BY [t_trade].[gs_security_id]"
extract_net_pos_t_1 = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim dim_pos_mv_id As Integer, dim_pos_mv_trading_account As Integer, dim_pos_mv_net_pos As Integer, _
    dim_pos_mv_market_value As Integer, dim_pos_mv_com As Integer, dim_pos_mv_ytd_pnl_reversal As Integer, _
    dim_pos_mv_ytd_crncy_code As Integer, dim_pos_mv_ytd_custom_reversal_ytd_pnl As Integer

For i = 0 To UBound(extract_net_pos_t_1, 2)
    If extract_net_pos_t_1(0, i) = "PB_ID" Then
        dim_pos_mv_id = i
    ElseIf extract_net_pos_t_1(0, i) = "NetPos" Then
        dim_pos_mv_net_pos = i
    ElseIf extract_net_pos_t_1(0, i) = "MarketValue" Then
        dim_pos_mv_market_value = i
    ElseIf extract_net_pos_t_1(0, i) = "COMT" Then
        dim_pos_mv_com = i
    ElseIf extract_net_pos_t_1(0, i) = "Reversal_YTD_PnL" Then
        dim_pos_mv_ytd_pnl_reversal = i
    ElseIf extract_net_pos_t_1(0, i) = "CUSTOM_reversal_ytd_pnl" Then
        dim_pos_mv_ytd_custom_reversal_ytd_pnl = i
    ElseIf extract_net_pos_t_1(0, i) = "crncy_code" Then
        dim_pos_mv_ytd_crncy_code = i
    End If
Next i


'remonte les prix en cloture -> GVA
'destruction des valeurs de + de 7 jours

Dim extract_close_price As Variant
sql_query = "SELECT gs_date, gs_product_id, price_close_gva FROM t_close_price WHERE gs_date>=" & FormatDateSQL(date_t_2)
extract_close_price = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim dim_histo_price_date As Integer, dim_histo_price_product As Integer, dim_histo_price_close_price_gva As Integer

For i = 0 To UBound(extract_close_price, 2)
    If extract_close_price(0, i) = "gs_date" Then
        dim_histo_price_date = i
    ElseIf extract_close_price(0, i) = "gs_product_id" Then
        dim_histo_price_product = i
    ElseIf extract_close_price(0, i) = "price_close_gva" Then
        dim_histo_price_close_price_gva = i
    End If
Next i


'charge la view all de folio
Dim view_folio_obj As New File_Folio
view_folio_obj.set_file_path = base_path & "\GS_Folio\" & folio_all_view

Dim matrix_folio As Variant
matrix_folio = view_folio_obj.get_content_as_a_matrix()

Dim dim_folio_id As Integer, dim_folio_qty_yesterday_close As Integer, dim_folio_price_close As Integer, _
    dim_folio_price_factor As Integer, dim_folio_ticker As Integer


For i = 0 To UBound(matrix_folio, 2)
    
    If matrix_folio(0, i) = "Identifier" Then
        dim_folio_id = i
    ElseIf matrix_folio(0, i) = "Qty - Yesterday's Close" Then
        dim_folio_qty_yesterday_close = i
    ElseIf matrix_folio(0, i) = "Yesterday's Close (Local)" Then
        dim_folio_price_close = i
    ElseIf matrix_folio(0, i) = "Price Factor" Then
        dim_folio_price_factor = i
    ElseIf matrix_folio(0, i) = "Market Data Symbol" Then
        dim_folio_ticker = i
    End If
Next i




'quotity
Dim l_equity_db_header As Integer, c_equity_db_id As Integer, c_equity_db_price_close As Integer
    l_equity_db_header = 25
    c_equity_db_id = 1
    c_equity_db_price_close = 27

Dim l_kronos_index_db_header As Integer, c_kronos_index_db_quotity As Integer, c_kronos_index_db_option_quotity As Integer

l_kronos_index_db_header = 25
    c_kronos_index_db_option_quotity = 112
    c_kronos_index_db_quotity = 113
    
Dim l_kronos_equity_db_header As Integer, c_kronos_equity_db_option_quotity As Integer
    l_kronos_equity_db_header = 25
        c_kronos_equity_db_option_quotity = 49


k = 0
Dim vec_internal_db_id() As Variant
Dim vec_internal_db_quotity() As Variant
Dim vec_internal_db_option_quotity() As Variant

For i = l_kronos_index_db_header + 2 To 500 Step 3
    If Worksheets("Index_Database").Cells(i, 1) = "" And Worksheets("Index_Database").Cells(i + 3, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_internal_db_id(k)
        ReDim Preserve vec_internal_db_quotity(k)
        ReDim Preserve vec_internal_db_option_quotity(k)
        
        vec_internal_db_id(k) = Worksheets("Index_Database").Cells(i, 1)
        vec_internal_db_quotity(k) = Worksheets("Index_Database").Cells(i, c_kronos_index_db_quotity)
        vec_internal_db_option_quotity(k) = Worksheets("Index_Database").Cells(i, c_kronos_index_db_option_quotity)
        
        k = k + 1
    End If
Next i

For i = l_kronos_equity_db_header + 2 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" And Worksheets("Equity_Database").Cells(i + 2, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_internal_db_id(k)
        ReDim Preserve vec_internal_db_quotity(k)
        ReDim Preserve vec_internal_db_option_quotity(k)
        
        vec_internal_db_id(k) = Worksheets("Equity_Database").Cells(i, 1)
        vec_internal_db_option_quotity(k) = Worksheets("Equity_Database").Cells(i, c_kronos_equity_db_option_quotity)
        
        k = k + 1
    End If
Next i


'currency
k = 0
Dim crncy_rate() As Variant
For i = 14 To 32
    If Worksheets("Parametres").Cells(i, 5) <> "" Then
        ReDim Preserve crncy_rate(k)
        crncy_rate(k) = Array(Worksheets("Parametres").Cells(i, 5).Value, Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 6).Value)
        k = k + 1
    End If
Next i


'matching
Dim pos_t_1 As Double
Dim pos_t_2 As Double

Dim output_report() As Variant

Dim dim_report_product As Integer, dim_report_underlying As Integer, dim_report_ytd_pnl_t_2 As Integer, dim_report_ytd_pnl_t_1 As Integer, dim_report_daily_pnl As Integer
dim_report_product = 0
dim_report_underlying = 1
dim_report_ytd_pnl_t_2 = 2
dim_report_ytd_pnl_t_1 = 3
dim_report_daily_pnl = 4


For i = 1 To UBound(extract_net_pos_t_1, 1)
    underlying_id = ""
    Description = ""
    price_close = 0
    folio_position = 0
    price_factor = 1
    price_close = 0
    pos_gva = 0
    instrument_id = 0
    crncy_txt = 0
    crncy_conv_rate = 0
    
    found_price = False
    
    found_in_folio = False
    found_in_gva = False
    
    ytd_pnl_t_1 = 0
    ytd_pnl_t_2 = 0
    
    price_t_1 = 0
    price_t_2 = 0
    
    pos_t_1 = 0
    pos_t_2 = 0
    
    'instrument + underlying
    For j = 0 To UBound(extract_bridge, 1)
        If extract_net_pos_t_1(i, dim_pos_mv_id) = extract_bridge(j, dim_bridge_gs_id) Then
            
            underlying_id = extract_bridge(j, dim_bridge_gs_underlying_id)
            instrument_id = extract_bridge(j, dim_bridge_instu_id)
            
            For m = 0 To UBound(extract_exception, 1)
                If extract_net_pos_t_1(i, dim_pos_mv_id) = extract_exception(m, dim_exception_product_id) Then
                    
                    If underlying_id <> extract_exception(m, dim_exception_product_id) And instrument_id = 1 Then
                        'le bridge est faux
                        underlying_id = extract_exception(m, dim_exception_product_id)
                        
                    End If
                    
                    Exit For
                End If
            Next m
            
            If IsNull(extract_bridge(j, dim_bridge_description)) = False Then
                Description = extract_bridge(j, dim_bridge_description)
            Else
                If IsNull(extract_bridge(j, dim_bridge_pict_exec_description)) = False Then
                    Description = extract_bridge(j, dim_bridge_pict_exec_description)
                End If
            End If
            
            Exit For
        End If
    Next j
    
    
    'remonte la quotity d'equity/index database
    For m = 1 To UBound(extract_instrument, 1)
        If extract_instrument(m, dim_instrument_id) = instrument_id Then
            If extract_instrument(m, dim_instrument_name) = "STOCK" Then
                price_factor = 1
                found_quotity_in_local_db = True
            ElseIf extract_instrument(m, dim_instrument_name) = "FUTURE" Then
                For n = 0 To UBound(vec_internal_db_id, 1)
                    If underlying_id = vec_internal_db_id(n) Then
                        price_factor = vec_internal_db_quotity(n)
                        found_quotity_in_local_db = True
                        Exit For
                    End If
                Next n
            ElseIf extract_instrument(m, dim_instrument_name) = "OPTION" Then
                For n = 0 To UBound(vec_internal_db_id, 1)
                    If underlying_id = vec_internal_db_id(n) Then
                        price_factor = vec_internal_db_option_quotity(n)
                        found_quotity_in_local_db = True
                        Exit For
                    End If
                Next n
            End If
            
            Exit For
        End If
    Next m
    
    
    'si price factor introuvable, prendre dans folio
    If found_quotity_in_local_db = False Then
        For j = 0 To UBound(matrix_folio, 1)
            If extract_net_pos_t_1(i, dim_pos_mv_id) = matrix_folio(j, dim_folio_id) Then
                
                
                If found_quotity_in_local_db = False Then
                    price_factor = matrix_folio(j, dim_folio_price_factor)
        
                End If
    
            End If
        Next j
    End If
    
    
    'correction des prices factor des exceptions
    For m = 0 To UBound(extract_exception_gva_price_factor, 1)
        If underlying_id = extract_exception_gva_price_factor(m, 0) Or extract_net_pos_t_1(i, dim_pos_mv_id) = extract_exception_gva_price_factor(m, 0) Then
            price_factor = price_factor * extract_exception_gva_price_factor(m, 1)
        End If
    Next m
    
    
    'prendre le prix dans gva exporté dans table
    For j = 0 To UBound(extract_close_price, 1)
        If extract_net_pos_t_1(i, dim_pos_mv_id) = extract_close_price(j, dim_histo_price_product) And date_t_1 = extract_close_price(j, dim_histo_price_date) Then
            price_t_1 = extract_close_price(j, dim_histo_price_close_price_gva)
            'extract_close_price = Array_remove_item(extract_close_price, j)
            Exit For
        Else
            If j = UBound(extract_close_price, 1) Then
                'MsgBox ("price not found")
            End If
        End If
    Next j
    
    'calcul du ytd pnl t-1
    ytd_pnl_t_1 = (-price_factor * extract_net_pos_t_1(i, dim_pos_mv_market_value) + (extract_net_pos_t_1(i, dim_pos_mv_net_pos) * price_t_1 * price_factor)) + extract_net_pos_t_1(i, dim_pos_mv_ytd_pnl_reversal) + extract_net_pos_t_1(i, dim_pos_mv_ytd_custom_reversal_ytd_pnl)
    
    
    
    ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ trouver la valoristation à t-2
    For j = 0 To UBound(extract_net_pos_t_2, 1)
        If extract_net_pos_t_1(i, dim_pos_mv_id) = extract_net_pos_t_2(j, dim_pos_mv_id) Then
            
            'remonte le prix de la table gva
            For m = 0 To UBound(extract_close_price, 1)
                If extract_net_pos_t_2(j, dim_pos_mv_id) = extract_close_price(m, dim_histo_price_product) And date_t_2 = extract_close_price(m, dim_histo_price_date) Then
                    price_t_2 = extract_close_price(m, dim_histo_price_close_price_gva)
                    'extract_close_price = Array_remove_item(extract_close_price, m)
                    
                    ytd_pnl_t_2 = (-price_factor * extract_net_pos_t_2(j, dim_pos_mv_market_value) + (extract_net_pos_t_2(j, dim_pos_mv_net_pos) * price_t_2 * price_factor)) + extract_net_pos_t_2(j, dim_pos_mv_ytd_pnl_reversal) + extract_net_pos_t_2(j, dim_pos_mv_ytd_custom_reversal_ytd_pnl)
                    pos_t_2 = extract_net_pos_t_2(j, dim_pos_mv_net_pos)
                    
                    Exit For
                End If
            Next m
            
            
            'retire la ligne pour acceler les calculs futurs
            'extract_net_pos_t_2 = Array_remove_item(extract_net_pos_t_2, j)
            Exit For
        Else
            If j = UBound(extract_net_pos_t_2, 1) Then
                'aucune pos trouvée la veille de la veille
                ytd_pnl_t_2 = 0
                pos_t_2 = 0
            End If
        End If
    Next j
    
    For j = 0 To UBound(crncy_rate, 1)
        If extract_net_pos_t_1(i, dim_pos_mv_ytd_crncy_code) = crncy_rate(j)(0) Then
            crncy_conv_rate = crncy_rate(j)(2)
        End If
    Next j
    
    ReDim Preserve output_report(i)
    'output_report(i) = Array(extract_net_pos_t_1(i, dim_pos_mv_id), underlying_id, description, pos_t_2, extract_net_pos_t_1(i, dim_pos_mv_net_pos), price_t_2, price_t_1, ytd_pnl_t_2, ytd_pnl_t_1, ytd_pnl_t_1 - ytd_pnl_t_2, crncy_conv_rate * (ytd_pnl_t_1 - ytd_pnl_t_2))
    output_report(i) = Array(extract_net_pos_t_1(i, dim_pos_mv_id), underlying_id, Description, pos_t_2, extract_net_pos_t_1(i, dim_pos_mv_net_pos), price_t_2, price_t_1, crncy_conv_rate * ytd_pnl_t_2, crncy_conv_rate * ytd_pnl_t_1, crncy_conv_rate * (ytd_pnl_t_1 - ytd_pnl_t_2))
    
Next i


For j = 0 To 10
    Worksheets("Cointrin").Columns(j + 200).Clear
Next j

'header
Worksheets("Cointrin").Cells(8, 200) = "product_id"
Worksheets("Cointrin").Cells(8, 201) = "underlying_id"
Worksheets("Cointrin").Cells(8, 202) = "description"
Worksheets("Cointrin").Cells(8, 203) = "pos_t-2"
    Worksheets("Cointrin").Columns(203).NumberFormat = "0"

Worksheets("Cointrin").Cells(8, 204) = "pos_t-1"
    Worksheets("Cointrin").Columns(204).NumberFormat = "0"

Worksheets("Cointrin").Cells(8, 205) = "close_price_t-2"
    Worksheets("Cointrin").Columns(205).NumberFormat = "#,##0.00"
    
Worksheets("Cointrin").Cells(8, 206) = "close_price_t-1"
    Worksheets("Cointrin").Columns(206).NumberFormat = "#,##0.00"
    
Worksheets("Cointrin").Cells(8, 207) = "ytd_pnl_t-2_EUR"
    Worksheets("Cointrin").Columns(207).NumberFormat = "0"
    
Worksheets("Cointrin").Cells(8, 208) = "ytd_pnl_t-1_EUR"
    Worksheets("Cointrin").Columns(208).NumberFormat = "0"
    
Worksheets("Cointrin").Cells(8, 209) = "daily_pnl_EUR"
    Worksheets("Cointrin").Columns(209).NumberFormat = "0"


'impression des résultats
For i = 1 To UBound(output_report, 1)
    For j = 0 To 9
        Worksheets("Cointrin").Cells(8 + i, j + 200) = output_report(i)(j)
    Next j
Next i



'seconde partie agregee
Dim vec_distinct_u_id() As Variant
ReDim Preserve vec_distinct_u_id(0)
k = 0
For i = 1 To UBound(output_report, 1)
    For j = 0 To UBound(vec_distinct_u_id, 1)
        If output_report(i)(1) = vec_distinct_u_id(j) Then
            Exit For
        Else
            If j = UBound(vec_distinct_u_id, 1) Then
                ReDim Preserve vec_distinct_u_id(k)
                vec_distinct_u_id(k) = output_report(i)(1)
                k = k + 1
            End If
        End If
    Next j
Next i



Dim agreg_report() As Variant
ReDim agreg_report(UBound(vec_distinct_u_id, 1))
Dim dim_report_underlying_id As Integer, dim_report_description As Integer, dim_report_daily_pnl_eur As Integer

For i = 0 To UBound(vec_distinct_u_id, 1)
    agreg_report(i) = Array(vec_distinct_u_id(i), "", 0)
Next i


dim_report_underlying_id = 0
dim_report_description = 1
dim_report_daily_pnl_eur = 2

Dim name_product As String
For i = 0 To UBound(vec_distinct_u_id, 1)
    
    name_product = ""
    
    For j = 1 To UBound(output_report, 1)
        If vec_distinct_u_id(i) = output_report(j)(1) Then
            
            If name_product = "" Then
                For k = 1 To UBound(extract_bridge, 1)
                    If extract_bridge(k, 0) = vec_distinct_u_id(i) Then
                        name_product = extract_bridge(k, 2)
                        
                        If name_product = "" Then
                            name_product = extract_bridge(k, 3)
                        End If
                        
                        Exit For
                    End If
                Next k
            End If
            
            If name_product = "" Then
                name_product = output_report(j)(2)
            End If
            
            'agreg_report(i) = Array(output_report(j)(1), name_product, agreg_report(i)(2) + output_report(j)(9))
            agreg_report(i)(1) = name_product
            agreg_report(i)(2) = agreg_report(i)(2) + output_report(j)(9)
        End If
    Next j
Next i


'impression du rapport
For j = 0 To 2
    Worksheets("Cointrin").Columns(j + 211).Clear
Next j


'header
Worksheets("Cointrin").Cells(8, 211) = "underlying_id"
Worksheets("Cointrin").Cells(8, 212) = "description"
Worksheets("Cointrin").Cells(8, 213) = "daily_PnL_EUR"
    Worksheets("Cointrin").Columns(213).NumberFormat = "#,##0"

For i = 0 To UBound(agreg_report, 1)
    Worksheets("Cointrin").Cells(8 + 1 + i, 211) = agreg_report(i)(0)
    Worksheets("Cointrin").Cells(8 + 1 + i, 212) = agreg_report(i)(1)
    Worksheets("Cointrin").Cells(8 + 1 + i, 213) = agreg_report(i)(2)
Next i


End Sub


Public Function get_trader_code() As Integer

Dim trader_txt As String, trader_prenom As String, trader_nom As String
trader_txt = Worksheets("Cointrin").Cells(5, 2)

trader_prenom = Left(trader_txt, InStr(trader_txt, " ") - 1)
trader_nom = Mid(trader_txt, Len(trader_prenom) + 2)

Dim sql_query As String
sql_query = "SELECT system_code FROM t_trader WHERE system_first_name=""" & trader_prenom & """ AND system_surname=""" & trader_nom & """"
Dim extract_trader As Variant
extract_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_trader, 1) > 0 Then
    get_trader_code = extract_trader(1, 0)
Else
    get_trader_code = 0
End If

End Function

'Public Function get_currency_code(ByVal crncy_name As String) As Integer
'
''remonte le currency code
'    sql_query = "SELECT system_code FROM t_currency WHERE system_name=""" & crncy_name & """"
'    Dim extract_currency As Variant, currency_code As Integer
'    extract_currency = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
'
'    If UBound(extract_currency, 1) > 0 Then
'        get_currency_code = extract_currency(1, 0)
'    Else
'        Exit Function
'    End If
'
'End Function


Public Function get_broker_code(ByVal gs_broker_code_name As String) As Integer

    'remonte le broker code
    sql_query = "SELECT system_code FROM t_broker WHERE gs_code=""" & gs_broker_code_name & """"
    Dim extract_broker_code As Variant, broker_code As Integer
    extract_broker_code = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    If UBound(extract_broker_code, 1) > 0 Then
        get_broker_code = extract_broker_code(1, 0)
    Else
        Exit Function
    End If

End Function


Public Function get_instrument_id(ByVal product_id As String) As Integer

'remonte le l'instrument id
    sql_query = "SELECT system_instrument_id FROM t_bridge WHERE gs_id=""" & product_id & """"
    Dim extract_instrument_id As Variant, instrument_id As Integer
    extract_instrument_id = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    If UBound(extract_instrument_id, 1) > 0 Then
        get_instrument_id = extract_instrument_id(1, 0)
    Else
        Exit Function
    End If

End Function


Public Function get_rate(ByVal crncy As Variant) As Double

Dim i As Integer

Application.Calculation = xlCalculationManual
Dim c_concern As Integer
If IsNumeric(crncy) Then
    'crcny_code
    c_concern = 5
Else
    'currency_txt
    c_concern = 1
End If



For i = 14 To 32
    If Worksheets("Parametres").Cells(i, c_concern) = crncy Then
        get_rate = Worksheets("Parametres").Cells(i, 6)
        Exit For
    End If
Next i

End Function




Sub reload_product_data_in_cointrin(ByVal product_id As String)

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, sql_query As String, trader_code As Integer

trader_code = get_trader_code()
If trader_code = 0 Then
    Exit Sub
End If



'remonte les market value + net pos
Dim extract_position_market_value As Variant

sql_query = "SELECT [t_trade].[gs_security_id] AS PB_ID, [t_trading_account].[system_trader_code] AS trader_code, [t_trade].[system_currency_code] AS crncy_code, SUM([t_trade].[gs_exec_qty]) AS NetPos, Sum([t_trade].[gs_exec_qty]*[t_trade].[gs_exec_price]) AS MarketValue, SUM([t_trade].[system_commission_local_currency]+[t_trade].[system_comm_reversal]) AS COMT, Sum([t_trade].[system_ytd_pnl_reversal]+[t_trade].[system_comm_reversal]) AS Reversal_YTD_PnL, SUM([t_trade].[system_custom_ytd_pnl_reversal]) AS CUSTOM_reversal_ytd_pnl"
sql_query = sql_query & " FROM t_trade INNER JOIN t_trading_account ON [t_trade].[gs_trading_account] = [t_trading_account].[gs_account_number]"
sql_query = sql_query & " WHERE [t_trade].[gs_security_id]=""" & product_id & """"
sql_query = sql_query & " GROUP BY [t_trade].[gs_security_id], [t_trading_account].[system_trader_code], [t_trade].[system_currency_code]"
sql_query = sql_query & " HAVING [t_trading_account].[system_trader_code]=" & trader_code
sql_query = sql_query & " ORDER BY [t_trade].[gs_security_id]"

extract_position_market_value = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim dim_pos_mv_id As Integer, dim_pos_mv_trading_account As Integer, dim_pos_mv_net_pos As Integer, _
    dim_pos_mv_market_value As Integer, dim_pos_mv_com As Integer, dim_pos_mv_ytd_pnl_reversal As Integer, _
    dim_pos_mv_ytd_crncy_code As Integer, dim_pos_mv_ytd_custom_reversal_ytd_pnl As Integer

For i = 0 To UBound(extract_position_market_value, 2)
    If extract_position_market_value(0, i) = "PB_ID" Then
        dim_pos_mv_id = i
    ElseIf extract_position_market_value(0, i) = "NetPos" Then
        dim_pos_mv_net_pos = i
    ElseIf extract_position_market_value(0, i) = "MarketValue" Then
        dim_pos_mv_market_value = i
    ElseIf extract_position_market_value(0, i) = "COMT" Then
        dim_pos_mv_com = i
    ElseIf extract_position_market_value(0, i) = "Reversal_YTD_PnL" Then
        dim_pos_mv_ytd_pnl_reversal = i
    ElseIf extract_position_market_value(0, i) = "CUSTOM_reversal_ytd_pnl" Then
        dim_pos_mv_ytd_custom_reversal_ytd_pnl = i
    ElseIf extract_position_market_value(0, i) = "crncy_code" Then
        dim_pos_mv_ytd_crncy_code = i
    End If
Next i



'remonte les colonnes du rapport
Dim l_cointrin_header As Integer
l_cointrin_header = 8

Dim c_cointrin_product_id As Integer, c_cointrin_net_pos As Integer, c_cointrin_close_price As Integer

c_cointrin_product_id = 0

For i = 1 To 50
    If Worksheets("Cointrin").Cells(l_cointrin_header, i) = "Identifier" And c_cointrin_product_id = 0 Then
        c_cointrin_product_id = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "net_position" Then
        c_cointrin_net_pos = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "gva_close_price" Then
        c_cointrin_close_price = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "ytd_pnl_gross" Then
        c_cointrin_ytd_pnl_gross = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "comm" Then
        c_cointrin_comm_local = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "ytd_pnl_net_local" Then
        c_cointrin_ytd_pnl_net_local = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "ytd_pnl_net_base" Then
        c_cointrin_ytd_pnl_net_base = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "folio_position" Then
        c_cointrin_folio_pos = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "delta_pos_folio" Then
        c_cointrin_delta_pos_folio = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "factor" Then
        c_cointrin_price_factor = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "comm_base" Then
        c_cointrin_comm_base = i
    End If
Next i


'remonte la ligne concernée
Dim l_concern As Integer
l_concern = 0
For i = l_cointrin_header + 1 To 32000
    If Worksheets("Cointrin").Cells(i, c_cointrin_product_id) = "" Then
        Exit Sub
    Else
        If Worksheets("Cointrin").Cells(i, c_cointrin_product_id) = product_id Then
            l_concern = i
            Exit For
        End If
    End If
Next i

If l_concern = 0 Then
    Exit Sub
End If


'remonte le price factor
price_factor = 1
price_factor = Worksheets("Cointrin").Cells(l_concern, c_cointrin_price_factor)
price_close = Worksheets("Cointrin").Cells(l_concern, c_cointrin_close_price)

If price_close = 0 Then
    price_close = CDbl(InputBox("price unavailable", "Error", 0))
End If


Dim ytd_pnl_gross As Double, commt_local As Double, ytd_pnl_net_local As Double, ytd_pnl_net_base As Double, comm_base As Double
ytd_pnl_gross = (-price_factor * extract_position_market_value(1, dim_pos_mv_market_value) + (extract_position_market_value(1, dim_pos_mv_net_pos) * price_close * price_factor)) + extract_position_market_value(1, dim_pos_mv_ytd_pnl_reversal) + extract_position_market_value(1, dim_pos_mv_ytd_custom_reversal_ytd_pnl)
commt_local = extract_position_market_value(1, dim_pos_mv_com)

Dim chg_rate As Double
chg_rate = get_rate(extract_position_market_value(1, dim_pos_mv_ytd_crncy_code))

ytd_pnl_net_local = ytd_pnl_gross - commt_local
ytd_pnl_net_base = ytd_pnl_net_local * chg_rate

comm_base = commt_local * chg_rate


'edition des valeurs en place
Worksheets("Cointrin").Cells(l_concern, c_cointrin_close_price) = price_close
Worksheets("Cointrin").Cells(l_concern, c_cointrin_net_pos) = extract_position_market_value(1, dim_pos_mv_net_pos)
Worksheets("Cointrin").Cells(l_concern, c_cointrin_ytd_pnl_gross) = ytd_pnl_gross
Worksheets("Cointrin").Cells(l_concern, c_cointrin_comm_local) = commt_local
Worksheets("Cointrin").Cells(l_concern, c_cointrin_comm_base) = comm_base
Worksheets("Cointrin").Cells(l_concern, c_cointrin_ytd_pnl_net_local) = ytd_pnl_net_local
Worksheets("Cointrin").Cells(l_concern, c_cointrin_ytd_pnl_net_base) = ytd_pnl_net_base

Worksheets("Cointrin").Cells(l_concern, c_cointrin_delta_pos_folio) = Abs(extract_position_market_value(1, dim_pos_mv_net_pos) - Worksheets("Cointrin").Cells(l_concern, c_cointrin_folio_pos))

If Abs(extract_position_market_value(1, dim_pos_mv_net_pos) - Worksheets("Cointrin").Cells(l_concern, c_cointrin_folio_pos)) > 0.1 Then
    Worksheets("Cointrin").Cells(l_concern, c_cointrin_delta_pos_folio).Interior.ColorIndex = 3
End If

Application.Calculation = xlCalculationAutomatic

End Sub


Sub load_trade_in_form(ByVal unique_id As String)

Dim sql_query As String
sql_query = "SELECT * FROM t_trade WHERE gs_unique_id=""" & unique_id & """"
Dim extract_trade As Variant
extract_trade = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_trade, 1) = 0 Then
    Exit Sub
End If

'repere les dim
For i = 0 To UBound(extract_trade, 2)
    If extract_trade(0, i) = "gs_unique_id" Then
        dim_trade_id = i
    ElseIf extract_trade(0, i) = "gs_security_id" Then
        dim_trade_security_id = i
    ElseIf extract_trade(0, i) = "gs_exec_qty" Then
        dim_trade_qty = i
    ElseIf extract_trade(0, i) = "gs_exec_price" Then
        dim_trade_price = i
    ElseIf extract_trade(0, i) = "gs_trading_account" Then
        dim_trade_trading_account = i
    ElseIf extract_trade(0, i) = "gs_user_id" Then
        dim_trade_user_id = i
    ElseIf extract_trade(0, i) = "gs_exec_broker" Then
        dim_trade_broker_txt = i
    ElseIf extract_trade(0, i) = "system_currency_code" Then
        dim_trade_crncy_code = i
    ElseIf extract_trade(0, i) = "system_broker_id" Then
        dim_trade_broker_code = i
    ElseIf extract_trade(0, i) = "system_commission_local_currency" Then
        dim_trade_commission_local = i
    ElseIf extract_trade(0, i) = "system_trader_code" Then
        dim_trade_trader_code = i
    End If
Next i


'charge les donnees du bridge
sql_query = "SELECT gs_id, gs_underlying_id, gs_description, gs_pict_exec_description FROM t_bridge WHERE gs_id=""" & extract_trade(1, dim_trade_security_id) & """"
Dim extract_bridge As Variant
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_bridge, 1) = 0 Then
    Exit Sub
End If


sql_query = "SELECT system_name FROM t_currency WHERE system_code=" & extract_trade(1, dim_trade_crncy_code)
Dim extract_currency As Variant
extract_currency = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_currency, 1) = 0 Then
    Exit Sub
End If

'mise en place des valeurs dans le form
frm_Cointrin_add_edit_trade.OB_edit.Value = True


frm_Cointrin_add_edit_trade.TB_unique_id.Value = extract_trade(1, dim_trade_id)


If extract_trade(1, dim_trade_id) <> extract_bridge(1, 1) Then
    sql_query = "SELECT gs_description, gs_pict_exec_description FROM t_bridge WHERE gs_id=""" & extract_bridge(1, 1) & """"
    Dim extract_underlying As Variant
    extract_underlying = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    
    If extract_underlying(1, 0) <> "" Then
        frm_Cointrin_add_edit_trade.CB_underlying.Value = extract_underlying(1, 0)
    ElseIf extract_underlying(1, 1) <> "" Then
        frm_Cointrin_add_edit_trade.CB_underlying.Value = extract_underlying(1, 1)
    End If
End If


'frm_Cointrin_add_edit_trade.CB_underlying.Value = extract_trade(1, dim_trade_id)
If extract_bridge(1, 2) <> "" Then
    frm_Cointrin_add_edit_trade.CB_product.Value = extract_bridge(1, 2)
    
    If extract_trade(1, dim_trade_id) = extract_bridge(1, 1) Then
        frm_Cointrin_add_edit_trade.CB_product.Value = extract_bridge(1, 2)
    End If
ElseIf extract_bridge(1, 3) <> "" Then
    frm_Cointrin_add_edit_trade.CB_product.Value = extract_bridge(1, 3)
    
    If extract_trade(1, dim_trade_id) = extract_bridge(1, 1) Then
        frm_Cointrin_add_edit_trade.CB_product.Value = extract_bridge(1, 3)
    End If
End If

frm_Cointrin_add_edit_trade.TB_underlying_id.Value = extract_bridge(1, 1)
frm_Cointrin_add_edit_trade.TB_product_id.Value = extract_trade(1, dim_trade_security_id)

frm_Cointrin_add_edit_trade.TB_qty.Value = extract_trade(1, dim_trade_qty)
frm_Cointrin_add_edit_trade.TB_price.Value = extract_trade(1, dim_trade_price)

frm_Cointrin_add_edit_trade.CB_broker.Value = extract_trade(1, dim_trade_broker_txt)
frm_Cointrin_add_edit_trade.TB_comissions.Value = extract_trade(1, dim_trade_commission_local)



frm_Cointrin_add_edit_trade.TB_crncy.Value = extract_currency(1, 0)

frm_Cointrin_add_edit_trade.TB_trader_id_r_plus.Value = extract_trade(1, dim_trade_user_id)
frm_Cointrin_add_edit_trade.TB_trader_code.Value = extract_trade(1, dim_trade_trader_code)
frm_Cointrin_add_edit_trade.TB_account.Value = extract_trade(1, dim_trade_trading_account)


frm_Cointrin_add_edit_trade.Show


End Sub


Sub load_tag_in_form()

If ActiveSheet.name = "Open" And ActiveCell.row > 25 Then
    Call load_tag(Cells(ActiveCell.row, 1))
End If

End Sub

Sub load_tag(ByVal product_id As String)

Dim i As Integer, j As Integer
Dim date_tmp As Date
date_tmp = FormatDateTime(Date, vbShortDate)

Dim c_ed_tag As Integer
c_ed_tag = 137
Dim tag_json As String
tag_json = ""

Application.Calculation = xlCalculationManual

'regarde s'il existe deja un tag pour la position desiree
For i = 27 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        MsgBox ("product_id=" & product_id & " doesn't exist in equity_db")
        Exit Sub
    Else
        If Worksheets("Equity_Database").Cells(i, 1) = product_id Then
            tag_json = Worksheets("Equity_Database").Cells(i, c_ed_tag)
            
            frm_Tag.CB_tag.Value = ""
            frm_Tag.TB_name.Value = CStr(Worksheets("Equity_Database").Cells(i, 2).Value)
            frm_Tag.TB_product_id.Value = CStr(Worksheets("Equity_Database").Cells(i, 1).Value)
            
            frm_Tag.TB_delta.Value = CStr(Round(Worksheets("Equity_Database").Cells(i, 6).Value, 0))
            frm_Tag.TB_result_ytd.Value = CStr(Round(Worksheets("Equity_Database").Cells(i, 14).Value, 2))
            
            'frm_tag.TB_control_date = date_tmp
            
            Exit For
        End If
    End If
Next i

If tag_json <> "" Then
    'un tag est present conversion de str en collection
    Dim oJSON As New JSONLib
    Dim oTags As Collection
    Set oTags = oJSON.parse(tag_json)
    Dim oTag As Collection
    
    frm_Tag.LB_tag_already.Clear
    k = 0
    For Each oTag In oTags
        
        frm_Tag.LB_tag_already.AddItem
        
        For i = 1 To oTag.count
            frm_Tag.LB_tag_already.list(k, i - 1) = oTag.Item(i)
        Next i
        
        k = k + 1
    Next
End If

frm_Tag.Show

End Sub


Public Function get_marketplace_redi_plus(ByVal ticker As String) As String

Dim matrix_marketplace As Variant
matrix_marketplace = Array(Array("VX", "VRTS"), Array("SW", "EBS"), Array("GY", "XETR"), Array("GR", "XETR"), _
    Array("FP", "PARE"), Array("NO", "OSLE"), Array("NA", "AMSE"), Array("IM", "MILE"), Array("LN", "LNSE"), _
    Array("FH", "HELE"), Array("DC", "CPHE"), Array("DC", "CPHE"), Array("SM", "SIBE"), Array("LI", "LNIO"), _
    Array("SS", "STOE"), Array("US", "SIGMA"), Array("SI", "SINE"))


ticker = UCase(ticker)

Dim debug_test As Variant
Dim marketplace_bbg As String
marketplace_bbg = ""

Dim reg As New RegExp
reg.Pattern = "[\s][A-Z]{2}[\s]"
Dim match As match
Dim matches As MatchCollection

Set matches = reg.Execute(ticker)
Dim str_marketplace_bbg As String

For Each match In matches
    'str_marketplace_bbg = match.value
    For i = 1 To Len(match.Value)
        If Mid(match.Value, i, 1) <> " " Then
            marketplace_bbg = marketplace_bbg & Mid(match.Value, i, 1)
        End If
    Next i
    
    Exit For
Next

If Len(marketplace_bbg) = 2 Then
Else
    get_marketplace_redi_plus = "-1"
    Exit Function
End If


For i = 0 To UBound(matrix_marketplace, 1)
    If marketplace_bbg = matrix_marketplace(i)(0) Then
        get_marketplace_redi_plus = matrix_marketplace(i)(1)
        Exit Function
    End If
Next i


End Function



Public Function get_marketplace_and_limit_style_redi_plus(ByVal ticker As String) As Variant

Dim matrix_marketplace_equity As Variant, matrix_marketplace_index As Variant
matrix_marketplace_equity = Array(Array("VX", "VRTX ", "Limit"), Array("SW", "EBS ", "Limit"), Array("GY", "XETR ", "Limit"), _
    Array("GR", "XETR ", "Limit"), Array("FP", "PARE ", "Limit"), Array("NO", "OSLE ", "Limit"), Array("NA", "AMSE ", "Limit"), _
    Array("IM", "MILE ", "Limit"), Array("LN", "LNSE ", "Limit"), Array("FH", "HELE ", "Limit"), Array("DC", "CPHE ", "Limit"), _
    Array("DC", "CPHE ", "Limit"), Array("SM", "SIBE ", "Limit"), Array("SQ", "SIBE ", "Limit"), Array("LI", "LNIO ", "Limit"), _
    Array("SS", "STOE ", "Limit"), Array("US", "SIGMA", "Smart Lmt"), Array("SI", "SINE ", "Limit"), Array("BB", "BRUE ", "Limit"))
    
matrix_marketplace_index = Array(Array("ES", "CME ", "Limit"), Array("VG", "EURXUS ", "Limit"), Array("GX", "EURXUS ", "Limit"), _
    Array("SM", "EURX ", "Limit"), Array("NX", "CME ", "Limit"), Array("HC", "HKFE ", "Limit"), Array("UX", "CFE ", "Limit"))


ticker = UCase(ticker)

Dim debug_test As Variant



If Right(ticker, 6) = "EQUITY" Then
    
    Dim reg As New RegExp
    Dim marketplace_bbg As String
    marketplace_bbg = ""

    reg.Pattern = "[\s][A-Z]{2}[\s]"
    Dim match As match
    Dim matches As MatchCollection
    
    Set matches = reg.Execute(ticker)
    Dim str_marketplace_bbg As String
    
    For Each match In matches
        'str_marketplace_bbg = match.value
        For i = 1 To Len(match.Value)
            If Mid(match.Value, i, 1) <> " " Then
                marketplace_bbg = marketplace_bbg & Mid(match.Value, i, 1)
            End If
        Next i
        
        Exit For
    Next
    
    
    For i = 0 To UBound(matrix_marketplace_equity, 1)
        If marketplace_bbg = matrix_marketplace_equity(i)(0) Then
            get_marketplace_and_limit_style_redi_plus = Array(matrix_marketplace_equity(i)(1), matrix_marketplace_equity(i)(2))
            Exit Function
        End If
    Next i
    
    If Len(marketplace_bbg) = 2 Then
    
    Else
        get_marketplace_and_limit_style_redi_plus = Array("-1", "-1")
        Exit Function
    End If
    
ElseIf Right(ticker, 5) = "INDEX" Then
    
    For i = 0 To UBound(matrix_marketplace_index, 1)
        If Left(ticker, Len(matrix_marketplace_index(i)(0))) = matrix_marketplace_index(i)(0) Then
            get_marketplace_and_limit_style_redi_plus = Array(matrix_marketplace_index(i)(1), matrix_marketplace_index(i)(2))
            Exit Function
        Else
            If i = UBound(matrix_marketplace_index, 1) Then
                get_marketplace_and_limit_style_redi_plus = Array("-1", "-1")
                Exit Function
            End If
        End If
    Next i
    
End If


End Function



Public Function get_symbol_redi_plus(ByVal ticker As String) As String

Dim i As Integer, j As Integer, k As Integer
ticker = UCase(ticker)

get_symbol_redi_plus = ticker

Dim match_index_ric As Variant
match_index_ric = Array(Array("ES", "ES"), Array("VG", "STXE"), Array("GX", "FDX"), Array("SM", "FSMI"), _
    Array("NX", "NKD"), Array("HC", "HCEI"), Array("UX", "VX"))

If Right(ticker, 6) = "EQUITY" Then
    
    get_symbol_redi_plus = Replace(UCase(ticker), " EQUITY", "")
    Exit Function
    
ElseIf Right(ticker, 5) = "INDEX" Then
    
    For i = 0 To UBound(match_index_ric, 1)
        If Left(ticker, Len(match_index_ric(i)(0))) = match_index_ric(i)(0) Then
            
            ticker = match_index_ric(i)(1) & Mid(ticker, Len(match_index_ric(i)(0)) + 1)
            get_symbol_redi_plus = Left(ticker, InStr(ticker, " ") - 1)
            
            Exit For
        End If
    Next i
    
End If


End Function


Public Function get_stocks_amount(ByVal ticker_or_product_id As String) As Double

Dim i As Integer, j As Integer, k As Integer

Dim c_concern As Integer

If Left(ticker_or_product_id, 2) <> "P=" Then
    c_concern = 47
Else
    c_concern = 1
End If


If UCase(Right(ticker_or_product_id, 6)) = "EQUITY" Or Left(ticker_or_product_id, 2) = "P=" Then
    For i = 27 To 32000 Step 2
        If Workbooks("Kronos.xls").Worksheets("Equity_Database").Cells(i, 1) = "" Then
            get_stocks_amount = 0
            Exit Function
        Else
            If UCase(Workbooks("Kronos.xls").Worksheets("Equity_Database").Cells(i, c_concern)) = UCase(ticker_or_product_id) Then
                If IsError(Workbooks("Kronos.xls").Worksheets("Equity_Database").Cells(i, 24)) Then
                    get_stocks_amount = 0
                Else
                    If IsNumeric(Workbooks("Kronos.xls").Worksheets("Equity_Database").Cells(i, 24)) Then
                        get_stocks_amount = Workbooks("Kronos.xls").Worksheets("Equity_Database").Cells(i, 24)
                    Else
                        get_stocks_amount = 0
                    End If
                End If
                
                Exit For
            End If
        End If
    Next i
ElseIf UCase(Right(ticker_or_product_id, 5)) = "INDEX" Then
    get_stocks_amount = 10000 'fictive pour eviter short sell
End If


End Function

Public Function get_side_redi_plus(ByVal ticker As String, ByVal qty As Double) As String

If qty > 0 Then
    get_side_redi_plus = "Buy"
    Exit Function
Else
    If get_stocks_amount(ticker) <= 0 Then
        get_side_redi_plus = "SS"
    Else
        get_side_redi_plus = "Sell"
    End If
    
    Exit Function
End If

End Function


Public Function get_trade_account_redi_plus(ByVal ticker As String) As String

'remonte les 2 accounts (cash + derivative)
Dim acc_equity As String, acc_derivative As String
acc_equity = Workbooks("Kronos.xls").Worksheets("FORMAT2").Cells(6, 23)
acc_derivative = Workbooks("Kronos.xls").Worksheets("FORMAT2").Cells(6, 24)

ticker = UCase(ticker)

If Right(ticker, 6) = "EQUITY" Then
    get_trade_account_redi_plus = acc_equity
    Exit Function
ElseIf Right(ticker, 5) = "INDEX" Then
    get_trade_account_redi_plus = acc_derivative
    Exit Function
End If

End Function



Public Function get_redi_info_from_book_format2() As Collection

Set get_redi_info_from_book_format2 = New Collection

    get_redi_info_from_book_format2.Add Worksheets("FORMAT2").Cells(7, 21), key_r_plus_user
    get_redi_info_from_book_format2.Add Worksheets("FORMAT2").Cells(8, 21), key_r_plus_password
    
    get_redi_info_from_book_format2.Add Worksheets("FORMAT2").Cells(6, 23), key_r_plus_cash_account
    get_redi_info_from_book_format2.Add Worksheets("FORMAT2").Cells(6, 24), key_r_plus_derivatives_account


End Function


Public Sub prepare_line_for_trading_with_r_plus()

Dim oBBG As New cls_Bloomberg_Sync


Dim ticker_to_trade As String
ticker_to_trade = find_security_on_the_line(ActiveSheet.name, ActiveCell.row, ActiveCell.column)

If ticker_to_trade <> "-1" Then
    
    'descend last price
    Dim bbg_last_price As Variant
    bbg_last_price = oBBG.bdp(Array(ticker_to_trade), Array("last_price"), output_format.of_vec_without_header)
    
    If IsNumeric(bbg_last_price(0)(0)) Then
        frm_redi_plus.TB_price.Value = Round(bbg_last_price(0)(0), 2)
    End If
    
    
    frm_redi_plus.TB_ticker.Value = ticker_to_trade
    frm_redi_plus.Show
End If

End Sub


Public Sub prepare_line_for_trading_with_r_plus_advanced()

Dim ticker_to_trade As String
ticker_to_trade = find_security_on_the_line(ActiveSheet.name, ActiveCell.row, ActiveCell.column)

If ticker_to_trade <> "-1" Then
    frm_redi_plus_advanced.TB_order_ticker.Value = ticker_to_trade
    frm_redi_plus_advanced.Show
End If

End Sub



Public Function find_security_on_the_line(ByVal active_worksheet As String, ByVal active_row As Integer, active_colum As Integer) As String

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer

Dim product_id As String
Dim underlying_id As String

Dim found_security As Boolean
found_security = False

Dim most_active_fut As String


If active_worksheet = "Open" And active_row > 25 Then

    'find_security_on_the_line = UCase(Worksheets(active_worksheet).Cells(active_row, 104))
    
    product_id = Worksheets(active_worksheet).Cells(active_row, 2)
    underlying_id = Worksheets(active_worksheet).Cells(active_row, 1)
    
    's'il s'agit d'un index -> trouver le FUT
    'If Right(UCase(find_security_on_the_line), 5) = "INDEX" Then
    If Worksheets("Open").Cells(active_row, 4) = "I" Then
        found_security = False
        
        For i = 27 To 500 Step 3
            If Worksheets("Index_Database").Cells(i, 1) = "" Then
                Exit For
            Else
                If Worksheets("Index_Database").Cells(i, 1) = underlying_id Then
                    
                    most_active_fut = UCase(Worksheets("Index_Database").Cells(i, 34))
                    
                    If Worksheets("Index_Database").Cells(i, 31) = product_id Then
                        find_security_on_the_line = UCase(Worksheets("Index_Database").Cells(i, 34))
                        
                        found_security = True
                        Exit For
                    
                    ElseIf Worksheets("Index_Database").Cells(i, 32) = product_id Then
                        find_security_on_the_line = UCase(Worksheets("Index_Database").Cells(i + 1, 34))
                        found_security = True
                        Exit For
                    Else
                        'l'appel a ete fait depuis une ligne d'option -> prendre le fut le plus recent
                        find_security_on_the_line = UCase(most_active_fut)
                        found_security = True
                    End If
                    
                End If
            End If
        Next i
    ElseIf Worksheets("Open").Cells(active_row, 4) = "E" Then
        
        found_security = False
        
        For i = 27 To 32000 Step 2
            If Worksheets("Equity_Database").Cells(i, 1) = "" Then
                Exit For
            Else
                If Worksheets("Equity_Database").Cells(i, 1) = underlying_id Then
                    find_security_on_the_line = UCase(Worksheets("Equity_Database").Cells(i, 47).Value)
                    found_security = True
                    Exit For
                End If
            End If
        Next i
        
    End If
    
    
Else
    For i = active_colum To 1 Step -1
        If Right(UCase(Worksheets(active_worksheet).Cells(active_row, i)), 6) = "EQUITY" Or Right(UCase(Worksheets(active_worksheet).Cells(active_row, i)), 5) = "INDEX" Then
            find_security_on_the_line = UCase(Worksheets(active_worksheet).Cells(active_row, i))
            found_security = True
            Exit For
        End If
    Next i
    
    If found_security = False Then
        
        For i = active_colum To 250
            If Right(UCase(Worksheets(active_worksheet).Cells(active_row, i)), 6) = "EQUITY" Or Right(UCase(Worksheets(active_worksheet).Cells(active_row, i)), 5) = "INDEX" Then
                find_security_on_the_line = UCase(Worksheets(active_worksheet).Cells(active_row, i))
                found_security = True
                Exit For
            End If
        Next i
        
    End If
    
    
    'si index utiliser le book pour trouver le most active fut
    If found_security = True Then
        If Right(UCase(find_security_on_the_line), 5) = "INDEX" Then
            
            found_security = False
            
            's assure que le book est ouvert
            Dim find_book As Boolean
                find_book = False
            Dim tmp_wrbk As Workbook
            For Each tmp_wrbk In Workbooks
                If UCase(tmp_wrbk.name) = UCase("Kronos.xls") Then
                    find_book = True
                    Exit For
                End If
            Next
            
            If find_book = True Then
                For i = 27 To 500 Step 3
                    If Workbooks("Kronos.xls").Worksheets("Index_Database").Cells(i, 1) = "" Then
                        Exit For
                    Else
                        If UCase(Workbooks("Kronos.xls").Worksheets("Index_Database").Cells(i, 110)) = UCase(find_security_on_the_line) Then
                            find_security_on_the_line = UCase(Workbooks("Kronos.xls").Worksheets("Index_Database").Cells(i, 34))
                            found_security = True
                            Exit For
                        End If
                    End If
                Next i
            Else
                found_security = False
            End If
        End If
    End If
    
    
    
End If

If found_security = False Then
        find_security_on_the_line = "-1"
    End If

End Function



Public Function get_side_redi_plus_optimize_with_lending(ByVal symbol As String, ByVal qty_to_trade As Long, ByVal current_position_in_equity_db As Long, ByVal qty_waiting_buy As Long, ByVal qty_waiting_sell As Long, ByVal qty_order_buy As Long, ByVal qty_order_sell As Long) As String



Dim i As Long, j As Long, k As Long

If qty_to_trade > 0 Then
    get_side_redi_plus_optimize_with_lending = "Buy"
Else
    
    get_side_redi_plus_optimize_with_lending = "SS"
    
    debug_test = current_position_in_equity_db + qty_waiting_sell + qty_to_trade
    If current_position_in_equity_db + qty_waiting_sell + qty_to_trade >= 0 Then
        get_side_redi_plus_optimize_with_lending = "Sell"
    End If
    
'    For i = 0 To UBound(lending_matrix, 1)
'        If lending_matrix(i)(dim_lending_ticker) = symbol Then
'            If lending_matrix(i)(dim_lending_actual_position) <= 0 Then
'                get_side_redi_plus_optimize_with_lending = "SS"
'            Else
'
'                If lending_matrix(i)(dim_lending_worst_case_locate) > 0 Then
'                    'si meme en faisant passer tous les autres la position reste long
'                    get_side_redi_plus_optimize_with_lending = "Sell" ' sell to cover
'                Else
'
'                    'une partie en sell et le sold en SS
'
'
'                End If
'
'            End If
'        End If
'    Next i
    
End If

End Function


Public Function generate_group_id_trade() As Double

Dim tmp_group As Double
Randomize
tmp_group = CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0))

generate_group_id_trade = tmp_group

End Function


Public Sub send_order_with_api_rplus()

Application.Calculation = xlCalculationManual

Dim ticker_index_tmp As String
Dim oJSON As New JSONLib

Dim match_index_ric As Variant
match_index_ric = Array(Array("ES", "ES"), Array("VG", "STXE"), Array("GX", "FDX"), Array("SM", "FSM"), Array("NX", "NKD"), Array("HC", "HCEI"))

Dim qty_shares As Double
qty_shares = Cells(10, 4)

Dim l_first_line As Integer
l_first_line = 14


Dim prefix_src_new_line As String
    prefix_src_new_line = "*** AUTO-GENERATED ***"


Dim c_ticker As Integer, c_price_type As Integer, c_side As Integer, c_qty As Integer, c_limit As Integer, c_TIF As Integer

c_ticker = 1
c_price_type = 4
c_side = 5
c_qty = 6
c_limit = 7
c_TIF = 8

Dim auto_order_folder As String
auto_order_folder = "auto_orders"



Dim ready_color As Integer
ready_color = 42


Dim trades() As Variant, color_lines() As Variant, org_ticker() As Variant
Dim dim_symbol As Integer, dim_side As Integer, dim_qty As Integer, dim_price As Integer, dim_price_type As Integer, dim_TIF As Integer

dim_symbol = 0
dim_side = 1
dim_qty = 2
dim_price = 3
dim_price_type = 4
dim_TIF = 5

Dim vec_trade_universal() As Variant


Dim tmp_ticker As String, tmp_side As String, ref_list As String, tmp_qty As Double

k = 0
For i = l_first_line To 1000
    If Cells(i, c_ticker) = "" And Cells(i + 1, c_ticker) = "" And Cells(i + 2, c_ticker) = "" And Cells(i + 3, c_ticker) = "" Then
        Exit For
    Else
        
        If Cells(i, c_ticker) <> "" And Cells(i, c_ticker).Interior.ColorIndex = ready_color And IsNumeric(Cells(i, c_qty)) And IsNumeric(Cells(i, c_limit)) And (UCase(Left(Cells(i, c_side), 1)) = "B" Or UCase(Left(Cells(i, c_side), 1)) = "S" Or UCase(Left(Cells(i, c_side), 1)) = "H") Then
            
            If Cells(i, c_limit) <> 0 And Cells(i, c_qty) <> 0 Then
            
                If UCase(Left(Cells(i, c_side), 1)) = "B" Or UCase(Left(Cells(i, c_side), 1)) = "C" Then
                    tmp_side = "Buy"
                ElseIf UCase(Left(Cells(i, c_side), 1)) = "S" Or UCase(Left(Cells(i, c_side), 1)) = "H" Then
                    If qty_shares <= 0 Then
                        If UCase(Right(Cells(i, c_ticker), 6)) = "EQUITY" Then
                            tmp_side = "SS"
                        ElseIf UCase(Right(Cells(i, c_ticker), 5)) = "INDEX" Then
                            tmp_side = "Sell"
                        End If
                    Else
                        tmp_side = "Sell"
                    End If
                End If
                
                
                ReDim Preserve color_lines(k)
                color_lines(k) = i
                
                ReDim Preserve trades(k)
                
                ReDim Preserve org_ticker(k)
                org_ticker(k) = UCase(Cells(i, c_ticker))
                
                If InStr(UCase(Cells(i, c_ticker)), "EQUITY") <> 0 Then
                    trades(k) = Array(Left(Cells(i, c_ticker), InStr(InStr(Cells(i, c_ticker), " ") + 1, Cells(i, c_ticker), " ") - 1), tmp_side, Cells(i, c_qty).Value, Cells(i, c_limit).Value, Cells(i, c_price_type).Value, Cells(i, c_TIF).Value)
                ElseIf InStr(UCase(Cells(i, c_ticker)), "INDEX") <> 0 Then
                    
                    ticker_index_tmp = Cells(i, c_ticker)
                    
                    'transforme le ticker bbg en ric reuters
                    For j = 0 To UBound(match_index_ric, 1)
                        If Left(ticker_index_tmp, Len(match_index_ric(j)(0))) = match_index_ric(j)(0) Then
                            
                            'cut le suffixe bbg
                            ticker_index_tmp = Mid(ticker_index_tmp, Len(match_index_ric(j)(0)) + 1)
                            ticker_index_tmp = match_index_ric(j)(1) & ticker_index_tmp
                            
                            trades(k) = Array(Left(ticker_index_tmp, InStr(ticker_index_tmp, " ") - 1), tmp_side, Cells(i, c_qty).Value, Cells(i, c_limit).Value, Cells(i, c_price_type).Value, Cells(i, c_TIF).Value)
                            
                            Exit For
                        End If
                    Next j
                    
                    
                End If
                
                
                If UCase(Left(Cells(i, c_side), 1)) = "B" Or UCase(Left(Cells(i, c_side), 1)) = "C" Then
                    tmp_qty = Cells(i, c_qty)
                ElseIf UCase(Left(Cells(i, c_side), 1)) = "S" Or UCase(Left(Cells(i, c_side), 1)) = "H" Then
                    tmp_qty = -Cells(i, c_qty)
                End If
                
                
                ReDim Preserve vec_trade_universal(k)
                
                If UCase(Cells(i, c_TIF)) = "STP" Or UCase(Cells(i, c_TIF)) = "STOP" Then
                    vec_trade_universal(k) = Array(Cells(i, c_ticker).Value, tmp_qty, Cells(i, c_limit).Value, Cells(i, c_limit).Value)
                Else
                    vec_trade_universal(k) = Array(Cells(i, c_ticker).Value, tmp_qty, Cells(i, c_limit).Value)
                End If
                
                k = k + 1
            
            End If
        End If
    End If
Next i

'second passage pour les warm-up premarket

Dim tmp_group_id As Double, tmp_price As Double, tmp_stop As Variant, tmp_vec_tag() As Variant
Dim tmp_last_ticker As String
    tmp_last_ticker = ""
Dim tmp_last_line As Integer, tmp_last_base_line As Integer
Dim tmp_last_src As String

For i = l_format2_header + 1 To 32000
    If Cells(i, c_ticker) = "" And Cells(i + 1, c_ticker) = "" And Cells(i + 2, c_ticker) = "" And Cells(i + 3, c_ticker) = "" Then
        Exit For
    Else
        If Cells(i, c_ticker) <> "" And Cells(i, c_ticker).Interior.ColorIndex = ready_color And IsNumeric(Cells(i, c_qty)) And IsNumeric(Cells(i, c_limit)) And (UCase(Left(Cells(i, c_side), 1)) = "B" Or UCase(Left(Cells(i, c_side), 1)) = "S" Or UCase(Left(Cells(i, c_side), 1)) = "H") Then
            
            If Cells(i, c_limit) <> 0 And Cells(i, c_qty) <> 0 Then
            
                If UCase(Left(Cells(i, c_side), 1)) = "B" Then
                    tmp_side = "Buy"
                ElseIf UCase(Left(Cells(i, c_side), 1)) = "S" Or UCase(Left(Cells(i, c_side), 1)) = "H" Then
                    If qty_shares <= 0 Then
                        If UCase(Right(Cells(i, c_ticker), 6)) = "EQUITY" Then
                            tmp_side = "SS"
                        ElseIf UCase(Right(Cells(i, c_ticker), 5)) = "INDEX" Then
                            tmp_side = "Sell"
                        End If
                    Else
                        tmp_side = "Sell"
                    End If
                End If
                
                
                ReDim Preserve color_lines(k)
                color_lines(k) = i
                
                ReDim Preserve trades(k)
                
                ReDim Preserve org_ticker(k)
                org_ticker(k) = UCase(Cells(i, c_ticker))
                
                If InStr(UCase(Cells(i, c_ticker)), "EQUITY") <> 0 Then
                    trades(k) = Array(Left(Cells(i, c_ticker), InStr(InStr(Cells(i, c_ticker), " ") + 1, Cells(i, c_ticker), " ") - 1), tmp_side, Cells(i, c_qty).Value, Cells(i, c_limit).Value, Cells(i, c_price_type).Value, Cells(i, c_TIF).Value)
                ElseIf InStr(UCase(Cells(i, c_ticker)), "INDEX") <> 0 Then
                    
                    ticker_index_tmp = Cells(i, c_ticker)
                    
                    'transforme le ticker bbg en ric reuters
                    For j = 0 To UBound(match_index_ric, 1)
                        If Left(ticker_index_tmp, Len(match_index_ric(j)(0))) = match_index_ric(j)(0) Then
                            
                            'cut le suffixe bbg
                            ticker_index_tmp = Mid(ticker_index_tmp, Len(match_index_ric(j)(0)) + 1)
                            ticker_index_tmp = match_index_ric(j)(1) & ticker_index_tmp
                            
                            trades(k) = Array(Left(ticker_index_tmp, InStr(ticker_index_tmp, " ") - 1), tmp_side, Cells(i, c_qty).Value, Cells(i, c_limit).Value, Cells(i, c_price_type).Value, Cells(i, c_TIF).Value)
                            
                            Exit For
                        End If
                    Next j
                    
                    
                End If
                
                If UCase(Left(Cells(i, c_side), 1)) = "B" Or UCase(Left(Cells(i, c_side), 1)) = "C" Then
                    tmp_qty = Cells(i, c_qty)
                ElseIf UCase(Left(Cells(i, c_side), 1)) = "S" Or UCase(Left(Cells(i, c_side), 1)) = "H" Then
                    tmp_qty = -Cells(i, c_qty)
                End If
                
                
                tmp_price = Cells(i, c_format2_price)
                
                
                If UCase(Cells(i, c_TIF)) = "STP" Or UCase(Cells(i, c_TIF)) = "STOP" Then
                    tmp_stop = Cells(i, c_format2_price).Value
                Else
                    tmp_stop = Empty
                End If
                
                
                m = 0
                ReDim Preserve tmp_vec_tag(m)
                If Cells(i, c_format2_source) <> prefix_src_new_line Then
                    tmp_vec_tag(m) = Replace(Cells(i, c_format2_source).Value, " ", "_")
                Else
                    'remonte
                    For j = i - 1 To l_format2_header + 1 Step -1
                        If Cells(i, c_format2_source) <> prefix_src_new_line Then
                            tmp_vec_tag(m) = Replace(Cells(j, c_format2_source).Value, " ", "_")
                            Exit For
                        End If
                    Next j
                End If
                m = m + 1
                
                
                ReDim Preserve tmp_vec_tag(m)
                If Cells(i, c_format2_source) <> prefix_src_new_line Then
                    tmp_vec_tag(m) = "base"
                    tmp_last_base_line = i
                Else
                    tmp_vec_tag(m) = Cells(i, c_format2_s3).Value
                End If
                m = m + 1
                
                
                For j = c_format2_s3 To c_format2_r3
                    If Cells(i, j) <> "" And IsNumeric(Cells(i, j)) Then
                        
                        If Round(Cells(i, j), 2) = Round(Cells(i, c_format2_price), 2) Then
                            ReDim Preserve tmp_vec_tag(m)
                            tmp_vec_tag(m) = Cells(l_format2_header, j).Value
                            m = m + 1
                        End If
                    End If
                Next j
                
                
                If Cells(i, c_format2_ticker).Value <> tmp_last_ticker Then
                    tmp_group_id = generate_group_id_trade
                Else
                    
                    If Cells(i, c_format2_source) <> prefix_src_new_line Then
                        tmp_group_id = generate_group_id_trade
                    Else
                        's agit - il vraiment d un group existant
                        For j = i - 1 To l_format2_header + 1 Step -1
                            If Cells(j, c_format2_source) <> prefix_src_new_line Then
                                
                                If j <> tmp_last_base_line Then
                                    tmp_last_base_line = j
                                    tmp_group_id = generate_group_id_trade
                                End If
                                
                                Exit For
                            End If
                        Next j
                    End If
                    
                End If
                
                
                
                ReDim Preserve vec_trade_universal(k)
                vec_trade_universal(k) = Array(Cells(i, c_format2_ticker).Value, tmp_qty, tmp_price, tmp_stop, tmp_group_id, oJSON.toString(tmp_vec_tag))
                k = k + 1
                
                tmp_last_ticker = Cells(i, c_format2_ticker).Value
                
            End If
        End If
    End If
Next i


Dim tmp_ref_wrb_name As String, tmp_ref_wrb_path As String
If k > 0 Then
    
    Dim exec_order As Variant
    exec_order = universal_trades_r_plus(vec_trade_universal)
    
    
    'retire la couleur
    For i = 0 To UBound(color_lines, 1)
        rows(color_lines(i)).Interior.ColorIndex = xlNone
    Next i
End If


Application.Calculation = xlCalculationAutomatic

End Sub




'reception vec : ticker / qty / price / stop (optional) / group_id / json_tag
Public Function universal_trades_r_plus(ByVal vec_trades As Variant) As Variant

Dim return_order() As Variant
ReDim return_order(UBound(vec_trades, 1))

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim dim_symbol As Integer, dim_side As Integer, dim_qty As Integer, dim_price As Integer, dim_exchange As Integer, _
    dim_price_type As Integer, dim_account As Integer

dim_symbol = 0
dim_qty = 1
dim_price = 2
dim_stop = 3
dim_group_id = 4
dim_json_tag = 5

Call moulinette_init_db

Dim redi_userid As String, redi_password As String
redi_userid = Workbooks("Kronos.xls").Worksheets("FORMAT2").Cells(7, 21)
redi_password = Decrypter(Workbooks("Kronos.xls").Worksheets("FORMAT2").Cells(8, 21))

Dim marketplace_and_limit_type As Variant

Dim hOrder As New RediLib.Order
    Dim myerr As Variant
    Dim retValueOrder As Boolean


'prepartion d'une list de summary pour avoir des msgbox recapitulatives / titre
Dim min_ticker  As String, min_pos As Long, tmp_vec As Variant
For i = 0 To UBound(vec_trades, 1)
    
    'order by symbol
    min_ticker = vec_trades(i)(dim_symbol)
    min_pos = i
    
    For j = i + 1 To UBound(vec_trades, 1)
        If vec_trades(j)(dim_symbol) < min_ticker Then
            min_ticker = vec_trades(j)(dim_symbol)
            min_pos = j
        End If
        
    Next j
    
    If min_pos <> i Then
        tmp_vec = vec_trades(i)
        vec_trades(i) = vec_trades(min_pos)
        vec_trades(min_pos) = tmp_vec
    End If
Next i


Dim msg_summary() As Variant
    dim_summary_ticker = 0
    dim_summary_qty_buy = 1
    dim_summary_qty_sell = 2
    dim_summary_avg_price_buy = 3
    dim_summary_avg_price_sell = 4
    dim_summary_already_show = 5
    dim_summary_answer = 6

Dim tmp_marketvalue As Double, last_ticker As String


Dim lending_report() As Variant
    dim_lending_ticker = 0
    dim_lending_actual_position = 1
    dim_lending_buy_in_redi_order_queue = 2
    dim_lending_sell_in_redi_order_queue = 3
    dim_lending_total_qty_buy = 4
    dim_lending_total_qty_sell = 5
    dim_lending_worst_case_locate = 6


k = -1
last_ticker = ""
Dim tmp_stat_ticker As Variant
    tmp_stat_ticker = Array("", 0, 0)
    
Dim waiting_buy As Long
Dim waiting_sell As Long

For i = 0 To UBound(vec_trades, 1)
    If vec_trades(i)(dim_symbol) <> last_ticker Then
        
        'stat du lending
        If k > -1 Then
            
            'check les hors lending et reajustement des ordres - split
            
            
        End If
        
        
        k = k + 1
        ReDim Preserve msg_summary(k)
        msg_summary(k) = Array(vec_trades(i)(dim_symbol), 0, 0, 0, 0, False, False)
        last_ticker = vec_trades(i)(dim_symbol)
        
        actual_qty = get_stocks_amount(vec_trades(i)(dim_symbol))
        waiting_buy = get_redi_amount_in_queue("orders", get_symbol_redi_plus(vec_trades(i)(dim_symbol)), "B")
        waiting_sell = get_redi_amount_in_queue("orders", get_symbol_redi_plus(vec_trades(i)(dim_symbol)), "S")
        
        ReDim Preserve lending_report(k)
        lending_report(k) = Array(vec_trades(i)(dim_symbol), actual_qty, waiting_buy, waiting_sell, 0, 0, actual_qty + waiting_sell)
        
    Else
        
    End If
        
    If vec_trades(i)(dim_qty) < 0 Then
        'a sell
        tmp_marketvalue = msg_summary(k)(dim_summary_qty_sell) * msg_summary(k)(dim_summary_avg_price_sell)
        msg_summary(k)(dim_summary_qty_sell) = msg_summary(k)(dim_summary_qty_sell) + vec_trades(i)(dim_qty)
            tmp_marketvalue = tmp_marketvalue + vec_trades(i)(dim_qty) * vec_trades(i)(dim_price)
        msg_summary(k)(dim_summary_avg_price_sell) = tmp_marketvalue / msg_summary(k)(dim_summary_qty_sell)
        
        lending_report(k)(dim_lending_total_qty_sell) = lending_report(k)(dim_lending_total_qty_sell) + vec_trades(i)(dim_qty)
        lending_report(k)(dim_lending_worst_case_locate) = lending_report(k)(dim_lending_worst_case_locate) + vec_trades(i)(dim_qty)
    Else
        'a buy
        tmp_marketvalue = msg_summary(k)(dim_summary_qty_buy) * msg_summary(k)(dim_summary_avg_price_buy)
        msg_summary(k)(dim_summary_qty_buy) = msg_summary(k)(dim_summary_qty_buy) + vec_trades(i)(dim_qty)
            tmp_marketvalue = tmp_marketvalue + vec_trades(i)(dim_qty) * vec_trades(i)(dim_price)
        msg_summary(k)(dim_summary_avg_price_buy) = tmp_marketvalue / msg_summary(k)(dim_summary_qty_buy)
        
        lending_report(k)(dim_lending_total_qty_buy) = lending_report(k)(dim_lending_total_qty_sell) + vec_trades(i)(dim_qty)
    End If
    
Next i



Dim tmp_id As String, tmp_group_id As Double, tmp_json As Variant, tmp_order_price As Double

Dim vec_export_db_moulinette() As Variant

With hOrder
    For i = 0 To UBound(vec_trades, 1)
    
        For j = 0 To UBound(msg_summary, 1)
            If msg_summary(j)(dim_summary_ticker) = vec_trades(i)(dim_symbol) Then
                
                .side = get_side_redi_plus_optimize_with_lending(vec_trades(i)(dim_symbol), vec_trades(i)(dim_qty), lending_report(j)(dim_lending_actual_position), lending_report(j)(dim_lending_buy_in_redi_order_queue), lending_report(j)(dim_lending_sell_in_redi_order_queue), lending_report(j)(dim_lending_total_qty_buy), lending_report(j)(dim_lending_total_qty_sell))
                
                If msg_summary(j)(dim_summary_already_show) = False Then
                    msg_summary(j)(dim_summary_already_show) = True
                    
                    custom_msg = "Send orders " & vec_trades(i)(dim_symbol) & vbCrLf
                    
                    custom_msg = custom_msg & vbCrLf
                    
                    'ajoute quelques infos relative au lending
                    If InStr(UCase(vec_trades(i)(dim_symbol)), "EQUITY") <> 0 Then
                        custom_msg = custom_msg & "Actual position in Equity DB=" & lending_report(j)(dim_lending_actual_position) & vbCrLf
                        
                        If lending_report(j)(dim_lending_buy_in_redi_order_queue) <> 0 Then
                            custom_msg = custom_msg & "Buy orders waiting in redi=" & lending_report(j)(dim_lending_buy_in_redi_order_queue) & vbCrLf
                        End If
                        
                        If lending_report(j)(dim_lending_sell_in_redi_order_queue) <> 0 Then
                            custom_msg = custom_msg & "Sell orders waiting in redi=" & lending_report(j)(dim_lending_sell_in_redi_order_queue) & vbCrLf
                        End If
                        
                        If lending_report(j)(dim_lending_worst_case_locate) < 0 Then
                            custom_msg = custom_msg & "Worst case scenario need to be locate" & lending_report(j)(dim_lending_worst_case_locate) & vbCrLf
                        End If
                        
                        
                        custom_msg = custom_msg & vbCrLf
                    End If
                    
                    If msg_summary(j)(dim_summary_qty_buy) <> 0 Then
                        custom_msg = custom_msg & "Buy=" & msg_summary(j)(dim_summary_qty_buy) & "@avg price=" & Round(msg_summary(j)(dim_summary_avg_price_buy), 2) & vbCrLf
                    End If
                    
                    If msg_summary(j)(dim_summary_qty_sell) <> 0 Then
                        custom_msg = custom_msg & "Sell=" & Abs(msg_summary(j)(dim_summary_qty_sell)) & "@avg_price=" & Round(msg_summary(j)(dim_summary_avg_price_sell), 2) & vbCrLf
                    End If
                    
                    
                    
                    
                    
                    
                    tmp_answer = MsgBox(custom_msg, vbYesNo, "Send orders ?")
                    
                    If tmp_answer = vbYes Then
                        
                        msg_summary(j)(dim_summary_answer) = True
                        
                        'effectue le elocate si necessaire
                        If lending_report(j)(dim_lending_worst_case_locate) < 0 And InStr(UCase(vec_trades(i)(dim_symbol)), "EQUITY") <> 0 Then
                            'MsgBox ("Necessary lending for " & UCase(vec_trades(i)(dim_symbol)) & "=" & lending_report(j)(dim_lending_worst_case_locate))
                            debug_test = redi_elocate(vec_trades(i)(dim_symbol), lending_report(j)(dim_lending_worst_case_locate))
                        End If
                        
                    End If
                End If
                
                If msg_summary(j)(dim_summary_answer) = False Then
                    GoTo next_order
                End If
                
                Exit For
            End If
        Next j
        
        .symbol = get_symbol_redi_plus(vec_trades(i)(dim_symbol))
        .quantity = Abs(vec_trades(i)(dim_qty))
        .price = vec_trades(i)(dim_price)
        
        '.side = get_side_redi_plus(vec_trades(i)(dim_symbol), vec_trades(i)(dim_qty))
        
        marketplace_and_limit_type = get_marketplace_and_limit_style_redi_plus(vec_trades(i)(dim_symbol))
        
        .Exchange = marketplace_and_limit_type(0)
        .PriceType = marketplace_and_limit_type(1)
            
            If UBound(vec_trades(i), 1) >= dim_stop Then
                If IsEmpty(vec_trades(i)(dim_stop)) = False Then
                    .PriceType = "Stop"
                    .StopPrice = vec_trades(i)(dim_stop)
                Else
                    .StopPrice = 0
                End If
            Else
                .StopPrice = 0
            End If
        
        .account = get_trade_account_redi_plus(vec_trades(i)(dim_symbol))
        .UserID = redi_userid
        .TIF = "Day"
        
        
        .Memo = "none"
        .Password = redi_password
        .Warning = False
        
        
        retValueOrder = .Submit(myerr)
        
        
        
        ReDim Preserve vec_export_db_moulinette(i)
        
        Randomize
        tmp_id = Right(year(Now), 2) & Right("0" & Month(Now), 2) & Right("0" & day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & CInt(1000 * Rnd())
        
        If UBound(vec_trades(i), 1) >= dim_group_id Then
            tmp_group_id = vec_trades(i)(dim_group_id)
        Else
            Randomize
            tmp_group_id = CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0))
        End If
        
        If UBound(vec_trades(i), 1) >= dim_json_tag Then
            tmp_json = encode_json_for_DB(vec_trades(i)(dim_json_tag))
        Else
            tmp_json = Empty
        End If
        
        If UBound(vec_trades(i), 1) >= dim_stop Then
            If IsEmpty(vec_trades(i)(dim_stop)) = False Then
                tmp_order_price = vec_trades(i)(dim_stop)
            Else
                tmp_order_price = vec_trades(i)(dim_price)
            End If
        Else
            tmp_order_price = vec_trades(i)(dim_price)
        End If
        
        
        vec_export_db_moulinette(i) = Array(tmp_id, tmp_group_id, vec_trades(i)(dim_symbol), .symbol, ToJulianDay(Now), .side, vec_trades(i)(dim_qty), tmp_order_price, tmp_json)
        k = k + 1
        
        
        
        return_order(i) = Array(retValueOrder, myerr)
next_order:
    Next i
End With


If k > 0 Then
    db_sqlite_insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_order_xls, vec_export_db_moulinette, Array(f_moulinette_order_xls_id, f_moulinette_order_xls_group_id, f_moulinette_order_xls_ticker, f_moulinette_order_xls_symbol_redi, f_moulinette_order_xls_datetime, f_moulinette_order_xls_side, f_moulinette_order_xls_order_qty, f_moulinette_order_xls_order_price, f_moulinette_order_xls_json_tag))
    
    
    'offline store pour restore home
    Dim offline_status As Variant
    offline_status = moulinette_wash_offline_xls_store()
    
    Dim next_line_offline_xls As Integer
        next_line_offline_xls = 0
        
    If IsEmpty(offline_status) Then
        next_line_offline_xls = 1
    Else
        next_line_offline_xls = 2 + UBound(offline_status, 1)
    End If
    
    
    For i = 0 To UBound(vec_export_db_moulinette, 1)
        
        For j = 0 To UBound(vec_export_db_moulinette(i), 1)
            Workbooks("Kronos.xls").Worksheets(sheet_offline).Cells(next_line_offline_xls, c_offline_order_xls_id + j) = vec_export_db_moulinette(i)(j)
        Next j
        
        next_line_offline_xls = next_line_offline_xls + 1
    Next i
    
    
End If



universal_trades_r_plus = return_order

Application.Calculation = xlCalculationAutomatic

End Function


Sub algo_clean_equity_database()

Dim debug_test As Variant
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim sql_query As String


c_equity_db_id = 1
c_equity_db_name = 2
c_equity_db_nbre_stock = 24
c_equity_db_ytd_pnl = 28
c_equity_db_net_trading_cash_flow = 29
c_equity_db_option_amount = 37
c_equity_db_ticker = 47
c_equity_db_crncy = 44



sql_query = "SELECT t_trade.gs_security_id, Last(t_trade.gs_date) AS last_trade_date, Count(t_trade.gs_unique_id) AS nbre_trades, SUM(t_trade.gs_exec_qty) as net_pos "
sql_query = sql_query & " FROM t_trade INNER JOIN t_bridge ON t_trade.gs_security_id=t_bridge.gs_id"
sql_query = sql_query & " WHERE t_bridge.system_instrument_id=1"
sql_query = sql_query & " GROUP BY t_trade.gs_security_id"
sql_query = sql_query & " HAVING SUM(t_trade.gs_exec_qty)=0"
sql_query = sql_query & " AND Last(t_trade.gs_date)<" & FormatDateSQL(Date - 30)
Dim extract_cointrin As Variant
extract_cointrin = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_cointrin, 1) = 0 Then
    Exit Sub
End If


Application.Calculation = xlCalculationManual


'destruction des entree dans database folio pour eviter leur recreation lors d'automatic
If UBound(extract_cointrin, 1) > 0 Then
    For i = 1 To UBound(extract_cointrin, 1)
        For j = 13 To 32000
            If Worksheets("Database_Folio").Cells(j, 10) = "" Then
                Exit For
            Else
                If Worksheets("Database_Folio").Cells(j, 10) = extract_cointrin(i, 0) Or Worksheets("Database_Folio").Cells(j, 11) = extract_cointrin(i, 0) Then
                    Worksheets("Database_Folio").rows(j).Delete
                    j = j - 1
                End If
            End If
        Next j
    Next i
End If



Dim vec_crncy() As Variant
k = 0
For i = 14 To 31
    ReDim Preserve vec_crncy(k)
    vec_crncy(k) = Array(Worksheets("Parametres").Cells(i, 5).Value, Worksheets("Parametres").Cells(i, 1).Value)
    k = k + 1
Next i



dim_line = 0
dim_product_id = 1
dim_compagny_name = 2
dim_ticker = 3
dim_currency_code = 4
dim_ytd_pnl = 5

Dim vec_group_by_currency()
ReDim vec_group_by_currency(0)
vec_group_by_currency(0) = Array("", 0)


'passe en revue les entree d'equity database
Dim vec_ticker_to_delete() As Variant
k = 0
n = 0
For i = 27 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        If Worksheets("Equity_Database").Cells(i, c_equity_db_nbre_stock) = 0 And Worksheets("Equity_Database").Cells(i, c_equity_db_net_trading_cash_flow) = 0 And Worksheets("Equity_Database").Cells(i, c_equity_db_option_amount) = 0 Then
            
            'pas de pos ouverte et pas de pos option
            For j = 1 To UBound(extract_cointrin, 1)
                If Worksheets("Equity_Database").Cells(i, 1) = extract_cointrin(j, 0) Then
                    'titre a suppp d'equity database
                    ReDim Preserve vec_ticker_to_delete(k)
                    vec_ticker_to_delete(k) = Array(i, Worksheets("Equity_Database").Cells(i, c_equity_db_id).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_name).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_crncy).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_ytd_pnl).Value)
                    
                    
                    'mise en place d'un code 99
                    Worksheets("Equity_Database").Cells(i, 4) = 99
                    
                    
                    For m = 0 To UBound(vec_group_by_currency, 1)
                        If vec_group_by_currency(m)(0) = vec_ticker_to_delete(k)(dim_currency_code) Then
                            vec_group_by_currency(m)(1) = vec_group_by_currency(m)(1) + vec_ticker_to_delete(k)(dim_ytd_pnl)
                            Exit For
                        Else
                            If m = UBound(vec_group_by_currency, 1) Then
                                ReDim Preserve vec_group_by_currency(n)
                                vec_group_by_currency(n) = Array(vec_ticker_to_delete(k)(dim_currency_code), vec_ticker_to_delete(k)(dim_ytd_pnl))
                                n = n + 1
                            End If
                        End If
                    Next m
                    
                    k = k + 1
                    
                    Exit For
                End If
            Next j
            
        End If
    End If
Next i



'complete la zone d'exe
For i = 0 To UBound(vec_group_by_currency, 1)
    For j = 19 To 50
        If Worksheets("Exe").Cells(j, 28) = "" Then
            'la devise n'a pas ete trouvee
            ref = Worksheets("Exe").Cells(j - 1, 28)
            
            n = 0
            For m = ref + 1 To vec_group_by_currency(i)(0)
                Worksheets("Exe").Cells(j + n, 28) = m
                
                For p = 0 To UBound(vec_crncy, 1)
                    If vec_crncy(p)(0) = m Then
                        Worksheets("Exe").Cells(j + n, 27) = vec_crncy(p)(1)
                        
                        If vec_group_by_currency(i)(0) = vec_crncy(p)(0) Then
                            Worksheets("Exe").Cells(j + n, 30) = vec_group_by_currency(i)(1)
                        Else
                            Worksheets("Exe").Cells(j + n, 30) = 0
                        End If
                        
                        
                        'ligne de calcul AE
                        Worksheets("Exe").Cells(j + n, 31).FormulaLocal = "=($AC" & j + n & "+$AD" & j + n & " + $AH" & j + n & " -$AI" & j + n & " -$AJ" & j + n & ")*Parametres!$F$" & vec_crncy(p)(0) + 13
                        
                        'ligne de calcul AF
                        Worksheets("Exe").Cells(j + n, 32).FormulaLocal = "=($AC" & j + n & "+$AD" & j + n & " + $AH" & j + n & " -$AI" & j + n & " -$AJ" & j + n & ")"
                        
                        n = n + 1
                    End If
                Next p
                
            Next m
            
            
            Exit For
        Else
            If Worksheets("Exe").Cells(j, 28) = vec_group_by_currency(i)(0) Then
                Worksheets("Exe").Cells(j, 30) = Worksheets("Exe").Cells(j, 30) + vec_group_by_currency(i)(1)
                Exit For
            End If
        End If
    Next j
insert_next_pnl:
Next i


'destruction des lignes code 99
Application.ScreenUpdating = False

For i = UBound(vec_ticker_to_delete, 1) To 0 Step -1
    Worksheets("Equity_Database").rows(vec_ticker_to_delete(i)(dim_line)).Delete
    Worksheets("Equity_Database").rows(vec_ticker_to_delete(i)(dim_line) - 1).Delete
Next i

Application.ScreenUpdating = True

End Sub


Sub algo_clean_database_folio()

Application.Calculation = xlCalculationManual

Dim vb_answer As Variant
vb_answer = MsgBox("Lancer la version longue (~1-2min) retirant également les actions peu traitées sans position ouverte ?", vbYesNo, "Version")

Dim debug_test As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim sql_query As String

Dim oReg As New VBScript_RegExp_55.RegExp
    oReg.Global = True
    
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection


Dim l_database_folio_header As Integer
l_database_folio_header = 12



Dim c_db_folio_product_id As Integer, c_db_folio_underlying_id As Integer, c_db_folio_description As Integer
c_db_folio_product_id = 10
c_db_folio_underlying_id = 11
c_db_folio_description = 6


Dim l_view_folio_header As Integer, c_view_folio_product_id As Integer
l_view_folio_header = 10
c_view_folio_product_id = 1

Dim views_folio As Variant
views_folio = Array("Futures_Folio", "Equities_Folio", "Options_Folio")


Dim vec_open_positions()
ReDim Preserve vec_open_positions(0)
vec_open_positions(0) = Array("", "", "")
k = 0
'remonte la liste des pos ouverte
If vb_answer = vbYes Then
    For m = 0 To UBound(views_folio, 1)
        For i = l_view_folio_header + 2 To 32000
            If Worksheets(views_folio(m)).Cells(i, c_view_folio_product_id) = "" Then
                Exit For
            Else
                For j = 0 To UBound(vec_open_positions, 1)
    
                    If Worksheets(views_folio(m)).Cells(i, c_view_folio_product_id) = vec_open_positions(j)(1) Then
                        Exit For
                    Else
    
                        If j = UBound(vec_open_positions, 1) Then
    
                            ReDim Preserve vec_open_positions(k)
                            vec_open_positions(k) = Array(Replace(views_folio(m), "_Folio", ""), Worksheets(views_folio(m)).Cells(i, c_view_folio_product_id).Value, "")
    
                            If Left(Worksheets(views_folio(m)).Cells(i, c_view_folio_product_id + 1), 2) <> "P=" Then
                                vec_open_positions(k)(2) = Worksheets(views_folio(m)).Cells(i, c_view_folio_product_id).Value
                            Else
                                vec_open_positions(k)(2) = Worksheets(views_folio(m)).Cells(i, c_view_folio_product_id + 1).Value
                            End If
    
                            k = k + 1
    
                        End If
                    End If
                Next j
            End If
        Next i
    Next m
End If

'remonte la somme des trades ainsi que la date du dernier trade pour chaque produit
sql_query = "SELECT t_trade.gs_security_id, Last(t_trade.gs_date) AS last_trade_date, Count(t_trade.gs_unique_id) AS nbre_trades "
sql_query = sql_query & " FROM t_trade "
sql_query = sql_query & " GROUP BY t_trade.gs_security_id"
Dim extract_cointrin As Variant

If vb_answer = vbYes Then
    extract_cointrin = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
End If

Dim vec_line_to_delete() As Variant


'passe en revue les entrées
Dim date_tmp_txt As String, date_tmp_txt_day As String, date_tmp_txt_month As String, date_tmp_txt_year As String
Dim date_tmp As Date

Dim is_actively_trading As Boolean, last_trade_is_old As Boolean, has_open_position As Boolean

Dim l_db_folio_last_line As Integer
k = 0
For i = l_database_folio_header + 1 To 32000
    
    is_actively_trading = True
    last_trade_is_old = False
    has_open_position = True
    
    If Worksheets("Database_Folio").Cells(i, c_db_folio_product_id) = "" And Worksheets("Database_Folio").Cells(i + 1, c_db_folio_product_id) = "" And Worksheets("Database_Folio").Cells(i + 2, c_db_folio_product_id) = "" Then
        l_db_folio_last_line = i - 1
        Exit For
    Else
        If Worksheets("Database_Folio").Cells(i, c_db_folio_product_id) = "" Or Worksheets("Database_Folio").Cells(i, c_db_folio_underlying_id) = "" Then
            'erreur sur l'entrée
            ReDim Preserve vec_line_to_delete(k)
            vec_line_to_delete(k) = i
            k = k + 1
        Else
            
            's'agit-il d'une option
            If UCase(Left(Worksheets("Database_Folio").Cells(i, c_db_folio_description), 5)) = "CALL/" Or UCase(Left(Worksheets("Database_Folio").Cells(i, c_db_folio_description), 5)) = "CALL " Or UCase(Left(Worksheets("Database_Folio").Cells(i, c_db_folio_description), 4)) = "PUT/" Or UCase(Left(Worksheets("Database_Folio").Cells(i, c_db_folio_description), 4)) = "PUT " Then
                
                date_tmp_txt = ""
                
                
                'extraction de la date à l'aide d'une expression reguliere
                oReg.Pattern = "[0-9]{2}/[0-9]{2}/[0-9]{4}$"
                debug_test = Worksheets("Database_Folio").Cells(i, c_db_folio_description)
                Set matches = oReg.Execute(Worksheets("Database_Folio").Cells(i, c_db_folio_description).Value)
                
                For Each match In matches
                    date_tmp_txt = match.Value
                    
                    date_tmp = Mid(date_tmp_txt, 4, 2) & "." & Left(date_tmp_txt, 2) & "." & Right(date_tmp_txt, 4)
                Next
                
                oReg.Pattern = "[0-9]{1,2}(\s)[A-Za-z]{3}(\s)[0-9]{4}"
                Set matches = oReg.Execute(Worksheets("Database_Folio").Cells(i, c_db_folio_description).Value)
                
                For Each match In matches
                    date_tmp_txt = match.Value
                    
                    date_tmp_txt_day = Left(date_tmp_txt, InStr(date_tmp_txt, " ") - 1)
                    date_tmp_txt_year = Right(date_tmp_txt, 4)
                    
                    date_tmp_txt_month = ""
                    
                    If UCase(Mid(match.Value, 4, 3)) = "JAN" Then
                        date_tmp_txt_month = "01"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "FEB" Then
                        date_tmp_txt_month = "02"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "MAR" Then
                        date_tmp_txt_month = "03"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "APR" Then
                        date_tmp_txt_month = "04"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "MAY" Then
                        date_tmp_txt_month = "05"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "JUN" Then
                        date_tmp_txt_month = "06"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "JUL" Then
                        date_tmp_txt_month = "07"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "AUG" Then
                        date_tmp_txt_month = "08"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "SEP" Then
                        date_tmp_txt_month = "09"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "OCT" Then
                        date_tmp_txt_month = "10"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "NOV" Then
                        date_tmp_txt_month = "11"
                    ElseIf UCase(Mid(match.Value, 4, 3)) = "DEC" Then
                        date_tmp_txt_month = "12"
                    Else
                        date_tmp_txt = ""
                    End If
                    
                    
                    If date_tmp_txt_month <> "" Then
                        date_tmp = date_tmp_txt_day & "." & date_tmp_txt_month & "." & date_tmp_txt_year
                    End If
                    
                Next
                
                If date_tmp_txt <> "" Then
                    If Date - date_tmp >= 10 Then
                        ReDim Preserve vec_line_to_delete(k)
                        vec_line_to_delete(k) = i
                        k = k + 1
                    End If
                End If
                
            
            Else
                
                
                If vb_answer = vbYes Then
                
                    'equity / future
                    For j = 0 To UBound(vec_open_positions, 1)
                        If Worksheets("Database_Folio").Cells(i, c_db_folio_product_id) = vec_open_positions(j)(1) Or Worksheets("Database_Folio").Cells(i, c_db_folio_product_id) = vec_open_positions(j)(2) Then
                            'pos sur l'actif sous-jacent ou un derive
                            Exit For
                        Else
                            If j = UBound(vec_open_positions, 1) Then
                                has_open_position = False
                            End If
                        End If
                    Next j
                    
                    
                    If has_open_position = False Then
                        
                        'consulte la derniere date d'une transaction sur cet actif
                        For j = 1 To UBound(extract_cointrin, 1)
                            If Worksheets("Database_Folio").Cells(i, c_db_folio_product_id) = extract_cointrin(j, 0) Then
                                
                                If Date - extract_cointrin(j, 1) > 90 Then
                                    last_trade_is_old = True
                                    
                                    ReDim Preserve vec_line_to_delete(k)
                                    vec_line_to_delete(k) = i
                                    k = k + 1
                                    
                                End If
                                
                                Exit For
                            End If
                        Next j
                        
                    End If
                
                End If
            
            End If
            
        End If
    End If
Next i


If k > 0 Then
    For i = UBound(vec_line_to_delete, 1) To 0 Step -1
        'Worksheets("Database_Folio").rows(vec_line_to_delete(i)).Interior.ColorIndex = 7
        Worksheets("Database_Folio").rows(vec_line_to_delete(i)).Delete
    Next i
End If

End Sub


Sub premarket_orders_warm_up()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim l_format2_header As Integer
l_format2_header = 100


Dim vec_currency() As Variant
k = 0
For i = 14 To 31
    ReDim Preserve vec_currency(k)
    vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value)
    k = k + 1
Next i


Dim l_equity_db_header As Integer, c_equity_db_valeur_eur As Integer, c_equity_db_delta As Integer, _
    c_equity_db_theta As Integer, c_equity_db_ticker As Integer, c_equity_db_crncy As Integer
    
    
l_equity_db_header = 25
c_equity_db_valeur_eur = 5
c_equity_db_delta = 6
c_equity_db_theta = 9
c_equity_db_ticker = 47
c_equity_db_crncy = 44


Dim vec_ticker() As Variant
Dim vec_ticker_details() As Variant
    Dim dim_ticker As Integer, dim_valeur_eur As Integer, dim_delta As Integer, dim_theta As Integer, dim_crncy As Integer, _
        dim_line As Integer
    
    dim_ticker = 0
    dim_valeur_eur = 1
    dim_delta = 2
    dim_theta = 3
    dim_crncy = 4
    dim_line = 5
    
    
k = 0
Dim tmp_crncy As String
For i = l_equity_db_header + 2 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        If IsError(Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur)) = False And IsError(Worksheets("Equity_Database").Cells(i, c_equity_db_delta)) = False And IsError(Worksheets("Equity_Database").Cells(i, c_equity_db_theta)) = False Then
            If IsNumeric(Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur)) And IsNumeric(Worksheets("Equity_Database").Cells(i, c_equity_db_delta)) And IsNumeric(Worksheets("Equity_Database").Cells(i, c_equity_db_theta)) Then
                
                For j = 0 To UBound(vec_currency, 1)
                    If vec_currency(j)(1) = Worksheets("Equity_Database").Cells(i, c_equity_db_crncy).Value Then
                        Exit For
                    End If
                Next j
                
                
                'long and theta-
                debug_test = Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur)
                If Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur) > 0.1 And Worksheets("Equity_Database").Cells(i, c_equity_db_theta) < 0 Then
                    
                    ReDim Preserve vec_ticker_details(k)
                    vec_ticker_details(k) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_delta).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_theta).Value, vec_currency(j)(0), i)
                    
                    k = k + 1
                End If
                
                
                'short and theta-
                If Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur) < -0.1 And Worksheets("Equity_Database").Cells(i, c_equity_db_theta) < 0 Then
                    
                    ReDim Preserve vec_ticker_details(k)
                    vec_ticker_details(k) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_delta).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_theta).Value, vec_currency(j)(0), i)
                    
                    k = k + 1
                End If
                
            End If
        End If
    End If
Next i

'sort devise
Dim min_value As Variant
Dim min_pos As Double
Dim tmp_var As Variant

For i = 0 To UBound(vec_ticker_details, 1)
    
    min_value = vec_ticker_details(i)(dim_crncy)
    min_pos = i
    
    
    For j = i + 1 To UBound(vec_ticker_details, 1)
        If vec_ticker_details(j)(dim_crncy) < min_value Then
            min_value = vec_ticker_details(j)(dim_crncy)
            min_pos = j
        End If
    Next j
    
    If i <> min_pos Then
        
        tmp_var = vec_ticker_details(i)
        vec_ticker_details(i) = vec_ticker_details(min_pos)
        vec_ticker_details(min_pos) = tmp_var
        
    End If
Next i


    'second trie par devise par ticker
    For i = 0 To UBound(vec_ticker_details, 1)
        
        min_value = vec_ticker_details(i)(dim_ticker)
        min_pos = i
        
        For j = i + 1 To UBound(vec_ticker_details, 1)
            If vec_ticker_details(j)(dim_crncy) = vec_ticker_details(i)(dim_crncy) Then
                
                If vec_ticker_details(j)(dim_ticker) < min_value Then
                    
                    min_value = vec_ticker_details(j)(dim_ticker)
                    min_pos = j
                    
                End If
                
                
            Else
                Exit For
            End If
        Next j
        
        If i <> min_pos Then
            tmp_var = vec_ticker_details(i)
            vec_ticker_details(i) = vec_ticker_details(min_pos)
            vec_ticker_details(min_pos) = tmp_var
        End If
        
    Next i


ReDim vec_ticker(UBound(vec_ticker_details, 1))

For i = 0 To UBound(vec_ticker_details, 1)
    vec_ticker(i) = vec_ticker_details(i)(dim_ticker)
Next i



'appel bbg pour le calcul des pivots
Dim output_bbg As Variant
Dim bbg_field As Variant
bbg_field = Array("PX_YEST_CLOSE", "PX_YEST_LOW", "PX_YEST_HIGH")

Dim dim_bbg_yest_close As Integer, dim_bbg_yest_low As Integer, dim_bbg_yest_high As Integer

For i = 0 To UBound(bbg_field, 1)
    If UCase(bbg_field(i)) = UCase("PX_YEST_CLOSE") Then
        dim_bbg_yest_close = i
    ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_LOW") Then
        dim_bbg_yest_low = i
    ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_HIGH") Then
        dim_bbg_yest_high = i
    End If
Next i
output_bbg = bbg_multi_tickers_and_multi_fields(vec_ticker, bbg_field)


'preparation des trades
Dim p As Double, s1 As Double, s2 As Double, s3 As Double, r1 As Double, r2 As Double, r3 As Double

k = 0
Dim vec_trades() As Variant

Dim dim_trade_ticker As Integer, dim_trade_qty As Integer, dim_trade_price As Integer, _
    dim_trade_crncy As Integer, dim_trade_line As Integer, dim_trade_s3 As Integer, dim_trade_s2 As Integer, _
    dim_trade_s1 As Integer, dim_trade_p As Integer, dim_trade_r1 As Integer, dim_trade_r2 As Integer, dim_trade_r3 As Integer
    
    dim_trade_ticker = 0
    dim_trade_qty = 1
    dim_trade_price = 2
    dim_trade_crncy = 3
    dim_trade_line = 4
    dim_trade_s3 = 5
    dim_trade_s2 = 6
    dim_trade_s1 = 7
    dim_trade_p = 8
    dim_trade_r1 = 9
    dim_trade_r2 = 10
    dim_trade_r3 = 11
    

For i = 0 To UBound(vec_ticker_details, 1)
    If IsNumeric(output_bbg(i, dim_bbg_yest_close)) And IsNumeric(output_bbg(i, dim_bbg_yest_low)) And IsNumeric(output_bbg(i, dim_bbg_yest_high)) Then
        p = Round((output_bbg(i, dim_bbg_yest_close) + output_bbg(i, dim_bbg_yest_low) + output_bbg(i, dim_bbg_yest_high)) / 3, 3)
        
        r1 = Round(2 * p - output_bbg(i, dim_bbg_yest_low), 3)
        s1 = Round(2 * p - output_bbg(i, dim_bbg_yest_high), 3)
        
        r2 = Round((p - s1) + r1, 3)
        s2 = Round(p - (r1 - s1), 3)
        
        r3 = Round((p - s2) + r2, 3)
        s3 = Round(p - (r2 - s2), 3)
        
        'long and theta-
        If vec_ticker_details(i)(dim_valeur_eur) > 0 And vec_ticker_details(i)(dim_theta) < 0 Then
            
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(vec_ticker_details(i)(dim_ticker), -Round(Abs(vec_ticker_details(i)(dim_delta)) / 3, 0), r1, vec_ticker_details(i)(dim_crncy), vec_ticker_details(i)(dim_line), s3, s2, s1, p, r1, r2, r3)
            k = k + 1
            
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(vec_ticker_details(i)(dim_ticker), -Round(Abs(vec_ticker_details(i)(dim_delta)) / 3, 0), r2, vec_ticker_details(i)(dim_crncy), vec_ticker_details(i)(dim_line), s3, s2, s1, p, r1, r2, r3)
            k = k + 1
        End If
        
        
        'short and theta-
        If vec_ticker_details(i)(dim_valeur_eur) < 0 And vec_ticker_details(i)(dim_theta) < 0 Then
            
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(vec_ticker_details(i)(dim_ticker), Round(Abs(vec_ticker_details(i)(dim_delta)) / 3, 0), s1, vec_ticker_details(i)(dim_crncy), vec_ticker_details(i)(dim_line), s3, s2, s1, p, r1, r2, r3)
            k = k + 1
            
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(vec_ticker_details(i)(dim_ticker), Round(Abs(vec_ticker_details(i)(dim_delta)) / 3, 0), s2, vec_ticker_details(i)(dim_crncy), vec_ticker_details(i)(dim_line), s3, s2, s1, p, r1, r2, r3)
            k = k + 1
        End If
    End If
Next i


k = l_format2_header

If UBound(vec_trades, 1) > 0 Then
    'clean area
    For i = l_format2_header To 32000
        If Worksheets("FORMAT2").Cells(i, 1) = "" Then
            Exit For
        Else
            Worksheets("FORMAT2").rows(i).Clear
        End If
    Next i
    
    
    Application.ReferenceStyle = xlA1
    
    For i = 0 To UBound(vec_trades, 1)
        
        
        Worksheets("FORMAT2").Cells(k, 1) = vec_trades(i)(dim_trade_ticker)
        Worksheets("FORMAT2").Cells(k, 2) = "C6414GSJ"
        Worksheets("FORMAT2").Cells(k, 3) = "TRADING"
        'Worksheets("FORMAT2").Cells(k, 3) = UCase(vec_trades(i)(dim_trade_crncy))
        Worksheets("FORMAT2").Cells(k, 4) = UCase("LMT")
        
        If vec_trades(i)(1) < 0 Then
            Worksheets("FORMAT2").Cells(k, 5) = UCase("S")
            Worksheets("FORMAT2").Cells(k, 5).Font.ColorIndex = 3
        Else
            Worksheets("FORMAT2").Cells(k, 5) = UCase("B")
            Worksheets("FORMAT2").Cells(k, 5).Font.ColorIndex = 4
        End If
        
        Worksheets("FORMAT2").Cells(k, 10).FormulaLocal = "=Equity_Database!V" & vec_trades(i)(dim_trade_line) 'spot
            Worksheets("FORMAT2").Cells(k, 10).Interior.ColorIndex = 35
        
        Worksheets("FORMAT2").Cells(k, 19).FormulaLocal = "=Equity_Database!E" & vec_trades(i)(dim_trade_line) 'valeur eur
            Worksheets("FORMAT2").Cells(k, 19).NumberFormat = "#,##0_ ;-#,##0 "
        
        Worksheets("FORMAT2").Cells(k, 20).FormulaLocal = "=Equity_Database!F" & vec_trades(i)(dim_trade_line) 'delta
            Worksheets("FORMAT2").Cells(k, 20).NumberFormat = "#,##0_ ;-#,##0 "
        
        Worksheets("FORMAT2").Cells(k, 21).FormulaLocal = "=Equity_Database!I" & vec_trades(i)(dim_trade_line) 'theta
            Worksheets("FORMAT2").Cells(k, 21).NumberFormat = "#,##0_ ;-#,##0 "
        
        Worksheets("FORMAT2").Cells(k, 6) = Abs(vec_trades(i)(dim_trade_qty))
            Worksheets("FORMAT2").Cells(k, 6).NumberFormat = "#,##0_ ;-#,##0 "
        
        Worksheets("FORMAT2").Cells(k, 7) = Round(vec_trades(i)(dim_trade_price), 2)
            Worksheets("FORMAT2").Cells(k, 7).NumberFormat = "#,##0.00"
            
        Worksheets("FORMAT2").Cells(k, 8) = UCase("DAY")
        'Worksheets("FORMAT2").Cells(k, 9) = UCase("PICG")
        Worksheets("FORMAT2").Cells(k, 9) = Worksheets("FORMAT2").CB_exec_broker.Value
        
        Worksheets("FORMAT2").Cells(k, 11) = vec_trades(i)(dim_trade_s3)
            Worksheets("FORMAT2").Cells(k, 11).NumberFormat = "#,##0.00"
            If vec_trades(i)(dim_trade_s3) = vec_trades(i)(dim_trade_price) Then
                If UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "B" Then
                    Worksheets("FORMAT2").Cells(k, 11).Interior.ColorIndex = 4
                ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "S" Then
                    Worksheets("FORMAT2").Cells(k, 11).Interior.ColorIndex = 3
                End If
            End If
            
        Worksheets("FORMAT2").Cells(k, 12) = vec_trades(i)(dim_trade_s2)
            Worksheets("FORMAT2").Cells(k, 12).NumberFormat = "#,##0.00"
            If vec_trades(i)(dim_trade_s2) = vec_trades(i)(dim_trade_price) Then
                If UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "B" Then
                    Worksheets("FORMAT2").Cells(k, 12).Interior.ColorIndex = 4
                ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "S" Then
                    Worksheets("FORMAT2").Cells(k, 12).Interior.ColorIndex = 3
                End If
            End If
            
        Worksheets("FORMAT2").Cells(k, 13) = vec_trades(i)(dim_trade_s1)
            Worksheets("FORMAT2").Cells(k, 13).NumberFormat = "#,##0.00"
            If vec_trades(i)(dim_trade_s1) = vec_trades(i)(dim_trade_price) Then
                If UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "B" Then
                    Worksheets("FORMAT2").Cells(k, 13).Interior.ColorIndex = 4
                ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "S" Then
                    Worksheets("FORMAT2").Cells(k, 13).Interior.ColorIndex = 3
                End If
            End If
            
        Worksheets("FORMAT2").Cells(k, 14) = vec_trades(i)(dim_trade_p)
            Worksheets("FORMAT2").Cells(k, 14).NumberFormat = "#,##0.00"
            If vec_trades(i)(dim_trade_p) = vec_trades(i)(dim_trade_price) Then
                If UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "B" Then
                    Worksheets("FORMAT2").Cells(k, 14).Interior.ColorIndex = 4
                ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "S" Then
                    Worksheets("FORMAT2").Cells(k, 14).Interior.ColorIndex = 3
                End If
            End If
            
        Worksheets("FORMAT2").Cells(k, 15) = vec_trades(i)(dim_trade_r1)
            Worksheets("FORMAT2").Cells(k, 15).NumberFormat = "#,##0.00"
            If vec_trades(i)(dim_trade_r1) = vec_trades(i)(dim_trade_price) Then
                If UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "B" Then
                    Worksheets("FORMAT2").Cells(k, 15).Interior.ColorIndex = 4
                ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "S" Then
                    Worksheets("FORMAT2").Cells(k, 15).Interior.ColorIndex = 3
                End If
            End If
        
        Worksheets("FORMAT2").Cells(k, 16) = vec_trades(i)(dim_trade_r2)
            Worksheets("FORMAT2").Cells(k, 16).NumberFormat = "#,##0.00"
            If vec_trades(i)(dim_trade_r2) = vec_trades(i)(dim_trade_price) Then
                If UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "B" Then
                    Worksheets("FORMAT2").Cells(k, 16).Interior.ColorIndex = 4
                ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "S" Then
                    Worksheets("FORMAT2").Cells(k, 16).Interior.ColorIndex = 3
                End If
            End If
            
        Worksheets("FORMAT2").Cells(k, 17) = vec_trades(i)(dim_trade_r3)
            Worksheets("FORMAT2").Cells(k, 17).NumberFormat = "#,##0.00"
            If vec_trades(i)(dim_trade_r3) = vec_trades(i)(dim_trade_price) Then
                If UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "B" Then
                    Worksheets("FORMAT2").Cells(k, 17).Interior.ColorIndex = 4
                ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, 5), 1)) = "S" Then
                    Worksheets("FORMAT2").Cells(k, 17).Interior.ColorIndex = 3
                End If
            End If
        
        k = k + 1
    Next i
    
    Sheets("FORMAT2").Activate
    Worksheets("FORMAT2").Cells(l_format2_header, 1).Activate
End If

Application.Calculation = xlCalculationAutomatic

End Sub


Sub algo_helper_signal_eqs(ByVal format2_header As Integer, ByVal vec_ticker As Variant, ByVal vec_field_bbg As Variant, ByVal data_bbg As Variant, ByVal db_central As Variant)

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer, u As Integer, v As Integer

Dim oJSON As New JSONLib

Dim prefix_src_new_line As String
    prefix_src_new_line = "*** AUTO-GENERATED ***"


Dim trade_colors() As Variant, trade_color_default As Integer
    trade_colors = Array(Array(Array("STOP", "STP"), 22), Array(Array("TARGET", "TGT"), 33))
    trade_color_default = 42

'Dim pivot_s3 As Double, pivot_s2 As Double, pivot_s1 As Double, pivot_p As Double, pivot_r1 As Double, pivot_r2 As Double, pivot_r3 As Double

Dim first_try_mount_central As Boolean
    first_try_mount_central = True

Dim vec_line_to_del() As Variant
Dim count_line_to_del As Integer, last_line_to_del As Integer
    count_line_to_del = 0
    last_line_to_del = 0


If Worksheets("FORMAT2").CB_format2_eqs.Value <> "" Then
    
    Dim list_eqs As Variant
    list_eqs = get_input_format2_vec_eqs()
    
    
    'mount * parametres dispo des eqs present en input
    If IsEmpty(list_eqs) = False Then
        
        dim_eqs_name = 0
        dim_eqs_order_type = 1
        dim_eqs_formulas = 2
            dim_eqs_formulas_name = 0
            dim_eqs_formulas_formula = 1 'or empty
            dim_eqs_formulas_side = 2 'or empty
            dim_eqs_formulas_vec_fields = 3 'or empty
            dim_eqs_formulas_qty_multiplier = 4 'or 1
        
        
        For i = 0 To UBound(list_eqs, 1)
            list_eqs(i) = Array(list_eqs(i), get_eqs_order_type(list_eqs(i)), get_eqs_formulas(list_eqs(i)))
        Next i
    End If
    
    
    'passe les lignes en revue
    Dim tmp_ticker As String, tmp_color_index As Integer
    For i = format2_header + 1 To 15000
        
        If Worksheets("FORMAT2").Cells(i, 1) = "" Then
            Exit For
        Else
            
            
            'remonte les valeurs cles pour gagner du temps
            
            'package custom
            Dim package_custom_var() As Variant
            package_custom_var = Array(Array("#S3", Worksheets("FORMAT2").Cells(i, c_format2_s3).Value), Array("#S2", Worksheets("FORMAT2").Cells(i, c_format2_s2).Value), Array("#S1", Worksheets("FORMAT2").Cells(i, c_format2_s1).Value), Array("#P", Worksheets("FORMAT2").Cells(i, c_format2_p).Value), Array("#R1", Worksheets("FORMAT2").Cells(i, c_format2_r1).Value), Array("#R2", Worksheets("FORMAT2").Cells(i, c_format2_r2).Value), Array("#R3", Worksheets("FORMAT2").Cells(i, c_format2_r3).Value))
            
            
            If IsEmpty(list_eqs) = False And Worksheets("FORMAT2").Cells(i, c_format2_source) <> "" And InStr(Worksheets("FORMAT2").Cells(i, c_format2_source), prefix_src_new_line) = 0 Then 'saute les lignes rajouter precement par la procedure
                
                For j = 0 To UBound(list_eqs, 1)
                    If format_eps_report_format2(list_eqs(j)(0)) = Worksheets("FORMAT2").Cells(i, c_format2_source) Then
                        
                        
                        ' #####################################
                        ' # provient d'un eqs only buy / sell #
                        ' #####################################
                        If IsEmpty(list_eqs(j)(dim_eqs_order_type)) Then
                        Else
                            
                            If InStr(UCase(list_eqs(j)(dim_eqs_order_type)), "BUY") <> 0 Then
                                
                                If Worksheets("FORMAT2").Cells(i, c_format2_side) = "S" Or Worksheets("FORMAT2").Cells(i, c_format2_side) = "H" Then
                                    ReDim Preserve vec_line_to_del(count_line_to_del)
                                    vec_line_to_del(count_line_to_del) = i
                                    last_line_to_del = i
                                    count_line_to_del = count_line_to_del + 1
                                End If
                                
                            ElseIf InStr(UCase(list_eqs(j)(dim_eqs_order_type)), "SELL") <> 0 Then
                                If Worksheets("FORMAT2").Cells(i, c_format2_side) = "B" Or Worksheets("FORMAT2").Cells(i, c_format2_side) = "C" Then
                                    ReDim Preserve vec_line_to_del(count_line_to_del)
                                    vec_line_to_del(count_line_to_del) = i
                                    last_line_to_del = i
                                    count_line_to_del = count_line_to_del + 1
                                End If
                            Else
                            End If
                            
                        End If
                        
                        
                        Dim tmp_formula As String
                        Dim tmp_vec_fields_formula() As Variant



                        ' #####################################
                        ' # si provient d un eqs avec formula #
                        ' #####################################
                        If IsEmpty(list_eqs(j)(dim_eqs_formulas)) Then
                        Else
                            
                            Application.ScreenUpdating = False
                            
                            For u = UBound(list_eqs(j)(dim_eqs_formulas), 1) To 0 Step -1
    
                                If IsEmpty(list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_formula)) Or last_line_to_del = i Then  'evite de faire les calculs sur une ligne qui va degager
                                Else
                                    tmp_formula = list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_formula)
                                    tmp_vec_fields_formula = list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_vec_fields)
    
    
                                    'on remplace les valeurs des variables champs par des scalaires
                                    For m = 0 To UBound(tmp_vec_fields_formula, 1)
    
                                        If Left(tmp_vec_fields_formula(m), 1) = "$" Then 'bbg field
    
                                            For n = 0 To UBound(vec_ticker, 1)
                                                If Worksheets("FORMAT2").Cells(i, c_format2_ticker) = vec_ticker(n) Then
    
                                                    For p = 0 To UBound(vec_field_bbg, 1)
                                                        If UCase(Mid(tmp_vec_fields_formula(m), 2)) = UCase(vec_field_bbg(p)) Then
    
                                                            If IsNumeric(data_bbg(n)(p)) Then
                                                                tmp_formula = Replace(UCase(tmp_formula), UCase(tmp_vec_fields_formula(m)), data_bbg(n)(p))
                                                            End If
    
                                                            Exit For
                                                        End If
                                                    Next p
    
    
                                                    Exit For
                                                End If
                                            Next n
    
                                        ElseIf Left(tmp_vec_fields_formula(m), 1) = "#" Then 'custom
    
                                            For n = 0 To UBound(package_custom_var, 1)
                                                If UCase(tmp_vec_fields_formula(m)) = UCase(package_custom_var(n)(0)) Then
    
                                                    If IsNumeric(package_custom_var(n)(1)) Then
                                                        tmp_formula = Replace(UCase(tmp_formula), UCase(tmp_vec_fields_formula(m)), package_custom_var(n)(1))
                                                    End If
    
                                                    Exit For
                                                End If
                                            Next n
    
                                        ElseIf Left(tmp_vec_fields_formula(m), 1) = "£" Then 'central
    
                                            If IsEmpty(db_central) And first_try_mount_central = True Then 'aucun filtre sur central
                                                db_central = mount_sqlite_central()
                                                first_try_mount_central = False
                                            End If
    
                                            If IsEmpty(db_central) = False Then
                                                For n = 0 To UBound(db_central(0), 1)
                                                    If UCase(Mid(tmp_vec_fields_formula(m), 2)) = UCase(db_central(0)(n)) Then
    
                                                        For p = 1 To UBound(db_central, 1)
    
                                                            If UCase(db_central(p)(0)) = UCase(Worksheets("FORMAT2").Cells(i, c_format2_ticker).Value) Then
    
                                                                If IsNull(db_central(p)(n)) = False And IsNumeric(db_central(p)(n)) Then
                                                                    tmp_formula = Replace(UCase(tmp_formula), UCase(tmp_vec_fields_formula(m)), db_central(p)(n))
                                                                End If
    
                                                                Exit For
                                                            End If
    
                                                        Next p
    
                                                        Exit For
                                                    End If
    
                                                Next n
                                            End If
                                        
                                        ElseIf Left(tmp_vec_fields_formula(m), 1) = "&" Then 'equity db
                                            
                                            If IsEmpty(check_formula_syntax_vec_equity_db_header) Then
                                                p = 0
                                                Dim vec_equity_db_header() As Variant
                                                For n = 1 To 250
                                                    If Worksheets("Equity_Database").Cells(25, n) <> "" Then
                                                        ReDim Preserve vec_equity_db_header(p)
                                                        vec_equity_db_header(p) = Array(n, Worksheets("Equity_Database").Cells(25, n).Value, "&" & Replace(UCase(Worksheets("Equity_Database").Cells(25, n).Value), " ", "_"))
                                                        p = p + 1
                                                    End If
                                                Next n
                                                
                                                p = 0
                                                Dim vec_equity_db_ticker() As Variant
                                                For n = 27 To 30000 Step 2
                                                    If Worksheets("Equity_Database").Cells(n, 47) = "" Then
                                                        Exit For
                                                    Else
                                                        ReDim Preserve vec_equity_db_ticker(p)
                                                        vec_equity_db_ticker(p) = Array(n, patch_ticker_marketplace(Worksheets("Equity_Database").Cells(n, 47).Value))
                                                        p = p + 1
                                                    End If
                                                Next n
                                                
                                                
                                                check_formula_syntax_vec_equity_db_header = vec_equity_db_header
                                                check_formula_syntax_vec_equity_db_securities = vec_equity_db_ticker
                                                
                                            End If
                                            
                                            
                                            For n = 0 To UBound(check_formula_syntax_vec_equity_db_header, 1)
                                                'match column header
                                                If check_formula_syntax_vec_equity_db_header(n)(2) = tmp_vec_fields_formula(m) Or Replace(check_formula_syntax_vec_equity_db_header(n)(2), "_", "") = Replace(tmp_vec_fields_formula(m), "_", "") Then
                                                    
                                                    For p = 0 To UBound(vec_equity_db_ticker, 1)
                                                        If UCase(vec_equity_db_ticker(p)(1)) = UCase(Worksheets("FORMAT2").Cells(i, c_format2_ticker).Value) Then
                                                            
                                                            'check si numeric
                                                            If IsError(Worksheets("Equity_Database").Cells(vec_equity_db_ticker(p)(0), check_formula_syntax_vec_equity_db_header(n)(0))) = False Then
                                                                If IsNumeric(Worksheets("Equity_Database").Cells(vec_equity_db_ticker(p)(0), check_formula_syntax_vec_equity_db_header(n)(0))) Then
                                                                
                                                                    tmp_formula = Replace(UCase(tmp_formula), UCase(tmp_vec_fields_formula(m)), Worksheets("Equity_Database").Cells(vec_equity_db_ticker(p)(0), check_formula_syntax_vec_equity_db_header(n)(0)))
                                                                
                                                                End If
                                                                
                                                            End If
                                                            
                                                            
                                                            Exit For
                                                        Else
                                                            If p = UBound(vec_equity_db_ticker, 1) Then
                                                                'not in DB -> mise a zero si pnl, delta etc.
                                                                tmp_formula = Replace(UCase(tmp_formula), UCase(tmp_vec_fields_formula(m)), 0)
                                                            End If
                                                        End If
                                                    Next p
                                                    
                                                    Exit For
                                                End If
                                            Next n
                                            
                                            
                                        End If
    
                                    Next m
    
    
                                    'tente de construire un valeur pour le stop
                                    If IsError(Evaluate(tmp_formula)) = False Then
    
                                        'rajoute une ligne juste apres
                                        Worksheets("FORMAT2").Cells(i + 1, 1).EntireRow.Insert
    
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_ticker) = Worksheets("FORMAT2").Cells(i, c_format2_ticker)
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_aim_account) = Worksheets("FORMAT2").Cells(i, c_format2_aim_account)
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_strategy) = Worksheets("FORMAT2").Cells(i, c_format2_strategy)
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_strategy) = Worksheets("FORMAT2").Cells(i, c_format2_strategy)
    
                                            If Left(UCase(Worksheets("FORMAT2").Cells(i, c_format2_side)), 1) = "B" Or Left(UCase(Worksheets("FORMAT2").Cells(i, c_format2_side)), 1) = "C" Then
                                                If InStr(UCase(list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_side)), "OPPO") <> 0 Then
                                                    Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = "S"
                                                Else
                                                    Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = Left(UCase(Worksheets("FORMAT2").Cells(i, c_format2_side)), 1)
                                                End If
                                            ElseIf Left(UCase(Worksheets("FORMAT2").Cells(i, c_format2_side)), 1) = "S" Or Left(UCase(Worksheets("FORMAT2").Cells(i, c_format2_side)), 1) = "H" Then
                                                If InStr(UCase(list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_side)), "OPPO") <> 0 Then
                                                    Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = "B"
                                                Else
                                                    Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = Left(UCase(Worksheets("FORMAT2").Cells(i, c_format2_side)), 1)
                                                End If
                                            End If
    
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_qty) = Round(list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_qty_multiplier) * Worksheets("FORMAT2").Cells(i, c_format2_qty), 0)
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_price) = Round(Evaluate(tmp_formula), 2)
                                            
                                            For m = 0 To 1
                                                If Worksheets("FORMAT2").Cells(i, c_format2_pre_market_start_column + m) <> "" Then
                                                    Worksheets("FORMAT2").Cells(i + 1, c_format2_pre_market_start_column + m).FormulaLocal = Replace(Worksheets("FORMAT2").Cells(i, c_format2_pre_market_start_column + m).FormulaLocal, i, i + 1)
                                                End If
                                            Next m
                                            
                                            
                                                'ajustement du conditonal formatting
                                                With Worksheets("FORMAT2").Cells(i + 1, c_format2_price)
                                                    
                                                    .FormatConditions.Delete
                                                    
                                                    If Worksheets("FORMAT2").Cells(i, c_format2_pre_market_start_column + 1) <> "" Then
                                                        'bid/ask us
                                                        
                                                        If Left(UCase(Worksheets("FORMAT2").Cells(i + 1, c_format2_side)), 1) = "S" Or Left(UCase(Worksheets("FORMAT2").Cells(i + 1, c_format2_side)), 1) = "H" Then 'sell
                                                            .FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column) & "$" & i + 1
                                                        Else 'buy
                                                            .FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column + 1) & "$" & i + 1
                                                        End If
                                                        
                                                    Else
                                                        'theo price europe
                                                        If Left(UCase(Worksheets("FORMAT2").Cells(i + 1, c_format2_side)), 1) = "S" Or Left(UCase(Worksheets("FORMAT2").Cells(i + 1, c_format2_side)), 1) = "H" < 0 Then 'sell
                                                            .FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column) & "$" & i + 1
                                                        Else 'buy
                                                            .FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column) & "$" & i + 1
                                                        End If
                                                    End If
                                                    
                                                    .FormatConditions(1).Interior.ColorIndex = 16
                                                    
                                                End With
                                            
                                            
    
                                            If InStr(UCase(list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_name)), "STOP") <> 0 Then
                                                Worksheets("FORMAT2").Cells(i + 1, c_format2_time_limit) = "STP"
                                            Else
                                                Worksheets("FORMAT2").Cells(i + 1, c_format2_time_limit) = "DAY"
                                            End If
    
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_broker) = Worksheets("FORMAT2").Cells(i, c_format2_broker)
                                            
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_last_price).FormulaLocal = "=BDP(A" & i + 1 & ";""LAST_PRICE"")"
                                            
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_s3) = UCase(list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_name))
    
                                            Worksheets("FORMAT2").Cells(i + 1, c_format2_source) = prefix_src_new_line
    
    
                                            For m = c_format2_ticker To c_format2_broker
                                                If Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = "S" Or Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = "H" Then
                                                    Worksheets("FORMAT2").Cells(i + 1, m).Font.ColorIndex = 3
                                                ElseIf Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = "B" Or Worksheets("FORMAT2").Cells(i + 1, c_format2_side) = "C" Then
                                                    Worksheets("FORMAT2").Cells(i + 1, m).Font.ColorIndex = 10
                                                End If
                                            Next m
    
                                            
                                            For p = 0 To UBound(trade_colors, 1)
                                                For q = 0 To UBound(trade_colors(p)(0), 1)
                                                    
                                                    If InStr(UCase(list_eqs(j)(dim_eqs_formulas)(u)(dim_eqs_formulas_name)), UCase(trade_colors(p)(0)(q))) <> 0 Then
                                                        tmp_color_index = trade_colors(p)(1)
                                                        GoTo apply_color_to_trade
                                                    Else
                                                        If p = UBound(trade_colors, 1) And q = UBound(trade_colors(p)(0), 1) Then
                                                            tmp_color_index = trade_color_default
                                                        End If
                                                    End If
                                                    
                                                    
                                                Next q
                                            Next p
apply_color_to_trade:
                                            For m = c_format2_s3 To c_format2_r3
                                                Worksheets("FORMAT2").Cells(i + 1, m).Interior.ColorIndex = tmp_color_index
                                            Next m
    
                                    End If
    
                                End If
    
                            Next u
                        End If

                        Exit For
                    End If
                Next j
                
            End If


        End If
    
    Next i
    
    
    If count_line_to_del > 0 Then
        For i = UBound(vec_line_to_del, 1) To 0 Step -1
            Worksheets("FORMAT2").rows(vec_line_to_del(i)).Delete
        Next i
    End If
    
End If


Application.ScreenUpdating = True


End Sub



Sub algo_helper_filter_order_format2(ByVal format2_header As Integer, ByVal vec_ticker As Variant, ByVal vec_field_bbg As Variant, ByVal data_bbg As Variant)

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Application.Calculation = xlCalculationManual


For m = 0 To UBound(vec_field_bbg, 1)
    
    If UCase(vec_field_bbg(m)) = "DMI_ADX" Then
        dim_bbg_adx = m
    ElseIf UCase(vec_field_bbg(m)) = "DMI_DIM" Then
        dim_bbg_dmi_dim = m
    ElseIf UCase(vec_field_bbg(m)) = "DMI_DIP" Then
        dim_bbg_dmi_dip = m
    ElseIf UCase(vec_field_bbg(m)) = "INTERVAL_BOLL_PERCENT_B" Then
        dim_bbg_int_boll = m
    End If
    
Next m



Dim last_ticker As String
last_ticker = ""

For i = format2_header + 1 To 5000
    
    If Worksheets("FORMAT2").Cells(i, 1) = "" Then
        Exit For
    Else
        
        If InStr(Worksheets("FORMAT2").Cells(i, c_format2_source), "AUTO") <> 0 Then
            Debug.Print "$algo_helper_filter_order_format2: break line " & i & " because auto generated"
        Else
        
            If last_ticker <> Worksheets("FORMAT2").Cells(i, 1) Then
                
                trend_up = False
                trend_down = False
                boll_break_up = False
                boll_break_down = False
                
                'repere la ligne dans les donnees api
                
                
                last_ticker = Worksheets("FORMAT2").Cells(i, 1)
                
                For j = 0 To UBound(vec_ticker, 1)
                    If vec_ticker(j) = last_ticker Then
                        
                        'rempli les variables
                        
                        If IsNumeric(data_bbg(j)(dim_bbg_adx)) And IsNumeric(data_bbg(j)(dim_bbg_dmi_dim)) And IsNumeric(data_bbg(j)(dim_bbg_dmi_dip)) And IsNumeric(data_bbg(j)(dim_bbg_int_boll)) Then
                            
                            If data_bbg(j)(dim_bbg_adx) >= 20 Then
                                
                                If data_bbg(j)(dim_bbg_dmi_dim) > data_bbg(j)(dim_bbg_dmi_dip) And (data_bbg(j)(dim_bbg_dmi_dim) - data_bbg(j)(dim_bbg_dmi_dip) >= 15) Then
                                    trend_down = True
                                    trend_up = False
                                ElseIf data_bbg(j)(dim_bbg_dmi_dim) < data_bbg(j)(dim_bbg_dmi_dip) And (data_bbg(j)(dim_bbg_dmi_dip) - data_bbg(j)(dim_bbg_dmi_dim) >= 15) Then
                                    trend_down = False
                                    trend_up = True
                                Else
                                    trend_up = False
                                    trend_down = False
                                End If
                                
                            Else
                                trend_up = False
                                trend_down = False
                            End If
                            
                            If data_bbg(j)(dim_bbg_int_boll) > 1 Then
                                boll_break_up = True
                                boll_break_down = False
                            ElseIf data_bbg(j)(dim_bbg_int_boll) < 0 Then
                                boll_break_up = False
                                boll_break_down = True
                            Else
                                boll_break_up = False
                                boll_break_down = False
                            End If
                            
                        End If
                        
                        Exit For
                    End If
                Next j
            
            
            End If
                
                
            'analyse de la ligne
            If boll_break_up And (trend_up = False And trend_down = False) Then
                
                ' -> evite les vente
                If Worksheets("FORMAT2").Cells(i, c_format2_side) = "S" Or Worksheets("FORMAT2").Cells(i, c_format2_side) = "H" Then
                    
                    For j = 1 To 9
                        Worksheets("FORMAT2").Cells(i, j).Interior.Pattern = xlGray8
                    Next j
                    
                    For j = 1 To 50
                        If InStr(Worksheets("FORMAT2").Cells(format2_header, j), "bol") <> 0 Then
                            Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex = 4
                            Exit For
                        End If
                    Next j
                    
                End If
                
            End If
            
            
            
            If boll_break_down And (trend_up = False And trend_down = False) Then
                
                ' -> evite les achats
                If Worksheets("FORMAT2").Cells(i, c_format2_side) = "B" Or Worksheets("FORMAT2").Cells(i, c_format2_side) = "C" Then
                    
                    For j = 1 To 9
                        Worksheets("FORMAT2").Cells(i, j).Interior.Pattern = xlGray8
                    Next j
                    
                    For j = 1 To 50
                        If InStr(Worksheets("FORMAT2").Cells(format2_header, j), "bol") <> 0 Then
                            Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex = 3
                            Exit For
                        End If
                    Next j
                    
                End If
                
            End If
            
            
            If trend_up And (boll_break_up = False And boll_break_down = False) Then
                ' -> follow the trend up
                
                ' -> evite les vente
                If Worksheets("FORMAT2").Cells(i, c_format2_side) = "S" Or Worksheets("FORMAT2").Cells(i, c_format2_side) = "H" Then
                    
                    For j = 1 To 9
                        Worksheets("FORMAT2").Cells(i, j).Interior.Pattern = xlGray8
                    Next j
                    
                    For j = 1 To 50
                        If InStr(Worksheets("FORMAT2").Cells(format2_header, j), "dmi") <> 0 Then
                            Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex = 4
                            Exit For
                        End If
                    Next j
                    
                End If
                
            End If
            
            
            If trend_down And (boll_break_up = False And boll_break_down = False) Then
                ' -> follow the trend down
                
                ' -> evite les achats
                If Worksheets("FORMAT2").Cells(i, c_format2_side) = "B" Or Worksheets("FORMAT2").Cells(i, c_format2_side) = "C" Then
                    
                    For j = 1 To 9
                        Worksheets("FORMAT2").Cells(i, j).Interior.Pattern = xlGray8
                    Next j
                    
                    For j = 1 To 50
                        If InStr(Worksheets("FORMAT2").Cells(format2_header, j), "dmi") <> 0 Then
                            Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex = 3
                            Exit For
                        End If
                    Next j
                    
                End If
                
            End If
        
        End If
    End If
        
Next i

End Sub



Public Sub load_format2_new_eqs_example_fields(Optional ByVal filter As Variant)

Dim oReg As New VBScript_RegExp_55.RegExp
Dim matches As VBScript_RegExp_55.MatchCollection
Dim match As VBScript_RegExp_55.match

    oReg.IgnoreCase = True
    oReg.Global = True
    
    

Dim fields_examples() As Variant
fields_examples = Array("$EQY_BOLLINGER_UPPER", "$EQY_BOLLINGER_LOWER", "$MOV_AVG_20D", "$PX_LAST", "$PX_YEST_CLOSE", _
    "#S3", "#S2", "#S1", "#P", "#R1", "#R2", "#R3", _
    "&Valeur_Euro", "&Delta", "&Vega_1%_ALL", "&Theta_ALL", "&Daily Result", "&Result_Total", "&NetPosition", "&PositionReval", "&StartCurReval", "&Net_Cash_flow", "&Nav_Position", "&Nav_Daily", "&Perso_rel_1d", _
    "£Rank_EPS", "£Rank_Overall", "£KRRI_OVERALL", "£Rank_EPS_4w_chg_curr_yr", "£Rank_EPS_4w_chg_nxt_yr", "£Rank_LT_Growth", "£Rank_MoneyFlow", "£Rank_Ratio_EPS_curr_yr_lst", "£Rank_Ratio_EPS_nxt_yr_curr_yr", "£Rank_ROE", "£Rank_RS_ST", "£Rank_RS_LT", "£Rank_SURP", "£Rank_GEO_GROWTH_5YR_EPS", "£Rank_R2_5YR_EPS", "£Rank_MONTHLY_CHG_EPS", _
    "£chg_Rank_EPS_1w", "£chg_Rank_EPS_1m", "£chg_Rank_EPS_4w_chg_curr_yr_1w", "£chg_Rank_EPS_4w_chg_curr_yr_1m", "£chg_Rank_EPS_4w_chg_nxt_yr_1w", "£chg_Rank_EPS_4w_chg_nxt_yr_1m", "£chg_Rank_MoneyFlow_1w", "£chg_Rank_MoneyFlow_1m", "£chg_Rank_RS_LT_1w", "£chg_Rank_RS_LT_1m", "£chg_Rank_RS_ST_1w", "£chg_Rank_RS_ST_1m", "£chg_Rank_MONTHLY_CHG_EPS_1w", "£chg_Rank_MONTHLY_CHG_EPS_1m")

frm_format2_add_eqs.LB_bbg_field_examples.Clear


For i = 0 To UBound(fields_examples, 1)
    If IsMissing(filter) Then
        frm_format2_add_eqs.LB_bbg_field_examples.AddItem fields_examples(i)
    Else
        oReg.Pattern = Replace(filter, "$", "\$")
        
        Set matches = oReg.Execute(fields_examples(i))
        
        For Each match In matches
            frm_format2_add_eqs.LB_bbg_field_examples.AddItem fields_examples(i)
            Exit For
        Next
        
    End If
Next i


End Sub

Public Sub load_format2_new_eqs_form()

Call load_format2_new_eqs_example_fields

frm_format2_add_eqs.Show

End Sub


Public Sub load_format2_edit_eqs_in_form(ByVal eqs_name As String)

Call load_format2_new_eqs_example_fields

Dim oJSON As New JSONLib, colJSON As Collection, colJSON_element As Variant

Dim i As Integer, j As Integer

For i = 1 To 500
    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
        MsgBox ("Not found !")
        Exit Sub
    Else
        If UCase(Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value) = UCase(eqs_name) Then
            
            frm_format2_add_eqs.TB_eqs_screen_name = Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value
            frm_format2_add_eqs.CB_eqs_screen_folder = Worksheets("FORMAT2").Cells(i, c_format2_eqs_folder).Value
            frm_format2_add_eqs.CB_order_type = Worksheets("FORMAT2").Cells(i, c_format2_eqs_order_type).Value
            
            
            Set colJSON = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_insert_username).Value)
            
            If UCase(colJSON.Item(1)) = UCase(Environ("UserName")) Then
                frm_format2_add_eqs.CB_eqs_screen_folder = colJSON.Item(2)
                
                If UCase(colJSON.Item(2)) <> Worksheets("FORMAT2").Cells(i, c_format2_eqs_folder).Value Then
                    frm_format2_add_eqs.CB_eqs_shared = True
                Else
                    frm_format2_add_eqs.CB_eqs_shared = False
                End If
                
            End If
            
            
            If Worksheets("FORMAT2").Cells(i, c_format2_eqs_stop_formula).Value <> "" Then
            
                Set colJSON = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_stop_formula).Value)
                
                If colJSON Is Nothing Then
                Else
                    k = 1
                    For Each colJSON_element In colJSON
                        If k = 2 Then
                            frm_format2_add_eqs.TB_stp_mgmt_formula = colJSON_element
                        End If
                        k = k + 1
                    Next
                End If
            
            End If
            
            
            
            If Worksheets("FORMAT2").Cells(i, c_format2_eqs_target_formula).Value <> "" Then
            
                Set colJSON = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_target_formula).Value)
                
                If colJSON Is Nothing Then
                Else
                    k = 1
                    For Each colJSON_element In colJSON
                        If k = 2 Then
                            frm_format2_add_eqs.TB_tgt_mgmt_formula = colJSON_element
                        End If
                        k = k + 1
                    Next
                End If
            
            End If



            For j = 1 To 3
                If Worksheets("FORMAT2").Cells(i, c_format2_eqs_custom_formula_start - 1 + j) = "" Then
                    Exit For
                Else
                    
                    frm_format2_add_eqs.Height = form_height_with_custom_area
                    
                    Set colJSON = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_custom_formula_start - 1 + j))
                    
                    If colJSON Is Nothing Then
                    Else
                        k = 1
                        For Each colJSON_element In colJSON
                            If k = 1 Then
                                frm_format2_add_eqs.Controls("TB_customformula" & j & "_name").Value = colJSON_element
                            ElseIf k = 2 Then
                                frm_format2_add_eqs.Controls("TB_customformula" & j & "_formula").Value = colJSON_element
                            ElseIf k = 3 Then
                                frm_format2_add_eqs.Controls("CB_customformula" & j & "_side").Value = colJSON_element
                            ElseIf k = 4 Then
                                frm_format2_add_eqs.Controls("CB_customformula" & j & "_qty_mul").Value = colJSON_element
                            End If
                            k = k + 1
                        Next
                    End If
                    
                End If
            Next j
            
            
            Exit For
            
        End If
    End If
Next i

frm_format2_add_eqs.Caption = "Edit EQS Settings"
frm_format2_add_eqs.btn_add.Caption = "EDIT"
frm_format2_add_eqs.Show

End Sub


Public Sub load_format2_edit_eqs_form()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer


frm_format2_edit_eqs.CB_eqs.Clear

Dim vec_eqs() As Variant
k = 0
For i = 3 To 500
    
    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
        Exit For
    Else
        ReDim Preserve vec_eqs(k)
        vec_eqs(k) = Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value
        k = k + 1
    End If
Next i

If k = 0 Then
    MsgBox ("no eqs setting in worksheets format2. -> Exit")
    Exit Sub
End If


'sort
Dim min_pos As Integer
Dim min_value As String

Dim tmp_sort_var As Variant

For i = 0 To UBound(vec_eqs, 1)
    
    min_pos = i
    min_value = UCase(vec_eqs(i))
    
    For j = i + 1 To UBound(vec_eqs, 1)
        If UCase(vec_eqs(j)) < min_value Then
            min_value = UCase(vec_eqs(j))
            min_pos = j
        End If
    Next j
    
    If i <> min_pos Then
        tmp_sort_var = vec_eqs(i)
        vec_eqs(i) = vec_eqs(min_pos)
        vec_eqs(min_pos) = tmp_sort_var
    End If
    
Next i

For i = 0 To UBound(vec_eqs, 1)
    frm_format2_edit_eqs.CB_eqs.AddItem vec_eqs(i)
Next i

If Worksheets("FORMAT2").CB_format2_eqs.Value <> "" And InStr(Worksheets("FORMAT2").CB_format2_eqs.Value, ",") = 0 Then
    frm_format2_edit_eqs.CB_eqs.Value = Worksheets("FORMAT2").CB_format2_eqs.Value
End If


frm_format2_edit_eqs.Show

End Sub



Public Sub load_format2_eqs_in_form()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer


frm_format2_choose_eqs.LB_eqs.Clear

Dim vec_eqs() As Variant
k = 0
For i = 3 To 500
    
    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
        Exit For
    Else
        ReDim Preserve vec_eqs(k)
        vec_eqs(k) = Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value
        k = k + 1
    End If
Next i

If k = 0 Then
    MsgBox ("no eqs setting in worksheets format2. -> Exit")
    Exit Sub
End If


'sort
Dim min_pos As Integer
Dim min_value As String

Dim tmp_sort_var As Variant

For i = 0 To UBound(vec_eqs, 1)
    
    min_pos = i
    min_value = UCase(vec_eqs(i))
    
    For j = i + 1 To UBound(vec_eqs, 1)
        If UCase(vec_eqs(j)) < min_value Then
            min_value = UCase(vec_eqs(j))
            min_pos = j
        End If
    Next j
    
    If i <> min_pos Then
        tmp_sort_var = vec_eqs(i)
        vec_eqs(i) = vec_eqs(min_pos)
        vec_eqs(min_pos) = tmp_sort_var
    End If
    
Next i

For i = 0 To UBound(vec_eqs, 1)
    frm_format2_choose_eqs.LB_eqs.AddItem vec_eqs(i)
Next i


frm_format2_choose_eqs.Show

End Sub


Public Function get_input_format2_vec_eqs() As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


Dim vec_eqs() As Variant
k = 0
If InStr(Worksheets("FORMAT2").CB_format2_eqs.Value, ",") <> 0 Then
    is_multi_eqs_prt = True
    Dim vec_count_vec_eqs() As Variant
    
    k = 0
    For i = 1 To Len(Worksheets("FORMAT2").CB_format2_eqs.Value)
        If Mid(Worksheets("FORMAT2").CB_format2_eqs.Value, i, 1) = "," Then
            ReDim Preserve vec_count_vec_eqs(k)
            vec_count_vec_eqs(k) = i
            k = k + 1
        End If
    Next i
    
    ReDim Preserve vec_eqs(0)
    vec_eqs(0) = Left(Worksheets("FORMAT2").CB_format2_eqs.Value, InStr(Worksheets("FORMAT2").CB_format2_eqs.Value, ",") - 1)
    
    k = 1
    For i = 0 To UBound(vec_count_vec_eqs, 1)
        
        If i = UBound(vec_count_vec_eqs, 1) Then ' de , a last char
            
            For j = 0 To UBound(vec_eqs, 1)
                If Mid(Worksheets("FORMAT2").CB_format2_eqs.Value, vec_count_vec_eqs(i) + 1, Len(Worksheets("FORMAT2").CB_format2_eqs.Value) - vec_count_vec_eqs(i)) = vec_eqs(j) Then
                    Exit For
                Else
                    If j = UBound(vec_eqs, 1) Then
                        ReDim Preserve vec_eqs(i + 1)
                        vec_eqs(k) = Mid(Worksheets("FORMAT2").CB_format2_eqs.Value, vec_count_vec_eqs(i) + 1, Len(Worksheets("FORMAT2").CB_format2_eqs.Value) - vec_count_vec_eqs(i))
                        k = k + 1
                    End If
                End If
            Next j
            
        Else ' de , a ,
            
            For j = 0 To UBound(vec_eqs, 1)
                If Mid(Worksheets("FORMAT2").CB_format2_eqs.Value, vec_count_vec_eqs(i) + 1, vec_count_vec_eqs(i + 1) - (vec_count_vec_eqs(i) + 1)) = vec_eqs(j) Then
                    Exit For
                Else
                    If j = UBound(vec_eqs, 1) Then
                        ReDim Preserve vec_eqs(i + 1)
                        vec_eqs(k) = Mid(Worksheets("FORMAT2").CB_format2_eqs.Value, vec_count_vec_eqs(i) + 1, vec_count_vec_eqs(i + 1) - (vec_count_vec_eqs(i) + 1))
                        k = k + 1
                    End If
                End If
            Next j
            
        End If
        
    Next i
Else
    ReDim Preserve vec_eqs(0)
    vec_eqs(0) = Worksheets("FORMAT2").CB_format2_eqs.Value
    k = 1
End If

If k > 0 Then
    get_input_format2_vec_eqs = vec_eqs
Else
    get_input_format2_vec_eqs = Empty
End If

End Function



Private Function format_eps_report_format2(ByVal eqs_name As String) As String

format_eps_report_format2 = "EQS: " & Replace(Replace(eqs_name, "central_", ""), "_", " ")

End Function


Public Function get_vec_simple_trade_from_tweet(ByVal tweet As String, Optional ByVal vec_hashtag As Variant, Optional ByVal vec_ticker As Variant, Optional ByVal vec_mention As Variant) As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
oReg.Global = True

get_vec_simple_trade_from_tweet = Empty


Dim find_buy_sell As Boolean, find_buy As Boolean, find_sell As Boolean
Dim find_ticker As Variant 'boolean + value if found
Dim find_side As Variant 'boolean + value if found
Dim find_stop As Variant 'boolean + value if found
Dim find_tgt As Variant 'boolean + value if found
Dim find_room As Variant 'boolean + value if found
Dim find_comment As Variant 'boolean + value if found
Dim find_mention As Variant 'boolean + value if found


find_buy_sell = False
find_buy = False
find_sell = False

    Dim array_buy_hashtags() As Variant, array_sell_hashtags() As Variant
        array_buy_hashtags = Array("#BUY", "#B", "#LONG")
        array_sell_hashtags = Array("#S", "#SELL", "#SHORT", "#SS")
find_ticker = False
find_side = False
find_stop = False
    Dim array_stop_hashtags() As Variant
        array_stop_hashtags = Array("#STP", "#STOP")
find_tgt = False
    Dim array_target_hashtags() As Variant
        array_target_hashtags = Array("#TGT", "#TARGET")
find_room = False
find_comment = False
find_mention = False


If IsMissing(vec_hashtag) = False Then
    If IsArray(vec_hashtag) And IsEmpty(vec_hashtag) = False Then
        For i = 0 To UBound(vec_hashtag, 1)
            
            'side
            For j = 0 To UBound(array_buy_hashtags, 1)
                If vec_hashtag(i) = array_buy_hashtags(j) Then
                    find_buy = True
                    find_side = "B"
                End If
            Next j
            
            For j = 0 To UBound(array_sell_hashtags, 1)
                If vec_hashtag(i) = array_sell_hashtags(j) Then
                    find_sell = True
                    find_side = "S"
                End If
            Next j
            
            'stop
            For j = 0 To UBound(array_stop_hashtags, 1)
                If vec_hashtag(i) = array_stop_hashtags(j) Then
                    
                    'regexp pour checker si bien suivi d un prix
                    oReg.Pattern = array_stop_hashtags(j) & "\s+\d+(\.\d+|)"
                    Set matches = oReg.Execute(tweet)
                    
                    For Each match In matches
                        find_stop = CDbl(Replace(match.Value, array_stop_hashtags(j), ""))
                        Exit For
                    Next
                    
                    
                    Exit For
                End If
            Next j
            
            
            'target
            For j = 0 To UBound(array_target_hashtags, 1)
                If vec_hashtag(i) = array_target_hashtags(j) Then
                    
                    'regexp pour checker si bien suivi d un prix
                    oReg.Pattern = array_target_hashtags(j) & "\s+\d+(\.\d+|)"
                    Set matches = oReg.Execute(tweet)
                    
                    For Each match In matches
                        find_tgt = CDbl(Replace(match.Value, array_target_hashtags(j), ""))
                        Exit For
                    Next
                    
                    
                    Exit For
                End If
            Next j
            
            
        Next i
    End If
Else
    'passe par le tweet
    
End If



If IsMissing(vec_ticker) = False Then
    If IsArray(vec_ticker) And IsEmpty(vec_ticker) = False Then
        find_ticker = get_clean_ticker_bloomberg(vec_ticker(0))
    End If
Else
    'passe par le tweet
End If


If IsMissing(vec_mention) = False Then
    If IsArray(vec_mention) And IsEmpty(vec_mention) = False Then
        find_mention = vec_mention(0)
    End If
Else
    'passe par le tweet
End If



Dim opposite_side As String

k = 0
Dim vec_simple_trade() As Variant

If find_ticker <> False And find_side <> False And find_mention <> False Then
    
    If find_side = "B" Then
        opposite_side = "S"
    ElseIf find_side = "S" Then
        opposite_side = "B"
    End If
    
    
    ReDim Preserve vec_simple_trade(k)
    vec_simple_trade(k) = Array(find_ticker, find_side, Empty, "base", Array(find_mention, tweet))
    k = k + 1
    
    
    If find_stop <> False Then
        ReDim Preserve vec_simple_trade(k)
        vec_simple_trade(k) = Array(find_ticker, opposite_side, find_stop, "STOP", Array(find_mention, tweet))
        k = k + 1
    End If
    
    If find_tgt <> False Then
        ReDim Preserve vec_simple_trade(k)
        vec_simple_trade(k) = Array(find_ticker, opposite_side, find_tgt, "TGT", Array(find_mention, tweet))
        k = k + 1
    End If
    
    
End If


If k > 0 Then
    get_vec_simple_trade_from_tweet = vec_simple_trade
Else
    get_vec_simple_trade_from_tweet = Empty
End If


End Function


Private Function get_input_format2_standardized_portfolio_from_external_sources() As Variant

Dim oBBG As New cls_Bloomberg_Sync
Dim oJSON As New JSONLib

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim datetime_start As Date, datetime_end As Date


Dim vec_simple_trades() As Variant


Dim vec_currency() As Variant
k = 0
For i = 14 To 31
    ReDim Preserve vec_currency(k)
    vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, Worksheets("Parametres").Cells(i, 6).Value)
    k = k + 1
Next i


Dim region As Variant
    region = Array(Array("Asia/Pacific", Array("JPY", "HKD", "AUD", "SGD", "TWD", "KRW", "INR", "THB", "CNY"), Array("PX_BID_ALL_SESSION", "PX_ASK_ALL_SESSION")), Array("Europe", Array("CHF", "EUR", "GBP", "SEK", "NOK", "DKK", "PLN"), Array("THEO_PRICE")), Array("America", Array("USD", "CAD", "BRL"), Array("PX_BID_ALL_SESSION", "PX_ASK_ALL_SESSION")))

Dim l_equity_db_header As Integer, c_equity_db_valeur_eur As Integer, c_equity_db_delta As Integer, _
    c_equity_db_theta As Integer, c_equity_db_ticker As Integer, c_equity_db_crncy As Integer, c_equity_db_tag As Integer
    
    
l_equity_db_header = 25
c_equity_db_valeur_eur = 5
c_equity_db_delta = 6
c_equity_db_theta = 9
c_equity_db_net_position = 26
c_equity_db_ticker = 47
c_equity_db_crncy = 44
c_equity_db_tag = 137
c_equity_db_perso_rel_1d = 138

Dim portfolio_list_ticker() As Variant
Dim vec_equity_db() As Variant
Dim vec_prt_api_ticker() As Variant

m = 0

Dim with_trades As Boolean
    with_trades = False

Dim is_tradator As Boolean
    is_tradator = False

Dim count_src As Integer
    count_src = -1

Dim tmp_controlOLE As OLEObject
Dim vec_tradtor_enable() As Variant
    Dim count_tradtor_enable As Integer
    count_tradtor_enable = 0
For Each tmp_controlOLE In Worksheets("FORMAT2").OLEObjects
    If InStr(tmp_controlOLE.name, "tradator") <> 0 Then
        
        If tmp_controlOLE.Object.Value = True And TypeOf tmp_controlOLE.Object Is msforms.CheckBox Then
            ReDim Preserve vec_tradtor_enable(count_tradtor_enable)
            vec_tradtor_enable(count_tradtor_enable) = "@" & Replace(UCase(tmp_controlOLE.name), UCase("CB_TRADATOR_"), "")
            count_tradtor_enable = count_tradtor_enable + 1
        End If
    End If
Next


If Worksheets("FORMAT2").CB_format2_portfolio.Value <> "" Or Worksheets("FORMAT2").CB_format2_twitter_basket.Value <> "" Or Worksheets("FORMAT2").CB_format2_eqs.Value <> "" Or count_tradtor_enable <> 0 Then
    
    Dim is_twitter_prt As Boolean, is_eqs_prt As Boolean, is_multi_eqs_prt As Boolean
    q = 0 'compteur appel api afin de completer devise si non present dans equity db
    
    If Worksheets("FORMAT2").CB_format2_portfolio.Value <> "" Then
        
        is_twitter_prt = False
        
        Dim is_merge_prt As Boolean
        is_merge_prt = False
        
        
        'charge les tickers
        If get_db_access_prt_path <> "-1" Then
            
            count_src = 1
            
            portfolio_list_ticker = load_portfolio_attribution(Worksheets("FORMAT2").CB_format2_portfolio.Value)
            
            'detect les dim
            dim_prt_ticker = -1
            dim_prt_crncy_txt = -1
            dim_prt_source = -1
            
            For i = 0 To UBound(portfolio_list_ticker, 2)
                If portfolio_list_ticker(0, i) = "txt_ticker_bbg" Then
                    dim_prt_ticker = i
                ElseIf portfolio_list_ticker(0, i) = "txt_currency_short_name" Then
                    dim_prt_crncy_txt = i
                ElseIf portfolio_list_ticker(0, i) = "txt_isin" Then
                    dim_prt_source = i
                End If
            Next i
            
            If dim_prt_crncy_txt = -1 Then
                is_merge_prt = True
                dim_prt_crncy_txt = 1
                dim_prt_source = dim_prt_crncy_txt + 1
            End If
            
            
            'patch les tickers names - patch_ticker_marketplace
            For i = 1 To UBound(portfolio_list_ticker, 1)
                portfolio_list_ticker(i, dim_prt_ticker) = UCase(patch_ticker_marketplace(portfolio_list_ticker(i, dim_prt_ticker)))
                portfolio_list_ticker(i, dim_prt_source) = "Portfolio: " & Worksheets("FORMAT2").CB_format2_portfolio.Value
            Next i
            
        Else
            MsgBox ("Unable to find portfolio database. Are you at home ?")
            get_input_format2_standardized_portfolio_from_external_sources = Empty
            Exit Function
        End If
    
    ElseIf Worksheets("FORMAT2").CB_format2_twitter_basket.Value <> "" Then
        
        is_twitter_prt = True
        
        Dim list_ticker_twitter_prt As Variant
        list_ticker_twitter_prt = get_list_tickers_from_tweeted_portfolio(Worksheets("FORMAT2").CB_format2_twitter_basket.Value)

        If IsEmpty(list_ticker_twitter_prt) Then
            MsgBox ("No tickers in tweets with @portfolio and " & Worksheets("FORMAT2").CB_format2_twitter_basket.Value)
            get_input_format2_standardized_portfolio_from_external_sources = Empty
            Exit Function
        Else
            
            count_src = 1
            
            dim_prt_ticker = 0
            dim_prt_crncy_txt = 1
            dim_prt_source = 2

            ReDim Preserve portfolio_list_ticker(UBound(list_ticker_twitter_prt, 1) + 1, 5)
            For i = 0 To UBound(list_ticker_twitter_prt, 1)
                portfolio_list_ticker(i + 1, dim_prt_ticker) = list_ticker_twitter_prt(i)
                portfolio_list_ticker(i + 1, dim_prt_source) = "Twitter basket: " & Worksheets("FORMAT2").CB_format2_twitter_basket.Value
            Next i
        End If
    
    ElseIf Worksheets("FORMAT2").CB_format2_eqs.Value <> "" Then
        
        Dim vec_prt_with_all_eqs() As Variant
        
        Dim vec_eqs() As Variant
        k = 0
        If InStr(Worksheets("FORMAT2").CB_format2_eqs.Value, ",") <> 0 Then
            is_multi_eqs_prt = True
            vec_eqs = get_input_format2_vec_eqs
            k = UBound(vec_eqs, 1) + 1
        Else
            is_multi_eqs_prt = False
            vec_eqs = Array(Worksheets("FORMAT2").CB_format2_eqs.Value)
        End If
        
        is_twitter_prt = True 'permet la recuperation des devises
        is_eqs_prt = True
        
        
        'repere les parametres dans format2 concernant l eqs
        Dim eqs_screen_name As String
        Dim eqs_screen_type As String
        Dim eqs_screen_folder As String
        
        
        k = 0
        
        Dim count_found_eqs_settings As Integer
            count_found_eqs_settings = 0
        
        Dim colPropUser As Collection
        Dim prop_user_id As String
        Dim prop_user_folder As String
        For j = 0 To UBound(vec_eqs, 1)
            
            For i = 3 To 500
                If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = vec_eqs(j) Then
                    
                    eqs_screen_name = vec_eqs(j)
                    
                    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_type).Value <> "" Then
                        eqs_screen_type = Worksheets("FORMAT2").Cells(i, c_format2_eqs_type).Value
                    Else
                        eqs_screen_type = "PRIVATE"
                    End If
                    
                    eqs_screen_folder = Worksheets("FORMAT2").Cells(i, c_format2_eqs_folder).Value
                    
                    Set colPropUser = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_insert_username).Value)
                    
                    If UCase(colPropUser.Item(1)) = UCase(Environ("UserName")) Then
                        eqs_screen_folder = colPropUser.Item(2)
                    End If
                    
                    count_found_eqs_settings = count_found_eqs_settings + 1
                    
                    Exit For
                Else
                    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
                        MsgBox ("EQS " & vec_eqs(j) & " not found ! ->Check next.")
                        GoTo next_download_eqs
                    End If
                End If
            Next i
            
            
            Dim list_ticker_eqs As Variant
            datetime_start = Now()
            
            
            list_ticker_eqs = oBBG.eqs(eqs_screen_name, eqs_screen_type, eqs_screen_folder)
            
            If IsEmpty(list_ticker_eqs) Then
                If UCase(colPropUser.Item(1)) <> UCase(Environ("UserName")) Then
                    'peut etre lancer depuis la maison mais quand meme proprio du screen
                    list_ticker_eqs = oBBG.eqs(eqs_screen_name, eqs_screen_type, colPropUser.Item(2))
                End If
            End If
            
            
            datetime_end = Now()
            Debug.Print "Time to get EQS from Bloomberg: "; Application.RoundDown(1440 * (datetime_end - datetime_start), 0.1) & " minute(s) and " & CInt(60 * (1440 * (datetime_end - datetime_start) - Application.RoundDown(1440 * (datetime_end - datetime_start), 0.1))) & " seconds"
            
            
            If IsEmpty(list_ticker_eqs) Then
                Debug.Print "Problem or screening empty with EQS: " & vec_eqs(j) & ". -> Go to next entry EQS"
                count_found_eqs_settings = count_found_eqs_settings - 1
                GoTo next_download_eqs
            Else

                Dim dim_eqs_ticker As Integer
                dim_eqs_ticker = 0
                For i = 0 To UBound(list_ticker_eqs(0), 1)
                    If UCase(list_ticker_eqs(0)(i)) = UCase("Ticker") Then
                        dim_eqs_ticker = i
                        Exit For
                    End If
                Next i
                
                For i = 1 To UBound(list_ticker_eqs, 1)
                    ReDim Preserve vec_prt_with_all_eqs(k)
                    vec_prt_with_all_eqs(k) = Array(list_ticker_eqs(i)(dim_eqs_ticker), format_eps_report_format2(vec_eqs(j)))
                    k = k + 1
                Next i
                
            End If
next_download_eqs:
        Next j
        
        If count_found_eqs_settings = 0 Then
            get_input_format2_standardized_portfolio_from_external_sources = Empty
            Exit Function
        Else
            count_src = count_found_eqs_settings
            
            'tranformation du vecteur en matrix
            
            dim_prt_ticker = 0
            dim_prt_crncy_txt = 1
            dim_prt_source = dim_prt_crncy_txt + 1
            
            ReDim Preserve portfolio_list_ticker(UBound(vec_prt_with_all_eqs, 1) + 1, 5)
            For i = 0 To UBound(vec_prt_with_all_eqs, 1)
                portfolio_list_ticker(i + 1, dim_prt_ticker) = vec_prt_with_all_eqs(i)(0)
                portfolio_list_ticker(i + 1, dim_prt_source) = vec_prt_with_all_eqs(i)(1)
            Next i
            
        End If
        
    ElseIf count_tradtor_enable > 0 Then
        
        dim_extract_twitter_tweet = 0
        dim_extract_twitter_tickers_twitter = 1
        dim_extract_twitter_hashtag = 2
        dim_extract_twitter_mention = 3
        
        
        Dim find_buy_sell As Boolean
        Dim find_ticker As Boolean
        Dim find_side As Boolean
        Dim find_stop As Boolean
        Dim find_tgt As Boolean
        Dim find_room As Boolean
        
        
        Dim extract_twitter As Variant, extract_merge_twitter() As Variant
        k = 0
        m = 0
        count_src = UBound(vec_tradtor_enable, 1) + 1
        For i = 0 To UBound(vec_tradtor_enable, 1)
            extract_twitter = get_specific_tweet_content(Array(f_tweet_text, f_tweet_json_tickers, f_tweet_json_hashtags, f_tweet_json_mentions), Array(vec_tradtor_enable(i)))
            
            If IsEmpty(extract_twitter) Then
                count_src = count_src - 1
            Else
                For j = 0 To UBound(extract_twitter(dim_extract_twitter_tweet), 1)
                    ReDim Preserve extract_merge_twitter(k)
                    
                    
                    'construction du vec simple trade
                    Dim tmp_tweet As String
                    Dim tmp_vec_hashtags As Variant
                        tmp_vec_hashtags = Empty
                    Dim tmp_vec_mentions As Variant
                        tmp_vec_mentions = Empty
                    Dim tmp_vec_tickers As Variant
                        tmp_vec_tickers = Empty
                    
                    Dim tmp_mani_tweet_vec() As Variant
                    
                    Dim tmp_row_vec_simple_trade As Variant
                    If IsEmpty(extract_twitter(dim_extract_twitter_tickers_twitter)(j)) Then
                    Else
                        
                        tmp_tweet = extract_twitter(dim_extract_twitter_tweet)(j)(0)
                        
                        n = 0
                        If IsEmpty(extract_twitter(dim_extract_twitter_hashtag)(j)) = False Then
                            For p = 0 To UBound(extract_twitter(dim_extract_twitter_hashtag)(j), 1)
                                ReDim Preserve tmp_mani_tweet_vec(n)
                                tmp_mani_tweet_vec(n) = extract_twitter(dim_extract_twitter_hashtag)(j)(p)
                                n = n + 1
                            Next p
                            
                            If n > 0 Then
                                tmp_vec_hashtags = tmp_mani_tweet_vec
                            End If
                            
                        End If
                        
                        
                        n = 0
                        If IsEmpty(extract_twitter(dim_extract_twitter_mention)(j)) = False Then
                            For p = 0 To UBound(extract_twitter(dim_extract_twitter_mention)(j), 1)
                                ReDim Preserve tmp_mani_tweet_vec(n)
                                tmp_mani_tweet_vec(n) = extract_twitter(dim_extract_twitter_mention)(j)(p)
                                n = n + 1
                            Next p
                            
                            If n > 0 Then
                                tmp_vec_mentions = tmp_mani_tweet_vec
                            End If
                            
                        End If
                        
                        
                        n = 0
                        If IsEmpty(extract_twitter(dim_extract_twitter_tickers_twitter)(j)) = False Then
                            For p = 0 To UBound(extract_twitter(dim_extract_twitter_tickers_twitter)(j), 1)
                                ReDim Preserve tmp_mani_tweet_vec(n)
                                tmp_mani_tweet_vec(n) = extract_twitter(dim_extract_twitter_tickers_twitter)(j)(p)
                                n = n + 1
                            Next p
                            
                            If n > 0 Then
                                tmp_vec_tickers = tmp_mani_tweet_vec
                            End If
                            
                        End If
                        
                        
                        tmp_row_vec_simple_trade = get_vec_simple_trade_from_tweet(tmp_tweet, tmp_vec_hashtags, tmp_vec_tickers, tmp_vec_mentions)
                        
                        If IsEmpty(tmp_row_vec_simple_trade) Or IsArray(tmp_row_vec_simple_trade) = False Then
                        Else
                            For p = 0 To UBound(tmp_row_vec_simple_trade, 1)
                                ReDim Preserve vec_simple_trades(m)
                                vec_simple_trades(m) = tmp_row_vec_simple_trade(p)
                                m = m + 1
                            Next p
                        End If
                        
                        
                    End If
                    
                    extract_merge_twitter(k) = Array(extract_twitter(dim_extract_twitter_tweet)(j), extract_twitter(dim_extract_twitter_tickers_twitter)(j), extract_twitter(dim_extract_twitter_hashtag)(j), extract_twitter(dim_extract_twitter_mention)(j), vec_tradtor_enable(i))
                    
                    
                    k = k + 1
                Next j
            End If
            
        Next i
        
        
        
        'distinct ticker pour filters
        Dim tmp_ticker_bbg As String
        Dim vec_distinct_ticker_twitter() As Variant
        k = 0
        For i = 0 To UBound(extract_merge_twitter, 1)
        
            
            If IsEmpty(extract_merge_twitter(i)(1)) Then
            Else
                
                'prend que 1 seul ticker par tweet
                tmp_ticker_bbg = UCase(patch_ticker_marketplace(get_clean_ticker_bloomberg(extract_merge_twitter(i)(1)(0))))
                
                If k = 0 Then
                    ReDim Preserve vec_distinct_ticker_twitter(k)
                    vec_distinct_ticker_twitter(k) = tmp_ticker_bbg
                    k = k + 1
                Else
                    For j = 0 To UBound(vec_distinct_ticker_twitter, 1)
                        If vec_distinct_ticker_twitter(j) = tmp_ticker_bbg Then
                            Exit For
                        Else
                            If j = UBound(vec_distinct_ticker_twitter, 1) Then
                                ReDim Preserve vec_distinct_ticker_twitter(k)
                                vec_distinct_ticker_twitter(k) = tmp_ticker_bbg
                                k = k + 1
                            End If
                        End If
                    Next j
                End If
            End If
        Next i
        
        
        dim_prt_ticker = 0
        dim_prt_crncy_txt = 1
        dim_prt_source = dim_prt_crncy_txt + 1
        
        If k > 0 Then
            
            is_tradator = True
            with_trades = True
            
            ReDim Preserve portfolio_list_ticker(UBound(vec_distinct_ticker_twitter, 1) + 1, 5)
            For i = 0 To UBound(vec_distinct_ticker_twitter, 1)
                portfolio_list_ticker(i + 1, dim_prt_ticker) = vec_distinct_ticker_twitter(i)
                portfolio_list_ticker(i + 1, dim_prt_source) = "@tradator"
            Next i
        Else
            get_input_format2_standardized_portfolio_from_external_sources = Empty
            Exit Function
        End If
        
    End If
    
    'mount equity database
    dim_vec_equity_db_line = 0
    dim_vec_equity_db_ticker = 1
    dim_vec_equity_db_crncy = 2
    dim_vec_equity_db_qty_stock = 3
    
    k = 0
    For i = l_equity_db_header + 2 To 32000 Step 2
        If Worksheets("Equity_Database").Cells(i, 1) = "" Then
            Exit For
        Else
            'doit encore eviter les codes 12
            If Worksheets("Equity_Database").Cells(i, 97) = "" And Worksheets("Equity_Database").Cells(i, 98) = "" Then
                ReDim Preserve vec_equity_db(k)
                vec_equity_db(k) = Array(i, patch_ticker_marketplace(UCase(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value)), Worksheets("Equity_Database").Cells(i, c_equity_db_crncy).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_net_position).Value)
                k = k + 1
            Else
                If Worksheets("Equity_Database").Cells(i, 1) = Worksheets("Equity_Database").Cells(i, 97) Then
                    ReDim Preserve vec_equity_db(k)
                    vec_equity_db(k) = Array(i, patch_ticker_marketplace(UCase(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value)), Worksheets("Equity_Database").Cells(i, c_equity_db_crncy).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_net_position).Value)
                    k = k + 1
                End If
            End If
        End If
    Next i
    
    
    'regarde si les tickers sont trouver dans la base - si merge profite de completer la devise pour alleger au max le vecteur api
    n = 0
    For i = 1 To UBound(portfolio_list_ticker, 1)
        
        For j = 0 To UBound(vec_equity_db, 1)
            If portfolio_list_ticker(i, dim_prt_ticker) = vec_equity_db(j)(dim_vec_equity_db_ticker) Then
                
                'remplace l'emplacement prevue pour la devise par un array contenant devise txt, devise code, line equity db
                If is_merge_prt = True Or is_twitter_prt = True Or is_tradator = True Then
                    For n = 0 To UBound(vec_currency, 1)
                        If vec_currency(n)(1) = vec_equity_db(j)(dim_vec_equity_db_crncy) Then
                            portfolio_list_ticker(i, dim_prt_crncy_txt) = Array(vec_equity_db(j)(dim_vec_equity_db_line), vec_currency(n)(0), vec_equity_db(j)(dim_vec_equity_db_crncy))
                            Exit For
                        End If
                    Next n
                Else
                    'l'emplacemenet contient deja la devise txt
                    portfolio_list_ticker(i, dim_prt_crncy_txt) = Array(vec_equity_db(j)(dim_vec_equity_db_line), portfolio_list_ticker(i, dim_prt_crncy_txt), vec_equity_db(j)(dim_vec_equity_db_crncy))
                End If
                
                Exit For
            Else
                If j = UBound(vec_equity_db, 1) Then
                    'impossible de mettre la main sur la ligne dans equity db -> envoi dans bbg api
                    If is_merge_prt = True Or is_twitter_prt = True Or is_tradator = True Then
                        
                        'un appel api va etre necessaire
                        portfolio_list_ticker(i, dim_prt_crncy_txt) = Array(-1, -1, -1)
                        
                        ReDim Preserve vec_prt_api_ticker(q)
                        vec_prt_api_ticker(q) = portfolio_list_ticker(i, dim_prt_ticker)
                        q = q + 1
                        
                    Else
                        For n = 0 To UBound(vec_currency, 1)
                            If vec_currency(n)(0) = portfolio_list_ticker(i, dim_prt_crncy_txt) Then
                                portfolio_list_ticker(i, dim_prt_crncy_txt) = Array(-1, vec_currency(n)(0), vec_currency(n)(1))
                                Exit For
                            End If
                        Next n
                    End If
                End If
            End If
        Next j
    Next i
    
    
    If q > 0 Then
       output_bdp_prt = oBBG.bdp(vec_prt_api_ticker, Array("CRNCY"), output_format.of_vec_without_header)
        
        For i = 1 To UBound(portfolio_list_ticker, 1)
            For j = 0 To UBound(vec_prt_api_ticker, 1)
                If portfolio_list_ticker(i, dim_prt_ticker) = vec_prt_api_ticker(j) Then
                    
                    For n = 0 To UBound(vec_currency, 1)
                        'If UCase(output_bdp_prt(j, 0)) = UCase(vec_currency(n)(0)) Then
                        If UCase(output_bdp_prt(j)(0)) = UCase(vec_currency(n)(0)) Then
                            portfolio_list_ticker(i, dim_prt_crncy_txt) = Array(-1, vec_currency(n)(0), vec_currency(n)(1))
                            Exit For
                        Else
                            If n = UBound(vec_currency, 1) Then
                                'titre invalid
                                portfolio_list_ticker(i, dim_prt_ticker) = False
                            End If
                        End If
                    Next n
                    
                    Exit For
                End If
            Next j
        Next i
        
    End If
    
    
    Dim final_standardized_output() As Variant
    ReDim final_standardized_output(UBound(portfolio_list_ticker, 1), 2)
    
        'header
        final_standardized_output(0, 0) = "ticker"
        final_standardized_output(0, 1) = "vec_crncy"
        final_standardized_output(0, 2) = "src:" & count_src
    
    
    For i = 1 To UBound(portfolio_list_ticker, 1)
        final_standardized_output(i, 0) = portfolio_list_ticker(i, dim_prt_ticker)
        final_standardized_output(i, 1) = portfolio_list_ticker(i, dim_prt_crncy_txt)
        final_standardized_output(i, 2) = portfolio_list_ticker(i, dim_prt_source)
    Next i
    
    If with_trades = False Then
        get_input_format2_standardized_portfolio_from_external_sources = final_standardized_output
    Else
        get_input_format2_standardized_portfolio_from_external_sources = Array(final_standardized_output, vec_simple_trades)
    End If
    
Else
    get_input_format2_standardized_portfolio_from_external_sources = Empty
End If

End Function



Public Function check_stp_tgt_formula_syntax(ByVal formula As String) As Boolean

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

check_stp_tgt_formula_syntax = True

'meme nbre de ( )
m = 0
n = 0
If InStr(formula, "(") <> 0 Or InStr(formula, ")") <> 0 Then
    
    For i = 1 To Len(formula)
        
        If Mid(formula, i, 1) = "(" Then
            m = m + 1
        ElseIf Mid(formula, i, 1) = ")" Then
            n = n + 1
        End If
    Next i
    
End If

If m <> n Then
    check_stp_tgt_formula_syntax = False
End If

End Function



Public Function get_eqs_formulas(ByVal eqs As Variant) As Variant

get_eqs_formulas = Empty

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim oJSON As New JSONLib, colJSON As Collection, colElement As Variant

k = 0

Dim matrix_formulas() As Variant
Dim tmp_formula() As Variant


For i = 1 To 500
    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
        Exit For
    ElseIf eqs = Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value Then
        
        If Worksheets("FORMAT2").Cells(i, c_format2_eqs_stop_formula).Value <> "" Then
            
            Set colJSON = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_stop_formula).Value)
            
            If colJSON Is Nothing Then
            Else
                ReDim Preserve matrix_formulas(k)
                matrix_formulas(k) = Array(colJSON.Item(1), colJSON.Item(2), colJSON.Item(3), get_eqs_fields_stop_target_formula(, , colJSON.Item(2)), colJSON.Item(4))
                k = k + 1
            End If
        End If
        
        If Worksheets("FORMAT2").Cells(i, c_format2_eqs_target_formula).Value <> "" Then
            Set colJSON = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_target_formula).Value)
            
            If colJSON Is Nothing Then
            Else
                ReDim Preserve matrix_formulas(k)
                matrix_formulas(k) = Array(colJSON.Item(1), colJSON.Item(2), colJSON.Item(3), get_eqs_fields_stop_target_formula(, , colJSON.Item(2)), colJSON.Item(4))
                k = k + 1
            End If
        End If
        
        For j = 1 To 3
            If Worksheets("FORMAT2").Cells(i, c_format2_eqs_custom_formula_start - 1 + j).Value <> "" Then
                Set colJSON = oJSON.parse(Worksheets("FORMAT2").Cells(i, c_format2_eqs_custom_formula_start - 1 + j).Value)
                
                If colJSON Is Nothing Then
                Else
                    ReDim Preserve matrix_formulas(k)
                    matrix_formulas(k) = Array(colJSON.Item(1), colJSON.Item(2), colJSON.Item(3), get_eqs_fields_stop_target_formula(, , colJSON.Item(2)), colJSON.Item(4))
                    k = k + 1
                End If
            Else
                Exit For
            End If
            
        Next j
        
    End If
Next i


If k > 0 Then
    get_eqs_formulas = matrix_formulas
End If


End Function


Public Function get_eqs_stop_target_formula(ByVal eqs As Variant, ByVal column_formula As Variant) As Variant

For i = 1 To 500
    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
        get_eqs_stop_target_formula = Empty
        Exit Function
    ElseIf eqs = Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value Then
        If Worksheets("FORMAT2").Cells(i, column_formula).Value <> "" Then
            get_eqs_stop_target_formula = Worksheets("FORMAT2").Cells(i, column_formula).Value
            Exit Function
        Else
            get_eqs_stop_target_formula = Empty
            Exit Function
        End If
    End If
Next i

End Function


Public Function get_eqs_order_type(ByVal eqs As String) As Variant

Dim i As Integer

get_eqs_order_type = Empty

For i = 1 To 500
    If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
        Exit Function
    ElseIf eqs = Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value Then
        If Worksheets("FORMAT2").Cells(i, c_format2_eqs_order_type).Value <> "" Then
            get_eqs_order_type = Worksheets("FORMAT2").Cells(i, c_format2_eqs_order_type).Value
            Exit For
        Else
            Exit Function
        End If
    End If
Next i


End Function


Public Function get_eqs_fields_stop_target_formula(Optional ByVal eqs As Variant, Optional ByVal column_formula As Variant, Optional ByVal formula As Variant) As Variant

Dim xl_fn() As Variant
    xl_fn = Array("IF", "MIN", "MAX", "SUM", "ABS", "AND", "OR", "AVERAGE", "INT", "ROUND", "RAND", "ISUMBER", "LEN", "LEFT", "RIGHT", "LN", "LOG", "MEDIAN", "MOD", "MOD", "DAY", "MONTH", "YEAR", "HOUR", "MINUTE", "SECOND", "NOW")
    

Dim i As Integer, j As Integer, k As Integer

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True


If IsMissing(formula) Then
    
    For i = 1 To 500
        If Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value = "" Then
            get_eqs_fields_stop_target_formula = Empty
            Exit Function
        ElseIf eqs = Worksheets("FORMAT2").Cells(i, c_format2_eqs_name).Value Then
            If Worksheets("FORMAT2").Cells(i, column_formula).Value <> "" Then
                formula = Worksheets("FORMAT2").Cells(i, column_formula).Value
                Exit For
            Else
                get_eqs_fields_stop_target_formula = Empty
                Exit Function
            End If
        End If
    Next i
    
End If

oReg.Pattern = "(|\$|£|#|&)[\d\w]{2,}"

Set matches = oReg.Execute(formula)

Dim vec_fields() As Variant
k = 0
For Each match In matches
    If IsNumeric(match.Value) Then
        'scalaire
    Else
        For j = 0 To UBound(xl_fn, 1)
            If UCase(xl_fn(j)) = UCase(match.Value) Then
                'xl fn
                Exit For
            Else
                If j = UBound(xl_fn, 1) Then
        
                    If k = 0 Then
                        ReDim Preserve vec_fields(k)
                        vec_fields(k) = UCase(match.Value)
                        k = k + 1
                    Else
                        For i = 0 To UBound(vec_fields, 1)
                            If UCase(vec_fields(i)) = UCase(match.Value) Then
                                Exit For
                            Else
                                If i = UBound(vec_fields, 1) Then
                                    ReDim Preserve vec_fields(k)
                                    vec_fields(k) = UCase(match.Value)
                                    k = k + 1
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
        Next j
    End If
Next

If k > 0 Then
    get_eqs_fields_stop_target_formula = vec_fields
Else
    get_eqs_fields_stop_target_formula = Empty
End If


End Function


Public Function get_last_word_in_formula(ByVal formula As String) As String

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer


Dim math_operator() As Variant
    math_operator = Array("+", "-", "*", "/", "^", "(", ")", ";", ",", " ", "<", ">", "=")

get_last_word_in_formula = ""

If formula = "" Then
    Exit Function
Else
    
    Dim pos_end_word As Integer, pos_start_word As Integer
    
    Dim find_start_word_pos As Integer
        find_start_word_pos = 0
    
    pos_start_word = 1
       For i = 1 To Len(formula)
        
        If i = Len(formula) Then
            pos_end_word = Len(formula)
            Exit For
        End If
        
        If Mid(StrReverse(formula), i, 1) = " " Then
            
            If find_start_word_pos = 0 Then
                pos_start_word = i
            Else
                pos_end_word = i
                Exit For
            End If
            
        Else
            
            If find_start_word_pos = 0 Then
                find_start_word_pos = i
            End If
            
            For j = 0 To UBound(math_operator, 1)
                
                debug_test = Mid(StrReverse(formula), i, 1)
                If Mid(StrReverse(formula), i, 1) = math_operator(j) Then
                    pos_end_word = i - 1
                    GoTo bypass_next_chr
                End If
            Next j
            
        End If
    Next i
bypass_next_chr:
    
    If find_start_word_pos = 0 Or find_start_word_pos = 1 Then
    
        If StrReverse(Left(StrReverse(formula), pos_end_word)) <> " " Then
            get_last_word_in_formula = Trim(StrReverse(Left(StrReverse(formula), pos_end_word)))
        End If
    Else
        If pos_end_word - find_start_word_pos < 0 Then
            Exit Function
        Else
            get_last_word_in_formula = StrReverse(Mid(StrReverse(formula), find_start_word_pos, pos_end_word - find_start_word_pos))
        End If
    End If
    
End If


End Function



Public Function patch_stp_tgt_formula_syntax(ByVal formula As String) As String

Dim xl_fn() As Variant
    xl_fn = Array("IF", "MIN", "MAX", "SUM", "ABS", "AND", "OR", "AVERAGE", "INT", "ROUND", "RAND", "ISUMBER", "LEN", "LEFT", "RIGHT", "LN", "LOG", "MEDIAN", "MOD", "MOD", "DAY", "MONTH", "YEAR", "HOUR", "MINUTE", "SECOND", "NOW")
    

patch_stp_tgt_formula_syntax = formula

formula = Replace(formula, " ", "")



Dim i As Integer, j As Integer, k As Integer

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True

oReg.Pattern = "(|\$|£|#|&)[\d\w]{2,}"

Set matches = oReg.Execute(formula)

For Each match In matches
    
    If InStr(match.Value, "$") <> 0 Then
        
    ElseIf InStr(match.Value, "£") <> 0 Then
        
    ElseIf InStr(match.Value, "#") <> 0 Then
        
    ElseIf InStr(match.Value, "&") <> 0 Then
        
    ElseIf IsNumeric(match.Value) Then
        
    Else
        
        'doit differencier les nom de fonction / les champs internes / bbg
        
        For i = 0 To UBound(xl_fn, 1)
            If UCase(xl_fn(i)) = UCase(match.Value) Then
                Exit For
            Else
                If i = UBound(xl_fn, 1) Then
                    
                    If IsEmpty(check_formula_syntax_vec_equity_db_header) Then
                        Dim vec_equity_db_column() As Variant
                        
                        k = 0
                        For j = 1 To 250
                            
                            If Worksheets("Equity_Database").Cells(25, j) <> "" Then
                                ReDim Preserve vec_equity_db_column(k)
                                vec_equity_db_column(k) = Worksheets("Equity_Database").Cells(25, j).Value
                                k = k + 1
                            End If
                        Next
                        
                        check_formula_syntax_vec_equity_db_header = vec_equity_db_column
                    End If
                    
                    'check fields internal
                    For j = 0 To UBound(check_formula_syntax_vec_equity_db_header, 1)
                        If UCase(Replace(check_formula_syntax_vec_equity_db_header(j), " ", "")) = Replace(UCase(match.Value), "_", "") Then
                            formula = Replace(formula, match.Value, "&" & UCase(Replace(check_formula_syntax_vec_equity_db_header(j), " ", "_")))
                            Exit For
                        Else
                            If j = UBound(check_formula_syntax_vec_equity_db_header, 1) Then
                                'par de l hypothese que champ bloomberg
                                formula = Replace(formula, match.Value, "$" & match.Value)
                            End If
                        End If
                    Next j
                    
                End If
            End If
        Next i
        
        
    End If
    
Next

'risque de double $
formula = Replace(formula, "$$", "$")

'evite formulation excel
formula = Replace(formula, ";", ",")
For i = 1 To Len(formula)
    If Mid(formula, i, 1) = " " Then
        'on continue
    ElseIf Mid(formula, i, 1) = "=" Then
        formula = Mid(formula, InStr(formula, "=") + 1)
        Exit For
    Else
        Exit For
    End If
Next i


patch_stp_tgt_formula_syntax = formula

End Function


Sub prepare_trades_from_call_broker()

Dim oBBG As New cls_Bloomberg_Sync

Dim rec_mention() As Variant
    rec_mention = Array("@downgrade", "@upgrade", "@ug", "@dg")

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

Dim oJSON As New JSONLib

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer


Dim tmp_vec() As Variant

Application.Calculation = xlCalculationManual

Dim base_amount As Double
    base_amount = CDbl(Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value)

Dim vec_currency() As Variant
k = 0
For i = 14 To 31
    ReDim Preserve vec_currency(k)
    vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, Worksheets("Parametres").Cells(i, 6).Value)
    k = k + 1
Next i

Dim region As Variant
region = get_region_based_on_crncy()

Dim vec_ticker() As Variant

'remonte les call du jour de twitter
Dim sql_query As String
sql_query = "SELECT " & f_tweet_id & ", " & f_tweet_datetime & ", " & f_tweet_text & ", " & f_tweet_json_tickers & ", " & f_tweet_json_hashtags & ", " & f_tweet_json_mentions
    sql_query = sql_query & " FROM " & t_tweet
    sql_query = sql_query & " WHERE " & f_tweet_datetime & ">=" & ToJulianDay(Date)
    sql_query = sql_query & " AND " & f_tweet_json_tickers & " IS NOT NULL"
    
    sql_query = sql_query & " AND ("
        For i = 0 To UBound(rec_mention, 1)
            If i = 0 Then
            Else
                sql_query = sql_query & " OR "
            End If

            sql_query = sql_query & f_tweet_json_mentions & " LIKE ""%" & UCase(rec_mention(i)) & "%"""

        Next i
    sql_query = sql_query & ")"
    
    sql_query = sql_query & " ORDER BY " & f_tweet_datetime & " DESC"

Dim extract_daily_tweet_rec As Variant
extract_daily_tweet_rec = sqlite3_query(twitter_get_db_path, sql_query)



If UBound(extract_daily_tweet_rec, 1) > 0 Then
    
    'detect des dim
    For i = 0 To UBound(extract_daily_tweet_rec(0), 1)
        If extract_daily_tweet_rec(0)(i) = f_tweet_id Then
            dim_extract_tweet_id = i
        ElseIf extract_daily_tweet_rec(0)(i) = f_tweet_datetime Then
            dim_extract_datetime = i
        ElseIf extract_daily_tweet_rec(0)(i) = f_tweet_text Then
            dim_extract_text = i
        ElseIf extract_daily_tweet_rec(0)(i) = f_tweet_json_tickers Then
            dim_extract_tickers = i
        ElseIf extract_daily_tweet_rec(0)(i) = f_tweet_json_hashtags Then
            dim_extract_hashtags = i
        ElseIf extract_daily_tweet_rec(0)(i) = f_tweet_json_mentions Then
            dim_extract_mentions = i
        End If
    Next i
    
    
    'second appel pour construire le vec_ticker()
    k = 0
    Set colTickers = oJSON.parse(decode_json_from_DB(extract_daily_tweet_rec(1)(dim_extract_tickers)))
    For Each tmp_ticker In colTickers
        ReDim Preserve vec_ticker(k)
        vec_ticker(k) = get_clean_ticker_bloomberg(tmp_ticker)
        k = k + 1
        Exit For 'only one
    Next
    
    For i = 2 To UBound(extract_daily_tweet_rec, 1)
        Set colTickers = oJSON.parse(decode_json_from_DB(extract_daily_tweet_rec(i)(dim_extract_tickers)))
        For Each tmp_ticker In colTickers
            
            For j = 0 To UBound(vec_ticker, 1)
                If vec_ticker(j) = get_clean_ticker_bloomberg(tmp_ticker) Then
                    Exit For
                Else
                    If j = UBound(vec_ticker, 1) Then
                        ReDim Preserve vec_ticker(k)
                        vec_ticker(k) = get_clean_ticker_bloomberg(tmp_ticker)
                        k = k + 1
                    End If
                End If
            Next j
            
            Exit For 'only one
        Next
    Next i
    
    
    'appel bbg pour crncy, les pivots
    Dim bbg_field() As Variant
    bbg_field = Array("CRNCY", "PX_YEST_CLOSE", "PX_YEST_LOW", "PX_YEST_HIGH", "LAST_PRICE")
        
        For i = 0 To UBound(bbg_field, 1)
            If bbg_field(i) = "CRNCY" Then
                dim_bbg_CRNCY = i
            ElseIf bbg_field(i) = "PX_YEST_CLOSE" Then
                dim_bbg_px_yest_close = i
            ElseIf bbg_field(i) = "PX_YEST_LOW" Then
                dim_bbg_px_yest_low = i
            ElseIf bbg_field(i) = "PX_YEST_HIGH" Then
                dim_bbg_px_yest_high = i
            ElseIf bbg_field(i) = "LAST_PRICE" Then
                dim_bbg_last_price = i
            End If
        Next i
        
    Dim output_bbg As Variant
    output_bbg = oBBG.bdp(vec_ticker, bbg_field, output_format.of_vec_without_header)
    
    
    'remonte equity database
    k = 0
    Dim vec_equity_db() As Variant
    For i = 27 To 5000 Step 2
        If Worksheets("Equity_Database").Cells(i, 1) = "" Then
            Exit For
        Else
            If IsError(Worksheets("Equity_Database").Cells(i, 24)) Then
                tmp_pos = 0
            Else
                tmp_pos = Worksheets("Equity_Database").Cells(i, 24)
            End If
            
            ReDim Preserve vec_equity_db(k)
            vec_equity_db(k) = Array(i, patch_ticker_marketplace(Worksheets("Equity_Database").Cells(i, 47)), tmp_pos)
            k = k + 1
        End If
    Next i
    
    
    Dim vec_trade() As Variant
    count_trade = 0
    For i = 1 To UBound(extract_daily_tweet_rec, 1)
        
        'conversion des json en vecteur
        Set colTickers = oJSON.parse(decode_json_from_DB(extract_daily_tweet_rec(i)(dim_extract_tickers)))
        If colTickers Is Nothing Then
        Else
            k = 0
            For Each tmp_ticker In colTickers
                ReDim Preserve tmp_vec(k)
                tmp_ticker = patch_ticker_marketplace(get_clean_ticker_bloomberg(tmp_ticker))
                k = k + 1
                Exit For 'only one
            Next
            
        End If
        
        
        If IsNull(extract_daily_tweet_rec(i)(dim_extract_hashtags)) = False Then
            Set colHashtags = oJSON.parse(decode_json_from_DB(extract_daily_tweet_rec(i)(dim_extract_hashtags)))
            If colHashtags Is Nothing Then
                vec_hashtag = Empty
            Else
                k = 0
                For Each tmp_hashtag In colHashtags
                    ReDim Preserve tmp_vec(k)
                    tmp_vec(k) = tmp_hashtag
                    k = k + 1
                Next
                
                vec_hashtag = tmp_vec
            
            End If
        End If
        
        
        Set colMentions = oJSON.parse(decode_json_from_DB(extract_daily_tweet_rec(i)(dim_extract_mentions)))
        If colMentions Is Nothing Then
            vec_mention = Empty
        Else
            k = 0
            For Each tmp_mention In colMentions
                ReDim Preserve tmp_vec(k)
                tmp_vec(k) = tmp_mention
                k = k + 1
            Next
            
            vec_mention = tmp_vec
        
        End If
        
        
        tmp_side = "long"
        tmp_broker = Empty
        For m = 0 To UBound(vec_mention, 1)
            For n = 0 To UBound(rec_mention, 1)
                If UCase(vec_mention(m)) = UCase(rec_mention(n)) Then
                    
                    If InStr(UCase(rec_mention(n)), "DG") <> 0 Or InStr(UCase(rec_mention(n)), "DOWNGRADE") <> 0 Then
                        tmp_side = "short"
                    Else
                        tmp_side = "long"
                    End If
                    
                    'Exit For
                Else
                    If n = UBound(rec_mention, 1) Then
                        tmp_broker = twitter_get_broker_id_from_mention(vec_mention(m), output_conv_twitter_broker.aim_exec_broker)
                        
                        If tmp_broker = "" Then
                            tmp_broker = Worksheets("FORMAT2").CB_exec_broker.Value
                        End If
                        
                    End If
                End If
            Next n
        Next m
        
        oReg.Pattern = "\sto\s[\w|\s|/]+(\s#TGT|\sby)"
        Set matches = oReg.Execute(extract_daily_tweet_rec(i)(dim_extract_text))
        
        
        tmp_rec = ""
        For Each match In matches
            tmp_rec = match.Value
            
            tmp_rec = Replace(tmp_rec, " to ", "")
            tmp_rec = Replace(tmp_rec, " #TGT", "")
            tmp_rec = Replace(tmp_rec, " by", "")
            
            Exit For
        Next
        
        
        oReg.Pattern = "#TGT\s[\d]+(\.[\d]+|)"
        Set matches = oReg.Execute(extract_daily_tweet_rec(i)(dim_extract_text))
        
        tmp_pt = ""
        For Each match In matches
            tmp_pt = match.Value
            tmp_pt = Replace(tmp_pt, "#TGT ", "")
            Exit For
        Next
        
        If IsNumeric(tmp_pt) Then
            tmp_pt = CDbl(tmp_pt)
        End If
        
        
        tmp_src = extract_daily_tweet_rec(i)(dim_extract_text)
        
        Dim p As Double, r1 As Double, s1 As Double, r2 As Double, s2 As Double, r3 As Double, s3 As Double
        
        For j = 0 To UBound(vec_ticker, 1)
            If vec_ticker(j) = tmp_ticker Then
                
                If IsNumeric(output_bbg(j)(dim_bbg_last_price)) And IsNumeric(output_bbg(j)(dim_bbg_px_yest_low)) And IsNumeric(output_bbg(j)(dim_bbg_px_yest_high)) And IsNumeric(output_bbg(j)(dim_bbg_px_yest_close)) Then
                    
                    tmp_crncy = UCase(output_bbg(j)(dim_bbg_CRNCY))
                        For m = 0 To UBound(region, 1)
                            For n = 0 To UBound(region(m)(1), 1)
                                If tmp_crncy = region(m)(1)(n) Then
                                    tmp_region = region(m)(0)
                                End If
                            Next n
                        Next m
                        
                        For m = 0 To UBound(vec_currency, 1)
                            If UCase(vec_currency(m)(0)) = UCase(tmp_crncy) Then
                                tmp_fx = vec_currency(m)(2)
                                Exit For
                            End If
                        Next m
                        
                    
                    p = Round((output_bbg(j)(dim_bbg_px_yest_close) + output_bbg(j)(dim_bbg_px_yest_low) + output_bbg(j)(dim_bbg_px_yest_high)) / 3, 3)
                        
                    r1 = Round(2 * p - output_bbg(j)(dim_bbg_px_yest_low), 3)
                    s1 = Round(2 * p - output_bbg(j)(dim_bbg_px_yest_high), 3)
                    
                    r2 = Round((p - s1) + r1, 3)
                    s2 = Round(p - (r1 - s1), 3)
                    
                    r3 = Round((p - s2) + r2, 3)
                    s3 = Round(p - (r2 - s2), 3)
                    
                    
                    'determination de la qty
                    tmp_qty = Round(base_amount / (output_bbg(j)(dim_bbg_px_yest_close) * tmp_fx), 0)
                    
                    
                    'match equity db
                    For m = 0 To UBound(vec_equity_db, 1)
                        If vec_equity_db(m)(1) = tmp_ticker Then
                            tmp_line_equity_db = vec_equity_db(m)(0)
                            tmp_pos_equity_db = vec_equity_db(m)(2)
                            Exit For
                        Else
                            If m = UBound(vec_equity_db, 1) Then
                                tmp_line_equity_db = -1
                                tmp_pos_equity_db = 0
                            End If
                        End If
                    Next m
                    
                
                
                    'construction des trades
                    factor_side = 0
                    Dim vec_exec() As Variant
                    If tmp_side = "long" Then
                        vec_exec = Array(s3, s2, s1)
                        factor_side = 1
                        If p < output_bbg(j)(dim_bbg_last_price) Then
                            ReDim Preserve vec_exec(UBound(vec_exec, 1) + 1)
                            vec_exec(UBound(vec_exec, 1)) = p
                        End If
                        
                    ElseIf tmp_side = "short" Then
                        vec_exec = Array(r1, r2, r3)
                        factor_side = -1
                        If p > output_bbg(j)(dim_bbg_last_price) Then
                            ReDim Preserve vec_exec(UBound(vec_exec, 1) + 1)
                            vec_exec(UBound(vec_exec, 1)) = p
                        End If
                    End If
                
                    
                    For m = 0 To UBound(vec_exec, 1)
                        ReDim Preserve vec_trade(count_trade)
                        vec_trade(count_trade) = Array(tmp_ticker, factor_side * tmp_qty, vec_exec(m), tmp_crncy, tmp_region, tmp_line_equity_db, s3, s2, s1, p, r1, r2, r3, output_bbg(j)(dim_bbg_last_price), tmp_pos_equity_db, tmp_src, "base", tmp_broker)
                        count_trade = count_trade + 1
                    Next m
                    
                    
                
                End If
                
                Exit For
            End If
        Next j
        
    Next i
Else
    MsgBox ("No call to process")
End If

If count_trade > 0 Then
    Call preparation_trades_with_filters(vec_trade)
End If

End Sub



Public Function get_region_based_on_crncy() As Variant

get_region_based_on_crncy = Array(Array("Asia/Pacific", Array("JPY", "HKD", "AUD", "SGD", "TWD", "KRW", "INR", "THB", "CNY"), Array("PX_BID_ALL_SESSION", "PX_ASK_ALL_SESSION")), Array("Europe", Array("CHF", "EUR", "GBP", "SEK", "NOK", "DKK", "PLN"), Array("THEO_PRICE")), Array("America", Array("USD", "CAD", "BRL"), Array("PX_BID_ALL_SESSION", "PX_ASK_ALL_SESSION")))

End Function

Private Sub test_preparation_trades_with_filters()

Call preparation_trades_with_filters

End Sub


Sub preparation_trades_with_filters(Optional ByVal override_vec_trade As Variant)

Dim datetime_start As Date, datetime_end As Date

datetime_start = Now()

Application.Calculation = xlCalculationManual

Dim oBBG As New cls_Bloomberg_Sync

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
    oReg.Global = True

Dim debug_test As Variant
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, q As Integer, r As Integer, s As Integer

Dim oJSON As New JSONLib
Dim oTags As Collection
Dim oTag As Collection

Dim tmp_control As OLEObject

Dim src_file As String
src_file = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim path_sqlite_central As String
Dim path_possibility_db_central As Variant
path_possibility_db_central = Array("Q:\front\Stouff\RSSE\" & db_sqlite_central, src_file & db_sqlite_central)

Dim extract_central As Variant

Dim need_central As Boolean
    need_central = False


Dim vec_currency() As Variant
k = 0
For i = 14 To 31
    ReDim Preserve vec_currency(k)
    vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, Worksheets("Parametres").Cells(i, 6).Value)
    k = k + 1
Next i


Dim region As Variant
    region = get_region_based_on_crncy()

Dim output_bbg As Variant
    Dim bbg_field() As Variant
    bbg_field = Array("PX_YEST_CLOSE", "PX_YEST_LOW", "PX_YEST_HIGH", "LAST_PRICE", "DMI_ADX", "DMI_DIM", "DMI_DIP", "INTERVAL_BOLL_PERCENT_B", "RT_SIMP_SEC_STATUS")


Dim vec_ticker() As Variant
Dim vec_trades As Variant


If IsMissing(override_vec_trade) = False Then
    
    'appel api sur les champs necessaires
    
    'mount extract_central
    
    If IsEmpty(override_vec_trade) Then
        Exit Sub
    Else
        ReDim Preserve vec_ticker(0)
        vec_ticker(0) = UCase(override_vec_trade(0)(dim_vec_trade_ticker))
        
        k = 1
        For i = 1 To UBound(override_vec_trade, 1)
            For j = 0 To UBound(vec_ticker, 1)
                If UCase(override_vec_trade(i)(dim_vec_trade_ticker)) = UCase(vec_ticker(j)) Then
                    Exit For
                Else
                    If j = UBound(vec_ticker, 1) Then
                        ReDim Preserve vec_ticker(k)
                        vec_ticker(k) = UCase(override_vec_trade(i)(dim_vec_trade_ticker))
                        k = k + 1
                    End If
                End If
            Next j
        Next i
        
        output_bbg = oBBG.bdp(vec_ticker, bbg_field, output_format.of_vec_without_header)
        vec_trades = override_vec_trade
        extract_central = mount_sqlite_central
        
        GoTo insert_vec_trade_in_format2
    End If
    
End If


' 0 custom 1 field
' 1 internal </> numeric    ref_column, condition, column equity db, corr factor
' 2 internal cb txt
' 3 internal cb value * TO DEV
' 4 api txt
' 5 api </> numeric
' 6 </> central
' 7 json
' 8 central chg rank </> + timebase


Dim dim_code_filter As Integer, dim_filter_name As Integer, dim_value As Integer, dim_column_internal_db As Integer, _
    dim_value_factor As Integer, dim_trigger_control_box As Integer, dim_vec_ticker_details_src_is_customfield As Integer
    
    dim_code_filter = 0
    dim_filter_name = 1
    dim_value = 2
    dim_column_internal_db = 3
    dim_value_factor = 4
    dim_trigger_control_box = 5
    dim_vec_ticker_details_src_is_customfield = 6


Dim list_filters() As Variant
list_filters = Array(Array(0, "portfolio", "value", Empty, Empty, "control", False), Array(0, "twitter_basket", "value", Empty, Empty, "control", False), Array(0, "region", "value", Empty, Empty, "control", False), _
    Array(1, "Valeur_Euro", "value", 0, 1000, "control", False), Array(1, "Delta", "value", 6, 1, "control", False), Array(1, "Vega_1%_ALL", "value", 8, 1000, "control", False), Array(1, "Theta_ALL", "value", 9, 1, "control", False), Array(1, "EQY_BETA", "value", 23, 1, "control", False), _
    Array(2, "sector", "value", 53, Empty, "control", False), Array(2, "industry", "value", 54, Empty, "control", False), _
    Array(5, "rel_1d", "value", Empty, Empty, "control", False), Array(5, "rel_5d", "value", Empty, Empty, "control", False), _
    Array(6, "rank_eps", "value", Empty, Empty, "control", False), Array(6, "rank_moneyflow", "value", Empty, Empty, "control", False), Array(6, "rank_overall", "value", Empty, Empty, "control", False), _
    Array(7, "tag", "value", Empty, Empty, "control", False), _
    Array(8, "chg_rank_eps_<%timebase%>", "value", Empty, Empty, "control", False), Array(8, "chg_rank_eps_4w_chg_curr_yr_<%timebase%>", "value", Empty, Empty, "control", False), Array(8, "chg_rank_eps_4w_chg_nxt_yr_<%timebase%>", "value", Empty, Empty, "control", False), Array(8, "chg_rank_moneyflow_<%timebase%>", "value", Empty, Empty, "control", False), Array(8, "chg_rank_overall_<%timebase%>", "value", Empty, Empty, "control", False))


Dim override_control_name() As Variant
    override_control_name = Array(Array("EQY_BETA", "format2_beta"), Array("Theta_ALL", "format2_theta"), Array("Vega_1%_ALL", "format2_vega"), Array("Valeur_Euro", "format2_valeur_eur"), Array("Delta", "format2_delta"), Array("Valeur_Euro", "format2_valeur_eur"), _
        Array("chg_rank_eps_4w_chg_curr_yr", "frmt_chg_rnk_4w_crr"), Array("chg_rank_eps_4w_chg_nxt_yr", "frmt_chg_rnk_4w_nxt"), Array("chg_rank_moneyflow", "frmt_chg_rank_mf"), Array("chg_rank_overall", "frmt_chg_rank_ovrll"))
  


Dim count_column_equity_db As Integer
    count_column_equity_db = 0
Dim vec_column_equity_db() As Variant
For i = 0 To UBound(list_filters, 1)
    If list_filters(i)(dim_code_filter) = 1 Or list_filters(i)(dim_code_filter) = 2 Then
        
        If list_filters(i)(dim_column_internal_db) = 0 Then
            
            'repere les colonnes dans equity db
            If count_column_equity_db = 0 Then
                'mount un vecteur pour gagner du temps pour les eventuels appels suivants
                For j = 1 To 250
                    ReDim Preserve vec_column_equity_db(j)
                    vec_column_equity_db(j) = Worksheets("Equity_Database").Cells(25, j).Value
                Next j
            End If
            
        End If
        
        
        For j = 1 To UBound(vec_column_equity_db, 1)
            If UCase(list_filters(i)(dim_filter_name)) = UCase(vec_column_equity_db(j)) Then
                list_filters(i)(dim_column_internal_db) = j
                Exit For
            End If
        Next j
        
    End If
Next i



'extrapoluation du vec filters avec les custom fields
Dim prefix_custom_control_name As String
For Each tmp_control In Worksheets("FORMAT2").OLEObjects
    
    oReg.Pattern = "CB_format2_customfield\d+_type"
    
    Set matches = oReg.Execute(tmp_control.name)
    
    Dim everything_fine_with_custom_field As Boolean
    
    For Each match In matches
        
        everything_fine_with_custom_field = False
        
        prefix_custom_control_name = Replace(Replace(match.Value, "_type", ""), "CB_", "")
        
        If Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_type").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_name").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_sens").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("TB_" & prefix_custom_control_name & "_value").Object.Value <> "" Then
            
            'rajoute au filtre
            If Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_type").Object.Value = "API" Then
                
                If Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_sens").Object.Value = "=" Then
                    
                    ReDim Preserve list_filters(UBound(list_filters, 1) + 1)
                    list_filters(UBound(list_filters, 1)) = Array(4, LCase(Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_name").Object.Value), "value", Empty, Empty, "control", True)
                    
                    everything_fine_with_custom_field = True
                    
                ElseIf Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_sens").Object.Value = "<" Or Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_sens").Object.Value = ">" And IsNumeric(Worksheets("FORMAT2").OLEObjects("TB_" & prefix_custom_control_name & "_value").Object.Value) Then
                    
                    ReDim Preserve list_filters(UBound(list_filters, 1) + 1)
                    list_filters(UBound(list_filters, 1)) = Array(5, LCase(Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_name").Object.Value), "value", Empty, Empty, "control", True)
                    
                    everything_fine_with_custom_field = True
                    
                End If
                
            ElseIf Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_type").Object.Value = "INTERNAL" Then
                
                'check si le champ est trouvable
                For i = 1 To 250
                    
                    If UCase(Worksheets("Equity_Database").Cells(25, i).Value) = UCase(Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_name").Object.Value) Then
                        
                        ReDim Preserve list_filters(UBound(list_filters, 1) + 1)
                        list_filters(UBound(list_filters, 1)) = Array(1, LCase(Worksheets("Equity_Database").Cells(25, i).Value), "value", i, 1, "control", True)
                        
                        everything_fine_with_custom_field = True
                        
                        Exit For
                    End If
                Next i
            ElseIf Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_type").Object.Value = "CENTRAL" Then
                
                If Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_name").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_sens").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("TB_" & prefix_custom_control_name & "_value").Object.Value <> "" Then
                    
                    If IsNumeric(Worksheets("FORMAT2").OLEObjects("TB_" & prefix_custom_control_name & "_value").Object.Value) Then
                    
                        'check que field existe bien
                        If IsEmpty(extract_central) Then
                            
                            For i = 0 To UBound(path_possibility_db_central, 1)
                                If exist_file(path_possibility_db_central(i)) Then
                                    path_sqlite_central = path_possibility_db_central(i)
                                    
                                    extract_central = sqlite3_query(path_sqlite_central, "SELECT * FROM t_custom_rank ORDER BY Ticker")
                                    
                                    If UBound(extract_central, 1) > 0 Then
                                        For j = 0 To UBound(extract_central, 1)
                                            extract_central(i)(0) = patch_ticker_marketplace(UCase(extract_central(i)(0)))
                                        Next j
                                    End If
                                    
                                    Exit For
                                End If
                            Next i
                            
                        End If
                        
                        
                        If IsEmpty(extract_central) = False Then
                            
                            For i = 0 To UBound(extract_central(0), 1)
                                If UCase(extract_central(0)(i)) = UCase(Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_name").Object.Value) Then
                                    
                                    ReDim Preserve list_filters(UBound(list_filters, 1) + 1)
                                    list_filters(UBound(list_filters, 1)) = Array(6, LCase(extract_central(0)(i)), "value", Empty, Empty, "control", True)
                                    everything_fine_with_custom_field = True
                                    need_central = True
                                    
                                    Exit For
                                End If
                            Next i
                            
                        End If
                    End If
                End If
                
            End If
            
            
            'mise en place de l override
            If everything_fine_with_custom_field = True Then
                ReDim Preserve override_control_name(UBound(override_control_name, 1) + 1)
                override_control_name(UBound(override_control_name, 1)) = Array(LCase(Worksheets("FORMAT2").OLEObjects("CB_" & prefix_custom_control_name & "_name").Object.Value), prefix_custom_control_name)
            End If
            
            Exit For
            
        End If
        
    Next
    
Next
    
    
    
    
    'construction prefix des control de la feuille
    Dim prefix_object_control_in_sheet() As Variant
    ReDim Preserve prefix_object_control_in_sheet(0)
    prefix_object_control_in_sheet(0) = "format2"
    
    For i = 0 To UBound(override_control_name, 1)
        For j = 0 To UBound(prefix_object_control_in_sheet, 1)
            If prefix_object_control_in_sheet(j) = Left(override_control_name(i)(1), InStr(override_control_name(i)(1), "_") - 1) Then
                Exit For
            Else
                If j = UBound(prefix_object_control_in_sheet, 1) Then
                    ReDim Preserve prefix_object_control_in_sheet(UBound(prefix_object_control_in_sheet, 1) + 1)
                    prefix_object_control_in_sheet(UBound(prefix_object_control_in_sheet, 1)) = Left(override_control_name(i)(1), InStr(override_control_name(i)(1), "_") - 1)
                End If
            End If
        Next j
    Next i
    

Dim l_equity_db_header As Integer, c_equity_db_valeur_eur As Integer, c_equity_db_delta As Integer, _
    c_equity_db_theta As Integer, c_equity_db_ticker As Integer, c_equity_db_crncy As Integer, c_equity_db_tag As Integer
    
    
l_equity_db_header = 25
c_equity_db_valeur_eur = 5
c_equity_db_delta = 6
c_equity_db_theta = 9
c_equity_db_net_position = 26
c_equity_db_ticker = 47
c_equity_db_crncy = 44
c_equity_db_tag = 137
c_equity_db_perso_rel_1d = 138


'mount db view
Dim base_control_name As String
Dim activate_filters() As Variant
ReDim activate_filters(0)
activate_filters(0) = Array(0, 0, 0, 0, 0, "control")
k = 0
For i = 0 To UBound(list_filters, 1)
        
    
    'possibilite d override du nom du champs si logic trop longue
    Dim found_control_without_override As Boolean
        found_control_without_override = False
    
    If list_filters(i)(dim_vec_ticker_details_src_is_customfield) = False Then
        
        'check en priorite * champs avant de regarder dans override
        For Each tmp_control In Worksheets("FORMAT2").OLEObjects
            If InStr(UCase(tmp_control.name), UCase("CB_format2_" & Replace(list_filters(i)(1), "_<%timebase%>", ""))) <> 0 Or InStr(UCase(tmp_control.name), UCase("TB_format2_" & Replace(list_filters(i)(1), "_<%timebase%>", ""))) <> 0 Then
                found_control_without_override = True
                base_control_name = "format2_" & Replace(list_filters(i)(1), "_<%timebase%>", "")
                Exit For
            End If
        Next
        
        If found_control_without_override = False Then
            GoTo override_control_name
        End If
        
    Else
override_control_name:
        For j = 0 To UBound(override_control_name, 1)
            If UCase(Replace(list_filters(i)(1), "_<%timebase%>", "")) = UCase(override_control_name(j)(0)) Then
                base_control_name = override_control_name(j)(1)
                Exit For
            Else
                If j = UBound(override_control_name, 1) Then
                    MsgBox ("problem with filter " & i)
                    Exit Sub
                End If
            End If
        Next j
    End If
    
    
    For Each tmp_control In Worksheets("FORMAT2").OLEObjects
        
        For u = 0 To UBound(prefix_object_control_in_sheet, 1)
            
            If InStr(tmp_control.name, prefix_object_control_in_sheet(u)) <> 0 Then
                
                If InStr(tmp_control.name, base_control_name) <> 0 And UCase(Left(tmp_control.name, 2)) <> "L_" Then
                    
                    If list_filters(i)(dim_code_filter) = 0 Then 'custom 1 field
                        
                        If Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name).Object.Value <> "" Then
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array("txt", Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name).Object.Value)
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                    End If
                                End If
                            Next j
                        End If
                        
                    ElseIf list_filters(i)(dim_code_filter) = 1 Then 'internal </> numeric
                        
                        If Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value <> "" And IsNumeric(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value) Then
                            
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array(Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value, CDbl(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value))
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                    End If
                                End If
                            Next j
                            
                        End If
                        
                    ElseIf list_filters(i)(dim_code_filter) = 2 Then 'internal cb txt
                        
                        If Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name).Object.Value <> "" Then
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array("txt", Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name).Object.Value)
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                    End If
                                End If
                            Next j
                        End If
                        
                    ElseIf list_filters(i)(dim_code_filter) = 3 Then 'internal cb value
                        
                    ElseIf list_filters(i)(dim_code_filter) = 4 Then 'api txt
                        
                        If Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value <> "" Then
                            
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array("=", Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value)
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                        
                                    End If
                                End If
                            Next j
                            
                        End If
                        
                    ElseIf list_filters(i)(dim_code_filter) = 5 Then 'api </> numeric
                    
                        If Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value <> "" And IsNumeric(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value) Then
                            
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array(Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value, CDbl(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value))
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                        
                                    End If
                                End If
                            Next j
                            
                        End If
                        
                        
                    ElseIf list_filters(i)(dim_code_filter) = 6 Then ' </> central
                        
                        
                        If Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value <> "" And IsNumeric(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value) Then
                            
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        need_central = True
                                        
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array(Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value, CDbl(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value))
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                    End If
                                End If
                            Next j
                            
                        End If
                        
                    ElseIf list_filters(i)(dim_code_filter) = 7 Then 'json tag
                        
                        If Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name).Object.Value <> "" Then
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array("txt", Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name).Object.Value)
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                    End If
                                End If
                            Next j
                        End If
                    
                    ElseIf list_filters(i)(dim_code_filter) = 8 Then 'central chg rank </> + timebase
                        
                        If Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_timebase").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value <> "" And Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value <> "" And IsNumeric(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value) Then
                            
                            timebase = Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_timebase").Object.Value
                            
                            'conversion de la timebase
                            list_filters(i)(1) = Replace(list_filters(i)(1), "<%timebase%>", Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_timebase").Object.Value)
                            
                            
                            'check si deja pas active
                            For j = 0 To UBound(activate_filters, 1)
                                If activate_filters(j)(dim_filter_name) = list_filters(i)(dim_filter_name) And base_control_name = activate_filters(j)(dim_trigger_control_box) Then
                                    Exit For
                                Else
                                    If j = UBound(activate_filters, 1) Then
                                        need_central = True
                                        
                                        ReDim Preserve activate_filters(k)
                                        list_filters(i)(dim_trigger_control_box) = base_control_name
                                        list_filters(i)(dim_value) = Array(Worksheets("FORMAT2").OLEObjects("CB_" & base_control_name & "_sens").Object.Value, CDbl(Worksheets("FORMAT2").OLEObjects("TB_" & base_control_name & "_value").Object.Value))
                                        activate_filters(k) = list_filters(i)
                                        k = k + 1
                                    End If
                                End If
                            Next j
                            
                        End If
                    End If
                    
                End If
            
            End If
            
        Next u
    Next
Next i



If need_central = True And IsEmpty(extract_central) Then
    For i = 0 To UBound(path_possibility_db_central, 1)
        If exist_file(path_possibility_db_central(i)) Then
            path_sqlite_central = path_possibility_db_central(i)
            
            extract_central = sqlite3_query(path_sqlite_central, "SELECT * FROM t_custom_rank ORDER BY Ticker")
            
            If UBound(extract_central, 1) > 0 Then
                For j = 0 To UBound(extract_central, 1)
                    extract_central(i)(0) = patch_ticker_marketplace(UCase(extract_central(i)(0)))
                Next j
            End If
            
            Exit For
        End If
    Next i
End If

Dim tmp_ticker As String
Dim find_in_central As Boolean
Dim field_bbg_api_filter() As Variant
    ReDim field_bbg_api_filter(0)
    field_bbg_api_filter(0) = ""
Dim count_field_bbg_api As Integer
    count_field_bbg_api = 0
Dim take_the_security As Boolean
Dim need_api As Boolean
Dim vec_ticker_details() As Variant, vec_ticker_tmp() As Variant


Dim take_the_security_tag As Boolean
Dim take_the_security_region As Boolean

Dim portfolio_list_ticker As Variant

Dim vec_prt_api_ticker() As Variant
Dim vec_equity_db() As Variant

Dim output_bdp_prt As Variant

need_api = False



Dim tmp_controlOLE As OLEObject
Dim vec_tradtor_enable() As Variant
    Dim count_tradtor_enable As Integer
    count_tradtor_enable = 0
For Each tmp_controlOLE In Worksheets("FORMAT2").OLEObjects
    If InStr(tmp_controlOLE.name, "tradator") <> 0 Then
        
        If tmp_controlOLE.Object.Value = True And TypeOf tmp_controlOLE.Object Is msforms.CheckBox Then
            ReDim Preserve vec_tradtor_enable(count_tradtor_enable)
            vec_tradtor_enable(count_tradtor_enable) = "@" & Replace(UCase(tmp_controlOLE.name), UCase("CB_TRADATOR_"), "")
            count_tradtor_enable = count_tradtor_enable + 1
        End If
    End If
Next


Dim vec_simple_trades As Variant
Dim get_trades_from_standardized As Boolean
    get_trades_from_standardized = False


If k > 0 Then
get_input_tickers:
    m = 0

    Dim count_src As Integer
        count_src = -1

    If Worksheets("FORMAT2").CB_format2_portfolio.Value <> "" Or Worksheets("FORMAT2").CB_format2_twitter_basket.Value <> "" Or Worksheets("FORMAT2").CB_format2_eqs.Value <> "" Or count_tradtor_enable <> 0 Then
        
        portfolio_list_ticker = get_input_format2_standardized_portfolio_from_external_sources
        
        If IsEmpty(portfolio_list_ticker) Then
            MsgBox ("problem with extern sources, -> Exit.")
            Exit Sub
        Else
            
            'resplit si double reception portfolio_list_ticker / vec_trades
            vec_simple_trades = Empty
            On Error GoTo portfolio_list_ticker_std
            If IsArray(portfolio_list_ticker(0)) Then
                
                vec_simple_trades = portfolio_list_ticker(1)
                portfolio_list_ticker = portfolio_list_ticker(0)
                get_trades_from_standardized = True
                
                GoTo portfolio_list_ticker_std
            Else
portfolio_list_ticker_std:
                On Error GoTo 0
                
                For i = 0 To UBound(portfolio_list_ticker, 2)
                    If UCase(portfolio_list_ticker(0, i)) = UCase("ticker") Then
                        dim_prt_ticker = i
                    ElseIf portfolio_list_ticker(0, i) = "vec_crncy" Then
                        dim_prt_crncy_txt = i
                    ElseIf InStr(portfolio_list_ticker(0, i), "src") <> 0 Then
                        dim_prt_source = i
                        count_src = CInt(Replace(portfolio_list_ticker(0, i), "src:", ""))
                    End If
                Next i
            End If
            
        End If
        
        
        'retour procedure dans preparation_trades_with_filters
        
        For i = 1 To UBound(portfolio_list_ticker, 1)
            
            take_the_security = False
            
            If portfolio_list_ticker(i, dim_prt_ticker) <> False Then
            
                take_the_security = True
                'passe dans les filtres
                For j = 0 To UBound(activate_filters, 1)
                    
                    If activate_filters(j)(0) = 0 Then 'custom
                        
                        If activate_filters(j)(dim_filter_name) = "region" Then
                            
                            take_the_security_region = False
                            
                            For q = 0 To UBound(region, 1)
                                For r = 0 To UBound(region(q)(1), 1)
                                    If region(q)(1)(r) = portfolio_list_ticker(i, dim_prt_crncy_txt)(1) Then
                                        If activate_filters(j)(dim_value)(1) = region(q)(0) Then
                                            take_the_security_region = True
                                            Exit For
                                        Else
                                            take_the_security_region = False
                                            Exit For
                                        End If
                                    End If
                                Next r
                            Next q
                                    
                            
                            If take_the_security_region = True Then
                            Else
                                take_the_security = False
                            End If
                            
                        End If
                        
                    ElseIf activate_filters(j)(0) = 1 Then 'internal num
                        
                        If portfolio_list_ticker(i, dim_prt_crncy_txt)(0) = -1 Then
                            'le titre n'est pas dans equity database
                            take_the_security = False
                        Else
                        
                            If IsError(Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), activate_filters(j)(dim_column_internal_db))) = False Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), activate_filters(j)(dim_column_internal_db))) Then
                                    If activate_filters(j)(dim_value)(0) = ">" Then
                                        If activate_filters(j)(dim_value_factor) * Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), activate_filters(j)(dim_column_internal_db)) >= activate_filters(j)(dim_value)(1) Then
                                        Else
                                            take_the_security = False
                                        End If
                                    ElseIf activate_filters(j)(dim_value)(0) = "<" Then
                                        If activate_filters(j)(dim_value_factor) * Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), activate_filters(j)(dim_column_internal_db)) < activate_filters(j)(dim_value)(1) Then
                                        
                                        Else
                                            take_the_security = False
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                        
                        
                    ElseIf activate_filters(j)(0) = 2 Then 'internal txt
                        
                        If portfolio_list_ticker(i, dim_prt_crncy_txt)(0) = -1 Then
                            'le titre n'est pas dans equity database
                            take_the_security = False
                        Else
                            
                            If UCase(Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), activate_filters(j)(dim_column_internal_db))) = UCase(activate_filters(j)(dim_value)(1)) Then
                            Else
                                take_the_security = False
                            End If
                            
                        End If
                        
                    ElseIf activate_filters(j)(0) = 4 Then 'api text =
                        need_api = True
                        
                        For n = 0 To UBound(field_bbg_api_filter, 1)
                            If field_bbg_api_filter(n) = activate_filters(j)(dim_filter_name) Then
                                Exit For
                            Else
                                If n = UBound(field_bbg_api_filter, 1) Then
                                    ReDim Preserve field_bbg_api_filter(count_field_bbg_api)
                                    field_bbg_api_filter(count_field_bbg_api) = activate_filters(j)(dim_filter_name)
                                    count_field_bbg_api = count_field_bbg_api + 1
                                End If
                            End If
                        Next n
                        
                        
                    ElseIf activate_filters(j)(0) = 5 Then 'api value </>
                        need_api = True
                        
                        For n = 0 To UBound(field_bbg_api_filter, 1)
                            If field_bbg_api_filter(n) = activate_filters(j)(dim_filter_name) Then
                                Exit For
                            Else
                                If n = UBound(field_bbg_api_filter, 1) Then
                                    ReDim Preserve field_bbg_api_filter(count_field_bbg_api)
                                    field_bbg_api_filter(count_field_bbg_api) = activate_filters(j)(dim_filter_name)
                                    count_field_bbg_api = count_field_bbg_api + 1
                                End If
                            End If
                        Next n
                        
                    ElseIf activate_filters(j)(0) = 6 Or activate_filters(j)(0) = 8 Then  'central
                        
                        tmp_ticker = patch_ticker_marketplace(UCase(portfolio_list_ticker(i, dim_prt_ticker)))
                            
                        find_in_central = False
                        
                        If take_the_security = True Then
                            If UBound(extract_central, 1) > 0 Then
                                For n = 1 To UBound(extract_central, 1) And find_in_central = False
                                    
                                        If UCase(extract_central(n)(0)) = UCase(tmp_ticker) Then
                                            For q = 0 To UBound(extract_central(0), 1)
                                                
                                                If UCase(extract_central(0)(q)) = UCase(activate_filters(j)(dim_filter_name)) Then
                                                    
                                                    If activate_filters(j)(dim_value)(0) = "<" Then
                                                        
                                                        If extract_central(n)(q) < activate_filters(j)(dim_value)(1) Then
                                                            find_in_central = True
                                                            GoTo check_next_filter_prt
                                                        Else
                                                            take_the_security = False
                                                            find_in_central = True
                                                            GoTo check_next_filter_prt
                                                        End If
                                                        
                                                    ElseIf activate_filters(j)(dim_value)(0) = ">" Then
                                                        
                                                        If extract_central(n)(q) >= activate_filters(j)(dim_value)(1) Then
                                                            find_in_central = True
                                                            GoTo check_next_filter_prt
                                                        Else
                                                            take_the_security = False
                                                            find_in_central = True
                                                            GoTo check_next_filter_prt
                                                        End If
                                                        
                                                    End If
                                                    
                                                    
                                                    Exit For
                                                End If
                                            Next q
                                        End If
                                        
                                Next n
                                
                                If find_in_central = False Then
                                    take_the_security = False
                                End If
                            End If
                        End If
                        
                    ElseIf activate_filters(j)(0) = 7 Then 'json
                        
                        If portfolio_list_ticker(i, dim_prt_crncy_txt)(0) = -1 Then
                            'le titre n'est pas dans equity database
                            take_the_security = False
                        Else
                            
                            If Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), c_equity_db_tag) = "" Then
                                take_the_security = False
                            Else
                                
                                Set oTags = oJSON.parse(Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), c_equity_db_tag).Value)
                                
                                take_the_security_tag = False
                                For Each oTag In oTags
                                debug_test = oTag.Item(2)
                                    If UCase(oTag.Item(2)) = "OPEN" Then
                                        
                                        If UCase(oTag.Item(3)) = UCase(activate_filters(j)(dim_value)(1)) Then
                                            take_the_security_tag = True
                                        End If
                                        
                                    End If
                                Next
                                
                                If take_the_security_tag = True Then
                                Else
                                    take_the_security = False
                                End If
                                
                            End If
                            
                            
                        End If
                        
                    End If
check_next_filter_prt:
                Next j
            End If
            
            
            If take_the_security = True Then
                ReDim Preserve vec_ticker_details(m)

                For n = 0 To UBound(vec_currency, 1)
                    If vec_currency(n)(0) = portfolio_list_ticker(i, dim_prt_crncy_txt)(1) Then

                        'rematch la region, necessaire au premarket price
                        For s = 0 To UBound(region, 1)
                            For q = 0 To UBound(region(s)(1), 1)
                                If UCase(region(s)(1)(q)) = UCase(vec_currency(n)(0)) Then
                                
                                    If portfolio_list_ticker(i, dim_prt_crncy_txt)(0) <> -1 Then
                                        vec_ticker_details(m) = Array(patch_ticker_marketplace(portfolio_list_ticker(i, dim_prt_ticker)), vec_currency(n)(0), portfolio_list_ticker(i, dim_prt_crncy_txt)(0), Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), c_equity_db_delta).Value, region(s)(0), Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), c_equity_db_theta).Value, Worksheets("Equity_Database").Cells(portfolio_list_ticker(i, dim_prt_crncy_txt)(0), c_equity_db_net_position).Value, portfolio_list_ticker(i, dim_prt_source))
                                    Else
                                        vec_ticker_details(m) = Array(patch_ticker_marketplace(portfolio_list_ticker(i, dim_prt_ticker)), vec_currency(n)(0), portfolio_list_ticker(i, dim_prt_crncy_txt)(0), 0, region(s)(0), 0, 0, portfolio_list_ticker(i, dim_prt_source))
                                    End If
                                    
                                    Exit For
                                End If
                            Next q
                        Next s


                        Exit For
                    End If
                Next n

                ReDim Preserve vec_ticker(m)
                vec_ticker(m) = patch_ticker_marketplace(portfolio_list_ticker(i, dim_prt_ticker))

                m = m + 1
            End If
            
            
        Next i
        
    Else
    ' pas de prefitre sur input on tape donc dans equity db
    ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ EQUITY DATABASE @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        For i = l_equity_db_header + 2 To 32000 Step 2
            
            count_src = 1
            
            take_the_security = True
            
            If Worksheets("Equity_Database").Cells(i, 1) = "" Then
                Exit For
            Else
                For j = 0 To UBound(activate_filters, 1)
                    If activate_filters(j)(0) = 0 Then
                        
                        If activate_filters(j)(dim_filter_name) = "region" Then
                            
                            take_the_security_region = False
                            
                            For n = 0 To UBound(vec_currency, 1)
                                If vec_currency(n)(1) = Worksheets("Equity_Database").Cells(i, c_equity_db_crncy) Then
                                    
                                    For q = 0 To UBound(region, 1)
                                        For r = 0 To UBound(region(q)(1), 1)
                                            If region(q)(1)(r) = vec_currency(n)(0) Then
                                                If activate_filters(j)(dim_value)(1) = region(q)(0) Then
                                                    take_the_security_region = True
                                                    Exit For
                                                Else
                                                    take_the_security_region = False
                                                    Exit For
                                                End If
                                            End If
                                        Next r
                                    Next q
                                    
                                    Exit For
                                End If
                            Next n
                            
                            If take_the_security_region = True Then
                            Else
                                take_the_security = False
                            End If
                            
                        End If
                        
                    ElseIf activate_filters(j)(0) = 1 Then
                        
                        If IsError(Worksheets("Equity_Database").Cells(i, activate_filters(j)(dim_column_internal_db))) = False Then
                            If IsNumeric(Worksheets("Equity_Database").Cells(i, activate_filters(j)(dim_column_internal_db))) Then
                                If activate_filters(j)(dim_value)(0) = ">" Then
                                    If activate_filters(j)(dim_value_factor) * Worksheets("Equity_Database").Cells(i, activate_filters(j)(dim_column_internal_db)) >= activate_filters(j)(dim_value)(1) Then
                                    Else
                                        take_the_security = False
                                    End If
                                ElseIf activate_filters(j)(dim_value)(0) = "<" Then
                                    If activate_filters(j)(dim_value_factor) * Worksheets("Equity_Database").Cells(i, activate_filters(j)(dim_column_internal_db)) < activate_filters(j)(dim_value)(1) Then
                                    
                                    Else
                                        take_the_security = False
                                    End If
                                End If
                            End If
                        End If
                        
                    ElseIf activate_filters(j)(0) = 2 Then
                        
                        If Worksheets("Equity_Database").Cells(i, activate_filters(j)(dim_column_internal_db)) = activate_filters(j)(dim_value)(1) Then
                        Else
                            take_the_security = False
                        End If
                        
                    ElseIf activate_filters(j)(0) = 3 Then
                    ElseIf activate_filters(j)(0) = 4 Then 'api text =
                        need_api = True
                        
                        For n = 0 To UBound(field_bbg_api_filter, 1)
                            If field_bbg_api_filter(n) = activate_filters(j)(dim_filter_name) Then
                                Exit For
                            Else
                                If n = UBound(field_bbg_api_filter, 1) Then
                                    ReDim Preserve field_bbg_api_filter(count_field_bbg_api)
                                    field_bbg_api_filter(count_field_bbg_api) = activate_filters(j)(dim_filter_name)
                                    count_field_bbg_api = count_field_bbg_api + 1
                                End If
                            End If
                        Next n
                        
                    ElseIf activate_filters(j)(0) = 5 Then 'api num </>
                        need_api = True
                        
                        For n = 0 To UBound(field_bbg_api_filter, 1)
                            If field_bbg_api_filter(n) = activate_filters(j)(dim_filter_name) Then
                                Exit For
                            Else
                                If n = UBound(field_bbg_api_filter, 1) Then
                                    ReDim Preserve field_bbg_api_filter(count_field_bbg_api)
                                    field_bbg_api_filter(count_field_bbg_api) = activate_filters(j)(dim_filter_name)
                                    count_field_bbg_api = count_field_bbg_api + 1
                                End If
                            End If
                        Next n
                        
                    ElseIf activate_filters(j)(0) = 6 Or activate_filters(j)(0) = 8 Then 'central
                        
                        tmp_ticker = patch_ticker_marketplace(UCase(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value))
                        
                        find_in_central = False
                        
                        If take_the_security = True Then
                            If UBound(extract_central, 1) > 0 Then
                                For n = 1 To UBound(extract_central, 1) And find_in_central = False
                                    
                                        If UCase(extract_central(n)(0)) = UCase(tmp_ticker) Then
                                            For q = 0 To UBound(extract_central(0), 1)
                                                
                                                If UCase(extract_central(0)(q)) = UCase(activate_filters(j)(dim_filter_name)) Then
                                                    
                                                    If activate_filters(j)(dim_value)(0) = "<" Then
                                                        
                                                        If extract_central(n)(q) < activate_filters(j)(dim_value)(1) Then
                                                            find_in_central = True
                                                            GoTo check_next_filter
                                                        Else
                                                            take_the_security = False
                                                            find_in_central = True
                                                            GoTo check_next_filter
                                                        End If
                                                        
                                                    ElseIf activate_filters(j)(dim_value)(0) = ">" Then
                                                        
                                                        If extract_central(n)(q) >= activate_filters(j)(dim_value)(1) Then
                                                            find_in_central = True
                                                            GoTo check_next_filter
                                                        Else
                                                            take_the_security = False
                                                            find_in_central = True
                                                            GoTo check_next_filter
                                                        End If
                                                        
                                                    End If
                                                    
                                                    
                                                    Exit For
                                                End If
                                            Next q
                                        End If
                                        
                                Next n
                                
                                If find_in_central = False Then
                                    take_the_security = False
                                End If
                            End If
                        End If
                    
                    ElseIf activate_filters(j)(0) = 7 Then 'json tag
                        
                        If Worksheets("Equity_Database").Cells(i, c_equity_db_tag) = "" Then
                            take_the_security = False
                        Else
                            
                            Set oTags = oJSON.parse(Worksheets("Equity_Database").Cells(i, c_equity_db_tag).Value)
                            
                            take_the_security_tag = False
                            For Each oTag In oTags
                            debug_test = oTag.Item(2)
                                If UCase(oTag.Item(2)) = "OPEN" Then
                                    
                                    If UCase(oTag.Item(3)) = UCase(activate_filters(j)(dim_value)(1)) Then
                                        take_the_security_tag = True
                                    End If
                                    
                                End If
                            Next
                            
                            If take_the_security_tag = True Then
                            Else
                                take_the_security = False
                            End If
                            
                        End If
                            
                        
                    End If
check_next_filter:
                Next j
                
                If take_the_security = True Then
                    ReDim Preserve vec_ticker_details(m)
                    
                    For n = 0 To UBound(vec_currency, 1)
                        If vec_currency(n)(1) = Worksheets("Equity_Database").Cells(i, c_equity_db_crncy).Value Then
                            
                            'rematch la region, necessaire au premarket price
                            For s = 0 To UBound(region, 1)
                                For q = 0 To UBound(region(s)(1), 1)
                                    If UCase(region(s)(1)(q)) = UCase(vec_currency(n)(0)) Then
                                        vec_ticker_details(m) = Array(patch_ticker_marketplace(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value), vec_currency(n)(0), i, Worksheets("Equity_Database").Cells(i, c_equity_db_delta).Value, region(s)(0), Worksheets("Equity_Database").Cells(i, c_equity_db_theta).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_net_position).Value, "Equity DB")
                                        Exit For
                                    End If
                                Next q
                            Next s
                            
                            
                            Exit For
                        End If
                    Next n
                    
                    ReDim Preserve vec_ticker(m)
                    vec_ticker(m) = patch_ticker_marketplace(UCase(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value))
                    
                    m = m + 1
                End If
            End If
            
        Next i
    
    End If
    
    
    'appel api si necessaire
    If need_api = True Then
        Dim output_bbg_api_filter As Variant
        output_bbg_api_filter = oBBG.bdp(vec_ticker, field_bbg_api_filter, output_format.of_vec_without_header)
        
        n = 0
        For i = 0 To UBound(output_bbg_api_filter, 1)
            
            take_the_security = True
            
            'repasse les filtres et check pour api
            For j = 0 To UBound(activate_filters, 1)
                If activate_filters(j)(0) = 4 Then
                    For m = 0 To UBound(field_bbg_api_filter, 1)
                        If activate_filters(j)(dim_filter_name) = field_bbg_api_filter(m) Then
                            If UCase(CStr(output_bbg_api_filter(i)(m))) = UCase(CStr(activate_filters(j)(dim_value)(1))) Then
                                debug_test = "Ok"
                            Else
                                take_the_security = False
                                Exit For
                            End If
                        End If
                    Next m
                    
                ElseIf activate_filters(j)(0) = 5 Then
                    For m = 0 To UBound(field_bbg_api_filter, 1)
                        If activate_filters(j)(dim_filter_name) = field_bbg_api_filter(m) Then
                            
                            If IsNumeric(output_bbg_api_filter(i)(m)) Then
                                If activate_filters(j)(dim_value)(0) = "<" Then
                                    
                                    If output_bbg_api_filter(i)(m) < activate_filters(j)(dim_value)(1) Then
                                    
                                    Else
                                        take_the_security = False
                                        Exit For
                                    End If
                                    
                                ElseIf activate_filters(j)(dim_value)(0) = ">" Then
                                    
                                    If output_bbg_api_filter(i)(m) >= activate_filters(j)(dim_value)(1) Then
                                    
                                    Else
                                        take_the_security = False
                                        Exit For
                                    End If
                                
                                End If
                            Else
                                take_the_security = False
                                Exit For
                            End If
                            
                            Exit For
                        End If
                    Next m
                End If
            Next j
            
            
            If take_the_security = True Then
                ReDim Preserve vec_ticker_tmp(n)
                vec_ticker_tmp(n) = vec_ticker_details(i)
                n = n + 1
            End If
        
        Next i
        
        If n > 0 Then
            vec_ticker_details = vec_ticker_tmp
        End If
        
        m = n 'compteur des pos remplissant les filtres
        
    End If
    
Else
    If Worksheets("FORMAT2").CB_format2_eqs.Value <> "" Then
        GoTo get_input_tickers 'bypass si eqs comme deja un filtre en lui meme
    Else
        answer = MsgBox("No filter selected. Continue with all entries in the input ?", vbYesNo, "No filter")
        
        If answer = vbYes Then
            GoTo get_input_tickers
        Else
            Exit Sub
        End If
    End If
End If



If m > 0 Then
    For i = 0 To UBound(vec_ticker_details, 1)
        
        min_value = vec_ticker_details(i)(dim_vec_ticker_details_crncy)
        min_pos = i
        
        
        For j = i + 1 To UBound(vec_ticker_details, 1)
            If vec_ticker_details(j)(dim_vec_ticker_details_crncy) < min_value Then
                min_value = vec_ticker_details(j)(dim_vec_ticker_details_crncy)
                min_pos = j
            End If
        Next j
        
        If i <> min_pos Then
            
            tmp_var = vec_ticker_details(i)
            vec_ticker_details(i) = vec_ticker_details(min_pos)
            vec_ticker_details(min_pos) = tmp_var
            
        End If
    Next i


    'second trie par devise par ticker
    For i = 0 To UBound(vec_ticker_details, 1)
        
        min_value = vec_ticker_details(i)(dim_vec_ticker_details_ticker)
        min_pos = i
        
        For j = i + 1 To UBound(vec_ticker_details, 1)
            If vec_ticker_details(j)(dim_vec_ticker_details_crncy) = vec_ticker_details(i)(dim_vec_ticker_details_crncy) Then
                
                If vec_ticker_details(j)(dim_vec_ticker_details_ticker) < min_value Then
                    
                    min_value = vec_ticker_details(j)(dim_vec_ticker_details_ticker)
                    min_pos = j
                    
                End If
                
                
            Else
                Exit For
            End If
        Next j
        
        If i <> min_pos Then
            tmp_var = vec_ticker_details(i)
            vec_ticker_details(i) = vec_ticker_details(min_pos)
            vec_ticker_details(min_pos) = tmp_var
        End If
        
    Next i
    
    
    
    ReDim vec_ticker(UBound(vec_ticker_details, 1))
    
    For i = 0 To UBound(vec_ticker_details, 1)
        vec_ticker(i) = vec_ticker_details(i)(dim_vec_ticker_details_ticker)
    Next i
    
    
    
    
    'appel bbg pour le calcul des pivots
        If Worksheets("FORMAT2").CB_format2_eqs.Value <> "" Then
            
            Dim list_eqs As Variant
            list_eqs = get_input_format2_vec_eqs()
            
            Dim tmp_vec_fields As Variant
            For i = 0 To UBound(list_eqs, 1)
                
                'stop
                tmp_vec_fields = get_eqs_fields_stop_target_formula(list_eqs(i), c_format2_eqs_stop_formula)
                
                'passe en revue les fields
                If IsEmpty(tmp_vec_fields) Then
                Else
                    For j = 0 To UBound(tmp_vec_fields, 1)
                        
                        If Left(tmp_vec_fields(j), 1) = "$" Then
                            
                            For m = 0 To UBound(bbg_field, 1)
                                
                                If UCase(Mid(tmp_vec_fields(j), 2)) = UCase(bbg_field(m)) Then
                                    Exit For
                                Else
                                    If m = UBound(bbg_field, 1) Then
                                        'add bbg field
                                        ReDim Preserve bbg_field(UBound(bbg_field, 1) + 1)
                                        bbg_field(UBound(bbg_field, 1)) = UCase(Mid(tmp_vec_fields(j), 2))
                                    End If
                                End If
                                
                            Next m
                            
                        End If
                        
                    Next j
                End If
                
                
                'target
                tmp_vec_fields = get_eqs_fields_stop_target_formula(list_eqs(i), c_format2_eqs_target_formula)
                
                'passe en revue les fields
                If IsEmpty(tmp_vec_fields) Then
                Else
                    For j = 0 To UBound(tmp_vec_fields, 1)
                        
                        If Left(tmp_vec_fields(j), 1) = "$" Then
                            
                            For m = 0 To UBound(bbg_field, 1)
                                
                                If UCase(Mid(tmp_vec_fields(j), 2)) = UCase(bbg_field(m)) Then
                                    Exit For
                                Else
                                    If m = UBound(bbg_field, 1) Then
                                        'add bbg field
                                        ReDim Preserve bbg_field(UBound(bbg_field, 1) + 1)
                                        bbg_field(UBound(bbg_field, 1)) = UCase(Mid(tmp_vec_fields(j), 2))
                                    End If
                                End If
                                
                            Next m
                            
                        End If
                        
                    Next j
                End If
                
            Next i
            
            
        End If
        
        
    Dim dim_bbg_yest_close As Integer, dim_bbg_yest_low As Integer, dim_bbg_yest_high As Integer, _
        dim_bbg_dmi_adx As Integer, dim_bbg_dmi_dim As Integer, dim_bbg_dim_dip As Integer, dim_bbg_interval_boll As Integer, _
        dim_bbg_exchange_status As Integer
    
    For i = 0 To UBound(bbg_field, 1)
        If UCase(bbg_field(i)) = UCase("PX_YEST_CLOSE") Then
            dim_bbg_yest_close = i
        ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_LOW") Then
            dim_bbg_yest_low = i
        ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_HIGH") Then
            dim_bbg_yest_high = i
        ElseIf UCase(bbg_field(i)) = UCase("LAST_PRICE") Then
            dim_bbg_last_price = i
        ElseIf UCase(bbg_field(i)) = UCase("DMI_ADX") Then
            dim_bbg_dmi_adx = i
        ElseIf UCase(bbg_field(i)) = UCase("DMI_DIM") Then
            dim_bbg_dmi_dim = i
        ElseIf UCase(bbg_field(i)) = UCase("DMI_DIP") Then
            dim_bbg_dim_dip = i
        ElseIf UCase(bbg_field(i)) = UCase("INTERVAL_BOLL_PERCENT_B") Then
            dim_bbg_interval_boll = i
        ElseIf UCase(bbg_field(i)) = UCase("RT_SIMP_SEC_STATUS") Then
            dim_bbg_exchange_status = i
        End If
    Next i
    output_bbg = oBBG.bdp(vec_ticker, bbg_field, output_format.of_vec_without_header)
    
    
    
    'preparation des trades
    Dim p As Double, s1 As Double, s2 As Double, s3 As Double, r1 As Double, r2 As Double, r3 As Double
    
    k = 0
        
    Dim qty_to_trade As Double
    
    Dim vec_change_rate() As Variant
    
    j = 0
    For i = 14 To 31
        If Worksheets("Parametres").Cells(i, 1) = "" Then
            Exit For
        Else
            ReDim Preserve vec_change_rate(j)
            vec_change_rate(j) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, CDbl(Worksheets("Parametres").Cells(i, 6).Value))
            j = j + 1
        End If
    Next i
        
    
    
    vec_trades = generate_vec_trades(vec_ticker_details, output_bbg, bbg_field, vec_change_rate, vec_simple_trades)
insert_vec_trade_in_format2:
    If IsEmpty(vec_trades) Then
        GoTo no_trade_to_create_based_on_filters_selected
    End If

    Dim l_last_trade As Integer
    
    If UBound(vec_trades, 1) > 0 Then
        
        l_last_trade = insert_vec_trades_into_format2(vec_trades, vec_currency, region)
        
        'mise en place gestion des stop
        Call algo_helper_signal_eqs(l_format2_header, vec_ticker, bbg_field, output_bbg, extract_central)
        
        Call color_format2_trades
        
        Call algo_helper_filter_order_format2(l_format2_header, vec_ticker, bbg_field, output_bbg)
        
        Sheets("FORMAT2").Activate
        Worksheets("FORMAT2").Cells(l_format2_header, 1).Activate
    
    End If
    
Else
no_trade_to_create_based_on_filters_selected:
    MsgBox ("No ticker based on criterias")
End If

datetime_end = Now()
Debug.Print Application.RoundDown(1440 * (datetime_end - datetime_start), 0.1) & " minute(s) and " & CInt(60 * (1440 * (datetime_end - datetime_start) - Application.RoundDown(1440 * (datetime_end - datetime_start), 0.1))) & " seconds"

Application.Calculation = xlCalculationAutomatic

End Sub



Private Function insert_vec_trades_into_format2(ByVal vec_trades As Variant, ByVal vec_currency As Variant, ByVal region As Variant) As Variant

c_equity_db_perso_rel_1d = 138

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Double, q As Integer

Application.Calculation = xlCalculationManual

If IsEmpty(vec_trades) Then
    Exit Function
End If


    k = l_format2_header + 1

    If UBound(vec_trades, 1) > 0 Then
        
        'clean area
        For i = l_format2_header To 32000
            If Worksheets("FORMAT2").Cells(i, 1) = "" Then
                Exit For
            Else
                Worksheets("FORMAT2").rows(i).Clear
            End If
        Next i
        
        Application.ReferenceStyle = xlA1
    
        Dim extract_ibd As Variant
        extract_ibd = mount_sqlite_central()
        
        'header
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_ticker) = "ticker"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_aim_account) = "aim_account"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_strategy) = "aim_strategy_tag"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_limit_type) = "limit_type"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_side) = "side"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_qty) = "qty"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_price) = "limit_price"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_time_limit) = "limit_time"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_broker) = "broker"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_last_price) = "last" 'from equity database
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_s3) = "s3"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_s2) = "s2"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_s1) = "s1"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_p) = "p"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_r1) = "r1"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_r2) = "r2"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_r3) = "r3"
        
        For i = 0 To 1
            Worksheets("FORMAT2").Cells(l_format2_header, c_format2_pre_market_start_column + i) = "prmrkt"
        Next i
        
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_valeur_eur) = "val eur"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_delta) = "delta"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_theta) = "theta"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_eps) = "eps"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_perso_rel_1d) = "perso_rel_1d"
        
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_dmi) = "dmi"
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_pct_bol) = "%bol"
        
        Worksheets("FORMAT2").Cells(l_format2_header, c_format2_source) = "src"
        
        
        Dim l_first_trade As Integer, l_last_trade As Integer
        l_first_trade = l_format2_header + 1
        
        Dim total_index_equiv_dollar As Double
        total_index_equiv_dollar = 0
        For i = 0 To UBound(vec_trades, 1)
            
            For j = 0 To UBound(vec_currency, 1)
                If vec_currency(j)(0) = vec_trades(i)(dim_vec_trade_crncy) Then
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        total_index_equiv_dollar = total_index_equiv_dollar + vec_currency(j)(2) * Abs(vec_trades(i)(dim_vec_trade_qty)) * Round(vec_trades(i)(dim_vec_trade_price), 2)
                        Exit For
                    Else
                        total_index_equiv_dollar = total_index_equiv_dollar + vec_currency(j)(2) * Abs(vec_trades(i)(dim_vec_trade_qty)) * Round(vec_trades(i)(dim_vec_trade_last_price), 2)
                    End If
                End If
            Next j
            
            
            Worksheets("FORMAT2").Cells(k, c_format2_ticker) = vec_trades(i)(dim_vec_trade_ticker)
            
            If Worksheets("Parametres").Cells(17, 18) = "" Then
                Worksheets("FORMAT2").Cells(k, c_format2_aim_account) = "C6414GSJ"
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_aim_account) = Worksheets("Parametres").Cells(17, 18).Value
            End If
            
            
            If Worksheets("FORMAT2").CB_strategy.Value = "" Then
                Worksheets("FORMAT2").Cells(k, c_format2_strategy) = "TRADING"
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_strategy) = Worksheets("FORMAT2").CB_strategy.Value
            End If
            
            Worksheets("FORMAT2").Cells(k, c_format2_limit_type) = UCase("LMT")
            
            'repere si existe note
            If IsEmpty(extract_ibd) = False Then
                For j = 0 To UBound(extract_ibd(0), 1)
                    If UCase(extract_ibd(0)(j)) = UCase("TICKER") Then
                        dim_central_ticker = j
                    ElseIf UCase(extract_ibd(0)(j)) = UCase("RANK_EPS") Then
                        dim_central_eps = j
                    End If
                Next j
                
                Dim tmp_ticker_for_central As String
                tmp_ticker_for_central = patch_ticker_marketplace(UCase(vec_trades(i)(dim_vec_trade_ticker)))
                For j = 0 To UBound(extract_ibd, 1)
                    If extract_ibd(j)(dim_central_ticker) = tmp_ticker_for_central Then
                        Worksheets("FORMAT2").Cells(k, c_format2_eps) = extract_ibd(j)(dim_central_eps)
                        Exit For
                    End If
                Next j
            End If
            
            If vec_trades(i)(1) < 0 Then
                
                If vec_trades(i)(dim_vec_trade_position_in_equity_db) - Abs(vec_trades(i)(1)) < 0 Then
                    Worksheets("FORMAT2").Cells(k, c_format2_side) = UCase("H")
                Else
                    Worksheets("FORMAT2").Cells(k, c_format2_side) = UCase("S")
                End If
                
                For j = c_format2_ticker To c_format2_last_price
                    Worksheets("FORMAT2").Cells(k, j).Font.ColorIndex = 3
                Next j
                
                For j = c_format2_valeur_eur To c_format2_theta
                    Worksheets("FORMAT2").Cells(k, j).Font.ColorIndex = 3
                Next j
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_side) = UCase("B")
                
                For j = c_format2_ticker To c_format2_last_price
                    Worksheets("FORMAT2").Cells(k, j).Font.ColorIndex = 10
                Next j
                
                For j = c_format2_valeur_eur To c_format2_theta
                    Worksheets("FORMAT2").Cells(k, j).Font.ColorIndex = 10
                Next j
            End If
            
            
            Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).Interior.ColorIndex = 37
            If vec_trades(i)(dim_vec_trade_line) <> -1 Then
                If Worksheets("Equity_Database").Cells(vec_trades(i)(dim_vec_trade_line), xlColumnValue("V")) <> 0 Then
                    Worksheets("FORMAT2").Cells(k, c_format2_last_price).FormulaLocal = "=Equity_Database!V" & vec_trades(i)(dim_vec_trade_line) 'spot
                    Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).FormulaLocal = "=Equity_Database!" & xlColumnValue(c_equity_db_perso_rel_1d) & vec_trades(i)(dim_vec_trade_line) 'perso rel 1d
                    
                    'cond format
                    If IsNumeric(Worksheets("Equity_Database").Cells(vec_trades(i)(dim_vec_trade_line), c_equity_db_perso_rel_1d)) Then
                        If Worksheets("Equity_Database").Cells(vec_trades(i)(dim_vec_trade_line), c_equity_db_perso_rel_1d) < -5 Then
                            Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).Interior.ColorIndex = 3
                        ElseIf Worksheets("Equity_Database").Cells(vec_trades(i)(dim_vec_trade_line), c_equity_db_perso_rel_1d) > 5 Then
                            Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).Interior.ColorIndex = 4
                        Else
                            Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).Interior.ColorIndex = 37
                        End If
                    End If
                    
                Else
                    Worksheets("FORMAT2").Cells(k, c_format2_last_price).FormulaLocal = "=BDP(" & xlColumnValue(c_format2_ticker) & k & ";""LAST_PRICE"")"
                    Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).FormulaLocal = "=BDP(" & xlColumnValue(c_format2_ticker) & k & ";""REL_1D"")"
                End If
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_last_price).FormulaLocal = "=BDP(" & xlColumnValue(c_format2_ticker) & k & ";""LAST_PRICE"")"
                Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).FormulaLocal = "=BDP(" & xlColumnValue(c_format2_ticker) & k & ";""REL_1D"")"
            End If
                Worksheets("FORMAT2").Cells(k, c_format2_last_price).Interior.ColorIndex = 35
                Worksheets("FORMAT2").Cells(k, c_format2_last_price).NumberFormat = "#,##0.00"
                
                Worksheets("FORMAT2").Cells(k, c_format2_perso_rel_1d).NumberFormat = "#,##0.00"
            
            
            'pre market -> US / teo_price -> EU
            For q = 0 To UBound(region, 1)
                If vec_trades(i)(dim_vec_trade_region) = region(q)(0) Then
                    
                    'passe en revue les champs
                    For s = 0 To UBound(region(q)(2), 1)
                        Worksheets("FORMAT2").Cells(k, c_format2_pre_market_start_column + s).FormulaLocal = "=BDP(" & xlColumnValue(c_format2_ticker) & k & ";""" & region(q)(2)(s) & """)"
                            Worksheets("FORMAT2").Cells(k, c_format2_pre_market_start_column + s).NumberFormat = "#,##0.00"
                    Next s
                    
                    With Worksheets("FORMAT2").Cells(k, c_format2_price)
                        
                        .FormatConditions.Delete
                        
                        If UBound(region(q)(2), 1) = 0 Then 'un seul champ, par ex europe
                            
                            If vec_trades(i)(dim_vec_trade_qty) < 0 Then 'sell
                                .FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column) & "$" & k
                            Else 'buy
                                .FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column) & "$" & k
                            End If
                        ElseIf UBound(region(q)(2), 1) = 1 Then '2 champs, bid & ask
                            
                            If vec_trades(i)(dim_vec_trade_qty) < 0 Then 'sell
                                .FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column) & "$" & k
                            Else 'buy
                                .FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & xlColumnValue(c_format2_pre_market_start_column + 1) & "$" & k
                            End If
                        End If
                        
                        .FormatConditions(1).Interior.ColorIndex = 16
                        
                    End With
                    
                    Exit For
                End If
            Next q
            
            If vec_trades(i)(dim_vec_trade_line) <> -1 Then
                Worksheets("FORMAT2").Cells(k, c_format2_valeur_eur).FormulaLocal = "=Equity_Database!E" & vec_trades(i)(dim_vec_trade_line) 'valeur eur
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_valeur_eur) = 0
            End If
                Worksheets("FORMAT2").Cells(k, c_format2_valeur_eur).NumberFormat = "#,##0_ ;-#,##0"
            
            If vec_trades(i)(dim_vec_trade_line) <> -1 Then
                Worksheets("FORMAT2").Cells(k, c_format2_delta).FormulaLocal = "=Equity_Database!F" & vec_trades(i)(dim_vec_trade_line) 'delta
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_delta) = 0
            End If
                Worksheets("FORMAT2").Cells(k, c_format2_delta).NumberFormat = "#,##0_ ;-#,##0"
            
            If vec_trades(i)(dim_vec_trade_line) <> -1 Then
                Worksheets("FORMAT2").Cells(k, c_format2_theta).FormulaLocal = "=Equity_Database!I" & vec_trades(i)(dim_vec_trade_line) 'theta
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_theta) = 0
            End If
                Worksheets("FORMAT2").Cells(k, c_format2_theta).NumberFormat = "#,##0_ ;-#,##0"
            
            Worksheets("FORMAT2").Cells(k, c_format2_qty) = Abs(vec_trades(i)(dim_vec_trade_qty))
                Worksheets("FORMAT2").Cells(k, c_format2_qty).NumberFormat = "0"
            
            If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                Worksheets("FORMAT2").Cells(k, c_format2_price) = Round(vec_trades(i)(dim_vec_trade_price), 2)
            Else
                If Left(vec_trades(i)(dim_vec_trade_price), 7) = "get_out" Then
                    
                    get_out_style = Mid(vec_trades(i)(dim_vec_trade_price), Len("get_out_") + 1)
                    
                    If get_out_style = "pct of last price" Then
                        
                        If Worksheets("FORMAT2").Cells(l_format2_get_out_pct_last_price, c_format2_get_out_pct_last_price) <> "" Then
                            If IsNumeric(Worksheets("FORMAT2").Cells(l_format2_get_out_pct_last_price, c_format2_get_out_pct_last_price)) = True Then
                                factor_corrector = CDbl(Worksheets("FORMAT2").Cells(l_format2_get_out_pct_last_price, c_format2_get_out_pct_last_price))
                            Else
                                factor_corrector = 0
                                Worksheets("FORMAT2").Cells(l_format2_get_out_pct_last_price, c_format2_get_out_pct_last_price) = 0
                            End If
                        Else
                            factor_corrector = 0
                            Worksheets("FORMAT2").Cells(l_format2_get_out_pct_last_price, c_format2_get_out_pct_last_price) = 0
                        End If
                        
                        If Worksheets("FORMAT2").Cells(k, c_format2_side) = "B" Or Worksheets("FORMAT2").Cells(k, c_format2_side) = "C" Then
                            Worksheets("FORMAT2").Cells(k, c_format2_price).FormulaLocal = "=ROUND((1+" & xlColumnValue(c_format2_get_out_pct_last_price) & l_format2_get_out_pct_last_price & ")*" & xlColumnValue(c_format2_last_price) & k & ";2)"
                        ElseIf Worksheets("FORMAT2").Cells(k, c_format2_side) = "S" Or Worksheets("FORMAT2").Cells(k, c_format2_side) = "H" Then
                            Worksheets("FORMAT2").Cells(k, c_format2_price).FormulaLocal = "=ROUND((1-" & xlColumnValue(c_format2_get_out_pct_last_price) & l_format2_get_out_pct_last_price & ")*" & xlColumnValue(c_format2_last_price) & k & ";2)"
                        End If
                        
                    ElseIf get_out_style = "near piv" Then
                        
                        If Worksheets("FORMAT2").Cells(k, c_format2_side) = "B" Or Worksheets("FORMAT2").Cells(k, c_format2_side) = "C" Then
                            'repere le support le plus proche du last
                            For j = dim_vec_trade_p To dim_vec_trade_s3 Step -1
                                If vec_trades(i)(j) < vec_trades(i)(dim_vec_trade_last_price) Then
                                    Worksheets("FORMAT2").Cells(k, c_format2_price) = vec_trades(i)(j)
                                    Exit For
                                Else
                                    If j = dim_vec_trade_s3 Then
                                        Worksheets("FORMAT2").Cells(k, c_format2_price) = 0.9999 * vec_trades(i)(dim_vec_trade_last_price)
                                    End If
                                End If
                            Next j
                            
                        ElseIf Worksheets("FORMAT2").Cells(k, c_format2_side) = "S" Or Worksheets("FORMAT2").Cells(k, c_format2_side) = "H" Then
                            'repere la resistance la plus proche du last
                            For j = dim_vec_trade_r1 To dim_vec_trade_r3
                                If vec_trades(i)(j) > vec_trades(i)(dim_vec_trade_last_price) Then
                                    Worksheets("FORMAT2").Cells(k, c_format2_price) = vec_trades(i)(j)
                                    Exit For
                                Else
                                    If j = dim_vec_trade_r3 Then
                                        'place sur le last
                                        Worksheets("FORMAT2").Cells(k, c_format2_price) = 1.0001 * vec_trades(i)(dim_vec_trade_last_price)
                                    End If
                                End If
                            Next j
                        End If
                        
                    End If
                End If
            End If
                Worksheets("FORMAT2").Cells(k, c_format2_price).NumberFormat = "#,##0.00"
            
            
            If UCase(vec_trades(i)(dim_vec_trade_order_type)) = UCase("base") Then
                Worksheets("FORMAT2").Cells(k, c_format2_time_limit) = UCase("DAY")
            ElseIf UCase(vec_trades(i)(dim_vec_trade_order_type)) = "STP" Or UCase(vec_trades(i)(dim_vec_trade_order_type)) = "STOP" Then
                Worksheets("FORMAT2").Cells(k, c_format2_time_limit) = UCase("STOP")
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_time_limit) = UCase("DAY")
            End If
            
            If IsEmpty(vec_trades(i)(dim_vec_trade_broker)) Then
                If Worksheets("FORMAT2").CB_exec_broker.Value <> "" Then
                    Worksheets("FORMAT2").Cells(k, c_format2_broker) = Worksheets("FORMAT2").CB_exec_broker.Value
                Else
                    Worksheets("FORMAT2").Cells(k, c_format2_broker) = "GOLDMAN"
                End If
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_broker) = vec_trades(i)(dim_vec_trade_broker)
            End If
            
            
            If UCase(vec_trades(i)(dim_vec_trade_order_type)) = UCase("base") Then
            
                Worksheets("FORMAT2").Cells(k, c_format2_s3) = vec_trades(i)(dim_vec_trade_s3)
                    Worksheets("FORMAT2").Cells(k, c_format2_s3).NumberFormat = "#,##0.00"
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        If Round(vec_trades(i)(dim_vec_trade_s3), 2) = Round(vec_trades(i)(dim_vec_trade_price), 2) Then
                            If UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "C" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_s3).Interior.ColorIndex = 4
                            ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "H" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_s3).Interior.ColorIndex = 3
                            End If
                        End If
                    End If
                    
                Worksheets("FORMAT2").Cells(k, c_format2_s2) = vec_trades(i)(dim_vec_trade_s2)
                    Worksheets("FORMAT2").Cells(k, c_format2_s2).NumberFormat = "#,##0.00"
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        If Round(vec_trades(i)(dim_vec_trade_s2), 2) = Round(vec_trades(i)(dim_vec_trade_price), 2) Then
                            If UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "C" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_s2).Interior.ColorIndex = 4
                            ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "H" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_s2).Interior.ColorIndex = 3
                            End If
                        End If
                    End If
                    
                Worksheets("FORMAT2").Cells(k, c_format2_s1) = vec_trades(i)(dim_vec_trade_s1)
                    Worksheets("FORMAT2").Cells(k, c_format2_s1).NumberFormat = "#,##0.00"
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        If Round(vec_trades(i)(dim_vec_trade_s1), 2) = Round(vec_trades(i)(dim_vec_trade_price), 2) Then
                            If UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "C" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_s1).Interior.ColorIndex = 4
                            ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "H" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_s1).Interior.ColorIndex = 3
                            End If
                            
                        End If
                    End If
                    
                Worksheets("FORMAT2").Cells(k, c_format2_p) = vec_trades(i)(dim_vec_trade_p)
                    Worksheets("FORMAT2").Cells(k, c_format2_p).NumberFormat = "#,##0.00"
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        If Round(vec_trades(i)(dim_vec_trade_p), 2) = Round(vec_trades(i)(dim_vec_trade_price), 2) Then
                            If UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "C" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_p).Interior.ColorIndex = 4
                            ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "H" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_p).Interior.ColorIndex = 3
                            End If
                        End If
                    End If
                    
                Worksheets("FORMAT2").Cells(k, c_format2_r1) = vec_trades(i)(dim_vec_trade_r1)
                    Worksheets("FORMAT2").Cells(k, c_format2_r1).NumberFormat = "#,##0.00"
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        If Round(vec_trades(i)(dim_vec_trade_r1), 2) = Round(vec_trades(i)(dim_vec_trade_price), 2) Then
                            If UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "C" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_r1).Interior.ColorIndex = 4
                            ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "H" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_r1).Interior.ColorIndex = 3
                            End If
                        End If
                    End If
                
                Worksheets("FORMAT2").Cells(k, c_format2_r2) = vec_trades(i)(dim_vec_trade_r2)
                    Worksheets("FORMAT2").Cells(k, c_format2_r2).NumberFormat = "#,##0.00"
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        If Round(vec_trades(i)(dim_vec_trade_r2), 2) = Round(vec_trades(i)(dim_vec_trade_price), 2) Then
                            If UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "C" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_r2).Interior.ColorIndex = 4
                            ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "H" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_r2).Interior.ColorIndex = 3
                            End If
                        End If
                    End If
                    
                Worksheets("FORMAT2").Cells(k, c_format2_r3) = vec_trades(i)(dim_vec_trade_r3)
                    Worksheets("FORMAT2").Cells(k, c_format2_r3).NumberFormat = "#,##0.00"
                    If IsNumeric(vec_trades(i)(dim_vec_trade_price)) Then
                        If Round(vec_trades(i)(dim_vec_trade_r3), 2) = Round(vec_trades(i)(dim_vec_trade_price), 2) Then
                            If UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "C" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_r3).Interior.ColorIndex = 4
                            ElseIf UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(k, c_format2_side), 1)) = "H" Then
                                Worksheets("FORMAT2").Cells(k, c_format2_r3).Interior.ColorIndex = 3
                            End If
                        End If
                    End If
            Else
                
                Worksheets("FORMAT2").Cells(k, c_format2_s3) = vec_trades(i)(dim_vec_trade_order_type)
                
            End If
                
                
            Worksheets("FORMAT2").Cells(k, c_format2_dmi).FormulaLocal = "=BDP(" & xlColumnValue(c_format2_ticker) & k & ";""DMI_ADX"")"
                Worksheets("FORMAT2").Cells(k, c_format2_dmi).NumberFormat = "#,##0"
            
            Worksheets("FORMAT2").Cells(k, c_format2_pct_bol).FormulaLocal = "=BDP(" & xlColumnValue(c_format2_ticker) & k & ";""INTERVAL_BOLL_PERCENT_B"")"
                Worksheets("FORMAT2").Cells(k, c_format2_pct_bol).NumberFormat = "0%"
            
            
            If UCase(vec_trades(i)(dim_vec_trade_order_type)) = UCase("base") Then
                If count_source > 1 Then
                    
                    If IsArray(vec_trades(i)(dim_vec_trade_source)) Then
                        For j = 0 To UBound(vec_trades(i)(dim_vec_trade_source), 1)
                            Worksheets("FORMAT2").Cells(k, c_format2_source + j) = vec_trades(i)(dim_vec_trade_source)(j)
                        Next j
                    Else
                        Worksheets("FORMAT2").Cells(k, c_format2_source) = vec_trades(i)(dim_vec_trade_source)
                    End If
                Else
                    If IsArray(vec_trades(i)(dim_vec_trade_source)) Then
                        For j = 0 To UBound(vec_trades(i)(dim_vec_trade_source), 1)
                            Worksheets("FORMAT2").Cells(k, c_format2_source + j) = vec_trades(i)(dim_vec_trade_source)(j)
                        Next j
                    Else
                        Worksheets("FORMAT2").Cells(k, c_format2_source) = vec_trades(i)(dim_vec_trade_source)
                    End If
                End If
            Else
                Worksheets("FORMAT2").Cells(k, c_format2_source) = "*** AUTO-GENERATED ***"
            End If
            
            l_last_trade = k
            
            k = k + 1
        Next i
        
        
        'ajuste si panier de brokers pour les trades
        If Worksheets("FORMAT2").CB_exec_broker.Value <> "" Then
            For i = 7 To 100
                If Worksheets("FORMAT2").CB_exec_broker.Value = Worksheets("FORMAT2").Cells(i, 27).Value Then
                    
                    'creation des pct par broker
                    Dim vec_basket_broker() As Variant
                    Dim total_weight As Double
                    
                    k = 0
                    total_weight = 0
                    For j = 28 To 70
                        If Worksheets("FORMAT2").Cells(6, j) = "" Then
                            Exit For
                        Else
                            If Worksheets("FORMAT2").Cells(i, j) <> "" And IsNumeric(Worksheets("FORMAT2").Cells(i, j)) Then
                                If Worksheets("FORMAT2").Cells(i, j) > 0 Then
                                    ReDim Preserve vec_basket_broker(k)
                                    vec_basket_broker(k) = Array(Worksheets("FORMAT2").Cells(6, j).Value, Worksheets("FORMAT2").Cells(i, j).Value, 0, 0)
                                    
                                    total_weight = total_weight + Worksheets("FORMAT2").Cells(i, j).Value
                                    
                                    k = k + 1
                                End If
                            End If
                        End If
                    Next j
                    
                    'conversion en pct
                    If k > 0 Then
                        For j = 0 To UBound(vec_basket_broker, 1)
                            vec_basket_broker(j)(1) = vec_basket_broker(j)(1) / total_weight
                            vec_basket_broker(j)(2) = 1.1 * (vec_basket_broker(j)(1) * total_index_equiv_dollar)
                        Next j
                    End If
                    
                    
                    
                    
                    For j = 0 To UBound(vec_trades, 1)
                        For m = 0 To UBound(vec_currency, 1)
                            If vec_currency(m)(0) = vec_trades(j)(dim_vec_trade_crncy) Then
                                
new_tirage_au_sort_broker:
                                tirage_au_sort_broker = CInt(Rnd() * UBound(vec_basket_broker, 1))
                                
                                If vec_basket_broker(tirage_au_sort_broker)(3) > vec_basket_broker(tirage_au_sort_broker)(2) Then
                                    
                                    vec_basket_broker(tirage_au_sort_broker) = vec_basket_broker(UBound(vec_basket_broker, 1))
                                    ReDim Preserve vec_basket_broker(UBound(vec_basket_broker, 1) - 1)
                                    
                                    GoTo new_tirage_au_sort_broker
                                Else
                                    
                                    Worksheets("FORMAT2").Cells(l_first_trade + j, c_format2_broker) = vec_basket_broker(tirage_au_sort_broker)(0)
                                    
                                    If IsNumeric(vec_trades(j)(dim_vec_trade_price)) = False Then 'gestion get out
                                        vec_basket_broker(tirage_au_sort_broker)(3) = vec_basket_broker(tirage_au_sort_broker)(3) + vec_currency(m)(2) * Abs(vec_trades(j)(dim_vec_trade_qty)) * Round(vec_trades(j)(dim_vec_trade_last_price), 2)
                                    Else
                                        vec_basket_broker(tirage_au_sort_broker)(3) = vec_basket_broker(tirage_au_sort_broker)(3) + vec_currency(m)(2) * Abs(vec_trades(j)(dim_vec_trade_qty)) * Round(vec_trades(j)(dim_vec_trade_price), 2)
                                    End If
                                    
                                End If
                                
                            End If
                        Next m
                    Next j
                    
                    Exit For
                End If
            Next i
        End If
    
    End If


insert_vec_trades_into_format2 = l_last_trade

End Function



Private Function transform_simple_trades_into_format2_vec_trades(ByVal vec_simple_trades As Variant, ByVal vec_ticker_details As Variant, ByVal output_bbg As Variant, ByVal bbg_field As Variant, ByVal vec_change_rate As Variant)

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


If IsEmpty(vec_simple_trades) Then
Else
    
    'passe en revue les trades en s assurant
    For i = 0 To UBound(vec_simple_trades, 1)
        
    Next i
    
End If

End Function


Private Function generate_vec_trades(ByVal vec_ticker_details As Variant, ByVal output_bbg As Variant, ByVal bbg_field As Variant, ByVal vec_change_rate As Variant, Optional ByVal vec_simple_trades As Variant) As Variant

generate_vec_trades = Empty

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, q As Integer, u As Integer, v As Integer

Dim min_value As Variant
Dim min_pos As Integer
Dim tmp_var() As Variant



'preparation des trades
Dim p As Double, s1 As Double, s2 As Double, s3 As Double, r1 As Double, r2 As Double, r3 As Double

k = 0
Dim vec_trades() As Variant
        
Dim qty_to_trade As Double

            
Dim dim_bbg_yest_close As Integer, dim_bbg_yest_low As Integer, dim_bbg_yest_high As Integer, _
    dim_bbg_dmi_adx As Integer, dim_bbg_dmi_dim As Integer, dim_bbg_dim_dip As Integer, dim_bbg_interval_boll As Integer, _
    dim_bbg_exchange_status As Integer


For i = 0 To UBound(bbg_field, 1)
    If UCase(bbg_field(i)) = UCase("PX_YEST_CLOSE") Then
        dim_bbg_yest_close = i
    ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_LOW") Then
        dim_bbg_yest_low = i
    ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_HIGH") Then
        dim_bbg_yest_high = i
    ElseIf UCase(bbg_field(i)) = UCase("LAST_PRICE") Then
        dim_bbg_last_price = i
    ElseIf UCase(bbg_field(i)) = UCase("DMI_ADX") Then
        dim_bbg_dmi_adx = i
    ElseIf UCase(bbg_field(i)) = UCase("DMI_DIM") Then
        dim_bbg_dmi_dim = i
    ElseIf UCase(bbg_field(i)) = UCase("DMI_DIP") Then
        dim_bbg_dim_dip = i
    ElseIf UCase(bbg_field(i)) = UCase("INTERVAL_BOLL_PERCENT_B") Then
        dim_bbg_interval_boll = i
    ElseIf UCase(bbg_field(i)) = UCase("RT_SIMP_SEC_STATUS") Then
        dim_bbg_exchange_status = i
    End If
Next i
    
Dim exchange_tradeable_security_status() As Variant
    exchange_tradeable_security_status = Array("OUT", "TALC", "TRAD", "TRAU")
        

Dim vec_trade_based_on_simple()
    Dim count_vec_trade_based_on_simple As Integer
    count_vec_trade_based_on_simple = 0

Dim need_resort As Boolean
    need_resort = False
    

If IsEmpty(vec_simple_trades) = False Then
    
    'passe en revue les trades et check que bien dispo dans vec_ticker_details
    For q = 0 To UBound(vec_simple_trades, 1)
        For i = 0 To UBound(vec_ticker_details, 1)
            If vec_simple_trades(q)(dim_vec_simple_trade_ticker) = vec_ticker_details(i)(dim_vec_ticker_details_ticker) Then
                
                'le ticker a bien passe tous les filtres
                
                If IsNumeric(output_bbg(i)(dim_bbg_yest_close)) And IsNumeric(output_bbg(i)(dim_bbg_yest_low)) And IsNumeric(output_bbg(i)(dim_bbg_yest_high)) Then
            
                    If Worksheets("FORMAT2").CB_check_tradeable_hour.Value = True Then
                        
                        For j = 0 To UBound(exchange_tradeable_security_status, 1)
                            If exchange_tradeable_security_status(j) = output_bbg(i)(dim_bbg_exchange_status) Then
                                Exit For
                            Else
                                If j = UBound(exchange_tradeable_security_status, 1) Then
                                    GoTo check_next_entry_vec_ticker_details_for_construct_trades_simple_trades
                                End If
                            End If
                        Next j
                        
                    End If
                    
                    
                    p = Round((output_bbg(i)(dim_bbg_yest_close) + output_bbg(i)(dim_bbg_yest_low) + output_bbg(i)(dim_bbg_yest_high)) / 3, 3)
                    
                    r1 = Round(2 * p - output_bbg(i)(dim_bbg_yest_low), 3)
                    s1 = Round(2 * p - output_bbg(i)(dim_bbg_yest_high), 3)
                    
                    r2 = Round((p - s1) + r1, 3)
                    s2 = Round(p - (r1 - s1), 3)
                    
                    r3 = Round((p - s2) + r2, 3)
                    s3 = Round(p - (r2 - s2), 3)
                    
                    last_price = output_bbg(i)(dim_bbg_last_price)
                    
                            
                    For j = 0 To UBound(vec_change_rate, 1)
                        If vec_change_rate(j)(dim_vec_change_rate_txt) = vec_ticker_details(i)(dim_vec_ticker_details_crncy) Then
                            
                            
                            If IsEmpty(vec_simple_trades(q)(dim_vec_simple_trade_qty)) Or Left(UCase(vec_simple_trades(q)(dim_vec_simple_trade_qty)), 1) = "B" Or Left(UCase(vec_simple_trades(q)(dim_vec_simple_trade_qty)), 1) = "S" Then
                                
                                If Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) = True And Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value) = True And Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value) = True Then
                                    
                                    'check l'interval de theta
                                    If Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) >= 0 And Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) < Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value)) Then
                                        qty_to_trade = Round(Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                                    ElseIf Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) >= Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value)) And Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) < Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value)) Then
                                        qty_to_trade = 2 * Round(Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                                    ElseIf Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) >= Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value)) Then
                                        qty_to_trade = 3 * Round(Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                                    End If
                                    
                                Else
                                    If Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value) Then
                                        If Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value > 0 Then
                                            qty_to_trade = Round(Abs(CDbl(Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                                        Else
                                            If IsError(vec_ticker_details(i)(dim_vec_ticker_details_delta)) = False Then
                                                qty_to_trade = Round(Abs(vec_ticker_details(i)(dim_vec_ticker_details_delta)) / 3, 0)
                                            Else
                                                qty_to_trade = 0
                                            End If
                                        End If
                                    Else
                                        If IsError(vec_ticker_details(i)(dim_vec_ticker_details_delta)) = False Then
                                            qty_to_trade = Round(Abs(vec_ticker_details(i)(dim_vec_ticker_details_delta)) / 3, 0)
                                        Else
                                            qty_to_trade = 0
                                        End If
                                    End If
                                End If
                                
                                
                                If IsNumeric(vec_simple_trades(q)(dim_vec_simple_trade_qty)) = False Then
                                    
                                    If Left(UCase(vec_simple_trades(q)(dim_vec_simple_trade_qty)), 1) = "B" Then
                                        qty_to_trade = Abs(qty_to_trade)
                                    ElseIf Left(UCase(vec_simple_trades(q)(dim_vec_simple_trade_qty)), 1) = "S" Then
                                        qty_to_trade = -Abs(qty_to_trade)
                                    End If
                                    
                                    vec_simple_trades(q)(dim_vec_simple_trade_qty) = qty_to_trade
                                    
                                End If
                                
                            Else
                                qty_to_trade = vec_simple_trades(q)(dim_vec_simple_trade_qty)
                            End If
                            
                            
                            'ajustement des prix
                            If qty_to_trade < 0 Then
                                
                                If IsEmpty(vec_simple_trades(q)(dim_vec_simple_trade_price)) Then
choose_sell_price:
                                    If r1 > last_price Then
                                        vec_simple_trades(q)(dim_vec_simple_trade_price) = r1
                                    Else
                                        
                                        If r2 > last_price Then
                                            vec_simple_trades(q)(dim_vec_simple_trade_price) = r2
                                        Else
                                            If r3 > last_price Then
                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = r3
                                            Else
                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = last_price * 1.005
                                            End If
                                        End If
                                        
                                    End If
                                    
                                Else
                                    
                                    
                                    
                                    'check que marge de manoeuvre
                                    If UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "BASE" Then
                                        If vec_simple_trades(q)(dim_vec_simple_trade_price) < last_price Then
                                            GoTo choose_sell_price
                                        End If
                                    ElseIf UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "STP" Or UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "STOP" Then
                                        
                                        'repere order base et s assure que prix en dessous
                                        For u = q - 1 To 0 Step -1
                                            If vec_simple_trades(q)(dim_vec_simple_trade_ticker) = vec_simple_trades(u)(dim_vec_simple_trade_ticker) And UCase(vec_simple_trades(u)(dim_vec_simple_trade_order_type)) = "BASE" Then
                                                
                                                'sell (stop) d un original order buy
                                                If vec_simple_trades(q)(dim_vec_simple_trade_price) >= vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                    'le prix doit etre remplace
                                                    
                                                    If s1 < last_price And s1 < vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                        vec_simple_trades(q)(dim_vec_simple_trade_price) = s1
                                                        Exit For
                                                    Else
                                                        If s2 < last_price And s2 < vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                            vec_simple_trades(q)(dim_vec_simple_trade_price) = s2
                                                            Exit For
                                                        Else
                                                            If s3 < last_price And s3 < vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = s3
                                                                Exit For
                                                            Else
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = 0.97 * vec_simple_trades(u)(dim_vec_simple_trade_price)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                End If
                                                
                                                Exit For
                                            End If
                                        Next u
                                        
                                    ElseIf UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "TGT" Or UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "TARGET" Then
                                        
                                        'repere order base et s assure que prix en dessous
                                        For u = q - 1 To 0 Step -1
                                            If vec_simple_trades(q)(dim_vec_simple_trade_ticker) = vec_simple_trades(u)(dim_vec_simple_trade_ticker) And UCase(vec_simple_trades(u)(dim_vec_simple_trade_order_type)) = "BASE" Then
                                                
                                                'sell (target) d un original order buy
                                                If vec_simple_trades(q)(dim_vec_simple_trade_price) <= vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                    'le prix doit etre remplace
                                                    
                                                    If r1 > last_price And r1 > vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                        vec_simple_trades(q)(dim_vec_simple_trade_price) = r1
                                                        Exit For
                                                    Else
                                                        If r2 > last_price And r2 > vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                            vec_simple_trades(q)(dim_vec_simple_trade_price) = r2
                                                            Exit For
                                                        Else
                                                            If r3 > last_price And r3 > vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = r3
                                                                Exit For
                                                            Else
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = 1.03 * vec_simple_trades(u)(dim_vec_simple_trade_price)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                    
                                                End If
                                                
                                                Exit For
                                            End If
                                        Next u
                                        
                                        
                                    Else
                                        GoTo choose_sell_price
                                    End If
                                    
                                    
                                    
                                    
                                End If
                                
                            ElseIf qty_to_trade > 0 Then
                                
                                If IsEmpty(vec_simple_trades(q)(dim_vec_simple_trade_price)) Then
choose_buy_price:
                                    If s1 < last_price Then
                                        vec_simple_trades(q)(dim_vec_simple_trade_price) = s1
                                    Else
                                        
                                        If s2 < last_price Then
                                            vec_simple_trades(q)(dim_vec_simple_trade_price) = s2
                                        Else
                                            If s3 < last_price Then
                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = s3
                                            Else
                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = 0.995 * last_price
                                            End If
                                        End If
                                        
                                    End If
                                    
                                Else
                                    
                                    'check que marge de manoeuvre
                                    If UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "BASE" Then
                                        If vec_simple_trades(q)(dim_vec_simple_trade_price) > last_price Then
                                            GoTo choose_buy_price
                                        End If
                                    ElseIf UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "STP" Or UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "STOP" Then
                                        
                                        'repere order base et s assure que prix en dessous
                                        For u = q - 1 To 0 Step -1
                                            If vec_simple_trades(q)(dim_vec_simple_trade_ticker) = vec_simple_trades(u)(dim_vec_simple_trade_ticker) And UCase(vec_simple_trades(u)(dim_vec_simple_trade_order_type)) = "BASE" Then
                                                
                                                'buy (stop) d un original order sell
                                                If vec_simple_trades(q)(dim_vec_simple_trade_price) <= vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                    'le prix doit etre remplace
                                                    
                                                    If r1 > last_price And r1 > vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                        vec_simple_trades(q)(dim_vec_simple_trade_price) = r1
                                                        Exit For
                                                    Else
                                                        If r2 > last_price And r2 > vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                            vec_simple_trades(q)(dim_vec_simple_trade_price) = r2
                                                            Exit For
                                                        Else
                                                            If r3 > last_price And r3 > vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = r3
                                                                Exit For
                                                            Else
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = 1.03 * vec_simple_trades(u)(dim_vec_simple_trade_price)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                End If
                                                
                                                Exit For
                                                
                                            End If
                                        Next u
                                        
                                    ElseIf UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "TGT" Or UCase(vec_simple_trades(q)(dim_vec_simple_trade_order_type)) = "TARGET" Then
                                        
                                        'repere order base et s assure que prix en dessous
                                        For u = q - 1 To 0 Step -1
                                            If vec_simple_trades(q)(dim_vec_simple_trade_ticker) = vec_simple_trades(u)(dim_vec_simple_trade_ticker) And UCase(vec_simple_trades(u)(dim_vec_simple_trade_order_type)) = "BASE" Then
                                                
                                                'buy (target) d un original order sell
                                                If vec_simple_trades(q)(dim_vec_simple_trade_price) >= vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                    'le prix doit etre remplace
                                                    
                                                    If s1 < last_price And s1 < vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                        vec_simple_trades(q)(dim_vec_simple_trade_price) = s1
                                                        Exit For
                                                    Else
                                                        If s2 < last_price And s2 < vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                            vec_simple_trades(q)(dim_vec_simple_trade_price) = s2
                                                            Exit For
                                                        Else
                                                            If s3 < last_price And s3 < vec_simple_trades(u)(dim_vec_simple_trade_price) Then
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = s3
                                                                Exit For
                                                            Else
                                                                vec_simple_trades(q)(dim_vec_simple_trade_price) = 0.97 * vec_simple_trades(u)(dim_vec_simple_trade_price)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                End If
                                                
                                                Exit For
                                            End If
                                        Next u
                                        
                                        
                                    Else
                                        GoTo choose_buy_price
                                    End If
                                    
                                End If
                                
                            End If
                            
                            
                            ReDim Preserve vec_trade_based_on_simple(count_vec_trade_based_on_simple)
                            vec_trade_based_on_simple(count_vec_trade_based_on_simple) = Array(vec_simple_trades(q)(dim_vec_simple_trade_ticker), qty_to_trade, vec_simple_trades(q)(dim_vec_simple_trade_price), vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_simple_trades(q)(dim_vec_simple_trade_src), vec_simple_trades(q)(dim_vec_simple_trade_order_type), Empty)
                            count_vec_trade_based_on_simple = count_vec_trade_based_on_simple + 1
                            
                            Exit For
                        End If
                    Next j
                
                End If
                
                
                Exit For
            Else
                If i = UBound(vec_ticker_details, 1) Then
                    Debug.Print "@generate_vec_trades: $ticker: " & vec_simple_trades(q)(dim_vec_simple_trade_ticker) & " didn't passe all filters"
                End If
            End If
        Next i
check_next_entry_vec_ticker_details_for_construct_trades_simple_trades:
    Next q
    
    need_resort = True
    
    GoTo normal_procedure_withtout_vec_simple_trades
    
Else
    
normal_procedure_withtout_vec_simple_trades:
    
    
    For i = 0 To UBound(vec_ticker_details, 1)
        If IsNumeric(output_bbg(i)(dim_bbg_yest_close)) And IsNumeric(output_bbg(i)(dim_bbg_yest_low)) And IsNumeric(output_bbg(i)(dim_bbg_yest_high)) Then
            
            If Worksheets("FORMAT2").CB_check_tradeable_hour.Value = True Then
                
                For j = 0 To UBound(exchange_tradeable_security_status, 1)
                    If exchange_tradeable_security_status(j) = output_bbg(i)(dim_bbg_exchange_status) Then
                        Exit For
                    Else
                        If j = UBound(exchange_tradeable_security_status, 1) Then
                            GoTo check_next_entry_vec_ticker_details_for_construct_trades
                        End If
                    End If
                Next j
                
            End If
            
            
            p = Round((output_bbg(i)(dim_bbg_yest_close) + output_bbg(i)(dim_bbg_yest_low) + output_bbg(i)(dim_bbg_yest_high)) / 3, 3)
            
            r1 = Round(2 * p - output_bbg(i)(dim_bbg_yest_low), 3)
            s1 = Round(2 * p - output_bbg(i)(dim_bbg_yest_high), 3)
            
            r2 = Round((p - s1) + r1, 3)
            s2 = Round(p - (r1 - s1), 3)
            
            r3 = Round((p - s2) + r2, 3)
            s3 = Round(p - (r2 - s2), 3)
            
            last_price = output_bbg(i)(dim_bbg_last_price)
            
                    
            For j = 0 To UBound(vec_change_rate, 1)
                If vec_change_rate(j)(dim_vec_change_rate_txt) = vec_ticker_details(i)(dim_vec_ticker_details_crncy) Then
                    
                    If Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) = True And Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value) = True And Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value) = True Then
                        
                        'check l'interval de theta
                        If Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) >= 0 And Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) < Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value)) Then
                            qty_to_trade = Round(Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                        ElseIf Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) >= Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x2.Value)) And Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) < Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value)) Then
                            qty_to_trade = 2 * Round(Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                        ElseIf Abs(vec_ticker_details(i)(dim_vec_ticker_details_theta)) >= Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta_limit_x3.Value)) Then
                            qty_to_trade = 3 * Round(Abs(CDbl(Worksheets("FORMAT2").TB_dynamic_qty_based_theta.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                        End If
                        
                    Else
                        If Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value <> "" And IsNumeric(Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value) Then
                            If Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value > 0 Then
                                qty_to_trade = Round(Abs(CDbl(Worksheets("FORMAT2").TB_valeur_eur_each_trade.Value) / (output_bbg(i)(dim_bbg_yest_close) * vec_change_rate(j)(dim_vec_change_rate_rate))), 0)
                            Else
                                If IsError(vec_ticker_details(i)(dim_vec_ticker_details_delta)) = False Then
                                    qty_to_trade = Round(Abs(vec_ticker_details(i)(dim_vec_ticker_details_delta)) / 3, 0)
                                Else
                                    qty_to_trade = 0
                                End If
                            End If
                        Else
                            If IsError(vec_ticker_details(i)(dim_vec_ticker_details_delta)) = False Then
                                qty_to_trade = Round(Abs(vec_ticker_details(i)(dim_vec_ticker_details_delta)) / 3, 0)
                            Else
                                qty_to_trade = 0
                            End If
                        End If
                    End If
                    
                    
                    Exit For
                End If
            Next j
            
                
                If Worksheets("FORMAT2").CB_buy_s3.Value = True And s3 < last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), Abs(qty_to_trade), s3, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                If Worksheets("FORMAT2").CB_buy_s2.Value = True And s2 < last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), Abs(qty_to_trade), s2, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                If Worksheets("FORMAT2").CB_buy_s1.Value = True And s1 < last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), Abs(qty_to_trade), s1, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                If Worksheets("FORMAT2").CB_buy_p.Value = True And p < last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), Abs(qty_to_trade), p, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                
                If Worksheets("FORMAT2").CB_sell_p.Value = True And p > last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), -Abs(qty_to_trade), p, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                
                If Worksheets("FORMAT2").CB_smart_p.Value = True Then
                    
                    ReDim Preserve vec_trades(k)
                    
                    If last_price > p Then
                        vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), Abs(qty_to_trade), p, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    Else
                        vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), -Abs(qty_to_trade), p, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    End If
                    
                    k = k + 1
                End If
                
                
                If Worksheets("FORMAT2").CB_sell_R1.Value = True And r1 > last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), -Abs(qty_to_trade), r1, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                
                If Worksheets("FORMAT2").CB_sell_R2.Value = True And r2 > last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), -Abs(qty_to_trade), r2, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                
                If Worksheets("FORMAT2").CB_sell_R3.Value = True And r3 > last_price Then
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), -Abs(qty_to_trade), r3, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                    k = k + 1
                End If
                
                
                If Worksheets("FORMAT2").CB_get_out.Value = True Then
                    
                    If vec_ticker_details(i)(dim_vec_ticker_details_line) > 0 Then 'y a-t-il vraiment une pos dans le book ?
                        If IsError(Worksheets("Equity_Database").Cells(vec_ticker_details(i)(dim_vec_ticker_details_line), c_equity_db_delta)) = False Then
                            
                            If IsNumeric(Worksheets("Equity_Database").Cells(vec_ticker_details(i)(dim_vec_ticker_details_line), c_equity_db_delta)) = True Then
                                
                                If Worksheets("Equity_Database").Cells(vec_ticker_details(i)(dim_vec_ticker_details_line), c_equity_db_delta) <> 0 Then
                                    ReDim Preserve vec_trades(k)
                                    
                                    'get_out_type (pct_last_price ou near piv)
                                    
                                    get_out_type = Worksheets("FORMAT2").CB_get_out_sens.Value
                                    vec_trades(k) = Array(vec_ticker_details(i)(dim_vec_ticker_details_ticker), -Round(Worksheets("Equity_Database").Cells(vec_ticker_details(i)(dim_vec_ticker_details_line), c_equity_db_delta), 0), "get_out_" & get_out_type, vec_ticker_details(i)(dim_vec_ticker_details_crncy), vec_ticker_details(i)(dim_vec_ticker_details_region), vec_ticker_details(i)(dim_vec_ticker_details_line), s3, s2, s1, p, r1, r2, r3, last_price, vec_ticker_details(i)(dim_vec_ticker_details_net_pos), vec_ticker_details(i)(dim_vec_ticker_details_src), "base", Empty)
                                    k = k + 1
                                    
                                End If
                            End If
                        End If
                    End If
                End If
                
        End If
check_next_entry_vec_ticker_details_for_construct_trades:
    Next i
End If



If need_resort = True And count_vec_trade_based_on_simple > 0 Then
    
    k = 0
    Dim final_vec_trade() As Variant
    
    For i = 0 To UBound(vec_trade_based_on_simple, 1)
        
        If i > 0 Then
            If vec_trade_based_on_simple(i)(dim_vec_trade_ticker) <> vec_trade_based_on_simple(i - 1)(dim_vec_trade_ticker) Then
                'rajoute transaction normals
                For j = 0 To UBound(vec_trades, 1)
                    If vec_trades(j)(dim_vec_trade_ticker) = vec_trade_based_on_simple(i - 1)(dim_vec_trade_ticker) Then
                        ReDim Preserve final_vec_trade(k)
                        final_vec_trade(k) = vec_trades(j)
                        k = k + 1
                    End If
                Next j
            End If
        End If
        
        ReDim Preserve final_vec_trade(k)
        final_vec_trade(k) = vec_trade_based_on_simple(i)
        k = k + 1
        
    Next i
    
    vec_trades = final_vec_trade
    
End If


If count_vec_trade_based_on_simple > 0 Or k > 0 Then
    generate_vec_trades = vec_trades
Else
    generate_vec_trades = Empty
End If


End Function





Private Sub color_format2_trades()

Dim color_stop As Integer, color_target As Integer
    color_stop = 26
    color_target = 33

Dim i As Integer


For i = l_format2_header + 1 To 30000
    
    If Worksheets("FORMAT2").Cells(i, c_format2_ticker).Value = "" Then
        Exit For
    Else
        If i Mod (2) = 0 Then
            
            For j = c_format2_ticker To c_format2_broker
                Worksheets("FORMAT2").Cells(i, j).Interior.Color = RGB(240, 240, 240)
            Next j
            
            Worksheets("FORMAT2").Cells(i, c_format2_last_price).Interior.ColorIndex = 20
            
            If InStr(Worksheets("FORMAT2").Cells(i, c_format2_source).Value, "*** AUTO-GENERATED ***") = 0 Then
                For j = c_format2_s3 To c_format2_r3
                    If Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex <> 3 And Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex <> 4 Then
                        Worksheets("FORMAT2").Cells(i, j).Interior.Color = RGB(240, 240, 240)
                    End If
                Next j
            Else
                
            End If
            
            For j = c_format2_pre_market_start_column To c_format2_eps
                Worksheets("FORMAT2").Cells(i, j).Interior.Color = RGB(240, 240, 240)
            Next j
            
            For j = c_format2_dmi To c_format2_source
                Worksheets("FORMAT2").Cells(i, j).Interior.Color = RGB(240, 240, 240)
            Next j
        
        End If
        
        
        
        'coloriage stop/target
        If InStr(UCase(Worksheets("FORMAT2").Cells(i, c_format2_s3)), "STOP") <> 0 Or InStr(Worksheets("FORMAT2").Cells(i, c_format2_s3), "STP") <> 0 Then
            For j = c_format2_s3 To c_format2_r3
                Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex = color_stop
            Next
        ElseIf InStr(UCase(Worksheets("FORMAT2").Cells(i, c_format2_s3)), "TARGET") <> 0 Or InStr(UCase(Worksheets("FORMAT2").Cells(i, c_format2_s3)), "TGT") <> 0 Then
            For j = c_format2_s3 To c_format2_r3
                Worksheets("FORMAT2").Cells(i, j).Interior.ColorIndex = color_target
            Next
        End If
        
    End If
Next i


End Sub


Public Sub show_form_vol_mgmt()

frm_Open_Volatility.Show

End Sub



Public Sub tactical_trading_in_open_helper()

Application.Calculation = xlCalculationManual

Dim nbre_column_to_color As Integer
    nbre_column_to_color = 6

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

'remonte les distinct id + ticker d'open
Dim l_open_header As Integer, c_open_underlying_id As Integer, c_open_underlying_ticker As Integer

    l_open_header = 25
    c_open_underlying_id = 1
    c_open_underlying_ticker = 104
    c_open_line_equity_db = 103

dim_underlying_id = 0
dim_underlying_ticker = 1
dim_equity_db_line = 2
dim_data = 3
dim_line_open = 4

Dim vec_dinstinct_open_product() As Variant
ReDim Preserve vec_dinstinct_open_product(0)
vec_dinstinct_open_product(0) = Array("", "") 'product id + ticker
k = 0
For i = l_open_header + 1 To 3000
    If Worksheets("Open").Cells(i, c_open_underlying_id) = "" And Worksheets("Open").Cells(i + 1, c_open_underlying_id) = "" And Worksheets("Open").Cells(i + 2, c_open_underlying_id) = "" Then
        Exit For
    Else
        If Worksheets("Open").Cells(i, c_open_underlying_id) <> "" Then
            If InStr(UCase(Worksheets("Open").Cells(i, c_open_underlying_ticker)), "INDEX") = 0 Then
                
                'retire si precedent msg color
                If Worksheets("Open").Cells(i, 1).Interior.ColorIndex <> Worksheets("Open").Cells(i, 250).Interior.ColorIndex Then
                    For j = 1 To nbre_column_to_color
                        Worksheets("Open").Cells(i, j).Interior.ColorIndex = Worksheets("Open").Cells(i, 250).Interior.ColorIndex
                    Next j
                End If
                
                start_looking = UBound(vec_dinstinct_open_product, 1) - 3
                
                If start_looking < 0 Then
                    start_looking = 0
                End If
                
                For j = start_looking To UBound(vec_dinstinct_open_product, 1)
                    If vec_dinstinct_open_product(j)(1) = UCase(Worksheets("Open").Cells(i, c_open_underlying_ticker)) Then
                        Exit For
                    Else
                        If j = UBound(vec_dinstinct_open_product, 1) Then 'nouveau produit
                            ReDim Preserve vec_dinstinct_open_product(k)
                            vec_dinstinct_open_product(k) = Array(Worksheets("Open").Cells(i, c_open_underlying_id).Value, UCase(Worksheets("Open").Cells(i, c_open_underlying_ticker).Value), Worksheets("Open").Cells(i, c_open_line_equity_db).Value, -1, i)
                            k = k + 1
                        End If
                    End If
                Next j
            End If
        End If
    End If
Next i





Dim col_equity_db As Variant, col_equity_db_bb As Variant
    col_equity_db = Array("Equities_Name", "Valeur_Euro", "Delta", "Daily Change", "Daily Result", "Result Total", "Equity_Spot", "EQY_BETA", "Equity_Close", "Total Number", "Currency", "Sector", "Industry", "Perso_rel_1d", "json_tag")
    col_equity_db_bb = Array("PX_YEST_LOW", "PX_YEST_HIGH", "PX_YEST_CLOSE", "MOV_AVG_20D", "MOV_AVG_50D", "MOV_AVG_200D", "EQY_BOLLINGER_UPPER", "EQY_BOLLINGER_MID", "EQY_BOLLINGER_LOWER", "RSI_14D", "HIGH_52WEEK", "LOW_52WEEK")
    col_central = Array("Rank_EPS_4w_chg_curr_yr", "Rank_EPS_4w_chg_nxt_yr", "Rank_MoneyFlow", "Rank_Ratio_EPS_curr_yr_lst", "Rank_Ratio_EPS_nxt_yr_curr_yr", "Rank_EPS", "Rank_Overall")
    
    
    
    

    
    'construction du vecteur header
    Dim l_equity_db_header As Integer, l_equity_db_bb As Integer
        l_equity_db_header = 25
        l_equity_db_bb = 10
    
    
    
    'lie les lignes et les id pour equity_database (pour eviter de tjr devoir passer par la worksheet)
    Dim vec_equity_db_id_and_line() As Variant
    k = 0
    For i = l_equity_db_header + 2 To 32000 Step 2
        If Worksheets("Equity_Database").Cells(i, 1) = "" Then
            Exit For
        Else
            ReDim Preserve vec_equity_db_id_and_line(k)
            vec_equity_db_id_and_line(k) = Array(i, Worksheets("Equity_Database").Cells(i, 1).Value)
            k = k + 1
        End If
    Next i
    
    
    'lie les lignes et les id pour equity_database_bb
    Dim vec_equity_db_bb_id_and_line() As Variant
    k = 0
    For i = l_equity_db_bb + 1 To 32000
        If Worksheets("Equity_Database_BB").Cells(i, 1) = "" Then
            Exit For
        Else
            ReDim Preserve vec_equity_db_bb_id_and_line(k)
            vec_equity_db_bb_id_and_line(k) = Array(i, Worksheets("Equity_Database_BB").Cells(i, 1).Value)
            k = k + 1
        End If
    Next i
    
    
    
    Dim tmp_vec_header() As Variant
    k = 0
    
    ReDim Preserve tmp_vec_header(2)
        tmp_vec_header(0) = Array("system", "underlying_id", 0)
        k = k + 1
        tmp_vec_header(1) = Array("system", "bloomberg_ticker", 0)
        k = k + 1
        tmp_vec_header(2) = Array("system", "line_equity_db", 0)
        k = k + 1
    
    For i = 0 To UBound(col_equity_db, 1)
        
        For j = 1 To 250
            If Worksheets("Equity_Database").Cells(l_equity_db_header, j) = col_equity_db(i) Then
                ReDim Preserve tmp_vec_header(k)
                tmp_vec_header(k) = Array("Equity_Database", col_equity_db(i), j)
                k = k + 1
                Exit For
            End If
        Next j
        
    Next i
    
    
    For i = 0 To UBound(col_equity_db_bb, 1)
        
        For j = 1 To 250
            If Worksheets("Equity_Database_BB").Cells(l_equity_db_bb, j) = col_equity_db_bb(i) Then
                ReDim Preserve tmp_vec_header(k)
                tmp_vec_header(k) = Array("Equity_Database_BB", col_equity_db_bb(i), j)
                k = k + 1
                Exit For
            End If
        Next j
        
    Next i
    
    
    For i = 0 To UBound(col_central, 1)
        ReDim Preserve tmp_vec_header(k)
        tmp_vec_header(k) = Array("CENTRAL", col_central(i), 0)
        k = k + 1
    Next i
    
    


'remonte les differents champs des differentes sheets
Dim tmp_vec() As Variant
k = 0
For i = 0 To UBound(vec_dinstinct_open_product, 1)
    
    ReDim tmp_vec(UBound(tmp_vec_header, 1))
    
    
    'mount product id + ticker
    tmp_vec(0) = vec_dinstinct_open_product(i)(dim_underlying_id)
    tmp_vec(1) = vec_dinstinct_open_product(i)(dim_underlying_ticker)
    tmp_vec(2) = vec_dinstinct_open_product(i)(dim_equity_db_line)
    
    'repere les donnees dans les differentes sheets
    If Worksheets("Equity_Database").Cells(vec_dinstinct_open_product(i)(dim_equity_db_line), 1) = vec_dinstinct_open_product(i)(dim_underlying_id) Then
        
        For m = 0 To UBound(tmp_vec_header, 1)
            If tmp_vec_header(m)(0) = "Equity_Database" Then
                If IsError(Worksheets("Equity_Database").Cells(vec_dinstinct_open_product(i)(dim_equity_db_line), tmp_vec_header(m)(2))) = False Then
                    tmp_vec(m) = Worksheets("Equity_Database").Cells(vec_dinstinct_open_product(i)(dim_equity_db_line), tmp_vec_header(m)(2))
                End If
            End If
        Next m
        
    Else
        
        'repere manuellement avec une boucle a travers equity db - pour faciliter et acceler le traitement, utilise un vec qui relie ligne et product id
        For j = 0 To UBound(vec_equity_db_id_and_line, 1)
            If vec_equity_db_id_and_line(j)(1) = vec_dinstinct_open_product(i)(dim_underlying_id) Then
                
                For m = 0 To UBound(tmp_vec_header, 1)
                    If tmp_vec_header(m)(0) = "Equity_Database" Then
                        If IsError(Worksheets("Equity_Database").Cells(vec_equity_db_id_and_line(j)(0), tmp_vec_header(m)(2))) = False Then
                            tmp_vec(m) = Worksheets("Equity_Database").Cells(vec_equity_db_id_and_line(j)(0), tmp_vec_header(m)(2))
                            vec_dinstinct_open_product(i)(dim_equity_db_line) = vec_equity_db_id_and_line(j)(0)
                        End If
                    End If
                Next m
                
                Exit For
            End If
        Next j
        
    End If
    
    
    
    For j = 0 To UBound(vec_equity_db_bb_id_and_line, 1)
        If vec_equity_db_bb_id_and_line(j)(1) = vec_dinstinct_open_product(i)(dim_underlying_id) Then
            
            For m = 0 To UBound(tmp_vec_header, 1)
                If tmp_vec_header(m)(0) = "Equity_Database_BB" Then
                    If IsError(Worksheets("Equity_Database_BB").Cells(vec_equity_db_bb_id_and_line(j)(0), tmp_vec_header(m)(2))) = False Then
                        tmp_vec(m) = Worksheets("Equity_Database_BB").Cells(vec_equity_db_bb_id_and_line(j)(0), tmp_vec_header(m)(2))
                    End If
                End If
            Next m
            
            Exit For
        End If
    Next j
    
    vec_dinstinct_open_product(i)(dim_data) = tmp_vec
    
Next i






's'assure que les donnes sont a jour
For i = 0 To UBound(tmp_vec_header, 1)
    If tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "PX_YEST_CLOSE" Then
        c_yest_close_price_bb = i
    ElseIf tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Close" Then
        c_yest_close_price = i
    End If
Next i

Dim is_data_equity_db_bb_updated As Boolean
is_data_equity_db_bb_updated = True
For i = 0 To 10
    If IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price_bb)) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price_bb) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price)) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) > 0 Then
        If vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price_bb) <> vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) Then
            is_data_equity_db_bb_updated = False
        End If
    End If
Next i

If is_data_equity_db_bb_updated = False Then
    MsgBox ("Data from worksheet Equtiy_Database_BB not updated. Run Bloomberg(API) -> Equity Database BB")
    Exit Sub
End If




Dim vec_message() As Variant
k = 0
'lance les differents filtres

dim_message_product_id = 0
dim_message_underlying_id = 1
dim_message_equity_db_line = 2
dim_message_open_line = 3
dim_message_msg = 4
dim_message_color = 5

'MA20
c_last_price = 0
c_yest_close_price = 0
c_mav20 = 0

For i = 0 To UBound(tmp_vec_header, 1)
    If tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Spot" Then
        c_last_price = i
    ElseIf tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Close" Then
        c_yest_close_price = i
    ElseIf tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "MOV_AVG_20D" Then
        c_mav20 = i
    End If
Next i

'repere ceux qui ont breakes
For i = 0 To UBound(vec_dinstinct_open_product, 1)
    If IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_last_price)) And vec_dinstinct_open_product(i)(dim_data)(c_last_price) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price)) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_mav20)) And vec_dinstinct_open_product(i)(dim_data)(c_mav20) > 0 Then
        
        'break up
        If vec_dinstinct_open_product(i)(dim_data)(c_last_price) > vec_dinstinct_open_product(i)(dim_data)(c_mav20) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) < vec_dinstinct_open_product(i)(dim_data)(c_mav20) Then
            ReDim Preserve vec_message(k)
            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "break ma20 up", 4)
            k = k + 1
        End If
        
        'break down
        If vec_dinstinct_open_product(i)(dim_data)(c_last_price) < vec_dinstinct_open_product(i)(dim_data)(c_mav20) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) > vec_dinstinct_open_product(i)(dim_data)(c_mav20) Then
            ReDim Preserve vec_message(k)
            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "break ma20 down", 7)
            k = k + 1
        End If
        
    End If
Next i



'MA200
c_last_price = 0
c_yest_close_price = 0
c_mav20 = 0

For i = 0 To UBound(tmp_vec_header, 1)
    If tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Spot" Then
        c_last_price = i
    ElseIf tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Close" Then
        c_yest_close_price = i
    ElseIf tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "MOV_AVG_200D" Then
        c_mav200 = i
    End If
Next i

'repere ceux qui ont breakes
For i = 0 To UBound(vec_dinstinct_open_product, 1)
    If IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_last_price)) And vec_dinstinct_open_product(i)(dim_data)(c_last_price) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price)) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_mav200)) And vec_dinstinct_open_product(i)(dim_data)(c_mav200) > 0 Then
        
        'break up
        If vec_dinstinct_open_product(i)(dim_data)(c_last_price) > vec_dinstinct_open_product(i)(dim_data)(c_mav200) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) < vec_dinstinct_open_product(i)(dim_data)(c_mav200) Then
            ReDim Preserve vec_message(k)
            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "break ma200 up", 43)
            k = k + 1
        End If
        
        'break down
        If vec_dinstinct_open_product(i)(dim_data)(c_last_price) < vec_dinstinct_open_product(i)(dim_data)(c_mav200) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close_price) > vec_dinstinct_open_product(i)(dim_data)(c_mav200) Then
            ReDim Preserve vec_message(k)
            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "break ma200 down", 3)
            k = k + 1
        End If
        
    End If
Next i



''rsi
'c_rsi = 0
'For i = 0 To UBound(tmp_vec_header, 1)
'    If tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "RSI_14D" Then
'        c_rsi = i
'    End If
'Next i
'
''repere ceux qui ont breakes
'For i = 0 To UBound(vec_dinstinct_open_product, 1)
'    If IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_rsi)) Then
'
'        If vec_dinstinct_open_product(i)(dim_data)(c_rsi) < 35 Then
'            ReDim Preserve vec_message(k)
'            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "rsi14 up")
'            k = k + 1
'        End If
'
'        If vec_dinstinct_open_product(i)(dim_data)(c_rsi) > 70 Then
'            ReDim Preserve vec_message(k)
'            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "rsi14 down")
'            k = k + 1
'        End If
'
'    End If
'Next i
'
'
'
''bollinger
'c_last_price = 0
'c_boll_upper = 0
'c_ma20 = 0
'c_yest_close = 0
'
'For i = 0 To UBound(tmp_vec_header, 1)
'    If tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "EQY_BOLLINGER_UPPER" Then
'        c_boll_upper = i
'    ElseIf tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "MOV_AVG_20D" Then
'        c_ma20 = i
'    ElseIf tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Spot" Then
'        c_last_price = i
'    ElseIf tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Close" Then
'        c_yest_close = i
'    End If
'Next i
'
'For i = 0 To UBound(vec_dinstinct_open_product, 1)
'    If IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_last_price)) And vec_dinstinct_open_product(i)(dim_data)(c_last_price) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_boll_upper)) And vec_dinstinct_open_product(i)(dim_data)(c_boll_upper) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_ma20)) And vec_dinstinct_open_product(i)(dim_data)(c_ma20) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_yest_close)) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close) > 0 Then
'
'        'pos la condition
'        If (vec_dinstinct_open_product(i)(dim_data)(c_last_price) > vec_dinstinct_open_product(i)(dim_data)(c_boll_upper) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close) < vec_dinstinct_open_product(i)(dim_data)(c_boll_upper)) Or (vec_dinstinct_open_product(i)(dim_data)(c_last_price) < vec_dinstinct_open_product(i)(dim_data)(c_ma20) And vec_dinstinct_open_product(i)(dim_data)(c_yest_close) > vec_dinstinct_open_product(i)(dim_data)(c_boll_upper)) Then
'            ReDim Preserve vec_message(k)
'            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "bollinger")
'            k = k + 1
'        End If
'
'    End If
'Next i
'
'
'
'' high / low
'c_last_price = 0
'c_high_52w = 0
'c_low_52w = 0
'
'For i = 0 To UBound(tmp_vec_header, 1)
'    If tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "HIGH_52WEEK" Then
'        c_high_52w = i
'    ElseIf tmp_vec_header(i)(0) = "Equity_Database_BB" And tmp_vec_header(i)(1) = "LOW_52WEEK" Then
'        c_low_52w = i
'    ElseIf tmp_vec_header(i)(0) = "Equity_Database" And tmp_vec_header(i)(1) = "Equity_Spot" Then
'        c_last_price = i
'    End If
'Next i
'
'For i = 0 To UBound(vec_dinstinct_open_product, 1)
'    If IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_last_price)) And vec_dinstinct_open_product(i)(dim_data)(c_last_price) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_high_52w)) And vec_dinstinct_open_product(i)(dim_data)(c_high_52w) > 0 And IsNumeric(vec_dinstinct_open_product(i)(dim_data)(c_low_52w)) And vec_dinstinct_open_product(i)(dim_data)(c_low_52w) > 0 Then
'
'        If vec_dinstinct_open_product(i)(dim_data)(c_last_price) / vec_dinstinct_open_product(i)(dim_data)(c_high_52w) > 0.96 Then
'            ReDim Preserve vec_message(k)
'            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "near 52w high")
'            k = k + 1
'        End If
'
'
'        If vec_dinstinct_open_product(i)(dim_data)(c_last_price) / vec_dinstinct_open_product(i)(dim_data)(c_low_52w) < 1.04 Then
'            ReDim Preserve vec_message(k)
'            vec_message(k) = Array(vec_dinstinct_open_product(i)(dim_underlying_id), vec_dinstinct_open_product(i)(dim_underlying_ticker), vec_dinstinct_open_product(i)(dim_equity_db_line), vec_dinstinct_open_product(i)(dim_line_open), "near 52w low")
'            k = k + 1
'        End If
'
'    End If
'Next i


'store des resultat dans equtiy db en JSON + repere les new
'Dim oJSON As New JSONLib

'coloriage dans open
If k > 0 Then
    For i = 0 To UBound(vec_message, 1)
        For j = 1 To nbre_column_to_color
            Worksheets("Open").Cells(vec_message(i)(dim_message_open_line), j).Interior.ColorIndex = vec_message(i)(dim_message_color)
        Next j
    Next i
End If


''impression d'un msg
'If k > 0 Then
'    With frm_Alerts.LV_alerts
'
'        .view = lvwReport
'        .FullRowSelect = True
'
'        With .ColumnHeaders
'            .Clear
'
'            .Add , , "Ticker", 100
'            .Add , , "Message", 100
'        End With
'
'        k = 1
'        For i = 0 To UBound(vec_message, 1)
'
'            .ListItems.Add , , vec_message(i)(1)
'
'                .ListItems(k).ListSubItems.Add , , vec_message(i)(4)
'
'            k = k + 1
'        Next i
'
'        frm_Alerts.Show
'
'    End With
'End If

Application.Calculation = xlCalculationAutomatic

End Sub


Public Sub load_cointrin_from_bloomberg()

Application.Calculation = xlCalculationManual

Worksheets("Cointrin").Cells.Clear

l_cointrin_header = 8

c_cointrin_product_id = 1
c_cointrin_underlying_id = 2
c_cointrin_description = 3
c_cointrin_currency = 4
c_cointrin_close_position = 5
c_cointrin_close_price = 6
c_cointrin_ytd_pnl_gross = 7
c_cointrin_commt = 8
c_cointrin_ytd_pnl_local = 9
c_cointrin_ytd_pnl_base = 10
c_cointrin_folio_close_position = 11
c_cointrin_gva_pnl_local = 12
c_cointrin_gva_close_position = 13
c_cointrin_delta_pos_folio = 14
c_cointrin_delta_pnl_gva = 15
c_cointrin_delta_pos_gva = 16
c_cointrin_factor = 17
c_cointrin_comm_base = 18


'header
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_product_id) = "Identifier"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_underlying_id) = "Identifier"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_description) = "gs_description"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_currency) = "currency"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_close_position) = "net_position"
    Worksheets("Cointrin").Cells(l_cointrin_header - 1, c_cointrin_close_position) = "IH:COINTRIN_CLOSE_POSITION:js_cntrn"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_close_price) = "gva_close_price"
    Worksheets("Cointrin").Cells(l_cointrin_header - 1, c_cointrin_close_price) = "IH:COINTRIN_CLOSE_PRICE:js_cntrn"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_ytd_pnl_gross) = "ytd_pnl_gross"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_commt) = "comm"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_ytd_pnl_local) = "ytd_pnl_net_local"
    Worksheets("Cointrin").Cells(l_cointrin_header - 1, c_cointrin_ytd_pnl_local) = "IH:COINTRIN_YTD_PNL_LOCAL_NET:js_cntrn"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_ytd_pnl_base) = "ytd_pnl_net_base"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_folio_close_position) = "folio_position"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_gva_pnl_local) = "GVA_pnl_local"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_gva_close_position) = "GVA_pos"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_delta_pos_folio) = "delta_pos_folio"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_delta_pnl_gva) = "delta_pnl_GVA"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_delta_pos_gva) = "delta_pos_GVA"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_factor) = "factor"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_comm_base) = "comm_base"


dim_ticker = 0
dim_product_id = 1
dim_close_position = 2
dim_close_price = 3
dim_ytd_pnl_local_net = 4


Dim list_ticker As Variant
list_ticker = get_list_distinct_ticker_from_local_database_to_export_bbg


k = l_cointrin_header + 1
For i = 0 To UBound(list_ticker, 1)
    If UCase(list_ticker(i)(0)) = "EQUITY" Then
        For j = 0 To UBound(list_ticker(i)(1), 1)
            
            Worksheets("Cointrin").Cells(k, c_cointrin_product_id) = list_ticker(i)(1)(j)(dim_product_id)
            Worksheets("Cointrin").Cells(k, c_cointrin_underlying_id) = list_ticker(i)(1)(j)(dim_product_id)
            Worksheets("Cointrin").Cells(k, c_cointrin_description) = list_ticker(i)(1)(j)(dim_ticker)
            Worksheets("Cointrin").Cells(k, c_cointrin_currency) = ""
            Worksheets("Cointrin").Cells(k, c_cointrin_close_position).FormulaLocal = "=IF(LEFT(BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_position) & "$" & l_cointrin_header - 1 & ");1)<>""#"";BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_position) & l_cointrin_header - 1 & ");0)"
            Worksheets("Cointrin").Cells(k, c_cointrin_close_price).FormulaLocal = "=IF(LEFT(BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_price) & "$" & l_cointrin_header - 1 & ");1)<>""#"";BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_price) & l_cointrin_header - 1 & ");0)"
            Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_gross) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_commt) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_local).FormulaLocal = "=IF(LEFT(BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_ytd_pnl_local) & "$" & l_cointrin_header - 1 & ");1)<>""#"";BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_ytd_pnl_local) & l_cointrin_header - 1 & ");0)"
            Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_base) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_folio_close_position) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_gva_pnl_local) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_gva_close_position) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_delta_pos_folio) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_delta_pnl_gva) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_delta_pos_gva) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_factor) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_comm_base) = 0
            
            k = k + 1
        Next j
    ElseIf UCase(list_ticker(i)(0)) = "INDEX" Then
        For j = 0 To UBound(list_ticker(i)(1), 1)
            
            Worksheets("Cointrin").Cells(k, c_cointrin_product_id) = list_ticker(i)(1)(j)(dim_product_id)
            Worksheets("Cointrin").Cells(k, c_cointrin_underlying_id) = list_ticker(i)(1)(j)(dim_product_id)
            Worksheets("Cointrin").Cells(k, c_cointrin_description) = transform_index_and_tracker(list_ticker(i)(1)(j)(dim_ticker))
            Worksheets("Cointrin").Cells(k, c_cointrin_currency) = ""
            Worksheets("Cointrin").Cells(k, c_cointrin_close_position).FormulaLocal = "=IF(LEFT(BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_position) & "$" & l_cointrin_header - 1 & ");1)<>""#"";BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_position) & l_cointrin_header - 1 & ");0)"
            Worksheets("Cointrin").Cells(k, c_cointrin_close_price).FormulaLocal = "=IF(LEFT(BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_price) & "$" & l_cointrin_header - 1 & ");1)<>""#"";BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_close_price) & l_cointrin_header - 1 & ");0)"
            Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_gross) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_commt) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_local).FormulaLocal = "=IF(LEFT(BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_ytd_pnl_local) & "$" & l_cointrin_header - 1 & ");1)<>""#"";BDP(" & xlColumnValue(c_cointrin_description) & k & ";$" & xlColumnValue(c_cointrin_ytd_pnl_local) & l_cointrin_header - 1 & ");0)"
            Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_base) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_folio_close_position) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_gva_pnl_local) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_gva_close_position) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_delta_pos_folio) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_delta_pnl_gva) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_delta_pos_gva) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_factor) = 0
            Worksheets("Cointrin").Cells(k, c_cointrin_comm_base) = 0
            
            k = k + 1
        Next j
    ElseIf UCase(list_ticker(i)(0)) = "FUTURE" Then
        For j = 0 To UBound(list_ticker(i)(1), 1)
            'prend directement les donnees de folio
        Next j
    End If
Next i

last_line_cointrin = k - 1

l_folio_future_header = 10
c_folio_future_product_id = 1
c_folio_future_underlying_id = 2
c_folio_future_description = 16
c_folio_future_qty_yest_close = 28
c_folio_future_close_price_local = 50

For i = l_folio_future_header + 2 To 300
    If Worksheets("Futures_Folio").Cells(i, c_folio_future_product_id) = "" Then
        Exit For
    Else
        
        Worksheets("Cointrin").Cells(k, c_cointrin_product_id) = Worksheets("Futures_Folio").Cells(i, c_folio_future_product_id)
        Worksheets("Cointrin").Cells(k, c_cointrin_underlying_id) = Worksheets("Futures_Folio").Cells(i, c_folio_future_underlying_id)
        Worksheets("Cointrin").Cells(k, c_cointrin_description) = Worksheets("Futures_Folio").Cells(i, c_folio_future_description)
        Worksheets("Cointrin").Cells(k, c_cointrin_currency) = ""
        Worksheets("Cointrin").Cells(k, c_cointrin_close_position) = Worksheets("Futures_Folio").Cells(i, c_folio_future_qty_yest_close)
        Worksheets("Cointrin").Cells(k, c_cointrin_close_price) = Worksheets("Futures_Folio").Cells(i, c_folio_future_close_price_local)
        Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_gross) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_commt) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_local) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_ytd_pnl_base) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_folio_close_position) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_gva_pnl_local) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_gva_close_position) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_delta_pos_folio) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_delta_pnl_gva) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_delta_pos_gva) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_factor) = 0
        Worksheets("Cointrin").Cells(k, c_cointrin_comm_base) = 0
        
        k = k + 1
        
    End If
Next i

Application.Calculation = xlCalculationAutomatic

answer = MsgBox("Freeze all datas in text ?", vbYesNo, "Cointrin")

If answer = vbYes Then
    'attend et fixe le contenu
    Application.Wait Now() + TimeValue("00:00:07")
    Application.Calculation = xlCalculationManual
    last_line_cointrin = k - 1
    
    For i = l_cointrin_header + 1 To last_line_cointrin
        Worksheets("Cointrin").Cells(i, c_cointrin_close_position) = Worksheets("Cointrin").Cells(i, c_cointrin_close_position).Value
        Worksheets("Cointrin").Cells(i, c_cointrin_close_price) = Worksheets("Cointrin").Cells(i, c_cointrin_close_price).Value
        Worksheets("Cointrin").Cells(i, c_cointrin_ytd_pnl_local) = Worksheets("Cointrin").Cells(i, c_cointrin_ytd_pnl_local).Value
    Next i
    
    Application.Calculation = xlCalculationAutomatic
End If

End Sub


Public Sub prepare_cointrin_for_export_bbg()

Dim list_ticker As Variant
list_ticker = get_list_distinct_ticker_from_local_database_to_export_bbg


'impression du rapport a exporter
'header
l_cointrin_header = 8
c_cointrin_ticker = 26
c_cointrin_product_id = 27
c_cointrin_close_position = 28
c_cointrin_close_price = 29
c_cointrin_ytd_pnl_local = 30

Worksheets("Cointrin").Columns(c_cointrin_ticker).Clear
Worksheets("Cointrin").Columns(c_cointrin_product_id).Clear
Worksheets("Cointrin").Columns(c_cointrin_close_position).Clear
Worksheets("Cointrin").Columns(c_cointrin_close_price).Clear
Worksheets("Cointrin").Columns(c_cointrin_ytd_pnl_local).Clear

Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_ticker) = "SECURITY"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_product_id) = "PRODUCT_ID_GS"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_close_position) = "COINTRIN_CLOSE_POSITION"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_close_price) = "COINTRIN_CLOSE_PRICE"
Worksheets("Cointrin").Cells(l_cointrin_header, c_cointrin_ytd_pnl_local) = "COINTRIN_YTD_PNL_LOCAL_NET"

k = l_cointrin_header + 1

For i = 0 To UBound(list_ticker, 1)
    If UCase(list_ticker(i)(0)) = "EQUITY" Then
        For j = 0 To UBound(list_ticker(i)(1), 1)
            For m = 0 To UBound(list_ticker(i)(1)(j), 1)
                Worksheets("Cointrin").Cells(k, c_cointrin_ticker + m) = list_ticker(i)(1)(j)(m)
            Next m
            k = k + 1
        Next j
    ElseIf UCase(list_ticker(i)(0)) = "INDEX" Then
        For j = 0 To UBound(list_ticker(i)(1), 1)
            Worksheets("Cointrin").Cells(k, c_cointrin_ticker) = transform_index_and_tracker(list_ticker(i)(1)(j)(0))
            For m = 1 To UBound(list_ticker(i)(1)(j), 1)
                Worksheets("Cointrin").Cells(k, c_cointrin_ticker + m) = list_ticker(i)(1)(j)(m)
            Next m
            k = k + 1
        Next j
    End If
Next i

End Sub


Public Function get_list_distinct_ticker_from_local_database_to_export_bbg() As Variant

Application.Calculation = xlCalculationManual

Dim i As Long, j As Long, k As Long
Dim tmp_ticker As String

Dim sheet_equity_db As String, sheet_index_db As String
    sheet_equity_db = "Equity_Database"
    sheet_index_db = "Index_Database"

Dim l_equity_db_header As Integer, l_index_db_header As Integer
    l_equity_db_header = 25
    l_index_db_header = 25


k = 0
Dim vec_crncy()
For i = 14 To 31
    ReDim Preserve vec_crncy(k)
    vec_crncy(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, Worksheets("Parametres").Cells(i, 6).Value)
    k = k + 1
Next i



Dim c_equity_db_product_id As Integer, c_equity_db_close_position As Integer, c_equity_db_close_price As Integer, c_equity_db_ytd_pnl_local As Integer, c_equity_db_ticker As Integer
c_equity_db_product_id = 1
c_equity_db_status = 4
c_equity_db_close_position = 25
c_equity_db_close_price = 27
c_equity_db_ytd_pnl_local = 28
c_equity_db_ticker = 47
c_equity_db_total_number_current = 26


dim_ticker = 0
dim_product_id = 1
dim_close_position = 2
dim_close_price = 3
dim_ytd_pnl_local_net = 4

Dim vec_equity_db() As Variant
ReDim vec_equity_db(0)
vec_equity_db(0) = Array("", "", 0, 0, 0)
k = 0
d = 0
Dim l_equity_db_last_line As Long
For i = l_equity_db_header + 2 To 32000 Step 2
    If Worksheets(sheet_equity_db).Cells(i, c_equity_db_product_id) = "" Then
        l_equity_db_last_line = i - 2
        Exit For
    Else
        
        tmp_ticker = UCase(Worksheets(sheet_equity_db).Cells(i, c_equity_db_ticker))
        
        Dim test_range As Range
        Set test_range = Worksheets(sheet_equity_db).Cells(i, c_equity_db_ticker)
        
        If test_range.comment Is Nothing Then
        Else
            If InStr(UCase(test_range.comment.Text), "EQUITY") <> 0 Then
                tmp_ticker = UCase(test_range.comment.Text)
            End If
        End If
        
        
        For j = 0 To UBound(vec_equity_db, 1)
            If vec_equity_db(j)(dim_ticker) = tmp_ticker Then
                
                If Abs(Worksheets(sheet_equity_db).Cells(i, c_equity_db_ytd_pnl_local)) > 10 Then
                    vec_equity_db(j)(dim_ytd_pnl_local_net) = vec_equity_db(j)(dim_ytd_pnl_local_net) + Round(Worksheets(sheet_equity_db).Cells(i, c_equity_db_ytd_pnl_local), 0)
                    
                    If Worksheets(sheet_equity_db).Cells(i, c_equity_db_close_position).Value <> 0 Then ' vec_equity_db(j)(dim_close_position) = 0
                        vec_equity_db(j)(dim_product_id) = Worksheets(sheet_equity_db).Cells(i, c_equity_db_product_id).Value
                        vec_equity_db(j)(dim_close_position) = Worksheets(sheet_equity_db).Cells(i, c_equity_db_close_position).Value
                        vec_equity_db(j)(dim_close_price) = Worksheets(sheet_equity_db).Cells(i, c_equity_db_close_price).Value
                    End If
                    
                Else
                End If
                'que faire avec les doublons
                d = d + 1
                Exit For
            Else
                If j = UBound(vec_equity_db, 1) Then
                    'ne pas prendre en compte tout les tickers
                    If Worksheets(sheet_equity_db).Cells(i, c_equity_db_total_number_current) = 0 And Worksheets(sheet_equity_db).Cells(i, c_equity_db_close_position) = 0 And Abs(Worksheets(sheet_equity_db).Cells(i, c_equity_db_ytd_pnl_local)) < 10 Then
                        e = e + 1
                    Else
                        ReDim Preserve vec_equity_db(k)
                        vec_equity_db(k) = Array(tmp_ticker, Worksheets(sheet_equity_db).Cells(i, c_equity_db_product_id).Value, Worksheets(sheet_equity_db).Cells(i, c_equity_db_close_position).Value, Worksheets(sheet_equity_db).Cells(i, c_equity_db_close_price).Value, Round(Worksheets(sheet_equity_db).Cells(i, c_equity_db_ytd_pnl_local).Value, 0))
                        k = k + 1
                    End If
                End If
            End If
        Next j
    End If
Next i


'meme system pour les index
Dim vec_index_db() As Variant
Dim vec_future_db() As Variant

ReDim vec_index_db(0)
ReDim vec_future_db(0)
vec_index_db(0) = Array("", "", 0, 0, 0)
vec_future_db(0) = Array("", "", 0, 0, 0)


c_index_db_index_product_id = 1
c_index_db_index_status = 4
c_index_db_index_ytd_pnl_base = 28
c_index_db_index_crncy = 107
c_index_db_index_ticker = 110
c_index_db_index_tracker = 185

c_index_db_future_product_id = 31
c_index_db_future_ticker = 34
c_index_db_future_close_position = 38
c_index_db_future_close_price = 43




c_index_db_total_number_current = 26

Dim ytd_pnl_index_local As Double
count_index = 0
count_fut = 0
Dim l_index_db_last_line As Long
For i = l_index_db_header + 2 To 500 Step 3
    
    If Worksheets(sheet_index_db).Cells(i, c_index_db_index_product_id) = "" Then
        Exit For
    Else
        
        If Worksheets(sheet_index_db).Cells(i, c_index_db_index_tracker) <> "" Then
        
            'ytd pnl des * derives
                'retraitement pour avoir la donne en local
                ytd_pnl_index_local = 0
                For j = 0 To UBound(vec_crncy, 1)
                    If Worksheets(sheet_index_db).Cells(i, c_index_db_index_crncy) = vec_crncy(j)(1) Then
                        ytd_pnl_index_local = Round(Worksheets(sheet_index_db).Cells(i, c_index_db_index_ytd_pnl_base) / vec_crncy(j)(2), 0)
                        Exit For
                    End If
                Next j
            
            ReDim Preserve vec_index_db(count_index)
            vec_index_db(count_index) = Array(UCase(Worksheets(sheet_index_db).Cells(i, c_index_db_index_ticker).Value), Worksheets(sheet_index_db).Cells(i, c_index_db_index_product_id).Value, 0, 0, ytd_pnl_index_local)
            count_index = count_index + 1
            
            'remonte les different fut de la ligne
            If Worksheets(sheet_index_db).Cells(i, c_index_db_future_product_id) <> "" And Worksheets(sheet_index_db).Cells(i, c_index_db_future_product_id) <> 0 And InStr(UCase(Worksheets(sheet_index_db).Cells(i, c_index_db_future_ticker).Value), "INDEX") <> 0 Then
                ReDim Preserve vec_future_db(count_fut)
                vec_future_db(count_fut) = Array(UCase(Worksheets(sheet_index_db).Cells(i, c_index_db_future_ticker).Value), Worksheets(sheet_index_db).Cells(i, c_index_db_future_product_id).Value, Worksheets(sheet_index_db).Cells(i, c_index_db_future_close_position).Value, Worksheets(sheet_index_db).Cells(i, c_index_db_future_close_price).Value, 0)
                count_fut = count_fut + 1
            End If
            
            If Worksheets(sheet_index_db).Cells(i, c_index_db_future_product_id + 1) <> "" And Worksheets(sheet_index_db).Cells(i, c_index_db_future_product_id + 1) <> 0 And InStr(UCase(Worksheets(sheet_index_db).Cells(i + 1, c_index_db_future_ticker).Value), "INDEX") <> 0 Then
                ReDim Preserve vec_future_db(count_fut)
                vec_future_db(count_fut) = Array(UCase(Worksheets(sheet_index_db).Cells(i + 1, c_index_db_future_ticker).Value), Worksheets(sheet_index_db).Cells(i, c_index_db_future_product_id + 1).Value, Worksheets(sheet_index_db).Cells(i + 1, c_index_db_future_close_position).Value, Worksheets(sheet_index_db).Cells(i + 1, c_index_db_future_close_price).Value, 0)
                count_fut = count_fut + 1
            End If
        
        End If
        
    End If
    
    
Next i

get_list_distinct_ticker_from_local_database_to_export_bbg = Array(Array("equity", vec_equity_db), Array("index", vec_index_db), Array("future", vec_future_db))

End Function


Public Function transform_index_and_tracker(ByVal ticker As String) As String

l_index_db_header = 25
c_index_db_index_ticker = 110
c_index_db_tracker = 185

If InStr(UCase(ticker), "EQUITY") <> 0 Then 'reception d'un ETF
    
    For i = l_index_db_header + 2 To 500 Step 3
        If Worksheets("Index_Database").Cells(i, 1) = "" Then
            Exit For
        Else
            If UCase(Worksheets("Index_Database").Cells(i, c_index_db_tracker)) = UCase(ticker) Then
                transform_index_and_tracker = UCase(Worksheets("Index_Database").Cells(i, c_index_db_index_ticker))
                Exit For
            End If
        End If
    Next i
    
ElseIf InStr(UCase(ticker), "INDEX") <> 0 Then 'reception d'un index
    
    For i = l_index_db_header + 2 To 500 Step 3
        If Worksheets("Index_Database").Cells(i, 1) = "" Then
            Exit For
        Else
            If UCase(Worksheets("Index_Database").Cells(i, c_index_db_index_ticker)) = UCase(ticker) Then
                transform_index_and_tracker = UCase(Worksheets("Index_Database").Cells(i, c_index_db_tracker))
                Exit For
            End If
        End If
    Next i
    
End If

End Function


'essayer de faire un script sans la db afin de pouvoir l'utiliser partout
Public Function get_cash_settlement_for_future_and_option_on_index(ByVal product_id As String) As Double
    
    get_cash_settlement_for_future_and_option_on_index = -1
    
    'exchange / array(url_weekly, url_monthly)
    compo_website = Array(Array("cboe", Array(Array("MONTHLY", "http://www.cboe.com/data/Settlement.aspx"), Array("WEEKLY", "http://www.cboe.com/Data/WeeklysSettlements.aspx?DIR=TTMDIDXSettleValWeeklys&FILE=weeklys.doc"))), Array("eurex", Array(Array("MONTHLY", "http://www.eurexchange.com/market/clearing/finalsettlement_en.html"))))
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long, r As Long, s As Long
    
    Dim XMLHttpRequest As New MSXML2.XMLHTTP
    Dim HTMLDoc As New HTMLDocument
    
    Dim tagTable As HTMLtable
    Dim tagTr As HTMLTableRow, tagTr_cs As HTMLTableRow
    Dim tagTd As HTMLTableCell, tagTd_index As HTMLTableCell, tagTd_trading_symbol As HTMLTableCell, tagTd_expiration_date As HTMLTableCell, tagTd_settlement_symbol As HTMLTableCell, tagTd_settlement_value As HTMLTableCell
    

    Application.Calculation = xlCalculationManual
    
    c_open_product_id = 2
    c_open_underlying_id = 1
    c_open_product_type = 6
    
    c_open_expiry_date = 23
    c_open_characteristics = 106
    
    c_open_ticker_option = 105
    c_open_ticker_underlying = 104
    
    Dim is_future As Boolean, is_option As Boolean
    
    Dim date_expiry_derivative As Date
    Dim find_underlying As Boolean
        find_underlying = False
    
    deriviate_periodicity = "MONTHLY"
    
    
    'essaie de reperer le derive dans open
    For i = 26 To 5000
        If Worksheets("Open").Cells(i, c_open_product_id) = "" And Worksheets("Open").Cells(i + 1, c_open_product_id) = "" And Worksheets("Open").Cells(i + 2, c_open_product_id) = "" Then
            Exit For
        Else
            If Worksheets("Open").Cells(i, c_open_product_id) = product_id Then
                underlying_id = Worksheets("Open").Cells(i, c_open_underlying_id)
                find_underlying = True
                date_expiry_derivative = Worksheets("Open").Cells(i, c_open_expiry_date)
                
                date_english_txt_year = year(date_expiry_derivative)
                date_english_txt_month = Month(date_expiry_derivative)
                    If Len(date_english_txt_month) = 1 Then
                        date_english_txt_month = "0" & date_english_txt_month
                    End If
                date_english_txt_day = day(date_expiry_derivative)
                    If Len(date_english_txt_day) = 1 Then
                        date_english_txt_day = "0" & date_english_txt_day
                    End If
                
                date_english_txt_day_alt = day(date_expiry_derivative) + 1
                    If Len(date_english_txt_day_alt) = 1 Then
                        date_english_txt_day_alt = "0" & date_english_txt_day_alt
                    End If
                
                date_english_txt = Array(date_english_txt_month & "/" & date_english_txt_day & "/" & Right(date_english_txt_year, 2), date_english_txt_month & "/" & date_english_txt_day & "/" & date_english_txt_year, date_english_txt_month & "/" & date_english_txt_day_alt & "/" & Right(date_english_txt_year, 2), date_english_txt_month & "/" & date_english_txt_day_alt & "/" & date_english_txt_year)
                
                
                'ajustement de la periodicity
                underlying_symbol = UCase(Left(Worksheets("Open").Cells(i, c_open_ticker_underlying), InStr(Worksheets("Open").Cells(i, c_open_ticker_underlying), " ") - 1))
                derivative_symbol = UCase(Left(Worksheets("Open").Cells(i, c_open_ticker_option), InStr(Worksheets("Open").Cells(i, c_open_ticker_option), " ") - 1))
                
                If underlying_symbol & "W" = derivative_symbol Then
                    deriviate_periodicity = "WEEKLY"
                End If
                
                
                If Month(date_expiry_derivative) = 1 Then
                    date_month_english_txt = "JANUARY"
                ElseIf Month(date_expiry_derivative) = 2 Then
                    date_month_english_txt = "FEBRURAY"
                ElseIf Month(date_expiry_derivative) = 3 Then
                    date_month_english_txt = "MARCH"
                ElseIf Month(date_expiry_derivative) = 4 Then
                    date_month_english_txt = "APRIL"
                ElseIf Month(date_expiry_derivative) = 5 Then
                    date_month_english_txt = "MAY"
                ElseIf Month(date_expiry_derivative) = 6 Then
                    date_month_english_txt = "JUNE"
                ElseIf Month(date_expiry_derivative) = 7 Then
                    date_month_english_txt = "JULY"
                ElseIf Month(date_expiry_derivative) = 8 Then
                    date_month_english_txt = "AUGUST"
                ElseIf Month(date_expiry_derivative) = 9 Then
                    date_month_english_txt = "SEPTEMBER"
                ElseIf Month(date_expiry_derivative) = 10 Then
                    date_month_english_txt = "OCTOBER"
                ElseIf Month(date_expiry_derivative) = 11 Then
                    date_month_english_txt = "NOVEMBER"
                ElseIf Month(date_expiry_derivative) = 12 Then
                    date_month_english_txt = "DECEMBER"
                End If
                
                If Worksheets("Open").Cells(i, c_open_product_type) = "C" Or Worksheets("Open").Cells(i, c_open_product_type) = "P" Then
                    is_future = False
                    is_option = True
                ElseIf Worksheets("Open").Cells(i, c_open_product_type) = "F" Then
                    is_future = True
                    is_option = False
                Else
                    get_cash_settlement_for_future_and_option_on_index = -1
                    Exit Function
                End If
                
                Exit For
            End If
        End If
    Next i
    
    'si introuvable essaie de le reperer dans cointrin afin d'obtenir son underlying_id
    
    
    
    
    'se balade dans index_db pour reperer les details
    c_index_db_product_id = 1
    c_index_db_source_cash_settlement = 186
    
    For i = 27 To 500 Step 3
        If Worksheets("Index_Database").Cells(i, c_index_db_product_id) = underlying_id Then
            
            If Worksheets("Index_Database").Cells(i, c_index_db_source_cash_settlement) <> "" Then
                src_cash_settlement = Worksheets("Index_Database").Cells(i, c_index_db_source_cash_settlement)
                
                website = Left(src_cash_settlement, InStr(src_cash_settlement, ";") - 1)
                    
                    ticker_future = Replace(src_cash_settlement, website & ";", "")
                ticker_future = Left(ticker_future, InStr(ticker_future, ";") - 1)
                
                
                option_future = Replace(src_cash_settlement, website & ";" & ticker_future & ";", "")
                
                
                If is_future = True Then
                    ticker_derivative = ticker_future
                ElseIf is_option = True Then
                    ticker_derivative = option_future
                End If
                
                
                For m = 0 To UBound(compo_website, 1)
                    If UCase(compo_website(m)(0)) = UCase(website) Then
                        
                        For n = 0 To UBound(compo_website(m)(1), 1)
                            
                            If UCase(compo_website(m)(1)(n)(0)) = UCase(deriviate_periodicity) Then
                                
                                'envoi de la web query
                                XMLHttpRequest.Open "GET", compo_website(m)(1)(n)(1), False
                                XMLHttpRequest.send
                                
                                HTMLDoc.body.innerHTML = XMLHttpRequest.responseText
                                
                                
                                
                                
                                If UCase(website) = UCase("cboe") Then
                                    
                                    If UCase(deriviate_periodicity) = UCase("weekly") Then
                                        
                                        'repere le bon tableau
                                        For Each tagTr In HTMLDoc.getElementsByTagName("tr")
                                            If InStr(UCase(tagTr.innerHTML), UCase(date_month_english_txt)) <> 0 And InStr(UCase(tagTr.innerHTML), year(date_expiry_derivative)) <> 0 Then
                                                
                                                'mount le table
                                                Set tagTable = tagTr.parentElement
                                                
                                                For Each tagTr_cs In tagTable.getElementsByTagName("tr")
                                                    
                                                    k = 0
                                                    For Each tagTd In tagTr_cs.getElementsByTagName("td")
                                                        
                                                        If k = 0 Then
                                                            Set tagTd_index = tagTd
                                                            k = k + 1
                                                        ElseIf k = 1 Then
                                                            Set tagTd_trading_symbol = tagTd
                                                            k = k + 1
                                                        ElseIf k = 2 Then
                                                            Set tagTd_expiration_date = tagTd
                                                            k = k + 1
                                                        ElseIf k = 3 Then
                                                            Set tagTd_settlement_symbol = tagTd
                                                            k = k + 1
                                                        ElseIf k = 4 Then
                                                            Set tagTd_settlement_value = tagTd
                                                            k = k + 1
                                                        End If
                                                        
                                                    Next
                                                    
                                                    If k > 4 Then
                                                        

                                                        If InStr(UCase(tagTd_index.innerText), UCase(ticker_derivative)) <> 0 Then
                                                            
                                                            'passe en revue les dates
                                                            For p = 0 To UBound(date_english_txt, 1)
                                                                
                                                                If InStr(tagTd_expiration_date.innerText, date_english_txt(p)) <> 0 Then
                                                                    
                                                                    'la bonne ligne a ete repere s'assure qu'une valeur numeric est dispo pour le cash settlement
                                                                    If IsNumeric(Trim(tagTd_settlement_value.innerText)) = True Then
                                                                        get_cash_settlement_for_future_and_option_on_index = CDbl(Trim(tagTd_settlement_value.innerText))
                                                                        Exit Function
                                                                    End If
                                                                    
                                                                    Exit For
                                                                End If
                                                                
                                                            Next p
                                                            
                                                        End If
                                                    End If
                                                    
                                                Next
                                                
                                                
                                                
                                                Exit For
                                            End If
                                        Next
                                        
                                    ElseIf UCase(deriviate_periodicity) = UCase("monthly") Then
                                        
                                        'repere le bon tableau
                                        For Each tagTr In HTMLDoc.getElementsByTagName("tr")
                                            If InStr(UCase(tagTr.innerHTML), UCase(date_month_english_txt)) <> 0 And InStr(UCase(tagTr.innerHTML), year(date_expiry_derivative)) <> 0 And InStr(UCase(tagTr.innerHTML), UCase("SETTLEMENT VALUES")) <> 0 Then
                                                
                                                'mount le table
                                                Set tagTable = tagTr.parentElement
                                                
                                                
                                                For Each tagTr_cs In tagTable.getElementsByTagName("tr")
                                                    
                                                    'repere le symbol
                                                    k = 0
                                                    
                                                    For Each tagTd In tagTr_cs.getElementsByTagName("td")
                                                        If k = 0 Then
                                                            Set tagTd_index = tagTd
                                                            k = k + 1
                                                        ElseIf k = 1 Then
                                                            Set tagTd_settlement_value = tagTd
                                                            k = k + 1
                                                        End If
                                                    Next
                                                    
                                                    If k > 1 Then
                                                        
                                                        If InStr(UCase(tagTd_index.innerText), UCase(ticker_derivative)) <> 0 Then
                                                            If IsNumeric(Trim(tagTd_settlement_value.innerText)) = True Then
                                                                get_cash_settlement_for_future_and_option_on_index = CDbl(Trim(tagTd_settlement_value.innerText))
                                                                Exit Function
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                Next
                                            End If
                                        Next
                                        
                                    End If
                                    
                                ElseIf UCase(website) = UCase("eurex") Then
                                    
                                    If UCase(deriviate_periodicity) = UCase("monthly") Then
                                        
                                        For Each tagTr In HTMLDoc.getElementsByTagName("tr")
                                            k = 0
                                            For Each tagTd In tagTr.getElementsByTagName("td")
                                                If k = 0 Then
                                                    Set tagTd_index = tagTd
                                                    k = k + 1
                                                ElseIf k = 1 Then
                                                    Set tagTd_expiration_date = tagTd
                                                    k = k + 1
                                                ElseIf k = 2 Then
                                                    Set tagTd_settlement_value = tagTd
                                                    k = k + 1
                                                End If
                                            Next
                                            
                                            If k > 2 Then
                                                If InStr(UCase(tagTd_index.innerText), UCase(ticker_derivative)) <> 0 And InStr(UCase(Trim(tagTd_expiration_date.innerText)), UCase(date_month_english_txt)) <> 0 Then
                                                    
                                                    If IsNumeric(Replace(Trim(tagTd_settlement_value.innerText), ",", ".")) Then
                                                    
                                                        get_cash_settlement_for_future_and_option_on_index = CDbl(Replace(Trim(tagTd_settlement_value.innerText), ",", "."))
                                                        Exit Function
                                                    
                                                    End If
                                                    
                                                End If
                                            End If
                                            
                                        Next
                                        
                                    End If
                                    
                                End If
                                
                                
                            End If
                            
                        Next n
                        
                        
                        Exit For
                    End If
                Next m
                
                
                
                Exit For
            Else
                get_cash_settlement_for_future_and_option_on_index = -1
                Exit Function
            End If
            
            Exit For
        End If
    Next i
    
    
    
End Function


Sub fix_expired_options_volatility_in_open()

Dim date_tmp As Date, date_expiry As Date

Dim i As Long, j As Long, k As Long

c_open_product_type = 6
c_open_volatility = 25
c_open_expiry_date = 23


Application.Calculation = xlCalculationManual

For i = 26 To 3500
    If Worksheets("Open").Cells(i, 1) = "" And Worksheets("Open").Cells(i + 1, 1) = "" And Worksheets("Open").Cells(i + 2, 1) = "" Then
        Exit For
    Else
        If Worksheets("Open").Cells(i, c_open_product_type) = "C" Or Worksheets("Open").Cells(i, c_open_product_type) = "P" Then
            
            date_expiry = Worksheets("Open").Cells(i, c_open_expiry_date)
            
            If date_expiry < Date Then
                
                If IsError(Worksheets("Open").Cells(i, c_open_volatility)) Or Left(Worksheets("Open").Cells(i, c_open_volatility), 1) = "#" Then
                    
                    Worksheets("Open").Cells(i, c_open_volatility) = 22.22
                    
                End If
                
            End If
            
        End If
    End If
Next i

Application.Calculation = xlCalculationAutomatic

End Sub


Sub check_and_desactivate_hshares_premium_in_exe()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer

Dim c_index_db_name As Integer, c_index_db_underlying_id As Integer
    c_index_db_underlying_id = 1
    c_index_db_name = 2
    
Dim l_exe_header As Integer, c_exe_underlying_id As Integer, c_exe_qty As Integer, c_exe_result_executed As Integer
    l_exe_header = 16
    c_exe_underlying_id = 0
    c_exe_qty = 0
    c_exe_result_executed = 0

Dim hshares_underlying_id As String
    hshares_underlying_id = ""
    
For i = 27 To 500 Step 3
    If InStr(UCase(Worksheets("Index_Database").Cells(i, c_index_db_name)), "HSHARES") <> 0 Then
        hshares_underlying_id = Worksheets("Index_Database").Cells(i, c_index_db_underlying_id)
        Exit For
    Else
        If Worksheets("Index_Database").Cells(i, c_index_db_underlying_id) = "" Then
            MsgBox ("hshares not found in index db")
            Exit Sub
        End If
            
    End If
Next i

If hshares_underlying_id = "" Then
    MsgBox ("hshares not found in index db")
    Exit Sub
End If


'passe en revue les entrees d'exe avec comme underyling hshares et dont qty <> 0

'detection des colonnes
For i = 1 To 15
    If Worksheets("Exe").Cells(l_exe_header, i) = "Identifier" And c_exe_underlying_id = 0 Then
        c_exe_underlying_id = i
    ElseIf Worksheets("Exe").Cells(l_exe_header, i) = "Nombre" And c_exe_qty = 0 Then
        c_exe_qty = i
    ElseIf Worksheets("Exe").Cells(l_exe_header, i) = "Result Executed" And c_exe_result_executed = 0 Then
        c_exe_result_executed = i
    End If
Next i


Dim tmp_qty As Variant
Dim rng_description  As Range

For i = l_exe_header + 1 To 5000
    If Worksheets("Exe").Cells(i, c_exe_underlying_id) = "" And Worksheets("Exe").Cells(i + 1, c_exe_underlying_id) = "" Then
        Exit For
    Else
        If Worksheets("Exe").Cells(i, c_exe_underlying_id) <> "" Then
            If Worksheets("Exe").Cells(i, c_exe_underlying_id) = hshares_underlying_id Then
                
                If Worksheets("Exe").Cells(i, c_exe_qty) <> 0 Then
                    
                    Worksheets("Exe").Activate
                    Worksheets("Exe").Cells(i, c_exe_qty).Select
                    
                    
                    answer = MsgBox("desactivate this line ?", vbYesNo)
                    
                    If answer = vbYes Then
                        tmp_qty = Worksheets("Exe").Cells(i, c_exe_qty)
                        Worksheets("Exe").Cells(i, c_exe_qty) = 0
                        
                        
                        'insertion d un comment avec la qty remplacee
                        Set rng_description = Worksheets("Exe").Cells(i, c_exe_qty)
                        
                        If rng_description.comment Is Nothing Then rng_description.AddComment
                        rng_description.comment.Visible = False
                        rng_description.comment.Text "desactivate " & Date & " org qty=" & tmp_qty
                        
                        
                    End If
                    
                End If
                
            End If
        End If
    End If
Next i

End Sub



Sub prepare_cointrin_for_new_year()

Application.Calculation = xlCalculationManual

Dim i As Long, j As Long, k As Long, m As Long, n As Long

'remonte les donnees du trader
Dim prenom As String, nom As String
prenom = Left(Worksheets("Cointrin").Cells(5, 2), InStr(Worksheets("Cointrin").Cells(5, 2), " ") - 1)
nom = Mid(Worksheets("Cointrin").Cells(5, 2), Len(prenom) + 2)

'remonte trader code, redi plus id
Dim sql_query As String
sql_query = "SELECT system_code, gs_UserID  FROM t_trader WHERE system_first_name=""" & prenom & """ AND system_surname=""" & nom & """"
Dim data_trader As Variant
data_trader = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim trader_code As Integer, trader_r_plus_id As String

If UBound(data_trader, 1) > 0 Then
    trader_code = data_trader(1, 0)
    trader_r_plus_id = data_trader(1, 1)
Else
    Exit Sub
End If


'remonte le trading account
sql_query = "SELECT gs_account_number FROM t_trading_account WHERE system_trader_code=" & trader_code & " AND gs_main_account=TRUE"
Dim extract_account As Variant
extract_account = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

Dim extract_main_trading_account As String

If UBound(extract_account, 1) > 0 Then
    extract_main_trading_account = extract_account(1, 0)
Else
    Exit Sub
End If


'remonte currency
sql_query = "SELECT system_code, system_name FROM t_currency ORDER BY system_code ASC"
Dim extract_currency As Variant
extract_currency = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)
    dim_currency_code = 0
    dim_currency_txt = 1



Dim cointrin_folder As String
    cointrin_folder = StrReverse(Mid(StrReverse(db_cointrin_trades_path), InStr(StrReverse(db_cointrin_trades_path), "\")))

Dim year_txt As String, month_txt As String, day_txt As String, hour_txt As String, minute_txt As String, second_txt As String
    year_txt = CStr(year(Date))
    month_txt = CStr(Month(Date))
        If Len(month_txt) = 1 Then
            month_txt = "0" & month_txt
        End If
    day_txt = CStr(day(Date))
        If Len(day_txt) = 1 Then
            day_txt = "0" & day_txt
        End If
    hour_txt = CStr(Hour(Now))
        If Len(hour_txt) = 1 Then
            hour_txt = "0" & hour_txt
        End If
    minute_txt = CStr(Minute(Now))
        If Len(minute_txt) = 1 Then
            minute_txt = "0" & minute_txt
        End If
    second_txt = CStr(Second(Now))
        If Len(second_txt) = 1 Then
            second_txt = "0" & second_txt
        End If
    
    
    

'backup de la compostion actuel
'FileCopy db_cointrin_trades_path, cointrin_folder & "backup\db_trades_auto_" & year_txt & month_txt & day_txt & "_" & hour_txt & minute_txt & second_txt & ".mdb"

Dim l_cointrin_header As Integer
Dim c_cointrin_product_id As Integer, c_cointrin_currency As Integer, c_cointrin_net_position As Integer, _
    c_cointrin_close_price As Integer, c_cointrin_ytd_pnl_local_gross As Integer, c_cointrin_comm_local As Integer, _
    c_cointrin_ytd_pnl_local_net As Integer, c_cointrin_description As Integer

l_cointrin_header = 8
c_cointrin_product_id = 0

For i = 1 To 250
    If Worksheets("Cointrin").Cells(l_cointrin_header, i) = "Identifier" And c_cointrin_product_id = 0 Then
        c_cointrin_product_id = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "currency" Then
        c_cointrin_currency = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "net_position" Then
        c_cointrin_net_position = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "gva_close_price" Then
        c_cointrin_close_price = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "ytd_pnl_gross" Then
        c_cointrin_ytd_pnl_local_gross = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "comm" Then
        c_cointrin_comm_local = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "ytd_pnl_net_local" Then
        c_cointrin_ytd_pnl_local_net = i
    ElseIf Worksheets("Cointrin").Cells(l_cointrin_header, i) = "gs_description" Then
        c_cointrin_description = i
    End If
Next i


Dim vec_entry() As Variant
    dim_date = 0
    dim_time = 1
    dim_unique_id_trade = 2
    dim_product_id = 3
    dim_net_pos = 4
    dim_close_price = 5
    dim_side = 6
    dim_trading_account = 7
    dim_redi_user = 8
    dim_crncy_code = 9
    dim_trader_code = 10
    
    
    
    

'remonte les donnees de cointrin
Dim l_cointrin_last_line As Integer
Dim tmp_side As String, tmp_currency_code As Variant, msg_currency As String

k = 0
For i = l_cointrin_header + 1 To 10000
    If Worksheets("Cointrin").Cells(i, c_cointrin_product_id) = "" Then
        l_cointrin_last_line = i - 1
        Exit For
    Else
        
        If Worksheets("Cointrin").Cells(i, c_cointrin_net_position) <> 0 Then
            
            'side
            If Worksheets("Cointrin").Cells(i, c_cointrin_net_position).Value < 0 Then
                tmp_side = "S"
            Else
                tmp_side = "B"
            End If
            
            'currency code
            If Worksheets("Cointrin").Cells(i, c_cointrin_currency).Value = "" Then
manual_currency_code:
                msg_currency = "Which code for " & Worksheets("Cointrin").Cells(i, c_cointrin_product_id) & " - " & Worksheets("Cointrin").Cells(i, c_cointrin_description) & vbCrLf & vbCrLf
                For j = 1 To UBound(extract_currency, 1)
                    msg_currency = msg_currency & "[" & extract_currency(j, dim_currency_code) & "] - " & extract_currency(j, dim_currency_txt) & vbCrLf
                Next j
                
                tmp_currency_code = CInt(InputBox(msg_currency, "which code ?", 3))
                
                If IsNumeric(tmp_currency_code) Then
                    tmp_currency_code = CInt(tmp_currency_code)
                Else
                    GoTo manual_currency_code
                End If
                
            Else
                For j = 1 To UBound(extract_currency, 1)
                    If extract_currency(j, dim_currency_txt) = Worksheets("Cointrin").Cells(i, c_cointrin_currency).Value Then
                        tmp_currency_code = extract_currency(j, dim_currency_code)
                        Exit For
                    End If
                Next j
            End If
            
            
            ReDim Preserve vec_entry(k)
            vec_entry(k) = Array(Date, _
                Time, _
                "open_new_year_" & Worksheets("Cointrin").Cells(i, c_cointrin_product_id).Value & "_" & year(Now) & Month(Now) & day(Now) & Hour(Now) & Minute(Now) & Second(Now) & "_" & Round(1000 * Rnd(), 0), _
                Worksheets("Cointrin").Cells(i, c_cointrin_product_id).Value, _
                Worksheets("Cointrin").Cells(i, c_cointrin_net_position).Value, _
                Round(Worksheets("Cointrin").Cells(i, c_cointrin_close_price).Value, 4), _
                tmp_side, _
                extract_main_trading_account, _
                trader_r_plus_id, _
                tmp_currency_code, _
                trader_code)
                
            k = k + 1
        End If
        
    End If
Next i
    

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With


'vide les entree de la db - pour le trader concerne uniquement !!!!
'sql_query = "SELECT [t_trade].[gs_unique_id] "
'    sql_query = sql_query & " FROM t_trade, t_trading_account "
'    sql_query = sql_query & " WHERE [t_trade].[gs_trading_account] = [t_trading_account].[gs_account_number] "
'    sql_query = sql_query & " AND [t_trading_account].[system_trader_code]=" & trader_code


'sql_query = "SELECT [t_trade].[gs_unique_id] "
'    sql_query = sql_query & " FROM t_trade "
'    sql_query = sql_query & " WHERE [t_trade].[gs_trading_account] IN ("
'        sql_query = sql_query & " SELECT gs_account_number "
'        sql_query = sql_query & " FROM t_trading_account"
'        sql_query = sql_query & " WHERE system_trader_code=" & trader_code
'    sql_query = sql_query & ")"

sql_query = "DELETE "
    sql_query = sql_query & " FROM t_trade "
    sql_query = sql_query & " WHERE [t_trade].[gs_trading_account] IN ("
        sql_query = sql_query & " SELECT gs_account_number "
        sql_query = sql_query & " FROM t_trading_account"
        sql_query = sql_query & " WHERE system_trader_code=" & trader_code
    sql_query = sql_query & ")"
    
'insertion des donnee sur les prix de cloture
conn.Execute sql_query



With rst
    
    .ActiveConnection = conn
    .Open "t_trade", LockType:=adLockOptimistic
    
    For i = 0 To UBound(vec_entry, 1)
        
        .AddNew
            
            .fields("gs_date") = vec_entry(i)(dim_date)
            .fields("gs_time") = vec_entry(i)(dim_time)
            .fields("gs_unique_id") = vec_entry(i)(dim_unique_id_trade)
            .fields("gs_security_id") = vec_entry(i)(dim_product_id)
            .fields("gs_exec_qty") = vec_entry(i)(dim_net_pos)
            .fields("gs_exec_price") = vec_entry(i)(dim_close_price)
            .fields("gs_order_qty") = vec_entry(i)(dim_net_pos)
            .fields("gs_side") = vec_entry(i)(dim_side)
            .fields("gs_side_detailed") = vec_entry(i)(dim_side)
            .fields("gs_trading_account") = vec_entry(i)(dim_trading_account)
            .fields("gs_user_id") = vec_entry(i)(dim_redi_user)
            .fields("gs_exec_broker") = ""
            .fields("gs_doneaway") = False
            .fields("gs_close_price") = vec_entry(i)(dim_close_price)
            .fields("system_ytd_pnl_reversal") = 0
            .fields("system_custom_ytd_pnl_reversal") = 0
            .fields("system_position_reversal") = 0
            .fields("system_comm_reversal") = 0
            .fields("system_currency_code") = vec_entry(i)(dim_crncy_code)
            .fields("system_broker_id") = 0
            .fields("system_commission_local_currency") = 0
            .fields("system_trader_code") = vec_entry(i)(dim_trader_code)
            .fields("system_need_update") = False
            .fields("system_exercise") = False
            
        .Update
        
    Next i
    
End With


'mise a zero des valeur dans la feuille cointrin
For i = l_cointrin_header + 1 To l_cointrin_last_line
    Worksheets("Cointrin").Cells(i, c_cointrin_ytd_pnl_local_gross) = 0
    Worksheets("Cointrin").Cells(i, c_cointrin_comm_local) = 0
    Worksheets("Cointrin").Cells(i, c_cointrin_ytd_pnl_local_net) = 0
Next i

End Sub

Sub import_yearly_pnl_from_csv()

Application.Calculation = xlCalculationManual

Dim i As Long, j As Long, k As Long, m As Long, n As Long

Dim year As Integer
    year = 2011

Dim csv_export_file As String
csv_export_file = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\"))) & "export_ytd_pnl_base_" & year & ".csv"



Dim extract_vec_from_csv_file As Variant
extract_vec_from_csv_file = csv_to_array(csv_export_file, 1)


'charge l etat d index database
For i = 27 To 1000 Step 3
    If Worksheets("Index_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        For j = 0 To UBound(extract_vec_from_csv_file, 1)
            If extract_vec_from_csv_file(j)(0) = "Index_Database" Then
                If extract_vec_from_csv_file(j)(1) = Worksheets("Index_Database").Cells(i, 1) Then
                    
                    'inscription de l'agreg pnl
                    Worksheets("Index_Database").Cells(i, 188) = Worksheets("Index_Database").Cells(i, 188) + extract_vec_from_csv_file(j)(3)
                    
                    Exit For
                End If
            End If
        Next j
    End If
Next i


'charge l etat d equity database
For i = 27 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        For j = 0 To UBound(extract_vec_from_csv_file, 1)
            If extract_vec_from_csv_file(j)(0) = "Equity_Database" Then
                If extract_vec_from_csv_file(j)(1) = Worksheets("Equity_Database").Cells(i, 1) Then
                    
                    'inscription de l'agreg pnl
                    Worksheets("Equity_Database").Cells(i, 158) = Worksheets("Equity_Database").Cells(i, 158) + extract_vec_from_csv_file(j)(3)
                    
                    Exit For
                End If
            End If
        Next j
    End If
Next i


End Sub

Sub export_yearly_pnl_to_csv()

Dim year As Integer
    year = 2011

'worksheet / column_product_id / column_ytd_pnl_base / start_line / step
Dim worksheets_to_export As Variant
    worksheets_to_export = Array(Array("Equity_Database", 1, 14, 47, 27, 2), Array("Index_Database", 1, 13, 110, 27, 3))



    dim_export_config_sheet = 0
    dim_export_config_column_product_id = 1
    dim_export_config_column_ytd_pnl_base = 2
    dim_export_config_column_ticker = 3
    dim_export_config_start_line = 4
    dim_export_config_step_between_2_lines = 5
    

Application.Calculation = xlCalculationManual

Dim i As Long, j As Long, k As Long, m As Long, n As Long

Dim c_equity_db_product_id As Integer, c_equity_db_ytd_pnl_base As Integer
    c_equity_db_product_id = 1
    c_equity_db_ytd_pnl_base = 14
    

Dim vec_position_and_ytd_pnl()

k = 0
ReDim Preserve vec_position_and_ytd_pnl(0)
vec_position_and_ytd_pnl(0) = Array("worksheet", "product_id", "ticker", "ytd_pnl_base")
k = k + 1

For i = 0 To UBound(worksheets_to_export, 1)
    For j = worksheets_to_export(i)(dim_export_config_start_line) To 32000 Step worksheets_to_export(i)(dim_export_config_step_between_2_lines)
        If Worksheets(worksheets_to_export(i)(dim_export_config_sheet)).Cells(j, worksheets_to_export(i)(dim_export_config_column_product_id)) = "" Then
            Exit For
        Else
            If IsError(Worksheets(worksheets_to_export(i)(dim_export_config_sheet)).Cells(j, worksheets_to_export(i)(dim_export_config_column_ytd_pnl_base))) = False Then
                If IsNumeric(Worksheets(worksheets_to_export(i)(dim_export_config_sheet)).Cells(j, worksheets_to_export(i)(dim_export_config_column_ytd_pnl_base))) = True Then
                    ReDim Preserve vec_position_and_ytd_pnl(k)
                    vec_position_and_ytd_pnl(k) = Array(worksheets_to_export(i)(dim_export_config_sheet), Worksheets(worksheets_to_export(i)(dim_export_config_sheet)).Cells(j, worksheets_to_export(i)(dim_export_config_column_product_id)).Value, Worksheets(worksheets_to_export(i)(dim_export_config_sheet)).Cells(j, worksheets_to_export(i)(dim_export_config_column_ticker)).Value, Round(Worksheets(worksheets_to_export(i)(dim_export_config_sheet)).Cells(j, worksheets_to_export(i)(dim_export_config_column_ytd_pnl_base)).Value, 3))
                    k = k + 1
                End If
            End If
        End If
    Next j
Next i


'export du csv
Dim csv_export_file As String
csv_export_file = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\"))) & "export_ytd_pnl_base_" & year & ".csv"

debug_test = array_to_csv(vec_position_and_ytd_pnl, csv_export_file)


End Sub


Public Function import_trades_360_portal_into_cointrin(ByVal vec_product As Variant, ByVal file_from_gs_portal As String)

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim sql_query As String


For i = 0 To UBound(vec_product, 1)
    If Left(vec_product(i), 2) <> "P=" Then
        vec_product(i) = "P=" & vec_product(i)
    End If
Next i

Dim gs_file As String
gs_file = Right(file_from_gs_portal, InStr(StrReverse(file_from_gs_portal), "\") - 1)

'monte les données de devises
Dim vec_currency() As Variant
k = 0

For i = 14 To 31
    ReDim Preserve vec_currency(k)
    vec_currency(k) = Array(Workbooks("Kronos.xls").Worksheets("Parametres").Cells(i, 1).Value, Workbooks("Kronos.xls").Worksheets("Parametres").Cells(i, 5).Value)
    k = k + 1
Next i


'remonte trader code
Dim trader_redi_txt As String, tmp_trader_code As Integer
trader_redi_txt = Workbooks("Kronos.xls").Worksheets("FORMAT2").Cells(7, 21).Value
sql_query = "SELECT system_code FROM t_trader WHERE gs_UserID=""" & trader_redi_txt & """"
Dim extract_trader_code As Variant
extract_trader_code = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_trader_code, 1) = 0 Then
    MsgBox ("problem with trader in DB. Exit")
    Exit Function
Else
    tmp_trader_code = extract_trader_code(1, 0)
End If



'detection des colonnes du fichier de GS
Dim sheet_gs_file As String
    sheet_gs_file = "Sheet1"
Dim l_gs_file_header As Integer
    l_gs_file_header = 3
Dim c_gs_file_account As Integer, c_gs_file_trade_type As Integer, c_gs_file_side As Integer, c_gs_file_qty As Integer, _
    c_gs_file_currency As Integer, c_gs_file_price As Integer, c_gs_file_comm As Integer, c_gs_file_product_id As Integer
    
Call open_file(file_from_gs_portal, True)

For i = 1 To 100
    If Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "" Then
        Exit For
    Else
        If Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Account" Then
            c_gs_file_account = i
        ElseIf Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Trade Type" Then
            c_gs_file_trade_type = i
        ElseIf Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Buy/ Sell" Then
            c_gs_file_side = i
        ElseIf Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Quantity" Then
            c_gs_file_qty = i
        ElseIf Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Contract Currency" Then
            c_gs_file_currency = i
        ElseIf Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Trade Price" Then
            c_gs_file_price = i
        ElseIf Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Clearing + Execution Commission (local)" Then
            c_gs_file_comm = i
        ElseIf Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(l_gs_file_header, i) = "Product ID" Then
            c_gs_file_product_id = i
        End If
    End If
Next i


'recupere les trades des produits contenu dans le vec_prodcut
Dim vec_trades() As Variant

Dim dim_trades_product_id As Integer, dim_trades_qty As Integer, dim_trades_price As Integer, dim_trades_side As Integer, _
    dim_trades_account As Integer, dim_trades_crncy_code As Integer, dim_trades_comm_local As Integer

dim_trades_product_id = 0
dim_trades_qty = 1
dim_trades_price = 2
dim_trades_side = 3
dim_trades_account = 4
dim_trades_crncy_code = 5
dim_trades_comm_local = 6

Dim tmp_account As String, tmp_side As String, tmp_qty As Double, tmp_product_id As String, tmp_price As Double, _
    tmp_currency_code As Integer, tmp_comm_local As Double

k = 0
For i = l_gs_file_header + 1 To 30000
    If Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_product_id) = "" Then
        Exit For
    Else
        For j = 0 To UBound(vec_product, 1)
            If Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_trade_type) = "New Trade" Then
                If "P=" & Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_product_id) = vec_product(j) Then
                    
                    tmp_product_id = vec_product(j)
                    tmp_account = Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_account)
                    tmp_side = Left(Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_side), 1)
                    tmp_qty = Abs(CDbl(Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_qty)))
                    
                    If UCase(tmp_side) = "B" Then
                        
                    ElseIf UCase(tmp_side) = "S" Then
                        tmp_qty = -tmp_qty
                    End If
                    
                    tmp_price = Abs(CDbl(Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_price)))
                    
                    For m = 0 To UBound(vec_currency, 1)
                        If vec_currency(m)(0) = UCase(Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_currency)) Then
                            tmp_currency_code = vec_currency(m)(1)
                            Exit For
                        End If
                    Next m
                    
                    tmp_comm_local = Abs(CDbl(Workbooks(gs_file).Worksheets(sheet_gs_file).Cells(i, c_gs_file_comm)))
                    
                    
                    ReDim Preserve vec_trades(k)
                    vec_trades(k) = Array(tmp_product_id, tmp_qty, tmp_price, tmp_side, tmp_account, tmp_currency_code, tmp_comm_local)
                    
                    k = k + 1
                    
                End If
            End If
        Next j
    End If
Next i

Workbooks(gs_file).Close False

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset


If k > 0 Then
    
    Dim insert_status As Boolean
    's'assure que les valeurs sont ouvertes dans bridge
    For i = 0 To UBound(vec_product, 1)
        insert_status = open_new_entry_in_cointrin_bridge(vec_product(i))
    Next i
    
    With conn
        .Provider = "Microsoft.JET.OLEDB.4.0"
        .Open db_cointrin_trades_path
    End With
    
    'insertion des donnnes dans cointrin
    Dim msg_msgbox As String
    msg_msgbox = ""
    
    With rst
        .ActiveConnection = conn
        .Open "t_trade", LockType:=adLockOptimistic
        
        For i = 0 To UBound(vec_trades, 1)
            .AddNew
                
                .fields("gs_date") = Date
                .fields("gs_time") = Time
                
                .fields("gs_unique_id") = "import_360_gs_portal_" & vec_trades(i)(dim_trades_product_id) & "_" & year(Date) & Month(Date) & day(Date) & "_" & Hour(Time) & Minute(Time) & Second(Time) & "_" & Round(10000 * Rnd(), 0)
                
                .fields("gs_security_id") = vec_trades(i)(dim_trades_product_id)
                
                .fields("gs_order_qty") = vec_trades(i)(dim_trades_qty)
                .fields("gs_exec_qty") = vec_trades(i)(dim_trades_qty)
                .fields("gs_exec_price") = vec_trades(i)(dim_trades_price)
                
                .fields("gs_side") = vec_trades(i)(dim_trades_side)
                .fields("gs_side_detailed") = vec_trades(i)(dim_trades_side)
                
                .fields("gs_trading_account") = vec_trades(i)(dim_trades_account)
                .fields("gs_user_id") = trader_redi_txt
                .fields("system_trader_code") = tmp_trader_code
                
                .fields("system_broker_id") = 2
                .fields("gs_exec_broker") = "GSCO"
                
                .fields("system_currency_code") = vec_trades(i)(dim_trades_crncy_code)
                
                .fields("system_commission_local_currency") = Abs(vec_trades(i)(dim_trades_comm_local))
                
                msg_msgbox = msg_msgbox & vec_trades(i)(dim_trades_product_id) & ": " & vec_trades(i)(dim_trades_qty) & " @ " & vec_trades(i)(dim_trades_price) & vbCrLf
                
                
            .Update
        Next i
        
        
        MsgBox (msg_msgbox)
        
    End With
    
    
End If

Application.Calculation = xlCalculationAutomatic

End Function


Public Sub open_entry_in_contrin_bridge()

If ActiveSheet.name = "Open" Then
    If ActiveCell.row > 25 Then
        If Worksheets("Open").Cells(ActiveCell.row, 2).Value <> 0 Then
            Call open_new_entry_in_cointrin_bridge(Worksheets("Open").Cells(ActiveCell.row, 2).Value)
        Else
            Call open_new_entry_in_cointrin_bridge(Worksheets("Open").Cells(ActiveCell.row, 1).Value)
        End If
    End If
End If

End Sub


Public Function open_new_entry_in_cointrin_bridge(ByVal gs_product_id As String) As Boolean

open_new_entry_in_cointrin_bridge = False

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim sql_query As String


sql_query = "SELECT gs_description FROM t_bridge WHERE gs_id=""" & gs_product_id & """"
Dim extract_bridge As Variant
extract_bridge = ADOImportFromAccessTable(db_cointrin_trades_path, , sql_query)

If UBound(extract_bridge, 1) > 0 Then
    MsgBox ("Already in bridge")
    Exit Function
End If

'passe a travers les differentes tables pour tenter de mettre la main sur l'underlying id
Dim l_folio_view_header As Integer
l_folio_view_header = 10

Dim folio_view As Variant
folio_view = Array(Array("Futures_Folio", "IDENTIFIER", "UNDERLYER PRODUCT ID", "UNDERLYING PRODUCT DESCRIPTION", 2), Array("Equities_Folio", "IDENTIFIER", "IDENTIFIER", "DESCRIPTION", 1), Array("Options_Folio", "Identifier_O", "UNDERLYER PRODUCT ID", "DESCRIPTION", 3))

Dim dim_folio_view_sheet As Integer, dim_folio_view_product As Integer, dim_folio_view_underlying As Integer, _
    dim_folio_view_name As Integer, dim_folio_view_instrument_id As Integer

dim_folio_view_sheet = 0
dim_folio_view_product = 1
dim_folio_view_underlying = 2
dim_folio_view_name = 3
dim_folio_view_instrument_id = 4


Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset

With conn
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .Open db_cointrin_trades_path
End With


Dim tmp_underlying_id As String, tmp_product_description As String

For i = 0 To UBound(folio_view, 1)
    For j = 1 To 250
        If Worksheets(folio_view(i)(dim_folio_view_sheet)).Cells(l_folio_view_header, j) = folio_view(i)(dim_folio_view_product) Then
            
            For m = l_folio_view_header + 2 To 5000
                
                If Worksheets(folio_view(i)(dim_folio_view_sheet)).Cells(m, j) = "" Then
                    GoTo bypass_sheet
                Else
                    If Worksheets(folio_view(i)(dim_folio_view_sheet)).Cells(m, j) = gs_product_id Then
                        
                        'underlying_id
                        For n = 1 To 250
                            If Worksheets(folio_view(i)(dim_folio_view_sheet)).Cells(l_folio_view_header, n) = folio_view(i)(dim_folio_view_underlying) Then
                                tmp_underlying_id = Worksheets(folio_view(i)(dim_folio_view_sheet)).Cells(m, n)
                                Exit For
                            End If
                        Next n
                        
                        'product description
                        For n = 1 To 250
                            If Worksheets(folio_view(i)(dim_folio_view_sheet)).Cells(l_folio_view_header, n) = folio_view(i)(dim_folio_view_name) Then
                                tmp_product_description = Worksheets(folio_view(i)(dim_folio_view_sheet)).Cells(m, n)
                                Exit For
                            End If
                        Next n
                        
                        
                        'setup de la nouvelle entree dans le bridge
                        With rst
                            .ActiveConnection = conn
                            .Open "t_bridge", LockType:=adLockOptimistic
                            
                            .AddNew
                                
                                .fields("gs_id") = gs_product_id
                                .fields("gs_underlying_id") = tmp_underlying_id
                                .fields("gs_description") = tmp_product_description
                                .fields("system_instrument_id") = folio_view(i)(dim_folio_view_instrument_id)
                                
                            
                            .Update
                            
                        End With
                        
                        Exit Function
                        
                    End If
                End If
            Next m
            
        End If
    Next j
bypass_sheet:
Next i


End Function


Public Function get_tickers_with_next_dividend(Optional ByVal next_days As Integer = 5, Optional ByVal show_msgbox As Boolean = True) As Variant

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer

Dim date_tmp As Date


Dim l_equity_db_header As Integer, c_equity_db_ticker As Integer, c_equity_db_dvd_date As Integer, c_equity_db_valeur_eur As Integer, _
    c_equity_db_option_amount As Integer

l_equity_db_header = 25
c_equity_db_valeur_eur = 5
c_equity_db_option_amount = 37
c_equity_db_ticker = 47
c_equity_db_dvd_date = 121


Dim vec_ticker_and_div_date() As Variant
    Dim dim_ticker As Integer, dim_date As Integer, dim_date_dbl As Integer, dim_valeur_eur As Integer, dim_option_amount As Integer
    
    dim_ticker = 0
    dim_date = 1
    dim_date_dbl = 2
    dim_valeur_eur = 3
    dim_option_amount = 4
    


k = 0
Dim tmp_valeur_eur As Double
For i = l_equity_db_header + 2 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        If IsError(Worksheets("Equity_Database").Cells(i, c_equity_db_dvd_date)) = False Then
            date_tmp = Worksheets("Equity_Database").Cells(i, c_equity_db_dvd_date)
            
            If date_tmp >= Date And date_tmp <= Date + next_days Then
                
                tmp_valeur_eur = 0
                
                If IsError(Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur)) = False Then
                    If IsNumeric(Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur)) Then
                        tmp_valeur_eur = Worksheets("Equity_Database").Cells(i, c_equity_db_valeur_eur)
                    End If
                End If
                
                ReDim Preserve vec_ticker_and_div_date(k)
                vec_ticker_and_div_date(k) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_ticker).Value, date_tmp, CDbl(date_tmp), tmp_valeur_eur, Worksheets("Equity_Database").Cells(i, c_equity_db_option_amount).Value)
                k = k + 1
            End If
            
        End If
    End If
Next i

Dim tmp_min As Double
Dim min_pos As Long
Dim tmp_vec As Variant

'tri date
For i = 0 To UBound(vec_ticker_and_div_date, 1)
    tmp_min = vec_ticker_and_div_date(i)(dim_date_dbl)
    min_pos = i
    
    For j = i + 1 To UBound(vec_ticker_and_div_date, 1)
        If vec_ticker_and_div_date(j)(dim_date_dbl) < tmp_min Then
            min_pos = j
            tmp_min = vec_ticker_and_div_date(j)(dim_date_dbl)
        End If
    Next j
    
    If i <> min_pos Then
        tmp_vec = vec_ticker_and_div_date(i)
        vec_ticker_and_div_date(i) = vec_ticker_and_div_date(min_pos)
        vec_ticker_and_div_date(min_pos) = tmp_vec
    End If
    
Next i

'tri ticker
Dim tmp_min_txt As String
Dim tmp_date_min As Double
For i = 0 To UBound(vec_ticker_and_div_date, 1)
    tmp_date_min = vec_ticker_and_div_date(i)(dim_date_dbl)
    tmp_min_txt = vec_ticker_and_div_date(i)(dim_ticker)
    min_pos = i
    
    For j = i + 1 To UBound(vec_ticker_and_div_date, 1)
        If vec_ticker_and_div_date(j)(dim_date_dbl) = tmp_date_min Then
            If vec_ticker_and_div_date(j)(dim_ticker) < tmp_min_txt Then
                tmp_min_txt = vec_ticker_and_div_date(j)(dim_ticker)
                min_pos = j
            End If
        End If
    Next j
    
    If min_pos <> i Then
        tmp_vec = vec_ticker_and_div_date(i)
        vec_ticker_and_div_date(i) = vec_ticker_and_div_date(min_pos)
        vec_ticker_and_div_date(min_pos) = tmp_vec
    End If
    
Next i


date_tmp = Date - 5
Dim msg_msgbox As String
msg_msgbox = ""
If k > 0 Then
    
    
    
    For i = 0 To UBound(vec_ticker_and_div_date, 1)
        
        If date_tmp <> vec_ticker_and_div_date(i)(dim_date) Then
            'new header
            msg_msgbox = msg_msgbox & vbCrLf
            msg_msgbox = msg_msgbox & vbCrLf
            msg_msgbox = msg_msgbox & "[" & vec_ticker_and_div_date(i)(dim_date) & "]" & vbCrLf
            
            date_tmp = vec_ticker_and_div_date(i)(dim_date)
            
        End If
        
        msg_msgbox = msg_msgbox & vec_ticker_and_div_date(i)(dim_ticker)
        
            If vec_ticker_and_div_date(i)(dim_valeur_eur) <> 0 Then
                msg_msgbox = msg_msgbox & " (" & Round(1000 * vec_ticker_and_div_date(i)(dim_valeur_eur), 0) & ")"
            End If
            
            
            If vec_ticker_and_div_date(i)(dim_option_amount) > 0 Then
                msg_msgbox = msg_msgbox & " *"
            End If
            
        
        msg_msgbox = msg_msgbox & vbCrLf
    Next i

    get_tickers_with_next_dividend = vec_ticker_and_div_date
    
Else
    msg_msgbox = "no div for the next " & next_days & " days."
    get_tickers_with_next_dividend = Empty
End If

If show_msgbox = True Then
    MsgBox (msg_msgbox)
End If


End Function


Public Sub update_option_in_open(ByVal product_id As String, Optional ByVal qty As Variant, Optional ByVal price As Variant)

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer

If IsMissing(qty) Or IsMissing(price) Then
    For i = 8 To 32000
        If Worksheets("Cointrin").Cells(i, 1) = "" Then
            MsgBox ("not found in cointrin. Unable to sync with open")
            Exit Sub
        Else
            If Worksheets("Cointrin").Cells(i, 1) = product_id Then
                qty = Worksheets("Cointrin").Cells(i, 5)
                price = Worksheets("Cointrin").Cells(i, 6)
                Exit For
            End If
        End If
    Next i
End If

For i = 25 To 5000
    If Worksheets("Open").Cells(i, 1) = "" And Worksheets("Open").Cells(i + 1, 1) = "" And Worksheets("Open").Cells(i + 2, 1) = "" And Worksheets("Open").Cells(i + 3, 1) = "" And Worksheets("Open").Cells(i + 4, 1) = "" And Worksheets("Open").Cells(i + 5, 1) = "" Then
    Else
        If Worksheets("Open").Cells(i, 2) = product_id Then
            Worksheets("Open").Cells(i, 44) = price
            Worksheets("Open").Cells(i, 45) = qty
            Exit Sub
        End If
    End If
Next i


End Sub


Public Sub prepare_emsx_oi()

Dim oJSON As New JSONLib

Dim prefix_src_new_line As String
    prefix_src_new_line = "*** AUTO-GENERATED ***"

Dim emsx_csv_filename As String
    emsx_csv_filename = "00emsx_oi.csv"

Dim ready_color As Integer
ready_color = 42

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim vec_trades_db_redi() As Variant

Dim vec_trades() As Variant
Dim tmp_vec_trade() As Variant
k = 0

Dim color_lines() As Variant

'check le haut du fichier aussi
Dim tmp_qty As Double, tmp_price As Double, tmp_stop As Variant, tmp_vec_tag() As Variant, tmp_group_id As Variant


For i = 14 To 70
    If Worksheets("FORMAT2").Cells(i, 1) = "" And Worksheets("FORMAT2").Cells(i + 1, 1) = "" And Worksheets("FORMAT2").Cells(i + 2, 1) = "" Then
        Exit For
    Else
        
        If Worksheets("FORMAT2").Cells(i, 1).Interior.ColorIndex = ready_color Then
            If IsError(Worksheets("FORMAT2").Cells(i, 6)) = False And IsError(Worksheets("FORMAT2").Cells(i, 7)) = False Then
                
                If IsNumeric(Worksheets("FORMAT2").Cells(i, 6)) And IsNumeric(Worksheets("FORMAT2").Cells(i, 7)) Then
                    
                    If Worksheets("FORMAT2").Cells(i, 6) > 0 Then
                    
                        For j = 1 To 9
                            ReDim Preserve tmp_vec_trade(j - 1)
                            tmp_vec_trade(j - 1) = Worksheets("FORMAT2").Cells(i, j).Value
                        Next j
                        
                        ReDim Preserve vec_trades(k)
                        vec_trades(k) = tmp_vec_trade
                        
                        ReDim Preserve color_lines(k)
                        color_lines(k) = i
                        
                        
                        
                        '################ DB REDI ############$
                        
                        
                        If UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "C" Then
                            tmp_qty = Worksheets("FORMAT2").Cells(i, c_format2_qty)
                        ElseIf UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "H" Then
                            tmp_qty = -Worksheets("FORMAT2").Cells(i, c_format2_qty)
                        End If
                        
                        tmp_price = Worksheets("FORMAT2").Cells(i, c_format2_price)
                        
                        
                        If UCase(Worksheets("FORMAT2").Cells(i, c_format2_time_limit)) = "STP" Or UCase(Worksheets("FORMAT2").Cells(i, c_format2_time_limit)) = "STOP" Then
                            tmp_stop = Worksheets("FORMAT2").Cells(i, c_format2_price).Value
                        Else
                            tmp_stop = Empty
                        End If
                        
                        
                        m = 0
                        ReDim Preserve tmp_vec_tag(m)
                        tmp_vec_tag(m) = "EMSX"
                        m = m + 1
                        
                        ReDim Preserve tmp_vec_tag(m)
                        If Worksheets("FORMAT2").Cells(i, c_format2_source) <> prefix_src_new_line Then
                            tmp_vec_tag(m) = Replace(Worksheets("FORMAT2").Cells(i, c_format2_source).Value, " ", "_")
                        Else
                            'remonte
                            For j = i - 1 To 14 Step -1
                                If Worksheets("FORMAT2").Cells(i, c_format2_source) <> prefix_src_new_line Then
                                    tmp_vec_tag(m) = Replace(Worksheets("FORMAT2").Cells(j, c_format2_source).Value, " ", "_")
                                    Exit For
                                End If
                            Next j
                        End If
                        m = m + 1
                        
                        
                        
                        ReDim Preserve tmp_vec_tag(m)
                        If Worksheets("FORMAT2").Cells(i, c_format2_source) <> prefix_src_new_line Then
                            tmp_vec_tag(m) = "base"
                            tmp_last_base_line = i
                        Else
                            tmp_vec_tag(m) = Worksheets("FORMAT2").Cells(i, c_format2_s3).Value
                        End If
                        m = m + 1
                        
                        
                        For j = c_format2_s3 To c_format2_r3
                            If Worksheets("FORMAT2").Cells(i, j) <> "" And IsNumeric(Worksheets("FORMAT2").Cells(i, j)) Then
                                
                                If Round(Worksheets("FORMAT2").Cells(i, j), 2) = Round(Worksheets("FORMAT2").Cells(i, c_format2_price), 2) Then
                                    ReDim Preserve tmp_vec_tag(m)
                                    tmp_vec_tag(m) = Worksheets("FORMAT2").Cells(l_format2_header, j).Value
                                    m = m + 1
                                End If
                            End If
                        Next j
                        
                        
                        
                        If Cells(i, c_format2_ticker).Value <> tmp_last_ticker Then
                            tmp_group_id = generate_group_id_trade
                        Else
                            
                            If Cells(i, c_format2_source) <> prefix_src_new_line Then
                                tmp_group_id = generate_group_id_trade
                            Else
                                's agit - il vraiment d un group existant
                                For j = i - 1 To 13 Step -1
                                    If Cells(j, c_format2_source) <> prefix_src_new_line Then
                                        
                                        If j <> tmp_last_base_line Then
                                            tmp_last_base_line = j
                                            tmp_group_id = generate_group_id_trade
                                        End If
                                        
                                        Exit For
                                    End If
                                Next j
                            End If
                            
                        End If
                        
                        
                        ReDim Preserve vec_trades_db_redi(k)
                        vec_trades_db_redi(k) = Array(Worksheets("FORMAT2").Cells(i, c_format2_ticker).Value, tmp_qty, tmp_price, tmp_stop, tmp_group_id, oJSON.toString(tmp_vec_tag))
                        
                        
                        k = k + 1
                    
                    End If
                    
                End If
                
            End If
        End If
    End If
Next i


For i = 100 To 5000
    If Worksheets("FORMAT2").Cells(i, 1) = "" Then
        Exit For
    Else
        
        If Worksheets("FORMAT2").Cells(i, 1).Interior.ColorIndex = ready_color Then
            If IsError(Worksheets("FORMAT2").Cells(i, 6)) = False And IsError(Worksheets("FORMAT2").Cells(i, 7)) = False Then
                
                If IsNumeric(Worksheets("FORMAT2").Cells(i, 6)) And IsNumeric(Worksheets("FORMAT2").Cells(i, 7)) Then
                    
                    If Worksheets("FORMAT2").Cells(i, 6) > 0 Then
                    
                        For j = 1 To 9
                            ReDim Preserve tmp_vec_trade(j - 1)
                            tmp_vec_trade(j - 1) = Worksheets("FORMAT2").Cells(i, j).Value
                        Next j
                        
                        ReDim Preserve vec_trades(k)
                        vec_trades(k) = tmp_vec_trade
                        
                        ReDim Preserve color_lines(k)
                        color_lines(k) = i
                        
                        
                        
                        '################ DB REDI ############$
                        
                        
                        If UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "B" Or UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "C" Then
                            tmp_qty = Worksheets("FORMAT2").Cells(i, c_format2_qty)
                        ElseIf UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "S" Or UCase(Left(Worksheets("FORMAT2").Cells(i, c_format2_side), 1)) = "H" Then
                            tmp_qty = -Worksheets("FORMAT2").Cells(i, c_format2_qty)
                        End If
                        
                        tmp_price = Worksheets("FORMAT2").Cells(i, c_format2_price)
                        
                        
                        If UCase(Worksheets("FORMAT2").Cells(i, c_format2_time_limit)) = "STP" Or UCase(Worksheets("FORMAT2").Cells(i, c_format2_time_limit)) = "STOP" Then
                            tmp_stop = Worksheets("FORMAT2").Cells(i, c_format2_price).Value
                        Else
                            tmp_stop = Empty
                        End If
                        
                        
                        
                        m = 0
                        ReDim Preserve tmp_vec_tag(m)
                        tmp_vec_tag(m) = "EMSX"
                        m = m + 1
                        
                        ReDim Preserve tmp_vec_tag(m)
                        If Worksheets("FORMAT2").Cells(i, c_format2_source) <> prefix_src_new_line Then
                            tmp_vec_tag(m) = Replace(Worksheets("FORMAT2").Cells(i, c_format2_source).Value, " ", "_")
                        Else
                            'remonte
                            For j = i - 1 To l_format2_header + 1 Step -1
                                If Worksheets("FORMAT2").Cells(i, c_format2_source) <> prefix_src_new_line Then
                                    tmp_vec_tag(m) = Replace(Worksheets("FORMAT2").Cells(j, c_format2_source).Value, " ", "_")
                                    Exit For
                                End If
                            Next j
                        End If
                        m = m + 1
                        
                        
                        
                        ReDim Preserve tmp_vec_tag(m)
                        If Worksheets("FORMAT2").Cells(i, c_format2_source) <> prefix_src_new_line Then
                            tmp_vec_tag(m) = "base"
                            tmp_last_base_line = i
                        Else
                            tmp_vec_tag(m) = Worksheets("FORMAT2").Cells(i, c_format2_s3).Value
                        End If
                        m = m + 1
                        
                        
                        For j = c_format2_s3 To c_format2_r3
                            If Worksheets("FORMAT2").Cells(i, j) <> "" And IsNumeric(Worksheets("FORMAT2").Cells(i, j)) Then
                                
                                If Round(Worksheets("FORMAT2").Cells(i, j), 2) = Round(Worksheets("FORMAT2").Cells(i, c_format2_price), 2) Then
                                    ReDim Preserve tmp_vec_tag(m)
                                    tmp_vec_tag(m) = Worksheets("FORMAT2").Cells(l_format2_header, j).Value
                                    m = m + 1
                                End If
                            End If
                        Next j
                        
                        
                        
                        If Cells(i, c_format2_ticker).Value <> tmp_last_ticker Then
                            tmp_group_id = generate_group_id_trade
                        Else
                            
                            If Cells(i, c_format2_source) <> prefix_src_new_line Then
                                tmp_group_id = generate_group_id_trade
                            Else
                                's agit - il vraiment d un group existant
                                For j = i - 1 To l_format2_header + 1 Step -1
                                    If Cells(j, c_format2_source) <> prefix_src_new_line Then
                                        
                                        If j <> tmp_last_base_line Then
                                            tmp_last_base_line = j
                                            tmp_group_id = generate_group_id_trade
                                        End If
                                        
                                        Exit For
                                    End If
                                Next j
                            End If
                            
                        End If
                        
                        
                        ReDim Preserve vec_trades_db_redi(k)
                        vec_trades_db_redi(k) = Array(Worksheets("FORMAT2").Cells(i, c_format2_ticker).Value, tmp_qty, tmp_price, tmp_stop, tmp_group_id, oJSON.toString(tmp_vec_tag))
                        
                        
                        k = k + 1
                    
                    End If
                    
                End If
                
            End If
        End If
    End If
Next i

If k > 0 Then
    debug_test = array_to_csv(vec_trades, StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\"))) & emsx_csv_filename, ",")
    
    Call moulinette_inject_EMSX_csv_into_xls_trades(vec_trades_db_redi)
    
    For i = 0 To UBound(color_lines, 1)
        rows(color_lines(i)).Interior.ColorIndex = xlNone
    Next i
    
Else
    MsgBox ("nothing found !")
    Exit Sub
End If

End Sub
