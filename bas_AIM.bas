Attribute VB_Name = "bas_AIM"
'headers
Public Const l_header_aim_view As Integer = 10

Public Const aim_view_eod As String = "AIM_EOD_JSO"
Public Const aim_view_equities As String = "AIM_Equities_JSO"
Public Const aim_view_futures As String = "AIM_Futures_JSO"
Public Const aim_view_options As String = "AIM_Options_JSO"

Public Const sheet_index_db_multi_accounts As String = "Index_DB_multi_accounts"

Public Const l_header_internal_open As Integer = 25
    Public Const c_internal_open_underlying_id As Integer = 1
    Public Const c_internal_open_product_id As Integer = 2
    Public Const c_internal_open_ticker_underying As Integer = 104
    Public Const c_internal_open_ticker_option As Integer = 105
Public Const l_header_internal_db_equity As Integer = 25
Public Const l_header_internal_db_index As Integer = 25


Public Enum aim_view_code
    EOD = 0
    equities = 1
    futures = 2
    Options = 3
End Enum

Public Enum aim_instrument_type
    equity = 0
    future = 1
    option_equity = 2
    option_index = 3
    index = 4
End Enum

Public Enum output_format_rtd
    vec_with_header = 0
    vec_without_header = 1
    range_with_header = 2
    range_without_header = 3
End Enum


Public Enum aim_mode
    manual_import_mav_excel_export = 0
    rtd_auto_db_pictet = 1
End Enum




Public Const c_header_aim_eod_product_id As String = "identifier"
Public Const c_header_aim_eod_underlying_id As String = "under_product_id"
Public Const c_header_aim_eod_instrument_type As String = "pos_inst_type"
Public Const c_header_aim_eod_ticker As String = "bby_code"
Public Const c_header_aim_eod_currency As String = "ccy"
Public Const c_header_aim_eod_close_qty As String = "qty_yesterday_close"
Public Const c_header_aim_eod_close_price As String = "yesterday_price"
Public Const c_header_aim_eod_local_close_ytd_pnl_gross As String = "ytd_pnl_local_gross"
Public Const c_header_aim_eod_local_close_commission_broker As String = "tra_cost_broker"
Public Const c_header_aim_eod_local_close_commission_total As String = "tra_cost_total"
Public Const c_header_aim_eod_local_close_dividend As String = "ytd_dividend"
Public Const c_header_aim_eod_local_close_ytd_pnl_net As String = "ytd_pnl_local_net"
Public Const c_header_aim_eod_aim_account As String = "aim_account"
Public Const c_header_aim_eod_buidt As String = "CUSTOM_buidt"
Public Const c_header_aim_eod_buidt_underlying As String = "CUSTOM_buidt_underlying"
Public Const c_header_aim_eod_buidt_close_qty As String = "CUSTOM_buidt_close_qty"
Public Const c_header_aim_eod_buidt_close_price As String = "CUSTOM_buidt_close_price"
Public Const c_header_aim_eod_built_local_close_ytd_pnl_net As String = "CUSTOM_buidt_ytd_pnl_local_net"
Public Const c_header_aim_eod_built_close_commission_total As String = "CUSTOM_tra_cost_total"
Public Const c_header_aim_eod_aim_prime_broker As String = "aim_prime_broker"
Public Const c_header_aim_eod_accpbpid As String = "CUSTOM_accpbpid"
Public Const c_header_aim_eod_accpbpid_underlying As String = "CUSTOM_accpbpid_underlying"
Public Const c_header_aim_eod_accpbpid_close_qty As String = "CUSTOM_accpbpid_close_qty"
Public Const c_header_aim_eod_accpbpid_close_price As String = "CUSTOM_accpbpid_close_price"
Public Const c_header_aim_eod_accpbpid_local_close_ytd_pnl As String = "CUSTOM_accpbpid_ytd_pnl_local_net"


Public Const c_header_aim_equities_product_id As String = "identifier"
Public Const c_header_aim_equities_ticker As String = "bby_code"
Public Const c_header_aim_equities_description As String = "description"
Public Const c_header_aim_equities_currency As String = "ccy"
Public Const c_header_aim_equities_current_qty As String = "qty_current"
Public Const c_header_aim_equities_close_qty As String = "qty_yesterday_close"
Public Const c_header_aim_equities_close_price As String = "vendor_close_price"
Public Const c_header_aim_equities_local_intraday_net_trading_cash_flow As String = "net_cash_local_with_comm"
Public Const c_header_aim_equities_local_intraday_commission As String = "intraday_commission_local"
Public Const c_header_aim_equities_local_intraday_dividend As String = "intraday_dividend"
Public Const c_header_aim_equities_aim_account As String = "aim_account"
Public Const c_header_aim_equities_aim_prime_broker As String = "aim_prime_broker"
Public Const c_header_aim_equities_accpbpid As String = "CUSTOM_accpbpid"
Public Const c_header_aim_equities_accpbpid_underlying As String = "CUSTOM_accpbpid_underlying"
Public Const c_header_aim_equities_accpbpid_current_qty As String = "CUSTOM_accpbpid_qty_current"
Public Const c_header_aim_equities_accpbpid_close_qty As String = "CUSTOM_accpbpid_qty_yesterday_close"
Public Const c_header_aim_equities_accpbpid_close_price As String = "CUSTOM_accpbpid_vendor_close_price"
Public Const c_header_aim_equities_accpbpid_local_intraday_net_trading_cash_flow As String = "CUSTOM_accpbpid_net_cash_local_with_comm"


Public Const c_header_aim_futures_product_id As String = "identifier"
Public Const c_header_aim_futures_underlying_id As String = "under_product_id"
Public Const c_header_aim_futures_ticker As String = "bby_code"
Public Const c_header_aim_futures_currency As String = "ccy"
Public Const c_header_aim_futures_current_qty As String = "qty_current"
Public Const c_header_aim_futures_close_qty As String = "qty_yesterday_close"
Public Const c_header_aim_futures_close_price As String = "yesterday_close_local"
Public Const c_header_aim_futures_local_intraday_net_trading_cash_flow As String = "net_cash_local_with_comm"
Public Const c_header_aim_futures_local_intraday_commission As String = "intraday_commission_local"
Public Const c_header_aim_futures_aim_account As String = "aim_account"
Public Const c_header_aim_futures_buidt As String = "CUSTOM_buidt"
Public Const c_header_aim_futures_buidt_underlying As String = "CUSTOM_buidt_underlying"
Public Const c_header_aim_futures_buidt_current_qty As String = "CUSTOM_buidt_current_qty"
Public Const c_header_aim_futures_buidt_close_qty As String = "CUSTOM_buidt_close_qty"
Public Const c_header_aim_futures_buidt_close_price As String = "CUSTOM_buidt_close_price"
Public Const c_header_aim_futures_buidt_local_intrady_net_trading_cash_flow As String = "CUSTOM_buidt_net_cash_local_with_comm"
Public Const c_header_aim_futures_buidt_local_intraday_commission As String = "CUSTOM_intraday_commission_local"
Public Const c_header_aim_futures_aim_prime_broker As String = "aim_prime_broker"
Public Const c_header_aim_futures_accpbpid As String = "CUSTOM_accpbpid"
Public Const c_header_aim_futures_accpbpid_underlying As String = "CUSTOM_accpbpid_underlying"
Public Const c_header_aim_futures_accpbpid_current_qty As String = "CUSTOM_accpbpid_qty_current"
Public Const c_header_aim_futures_accpbpid_close_qty As String = "CUSTOM_accpbpid_qty_yesterday_close"
Public Const c_header_aim_futures_accpbpid_close_price As String = "CUSTOM_accpbpid_yesterday_close_local"
Public Const c_header_aim_futures_accpbpid_local_intraday_net_trading_cash_flow As String = "CUSTOM_accpbpid_net_cash_local_with_comm"


Public Const c_header_aim_options_product_id As String = "identifier"
Public Const c_header_aim_options_underlying_id As String = "under_product_id"
Public Const c_header_aim_options_ticker_with_yellow_key As String = "bby_code"
Public Const c_header_aim_options_ticker_without_yellow_key As String = "ticker"
Public Const c_header_aim_options_investment_class As String = "investment_class"
Public Const c_header_aim_options_currency As String = "ccy"
Public Const c_header_aim_options_current_qty As String = "qty_current"
Public Const c_header_aim_options_close_qty As String = "qty_yesterday_close"
Public Const c_header_aim_options_close_price_mid As String = "yesterday_mid_price"
Public Const c_header_aim_options_local_intraday_net_trading_cash_flow As String = "net_cash_local_with_comm"
Public Const c_header_aim_options_local_intraday_commission As String = "intraday_commission_local"
Public Const c_header_aim_options_last_transaction_code As String = "transaction_code"
Public Const c_header_aim_options_aim_account As String = "aim_account"
Public Const c_header_aim_options_buidt As String = "CUSTOM_buidt"
Public Const c_header_aim_options_buidt_underlying As String = "CUSTOM_buidt_underlying"
Public Const c_header_aim_options_buidt_current_qty As String = "CUSTOM_buidt_current_qty"
Public Const c_header_aim_options_buidt_close_qty As String = "CUSTOM_buidt_close_qty"
Public Const c_header_aim_options_buidt_close_price As String = "CUSTOM_buidt_close_price"
Public Const c_header_aim_options_buidt_local_intraday_net_trading_cash_flow As String = "CUSTOM_buidt_net_cash_local_with_comm"
Public Const c_header_aim_options_buidt_local_intraday_commission As String = "CUSTOM_intraday_commission_local"
Public Const c_header_aim_options_aim_prime_broker As String = "aim_prime_broker"
Public Const c_header_aim_options_accpbpid As String = "CUSTOM_accpbpid"
Public Const c_header_aim_options_accpbpid_underlying As String = "CUSTOM_accpbpid_underlying"
Public Const c_header_aim_options_accpbpid_current_qty As String = "CUSTOM_accpbpid_qty_current"
Public Const c_header_aim_options_accpbpid_close_qty As String = "CUSTOM_accpbpid_qty_yesterday_close"
Public Const c_header_aim_options_accpbpid_close_price_mid As String = "CUSTOM_accpbpid_yesterday_mid_price"
Public Const c_header_aim_options_accpbpid_local_intraday_net_trading_cash_flow As String = "CUSTOM_accpbpid_net_cash_local_with_comm"



'column index
Public c_aim_eod_product_id As Integer
Public c_aim_eod_underlying_id As Integer
Public c_aim_eod_instrument_type As Integer
Public c_aim_eod_ticker As Integer
Public c_aim_eod_currency As Integer
Public c_aim_eod_close_qty As Integer
Public c_aim_eod_close_price As Integer
Public c_aim_eod_local_close_ytd_pnl_gross As Integer
Public c_aim_eod_local_close_commission_broker As Integer
Public c_aim_eod_local_close_commission_total As Integer
Public c_aim_eod_local_close_dividend As Integer
Public c_aim_eod_local_close_ytd_pnl_net As Integer
Public c_aim_eod_aim_account As Integer
Public c_aim_eod_buidt As Integer
Public c_aim_eod_buidt_underlying  As Integer
Public c_aim_eod_buidt_close_qty As Integer
Public c_aim_eod_buidt_close_price As Integer
Public c_aim_eod_built_local_close_ytd_pnl_net As Integer
Public c_aim_eod_built_close_commission_total As Integer
Public c_aim_eod_aim_prime_broker As Integer
Public c_aim_eod_accpbpid As Integer
Public c_aim_eod_accpbpid_underlying As Integer
Public c_aim_eod_accpbpid_close_qty As Integer
Public c_aim_eod_accpbpid_close_price As Integer
Public c_aim_eod_accpbpid_local_close_ytd_pnl As Integer



Public c_aim_equities_product_id As Integer
Public c_aim_equities_ticker As Integer
Public c_aim_equities_description As Integer
Public c_aim_equities_currency As Integer
Public c_aim_equities_current_qty As Integer
Public c_aim_equities_close_qty As Integer
Public c_aim_equities_close_price As Integer
Public c_aim_equities_local_intraday_net_trading_cash_flow As Integer
Public c_aim_equities_local_intraday_commission As Integer
Public c_aim_equities_local_intraday_dividend As Integer
Public c_aim_equities_aim_account As Integer
Public c_aim_equities_aim_prime_broker As Integer
Public c_aim_equities_accpbpid As Integer
Public c_aim_equities_accpbpid_underlying As Integer
Public c_aim_equities_accpbpid_current_qty As Integer
Public c_aim_equities_accpbpid_close_qty As Integer
Public c_aim_equities_accpbpid_close_price As Integer
Public c_aim_equities_accpbpid_local_intraday_net_trading_cash_flow As Integer



Public c_aim_futures_product_id As Integer
Public c_aim_futures_underlying_id As Integer
Public c_aim_futures_ticker As Integer
Public c_aim_futures_currency As Integer
Public c_aim_futures_current_qty As Integer
Public c_aim_futures_close_qty As Integer
Public c_aim_futures_close_price As Integer
Public c_aim_futures_local_intraday_net_trading_cash_flow As Integer
Public c_aim_futures_local_intraday_commission As Integer
Public c_aim_futures_aim_account As Integer
Public c_aim_futures_buidt As Integer
Public c_aim_futures_buidt_underlying As Integer
Public c_aim_futures_buidt_current_qty As Integer
Public c_aim_futures_buidt_close_qty As Integer
Public c_aim_futures_buidt_close_price As Integer
Public c_aim_futures_buidt_local_intrady_net_trading_cash_flow As Integer
Public c_aim_futures_buidt_local_intraday_commission As Integer
Public c_aim_futures_aim_prime_broker As Integer
Public c_aim_futures_accpbpid As Integer
Public c_aim_futures_accpbpid_underlying As Integer
Public c_aim_futures_accpbpid_current_qty As Integer
Public c_aim_futures_accpbpid_close_qty As Integer
Public c_aim_futures_accpbpid_close_price As Integer
Public c_aim_futures_accpbpid_local_intraday_net_trading_cash_flow As Integer




Public c_aim_options_product_id As Integer
Public c_aim_options_underlying_id As Integer
Public c_aim_options_ticker_with_yellow_key As Integer
Public c_aim_options_ticker_without_yellow_key As Integer
Public c_aim_options_investment_class As Integer
Public c_aim_options_currency As Integer
Public c_aim_options_current_qty As Integer
Public c_aim_options_close_qty As Integer
Public c_aim_options_close_price_mid As Integer
Public c_aim_options_local_intraday_net_trading_cash_flow As Integer
Public c_aim_options_local_intraday_commission As Integer
Public c_aim_options_last_transaction_code As Integer
Public c_aim_options_aim_account As Integer
Public c_aim_options_buidt As Integer
Public c_aim_options_buidt_underlying As Integer
Public c_aim_options_buidt_current_qty As Integer
Public c_aim_options_buidt_close_qty As Integer
Public c_aim_options_buidt_close_price As Integer
Public c_aim_options_buidt_local_intraday_net_trading_cash_flow As Integer
Public c_aim_options_buidt_local_intraday_commission As Integer
Public c_aim_options_aim_prime_broker As Integer
Public c_aim_options_accpbpid As Integer
Public c_aim_options_accpbpid_underlying As Integer
Public c_aim_options_accpbpid_current_qty As Integer
Public c_aim_options_accpbpid_close_qty As Integer
Public c_aim_options_accpbpid_close_price_mid As Integer
Public c_aim_options_accpbpid_local_intraday_net_trading_cash_flow As Integer


Public Const filename_import_bbg_mav  As String = "Excel_export.xls"

Public Const db_bridge_import_mav As String = "db_bridge_import_mav.sqlt3"
    Public Const t_bridge_import_mav As String = "t_bridge_import_mav"
        Public Const f_bridge_import_mav_product_id As String = "f_bridge_import_mav_product_id"
        Public Const f_bridge_import_mav_underlying_id As String = "f_bridge_import_mav_underlying_id"
        Public Const f_brdige_import_mav_asset_type As String = "f_brdige_import_mav_asset_type"
    Public Const t_helper_import_mav As String = "t_helper_import_mav"
        Public Const f_helper_import_mav_text_1 As String = "f_helper_import_mav_text_1"
        Public Const f_helper_import_mav_text_2 As String = "f_helper_import_mav_text_2"
        Public Const f_helper_import_mav_text_3 As String = "f_helper_import_mav_text_3"
        Public Const f_helper_import_mav_numeric_1 As String = "f_helper_import_mav_numeric_1"
        Public Const f_helper_import_mav_numeric_2 As String = "f_helper_import_mav_numeric_2"
        Public Const f_helper_import_mav_numeric_3 As String = "f_helper_import_mav_numeric_3"
    


Public Sub aim_account_switch_account_open(ByVal account As String)

Call aim_update_open_summary(account)
Call aim_hide_lines_open_apart_of_account(Array(account))

End Sub


Private Sub aim_hide_lines_open_apart_of_account(ByVal vec_account As Variant)

Application.Calculation = xlCalculationManual

Dim c_open_account As Integer
    c_open_account = 92

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim open_threshold As Integer
    open_threshold = 25
    Dim is_last_line As Boolean


Worksheets("Open").rows("27:5000").Hidden = False

Application.ScreenUpdating = False

For i = 26 To 3000
    
    is_last_line = True
    
    For j = 0 To open_threshold
        If Worksheets("Open").Cells(i + j, 1) <> "" Then
            is_last_line = False
            Exit For
        End If
    Next j
    
    If is_last_line = True Then
        Exit For
    End If
    
    
    If Worksheets("Open").Cells(i, 1) <> "" Then
        
        For j = 0 To UBound(vec_account, 1)
            If vec_account(j) = Worksheets("Open").Cells(i, c_open_account) Or vec_account(j) = "" Then
                Exit For
            Else
                If j = UBound(vec_account, 1) Then
                    Worksheets("Open").rows(i).Hidden = True
                End If
            End If
        Next j
        
    End If
    
Next i

Application.ScreenUpdating = True

End Sub


Private Sub aim_update_open_summary(ByVal account As String)

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer

Dim array_summary_data_kronos_monitor() As Variant
array_summary_data_kronos_monitor = Array(Array("", 38), Array("C6414GSL", 73))


Dim vec_currency() As Variant
k = 0
For i = 14 To 32
    If Worksheets("Parametres").Cells(i, 1).Value = "" Then
        Exit For
    Else
        ReDim Preserve vec_currency(k)
        vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, i)
        k = k + 1
    End If
Next i


For p = 0 To UBound(array_summary_data_kronos_monitor, 1)
    If array_summary_data_kronos_monitor(p)(0) = account Then
        
        For i = 11 To 17
            For j = 0 To UBound(vec_currency, 1)
                If UCase(vec_currency(j)(0)) = UCase(Worksheets("Open").Cells(i, 22).Value) Then
                    
                    'trouve la colonne avec le bon currency code
                    For q = 29 To 60
                        If Worksheets("Kronos_Monitor").Cells(array_summary_data_kronos_monitor(p)(1), q) = vec_currency(j)(1) Or Worksheets("Kronos_Monitor").Cells(array_summary_data_kronos_monitor(p)(1), q) = array_summary_data_kronos_monitor(p)(0) & "_" & vec_currency(j)(1) Then
                            
                            'vega
                            Worksheets("Open").Cells(i, 24).FormulaLocal = "=Kronos_Monitor!R" & array_summary_data_kronos_monitor(p)(1) + 3 & "C" & q & "/1000"
                            
                            'theta
                            Worksheets("Open").Cells(i, 25).FormulaLocal = "=Kronos_Monitor!R" & array_summary_data_kronos_monitor(p)(1) + 2 & "C" & q
                            
                            'valeur eur
                            Worksheets("Open").Cells(i, 26).FormulaLocal = "=Kronos_Monitor!R" & array_summary_data_kronos_monitor(p)(1) + 1 & "C" & q & "/nav!R1C13"
                            
                            
                            Exit For
                        End If
                    Next q
                    
                    
                    Exit For
                End If
            Next j
        Next i
        
        
    End If
Next p




End Sub



Private Sub deploy_account_and_currency()

Dim i As Integer, j As Integer

Application.Calculation = xlCalculationManual

For i = 27 To 5000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        Worksheets("Equity_Database").Cells(i, 104).FormulaLocal = "=IF(RC103<>"""";RC103;Parametres!R17C18) & ""_"" &RC44"
    End If
    
Next i

End Sub



Public Sub aim_static_EOD()


If Workbooks("Kronos.xls").Worksheets("Parametres").Cells(16, 19) = 1 Or IsDate(Workbooks("Kronos.xls").Worksheets("Parametres").Cells(16, 19)) Then
Else
    Exit Sub
End If

Call aim_assign_column(Array(aim_view_code.EOD))

Dim oBBG As New cls_Bloomberg_Sync


Dim underyling_override As Variant
underyling_override = Array(Array("SX5ED", "DSX5E"))

Dim vec_asset_product() As Variant
    vec_asset_product = Array("Equity", "Equity Option", "Index Future", "Index Option")

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
oReg.Global = True

oReg.Pattern = "grid(\d){1,}.xls"

Dim file_static_eod As String
Dim find_static_eod As Boolean
    find_static_eod = False

Dim tmp_wrbk As Workbook
For Each tmp_wrbk In Workbooks
    
    Set matches = oReg.Execute(tmp_wrbk.name)
    
    For Each match In matches
        file_static_eod = tmp_wrbk.name
        find_static_eod = True
        Exit For
    Next
    
Next

If find_static_eod = False Then
    MsgBox ("file not found with export mav.")
    Exit Sub
End If


For i = 2 To 50
    If Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "" Then
        c_static_eod_ticker = 1
        Exit For
    Else
        If Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "Asset Type" Then
            c_static_eod_asset_type = i
        ElseIf Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "BB Unique Id" Then
            c_static_eod_buid = i
        ElseIf Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "Position" Then
            c_static_eod_qty = i
        ElseIf Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "Price" Then
            c_static_eod_close_price = i
        ElseIf Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "YTD_PnL_Local" Then
            c_static_eod_ytd_pnl_local = i
        ElseIf Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "YTD_transac_costs_local" Then
            c_static_eod_ytd_comm = i
        ElseIf Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "Prime" Then
            c_static_eod_pb = i
        ElseIf Workbooks(file_static_eod).Worksheets(1).Cells(1, i) = "Currency" Then
            c_static_eod_currency = i
        End If
    End If
Next i


Dim tmp_last_account As String
    tmp_last_account = Workbooks(file_static_eod).Worksheets(1).Cells(4, c_static_eod_ticker).Value


Dim dico_mav As New Dictionary
Dim dico_close_price_equity As New Dictionary

Dim is_new_account As Boolean


    Dim tmp_vec_row() As Variant
        
        dim_dico_account = 0
        dim_dico_pb = 1
        dim_dico_buid = 2
        dim_dico_uid = 3
        dim_dico_ticker = 4
        dim_dico_asset_type = 5
        dim_dico_qty = 6
        dim_dico_price = 7
        dim_dico_ytd = 8
        dim_dico_comm = 9
        dim_dico_currency = 10
        
        Dim vec_ticker() As Variant
        k = 0 'count ticker
        
        count_entry_mav = 0


For i = 5 To 12000
    
    If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ticker).Value = "" Then
        
        Exit For
    End If
    
    
    'new account  ?
    is_new_account = True
    For j = c_static_eod_asset_type To c_static_eod_pb
        If Workbooks(file_static_eod).Worksheets(1).Cells(i, j) <> "" Then
            is_new_account = False
            Exit For
        End If
    Next j
    
    
    Dim tmp_qty As Double, tmp_ytd_pnl As Double, tmp_ytd_comm As Double
    
    If is_new_account = True Then
        tmp_last_account = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ticker).Value
    Else
        
        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_pnl_local).Value <> "" Or Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_comm).Value <> "" Then

            For j = 0 To UBound(vec_asset_product, 1)

                If vec_asset_product(j) = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_asset_type).Value Then
                    'check si valeur existe deja
                    
                    If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_pb) <> "" Then
                        tmp_pb = CStr(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_pb))
                    Else
                        tmp_pb = "N/A"
                    End If
                    
                    
                    'debug_test = tmp_last_account & "_" & tmp_pb & "_" & Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_buid).Value
                    If dico_mav.Exists(tmp_last_account & "_" & tmp_pb & "_" & Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_buid).Value) Then
                        'edit pos

                        tmp_vec_row = dico_mav.Item(tmp_last_account & "_" & tmp_pb & "_" & Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_buid).Value)


                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_qty).Value <> "" Then
                            tmp_qty = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_qty).Value
                        Else
                            tmp_qty = 0
                        End If


                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_pnl_local).Value <> "" Then
                            tmp_ytd_pnl = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_pnl_local).Value
                        Else
                            tmp_ytd_pnl = 0
                        End If


                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_comm).Value <> "" Then
                            tmp_ytd_comm = Abs(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_comm).Value)
                        Else
                            tmp_ytd_comm = 0
                        End If

                        tmp_vec_row(dim_dico_qty) = tmp_vec_row(dim_dico_qty) + tmp_qty
                        tmp_vec_row(dim_dico_ytd) = tmp_vec_row(dim_dico_ytd) + tmp_ytd_pnl
                        tmp_vec_row(dim_dico_comm) = tmp_vec_row(dim_dico_comm) + tmp_ytd_comm
                        
                        
                        dico_mav.Item(tmp_last_account & "_" & Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_pb).Value & "_" & Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_buid).Value) = tmp_vec_row

                    Else

                        'new entry
                        tmp_account = CStr(tmp_last_account)
                        
                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_pb) <> "" Then
                            tmp_pb = CStr(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_pb))
                        Else
                            tmp_pb = "N/A"
                        End If
                        
                        tmp_buid = CStr(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_buid))
                        tmp_uid = ""
                        
                        
                        tmp_ticker = UCase(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ticker).Value)
                            tmp_ticker = Replace(tmp_ticker, "[", "")
                            tmp_ticker = Replace(tmp_ticker, "]", "")
                        If InStr(UCase(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_asset_type).Value), "INDEX") <> 0 Then
                            tmp_ticker = tmp_ticker & " INDEX"
                        ElseIf InStr(UCase(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_asset_type).Value), "EQUITY") <> 0 Then
                            tmp_ticker = tmp_ticker & " EQUITY"
                        End If

                        tmp_asset_type = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_asset_type).Value

                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_qty).Value <> "" Then
                            tmp_qty = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_qty).Value
                        Else
                            tmp_qty = 0
                        End If


                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_close_price).Value <> "" Then
                            tmp_price = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_close_price).Value
                            
                            If tmp_asset_type = "Equity" Then
                                If dico_close_price_equity.Exists(tmp_buid) Then
                                Else
                                    dico_close_price_equity.Add tmp_buid, tmp_price
                                End If
                            End If
                        Else
                            tmp_price = 0
                            
                            'si de type equity, il faut trouver un prix pour eviter de faire sauter les close d equity database
                            If tmp_asset_type = "Equity" Then
                                If dico_close_price_equity.Exists(tmp_buid) Then
                                    tmp_price = dico_close_price_equity.Item(tmp_buid)
                                Else
                                    debug_test = "stop"
                                End If
                            End If
                            
                        End If
                        
                        
                        If UCase(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_currency).Value) = "GBP" Then
                            tmp_price = tmp_price / 100
                        End If
                        


                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_pnl_local).Value <> "" Then
                            tmp_ytd_pnl = Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_pnl_local).Value
                        Else
                            tmp_ytd_pnl = 0
                        End If


                        If Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_comm).Value <> "" Then
                            tmp_ytd_comm = Abs(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_ytd_comm).Value)
                        Else
                            tmp_ytd_comm = 0
                        End If
                        
                        
                        tmp_currency = UCase(Workbooks(file_static_eod).Worksheets(1).Cells(i, c_static_eod_currency).Value)

                        tmp_vec_row = Array(tmp_account, tmp_pb, tmp_buid, tmp_uid, tmp_ticker, tmp_asset_type, tmp_qty, tmp_price, tmp_ytd_pnl, tmp_ytd_comm, tmp_currency)

                        dico_mav.Add tmp_account & "_" & tmp_pb & "_" & tmp_buid, tmp_vec_row

                        count_entry_mav = count_entry_mav + 1

                        If k = 0 Then
                            ReDim Preserve vec_ticker(k)
                            vec_ticker(k) = "/buid/" & tmp_buid
                            k = k + 1
                        Else
                            For m = 0 To UBound(vec_ticker, 1)
                                If "/buid/" & tmp_buid = vec_ticker(m) Then
                                    Exit For
                                Else
                                    If m = UBound(vec_ticker, 1) Then
                                        ReDim Preserve vec_ticker(k)
                                        vec_ticker(k) = "/buid/" & tmp_buid
                                        k = k + 1
                                    End If
                                End If
                            Next m
                        End If


                    End If


                    Exit For
                End If

            Next j

        End If
        
    End If
    
Next i


Workbooks(file_static_eod).Close False

'appel bbg pour les underlying id
Dim data_bbg As Variant
data_bbg = oBBG.bdp(vec_ticker, Array("UNDL_ID_BB_UNIQUE", "UNDL_SPOT_TICKER"), output_format.of_vec_without_header)



If count_entry_mav > 0 Then
    
    'on vide aim_eod et en remplace par le contenu du dico
    Workbooks("Kronos.xls").Worksheets(aim_view_eod).Range("A11:Y8000").Clear
    
    k = 1
    
    m = 1
    For Each tmp_id_dico In dico_mav
        
        tmp_vec_row = dico_mav.Item(tmp_id_dico)
        
        tmp_account = tmp_vec_row(dim_dico_account)
        tmp_pb = tmp_vec_row(dim_dico_pb)
        tmp_buid = tmp_vec_row(dim_dico_buid)
        tmp_ticker = tmp_vec_row(dim_dico_ticker)
        tmp_asset_type = tmp_vec_row(dim_dico_asset_type)
        tmp_qty = tmp_vec_row(dim_dico_qty)
        tmp_price = tmp_vec_row(dim_dico_price)
        tmp_ytd_pnl = tmp_vec_row(dim_dico_ytd)
        tmp_ytd_comm = tmp_vec_row(dim_dico_comm)
        
        
        
        If tmp_asset_type = "Equity" Then
            tmp_uid = tmp_buid
        Else
            For j = 0 To UBound(vec_ticker, 1)
                If vec_ticker(j) = "/buid/" & tmp_buid Then
                    If InStr(tmp_asset_type, "Option") <> 0 Then
                        tmp_uid = data_bbg(j)(0)
                    ElseIf InStr(tmp_asset_type, "Future") <> 0 Then
                        tmp_uid = "EI09" & data_bbg(j)(1)
                    End If
                    
                    Exit For
                End If
            Next j
                    
        End If
        
        
        For j = 0 To UBound(underyling_override, 1)
            If InStr(tmp_uid, underyling_override(j)(0)) <> 0 Then
                tmp_uid = Replace(tmp_uid, underyling_override(j)(0), underyling_override(j)(1))
            End If
        Next j
        
        
        
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, 1) = k
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_product_id) = tmp_buid
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_underlying_id) = tmp_uid
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_instrument_type) = tmp_asset_type
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_ticker) = tmp_ticker
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_close_qty) = tmp_qty
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_close_price) = tmp_price
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_local_close_ytd_pnl_gross) = tmp_ytd_pnl
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_local_close_commission_total) = tmp_ytd_comm
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_local_close_dividend) = 0
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_local_close_ytd_pnl_net) = tmp_ytd_pnl - tmp_ytd_comm
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_aim_account) = tmp_account
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_buidt) = tmp_account & "_" & tmp_buid
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_buidt_underlying) = tmp_account & "_" & tmp_uid
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_buidt_close_qty) = tmp_qty
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_buidt_close_price) = tmp_price
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_built_local_close_ytd_pnl_net) = tmp_ytd_pnl - tmp_ytd_comm
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_built_close_commission_total) = tmp_ytd_comm
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_aim_prime_broker) = tmp_pb
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_accpbpid) = tmp_account & "_" & tmp_pb & "_" & tmp_buid
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_accpbpid_underlying) = tmp_account & "_" & tmp_pb & "_" & tmp_uid
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_accpbpid_close_qty) = tmp_qty
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_accpbpid_close_price) = tmp_price
        Workbooks("Kronos.xls").Worksheets(aim_view_eod).Cells(k + l_header_aim_view, c_aim_eod_accpbpid_local_close_ytd_pnl) = tmp_ytd_pnl - tmp_ytd_comm
        
        
        k = k + 1
        
    Next
    
    
    Workbooks("Kronos.xls").Worksheets("Parametres").Cells(16, 19) = Now()
    
End If


End Sub


Public Function aim_get_worksheet_xls_name_from_view_code(ByVal view_code As Integer) As String

Debug.Print "aim_get_worksheet_xls_name_from_view_code: " & "INPUTs " & "#view_code=" & view_code

If view_code = aim_view_code.EOD Then
    aim_get_worksheet_xls_name_from_view_code = aim_view_eod
ElseIf view_code = aim_view_code.equities Then
    aim_get_worksheet_xls_name_from_view_code = aim_view_equities
ElseIf view_code = aim_view_code.futures Then
    aim_get_worksheet_xls_name_from_view_code = aim_view_futures
ElseIf view_code = aim_view_code.Options Then
    aim_get_worksheet_xls_name_from_view_code = aim_view_options
Else
    aim_get_worksheet_xls_name_from_view_code = "#N/A view"
End If

Debug.Print "aim_get_worksheet_xls_name_from_view_code: " & "OUTPUT " & "@" & aim_get_worksheet_xls_name_from_view_code

End Function


Public Function aim_get_view_column_index_from_header(ByVal view_code As Integer, ByVal header As String) As Integer

Debug.Print "aim_get_view_column_index_from_header: " & "INPUTs " & "#view_code=" & view_code & ", #header=" & header

aim_get_view_column_index_from_header = 0

Dim status_assign_column As Variant
status_assign_column = aim_assign_column(Array(view_code))


For i = 0 To UBound(status_assign_column(0), 1)
    If status_assign_column(0)(i)(0) = header Then
        aim_get_view_column_index_from_header = status_assign_column(0)(i)(1)
        Debug.Print "aim_get_view_column_index_from_header: " & "OUTPUT " & "@" & status_assign_column(0)(i)(1)
        Exit Function
    End If
Next i

End Function


Public Function aim_get_product_column_from_view_code(ByVal view_code As Integer) As Integer

Debug.Print "aim_get_product_column_from_view_code: " & "INPUT " & "#view_code=" & view_code

aim_get_product_column_from_view_code = 0

Dim status_assign_column As Variant
status_assign_column = aim_assign_column(Array(view_code))

If view_code = aim_view_code.EOD Then
    aim_get_product_column_from_view_code = c_aim_eod_product_id
ElseIf view_code = aim_view_code.equities Then
    aim_get_product_column_from_view_code = c_aim_equities_product_id
ElseIf view_code = aim_view_code.futures Then
    aim_get_product_column_from_view_code = c_aim_futures_product_id
ElseIf view_code = aim_view_code.Options Then
    aim_get_product_column_from_view_code = c_aim_options_product_id
End If

Debug.Print "aim_get_product_column_from_view_code: " & "OUTPUT " & "@" & aim_get_product_column_from_view_code

End Function


Public Function aim_get_underlying_id_column_from_view_code(ByVal view_code As Integer) As Integer

Debug.Print "aim_get_underlying_id_column_from_view_code: " & "INPUT " & "#view_code=" & view_code

aim_get_underlying_id_column_from_view_code = 0

Dim status_assign_column As Variant
status_assign_column = aim_assign_column(Array(view_code))

If view_code = aim_view_code.EOD Then
    aim_get_underlying_id_column_from_view_code = c_aim_eod_underlying_id
ElseIf view_code = aim_view_code.equities Then
    aim_get_underlying_id_column_from_view_code = c_aim_equities_product_id
ElseIf view_code = aim_view_code.futures Then
    aim_get_underlying_id_column_from_view_code = c_aim_futures_underlying_id
ElseIf view_code = aim_view_code.Options Then
    aim_get_underlying_id_column_from_view_code = c_aim_options_underlying_id
End If

Debug.Print "aim_get_underlying_id_column_from_view_code: " & "OUTPUT " & "@" & aim_get_underlying_id_column_from_view_code

End Function


Public Function aim_get_working_mode() As Variant

aim_get_working_mode = Worksheets("Parametres").Cells(15, 19).Value

End Function


Public Sub aim_autocomplete_formula_view()

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Debug.Print "aim_autocomplete_formula_view: #NO INPUT"

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
oReg.Global = True


If aim_get_working_mode = aim_mode.manual_import_mav_excel_export Then
    
    Application.Calculation = xlManual
    
    'import manual de bloomberg
        
    
    'passe en revue les fichiers ouverts et si trouve grid -> sauve pour l import futur
    oReg.Pattern = "grid(\d){1,}.xls"
    
    Dim tmp_wrbk As Workbook
    For Each tmp_wrbk In Workbooks
        
        Set matches = oReg.Execute(tmp_wrbk.name)
        
        For Each match In matches
            
            Application.DisplayAlerts = False
            Workbooks(tmp_wrbk.name).SaveAs base_path & filename_import_bbg_mav
            Application.DisplayAlerts = True
            Workbooks(tmp_wrbk.name).Close False
        Next
        
    Next
    
    
    Call aim_autocomplete_formula_with_import_bloomberg_mav
ElseIf aim_get_working_mode = aim_mode.rtd_auto_db_pictet Or aim_get_working_mode = "" Then
    Call aim_autocomplete_formula_with_link_db_pictet_kronos
End If



End Sub


Private Sub aim_init_db_helper_import_bloomberg_mav()

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))


If exist_file(base_path & db_bridge_import_mav) = False Then
    
    'init de la db
    create_table_status = sqlite3_create_db(base_path & db_bridge_import_mav)
End If


'creation des tables
If sqlite3_check_if_table_already_exist(base_path & db_bridge_import_mav, t_bridge_import_mav) = False Then
    create_table_query = sqlite3_get_query_create_table(t_bridge_import_mav, Array(Array(f_bridge_import_mav_product_id, "TEXT", ""), Array(f_bridge_import_mav_underlying_id, "TEXT", ""), Array(f_brdige_import_mav_asset_type, "TEXT", "")), Array(Array(f_bridge_import_mav_product_id, "ASC")))
    create_table_status = sqlite3_create_tables(base_path & db_bridge_import_mav, Array(create_table_query))
End If


If sqlite3_check_if_table_already_exist(base_path & db_bridge_import_mav, t_helper_import_mav) = False Then
    create_table_query = sqlite3_get_query_create_table(t_helper_import_mav, Array(Array(f_helper_import_mav_text_1, "TEXT", ""), Array(f_helper_import_mav_text_2, "TEXT", ""), Array(f_helper_import_mav_text_3, "TEXT", ""), Array(f_helper_import_mav_numeric_1, "NUMERIC", ""), Array(f_helper_import_mav_numeric_2, "NUMERIC", ""), Array(f_helper_import_mav_numeric_3, "NUMERIC", "")))
    create_table_status = sqlite3_create_tables(base_path & db_bridge_import_mav, Array(create_table_query))
End If


End Sub


Private Sub aim_manipulate_db_helper_bloomberg_mav()

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

debug_test = sqlite3_query(base_path & db_bridge_import_mav, "SELECT DISTINCT " & f_brdige_import_mav_asset_type & " FROM " & t_bridge_import_mav)

Dim sql_query As String
sql_query = "DELETE FROM " & t_bridge_import_mav & " WHERE " & f_brdige_import_mav_asset_type & "=""Index Future"" OR " & f_brdige_import_mav_asset_type & "=""Index Option"""

exec_query = sqlite3_query(base_path & db_bridge_import_mav, sql_query)

End Sub


Private Sub aim_check_db_helper_bloomberg_mav()

Dim i As Integer, j As Integer, k As Integer
Dim sql_query As String

exec_query = sqlite3_query(base_path & db_bridge_import_mav, "DELETE FROM " & t_bridge_import_mav & " WHERE " & f_brdige_import_mav_asset_type & "=""Index Future""")

End Sub

Private Sub aim_autocomplete_formula_with_import_bloomberg_mav()

Dim underyling_override As Variant
underyling_override = Array(Array("SX5ED", "DSX5E"))

datetime_start = Now()

Dim oBBG As New cls_Bloomberg_Sync

Application.Calculation = xlCalculationManual


Dim asset_type_to_import As Variant
    asset_type_to_import = Array("Equity", "Equity Option", "Index Future", "Index Option")




Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim debug_test As Variant
Dim sql_query As String

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Call aim_init_db_helper_import_bloomberg_mav


'charge la totatite du fichier excel d export
If exist_file(base_path & filename_import_bbg_mav) Then
Else
    MsgBox ("fichier introuvable: " & base_path & filename_import_bbg_mav & " -> Exit")
    Exit Sub
End If


open_file base_path & filename_import_bbg_mav, True


'dection des colonne
For i = 2 To 20
    If Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(1, i) = "Asset Type" Then
        c_mav_asset_type = i
    ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(1, i) = "BB Unique Id" Then
        c_mav_buid = i
    ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(1, i) = "Units" Then
        c_mav_position = i
    ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(1, i) = "Prime" Then
        c_mav_pb = i
    End If
Next i


'extraction integral

Dim dico_mav As New Dictionary

Dim vec_mav() As Variant
    ReDim vec_mav(0)
    vec_mav(0) = Array("", "", 0)
k = 0
Dim import_threshold As Integer, is_last_line As Boolean
    import_threshold = 50
For i = 1 To 12000
    
    is_last_line = True
    
    For j = 0 To import_threshold
        If Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i + j, 1) <> "" Then
            is_last_line = False
            Exit For
        End If
    Next j
    
    If is_last_line = True Then
        Exit For
    Else
        Dim tmp_ticker As String
        If Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1) <> "" And Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_position) <> "" And Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_pb) <> "" Then
            
            'check asset type
            For j = 0 To UBound(asset_type_to_import, 1)
                
                If Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type) = asset_type_to_import(j) Then
                    
                    Dim tmp_dico_item As Variant
                    If dico_mav.Exists(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_buid).Value) Then
                        
                        'adjust net qty
                        tmp_dico_item = dico_mav(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_buid).Value)
                            tmp_dico_item(1) = tmp_dico_item(1) + CDbl(Replace(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_position).Value, "'", ""))
                        
                        dico_mav(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_buid).Value) = tmp_dico_item
                        
                    Else
                        
                        'creation de la nouvelle entree
                        
                        If Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Equity" Then
                            tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " EQUITY"
                        ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Equity Option" Then
                            tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " EQUITY"
                        ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Index Option" Then
                            tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " INDEX"
                        ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Index Future" Then
                            tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " INDEX"
                        End If
                        
                        dico_mav.Add Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_buid).Value, Array(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_buid).Value, CDbl(Replace(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_position).Value, "'", "")), Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value, Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_pb).Value, tmp_ticker)
                        
                        k = k + 1
                        
                    End If
                    
                    
                    'deprecie car trop lent
'                    'check si deja pos -> sum
'                    For m = 0 To UBound(vec_mav, 1)
'                        If Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_buid) = vec_mav(m)(0) Then
'                            vec_mav(m)(1) = vec_mav(m)(1) + CDbl(Replace(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_position).Value, "'", ""))
'                            Exit For
'                        Else
'                            If m = UBound(vec_mav, 1) Then
'
'                                If k = 0 Then
'                                Else
'                                    ReDim Preserve vec_mav(k)
'                                End If
'
'
'                                If Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Equity" Then
'                                    tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " EQUITY"
'                                ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Equity Option" Then
'                                    tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " EQUITY"
'                                ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Index Option" Then
'                                    tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " INDEX"
'                                ElseIf Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value = "Index Future" Then
'                                    tmp_ticker = Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, 1).Value & " INDEX"
'                                End If
'
'                                vec_mav(k) = Array(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_buid).Value, CDbl(Replace(Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_position).Value, "'", "")), Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_asset_type).Value, Workbooks(filename_import_bbg_mav).Worksheets(1).Cells(i, c_mav_pb).Value, tmp_ticker)
'
'                                k = k + 1
'                            End If
'                        End If
'                    Next m
                    
                End If
                
            Next j
            
        End If
        
    End If
    
Next i

Workbooks(filename_import_bbg_mav).Close False

If k = 0 Then
    MsgBox ("nothing import file ! -> Exit.")
    Exit Sub
Else
    
    'transforme le dico en array
    k = 0
    For Each tmp_item In dico_mav
        ReDim Preserve vec_mav(k)
        vec_mav(k) = dico_mav(tmp_item)
        k = k + 1
    Next
    
End If

'inject tous les products id dans le helper pour reperer ceux dont on connait pas encore l underyling
Dim vec_helper_product_id()
For i = 0 To UBound(vec_mav, 1)
    ReDim Preserve vec_helper_product_id(i)
    vec_helper_product_id(i) = Array(vec_mav(i)(0), vec_mav(i)(1), vec_mav(i)(3), vec_mav(i)(4))
Next i
    
    exec_query = sqlite3_query(base_path & db_bridge_import_mav, "DELETE FROM " & t_helper_import_mav)
    
    insert_status = sqlite3_insert_with_transaction(base_path & db_bridge_import_mav, t_helper_import_mav, vec_helper_product_id, Array(f_helper_import_mav_text_1, f_helper_import_mav_numeric_1, f_helper_import_mav_text_2, f_helper_import_mav_text_3))
    
    sql_query = "SELECT " & f_helper_import_mav_text_1
        sql_query = sql_query & " FROM " & t_helper_import_mav
        sql_query = sql_query & " WHERE " & f_helper_import_mav_text_1 & " NOT IN ( "
            
            sql_query = sql_query & " SELECT " & f_bridge_import_mav_product_id & " FROM " & t_bridge_import_mav
            
        sql_query = sql_query & " ) "
    
    Dim extract_new_product_id As Variant
    extract_new_product_id = sqlite3_query(base_path & db_bridge_import_mav, sql_query)
    
    
    If UBound(extract_new_product_id, 1) > 0 Then
        
        'complete la table helper
        Dim vec_ticker() As Variant
        For i = 1 To UBound(extract_new_product_id, 1)
            ReDim Preserve vec_ticker(i - 1)
            vec_ticker(i - 1) = "/buid/" & extract_new_product_id(i)(0)
        Next i
        
        
        Dim field_bbg As Variant
            field_bbg = Array("UNDL_ID_BB_UNIQUE", "UNDL_SPOT_TICKER")
            
            Dim output_bbg As Variant
            output_bbg = oBBG.bdp(vec_ticker, field_bbg, output_format.of_vec_without_header)
            
        
        Dim vec_insert_bridge() As Variant
        k = 0
        For i = 0 To UBound(output_bbg, 1)
            
            'match pour connaitre asset type
            For j = 0 To UBound(vec_mav, 1)
                
                Dim tmp_underlying As String
                
                If vec_ticker(i) = "/buid/" & vec_mav(j)(0) Then
                    
                    If vec_mav(j)(2) = "Equity" Then
                        ReDim Preserve vec_insert_bridge(k)
                        vec_insert_bridge(k) = Array(vec_mav(j)(0), vec_mav(j)(0), vec_mav(j)(2))
                        k = k + 1
                    ElseIf vec_mav(j)(2) = "Index Future" And Left(output_bbg(i)(1), 1) <> "#" Then
                        ReDim Preserve vec_insert_bridge(k)
                        
                        tmp_underlying = "EI09" & output_bbg(i)(1)
                        
                        For m = 0 To UBound(underyling_override, 1)
                            tmp_underlying = Replace(tmp_underlying, underyling_override(m)(0), underyling_override(m)(1))
                        Next m
                        
                        vec_insert_bridge(k) = Array(vec_mav(j)(0), tmp_underlying, vec_mav(j)(2))
                        k = k + 1
                    ElseIf (vec_mav(j)(2) = "Equity Option" Or vec_mav(j)(2) = "Index Option") And Left(output_bbg(i)(0), 1) <> "#" Then
                        
                        tmp_underlying = output_bbg(i)(0)
                        
                        For m = 0 To UBound(underyling_override, 1)
                            tmp_underlying = Replace(tmp_underlying, underyling_override(m)(0), underyling_override(m)(1))
                        Next m
                        
                        ReDim Preserve vec_insert_bridge(k)
                        vec_insert_bridge(k) = Array(vec_mav(j)(0), tmp_underlying, vec_mav(j)(2))
                        k = k + 1
                    Else
                        debug_test = "bug"
                        
                    End If
                    
                    Exit For
                End If
                
            Next j
            
        Next i
        
        
        If k > 0 Then
            insert_status = sqlite3_insert_with_transaction(base_path & db_bridge_import_mav, t_bridge_import_mav, vec_insert_bridge, Array(f_bridge_import_mav_product_id, f_bridge_import_mav_underlying_id, f_brdige_import_mav_asset_type))
        End If
        
        
    End If
    
    
'nouvelle extraction avec toutes les donnees
sql_query = "SELECT " & f_bridge_import_mav_product_id & ", " & f_bridge_import_mav_underlying_id & ", " & f_brdige_import_mav_asset_type & ", " & f_helper_import_mav_numeric_1 & " AS current_position, " & f_helper_import_mav_text_2 & " AS pb, " & f_helper_import_mav_text_3 & " AS ticker "
    sql_query = sql_query & " FROM " & t_bridge_import_mav & ", " & t_helper_import_mav
    sql_query = sql_query & " WHERE " & f_bridge_import_mav_product_id & "=" & f_helper_import_mav_text_1
    
Dim extract_data_to_insert_in_aim_view As Variant
extract_data_to_insert_in_aim_view = sqlite3_query(base_path & db_bridge_import_mav, sql_query)

    'count chaque type pour la dimension des matrix/asset type
    sql_query = "SELECT " & f_brdige_import_mav_asset_type & ", COUNT(" & f_bridge_import_mav_product_id & ")"
        sql_query = sql_query & " FROM " & t_bridge_import_mav & ", " & t_helper_import_mav
        sql_query = sql_query & " WHERE " & f_bridge_import_mav_product_id & "=" & f_helper_import_mav_text_1
        sql_query = sql_query & " GROUP BY " & f_brdige_import_mav_asset_type
        
Dim extract_count_data_to_insert_in_aim_view As Variant
extract_count_data_to_insert_in_aim_view = sqlite3_query(base_path & db_bridge_import_mav, sql_query)


'vide la sheet EOD
Worksheets(aim_view_eod).Range("B" & l_header_aim_view + 1 & ":AF5000").Clear

Worksheets(aim_view_equities).Range("B" & l_header_aim_view + 1 & ":AF5000").Clear
Worksheets(aim_view_futures).Range("B" & l_header_aim_view + 1 & ":AF5000").Clear
Worksheets(aim_view_options).Range("B" & l_header_aim_view + 1 & ":AF5000").Clear

line_equity = 0
line_option = 0
line_future = 0


Dim aim_account As String
    aim_account = Worksheets("Parametres").Cells(17, 18)

Dim matrix_aim_view_equities() As Variant, matrix_aim_view_futures() As Variant, matrix_aim_view_options() As Variant
    
    For i = 1 To UBound(extract_count_data_to_insert_in_aim_view, 1)
        If extract_count_data_to_insert_in_aim_view(i)(0) = "Equity" Then
            line_equity = line_equity + extract_count_data_to_insert_in_aim_view(i)(1)
        ElseIf InStr(extract_count_data_to_insert_in_aim_view(i)(0), "Future") <> 0 Then
            line_future = line_future + extract_count_data_to_insert_in_aim_view(i)(1)
        ElseIf InStr(extract_count_data_to_insert_in_aim_view(i)(0), "Option") <> 0 Then
            line_option = line_option + extract_count_data_to_insert_in_aim_view(i)(1)
        End If
    Next i
    
    
    If line_equity > 0 Then
        ReDim matrix_aim_view_equities(1 To line_equity, 1 To 19)
    End If
    
    If line_future > 0 Then
        ReDim matrix_aim_view_futures(1 To line_future, 1 To 25)
    End If
    
    If line_option > 0 Then
        ReDim matrix_aim_view_options(1 To line_option, 1 To 29)
    End If
    
    
line_equity = 0
line_option = 0
line_future = 0

For i = 1 To UBound(extract_data_to_insert_in_aim_view, 1)
    
    If extract_data_to_insert_in_aim_view(i)(2) = "Equity" Then
        
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 2) = extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 3) = extract_data_to_insert_in_aim_view(i)(5)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 6) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 7) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 8) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 9) = 0
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 10) = 0
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 11) = 0
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 12) = aim_account
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 13) = extract_data_to_insert_in_aim_view(i)(4)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 14) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 15) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 16) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 17) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 18) = 0
'        Worksheets(aim_view_equities).Cells(l_header_aim_view + 1 + line_equity, 19) = 0
        
        matrix_aim_view_equities(line_equity + 1, 2) = extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_equities(line_equity + 1, 3) = extract_data_to_insert_in_aim_view(i)(5)
        matrix_aim_view_equities(line_equity + 1, 6) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_equities(line_equity + 1, 7) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_equities(line_equity + 1, 8) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_equities(line_equity + 1, 9) = 0
        matrix_aim_view_equities(line_equity + 1, 10) = 0
        matrix_aim_view_equities(line_equity + 1, 11) = 0
        matrix_aim_view_equities(line_equity + 1, 12) = aim_account
        matrix_aim_view_equities(line_equity + 1, 13) = extract_data_to_insert_in_aim_view(i)(4)
        matrix_aim_view_equities(line_equity + 1, 14) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_equities(line_equity + 1, 15) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_equities(line_equity + 1, 16) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_equities(line_equity + 1, 17) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_equities(line_equity + 1, 18) = 0
        matrix_aim_view_equities(line_equity + 1, 19) = 0
        
        line_equity = line_equity + 1
        
    ElseIf InStr(extract_data_to_insert_in_aim_view(i)(2), "Future") <> 0 Then
        
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 2) = extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 3) = extract_data_to_insert_in_aim_view(i)(1)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 4) = extract_data_to_insert_in_aim_view(i)(5)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 6) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 7) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 8) = 0
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 9) = 0
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 10) = 0
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 11) = aim_account
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 12) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 13) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(1)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 14) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 15) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 16) = 0
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 17) = 0
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 18) = 0
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 19) = extract_data_to_insert_in_aim_view(i)(4)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 20) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 21) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(1)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 22) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 23) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 24) = 0
'        Worksheets(aim_view_futures).Cells(l_header_aim_view + 1 + line_future, 25) = 0
        

        tmp_underlying = extract_data_to_insert_in_aim_view(i)(1)
                        
        For m = 0 To UBound(underyling_override, 1)
            tmp_underlying = Replace(tmp_underlying, underyling_override(m)(0), underyling_override(m)(1))
        Next m
        
        
        matrix_aim_view_futures(line_future + 1, 2) = extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_futures(line_future + 1, 3) = tmp_underlying
        matrix_aim_view_futures(line_future + 1, 4) = extract_data_to_insert_in_aim_view(i)(5)
        matrix_aim_view_futures(line_future + 1, 6) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_futures(line_future + 1, 7) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_futures(line_future + 1, 8) = 0
        matrix_aim_view_futures(line_future + 1, 9) = 0
        matrix_aim_view_futures(line_future + 1, 10) = 0
        matrix_aim_view_futures(line_future + 1, 11) = aim_account
        matrix_aim_view_futures(line_future + 1, 12) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_futures(line_future + 1, 13) = aim_account & "_" & tmp_underlying
        matrix_aim_view_futures(line_future + 1, 14) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_futures(line_future + 1, 15) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_futures(line_future + 1, 16) = 0
        matrix_aim_view_futures(line_future + 1, 17) = 0
        matrix_aim_view_futures(line_future + 1, 18) = 0
        matrix_aim_view_futures(line_future + 1, 19) = extract_data_to_insert_in_aim_view(i)(4)
        matrix_aim_view_futures(line_future + 1, 20) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_futures(line_future + 1, 21) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & tmp_underlying
        matrix_aim_view_futures(line_future + 1, 22) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_futures(line_future + 1, 23) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_futures(line_future + 1, 24) = 0
        matrix_aim_view_futures(line_future + 1, 25) = 0
        
        
        line_future = line_future + 1
        
    ElseIf InStr(extract_data_to_insert_in_aim_view(i)(2), "Option") <> 0 Then
        
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 2) = extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 3) = extract_data_to_insert_in_aim_view(i)(1)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 4) = extract_data_to_insert_in_aim_view(i)(5)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 5) = extract_data_to_insert_in_aim_view(i)(5)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 8) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 9) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 10) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 11) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 12) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 13) = ""
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 14) = aim_account
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 15) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 16) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(1)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 17) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 18) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 19) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 20) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 21) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 22) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 23) = extract_data_to_insert_in_aim_view(i)(4)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 24) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 25) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(1)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 26) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 27) = extract_data_to_insert_in_aim_view(i)(3)
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 28) = 0
'        Worksheets(aim_view_options).Cells(l_header_aim_view + 1 + line_option, 29) = 0
        
        tmp_underlying = extract_data_to_insert_in_aim_view(i)(1)
                        
        For m = 0 To UBound(underyling_override, 1)
            tmp_underlying = Replace(tmp_underlying, underyling_override(m)(0), underyling_override(m)(1))
        Next m
        
        matrix_aim_view_options(line_option + 1, 2) = extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_options(line_option + 1, 3) = tmp_underlying
        matrix_aim_view_options(line_option + 1, 4) = extract_data_to_insert_in_aim_view(i)(5)
        matrix_aim_view_options(line_option + 1, 5) = extract_data_to_insert_in_aim_view(i)(5)
        matrix_aim_view_options(line_option + 1, 8) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_options(line_option + 1, 9) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_options(line_option + 1, 10) = 0
        matrix_aim_view_options(line_option + 1, 11) = 0
        matrix_aim_view_options(line_option + 1, 12) = 0
        matrix_aim_view_options(line_option + 1, 13) = ""
        matrix_aim_view_options(line_option + 1, 14) = aim_account
        matrix_aim_view_options(line_option + 1, 15) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_options(line_option + 1, 16) = aim_account & "_" & tmp_underlying
        matrix_aim_view_options(line_option + 1, 17) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_options(line_option + 1, 18) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_options(line_option + 1, 19) = 0
        matrix_aim_view_options(line_option + 1, 20) = 0
        matrix_aim_view_options(line_option + 1, 21) = 0
        matrix_aim_view_options(line_option + 1, 22) = 0
        matrix_aim_view_options(line_option + 1, 23) = extract_data_to_insert_in_aim_view(i)(4)
        matrix_aim_view_options(line_option + 1, 24) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & extract_data_to_insert_in_aim_view(i)(0)
        matrix_aim_view_options(line_option + 1, 25) = aim_account & "_" & extract_data_to_insert_in_aim_view(i)(4) & "_" & tmp_underlying
        matrix_aim_view_options(line_option + 1, 26) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_options(line_option + 1, 27) = extract_data_to_insert_in_aim_view(i)(3)
        matrix_aim_view_options(line_option + 1, 28) = 0
        matrix_aim_view_options(line_option + 1, 29) = 0
        
        
        line_option = line_option + 1
    Else
        debug_test = "bug"
        
    End If
    
    
Next i


If line_equity > 0 Then
    Worksheets(aim_view_equities).Range("A" & l_header_aim_view + 1 & ":" & xlColumnValue(UBound(matrix_aim_view_equities, 2)) & l_header_aim_view + UBound(matrix_aim_view_equities, 1)) = matrix_aim_view_equities
End If

If line_future > 0 Then
    Worksheets(aim_view_futures).Range("A" & l_header_aim_view + 1 & ":" & xlColumnValue(UBound(matrix_aim_view_futures, 2)) & l_header_aim_view + UBound(matrix_aim_view_futures, 1)) = matrix_aim_view_futures
End If

If line_option > 0 Then
    Worksheets(aim_view_options).Range("A" & l_header_aim_view + 1 & ":" & xlColumnValue(UBound(matrix_aim_view_options, 2)) & l_header_aim_view + UBound(matrix_aim_view_options, 1)) = matrix_aim_view_options
End If


datetime_end = Now()
Debug.Print Application.RoundDown(1440 * (datetime_end - datetime_start), 0.1) & " minute(s) and " & CInt(60 * (1440 * (datetime_end - datetime_start) - Application.RoundDown(1440 * (datetime_end - datetime_start), 0.1))) & " seconds"


End Sub


Private Sub aim_autocomplete_formula_with_link_db_pictet_kronos()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim settings_view As Variant
settings_view = Array(Array(aim_view_code.EOD, 3000), Array(aim_view_code.equities, 750), Array(aim_view_code.futures, 15), Array(aim_view_code.Options, 600))
'settings_view = Array(Array(aim_view_code.EOD, 30), Array(aim_view_code.equities, 750), Array(aim_view_code.futures, 15), Array(aim_view_code.Options, 600))

loop_to_check_sufficient_formula:
Application.Calculation = xlCalculationManual

Dim tmp_date As Date

Dim tmp_sheet As String
Dim tmp_range As Range
Dim tmp_last_line_data As Integer

Dim need_to_insert_formula As Boolean, need_new_passage_through_the_loop As Boolean


For i = 0 To UBound(settings_view, 1)
new_passage_through_the_loop:
    
    tmp_sheet = aim_get_worksheet_xls_name_from_view_code(settings_view(i)(0))
        Debug.Print "aim_autocomplete_formula_view: " & "#sheet=" & tmp_sheet
    need_to_insert_formula = False
    need_new_passage_through_the_loop = False
    
    
    If tmp_sheet = "AIM_EOD_JSO" Then
        If Worksheets("Parametres").Cells(16, 19) <> "" And Worksheets("Parametres").Cells(16, 19) <> 0 Then
            
            If Worksheets("Parametres").Cells(16, 19) = 1 Then
                Call aim_static_EOD
                GoTo bypass_auto_formula
            ElseIf IsDate(Worksheets("Parametres").Cells(16, 19)) Then
                tmp_date = Worksheets("Parametres").Cells(16, 19)
                
                If tmp_date < Date Then
                    Call aim_static_EOD
                End If
                
                GoTo bypass_auto_formula
                
            End If
            
        Else
        End If
    End If
    
    
    For j = l_header_aim_view To 32000
        
        If j = l_header_aim_view + 1 Then
            Set tmp_range = Worksheets(tmp_sheet).Cells(j, 2)
            
            If tmp_range.HasFormula = False Then
                'la derniere mise a jour etait avec les donnees static
                tmp_last_line_data = l_header_aim_view
                
                'destruction de toute les donnnes
                Worksheets(tmp_sheet).Range("A" & l_header_aim_view + 1 & ":IV8000").Clear
                
                Exit For
            End If
        End If
        
        If Worksheets(tmp_sheet).Cells(j, 2).Value = "" Then
            'last line de data
            tmp_last_line_data = j - 1
            
            Debug.Print "aim_autocomplete_formula_view: " & "#sheet=" & tmp_sheet & ", #actual_last_line_with_data" & tmp_last_line_data
            
            Exit For
        End If
    Next j
    
    
    
    For j = tmp_last_line_data + 1 To tmp_last_line_data + settings_view(i)(1)
        Set tmp_range = Worksheets(tmp_sheet).Cells(j, 2)
        
        'Debug.Print "aim_autocomplete_formula_view: " & "#sheet=" & tmp_sheet & ", #check_cell_formula_line=" & j
        
        If tmp_range.HasFormula = False Then
            need_to_insert_formula = True
            
            Worksheets(tmp_sheet).Cells(j, 1) = j - l_header_aim_view
            
            For m = 2 To 250
                If Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "" Then
                    Exit For
                Else
                    If InStr(UCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, m)), "CUSTOM") <> 0 Then
                        
                        If Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_buidt" Then
                            
                            c_sheet_identifier = 0
                            c_sheet_aim_account = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "identifier") <> 0 Then
                                    c_sheet_identifier = n
                                ElseIf InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "aim_account") <> 0 Then
                                    c_sheet_aim_account = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=IF(RC" & c_sheet_identifier & "<>"""",RC" & c_sheet_aim_account & " & ""_"" & RC" & c_sheet_identifier & ","""")"
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_buidt_underlying" Then
                            
                            c_sheet_uid = 0
                            c_sheet_aim_account = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "under_product_id") <> 0 Then
                                    c_sheet_uid = n
                                ElseIf InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "aim_account") <> 0 Then
                                    c_sheet_aim_account = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=IF(RC" & c_sheet_uid & "<>"""",RC" & c_sheet_aim_account & " & ""_"" & RC" & c_sheet_uid & ","""")"
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_buidt_current_qty" Then
                            
                            c_sheet_qty_current = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "qty_current") <> 0 Then
                                    c_sheet_qty_current = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_qty_current
                        
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_buidt_close_qty" Then
                            
                            c_sheet_qty_yesterday_close = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "qty_yesterday_close") <> 0 Then
                                    c_sheet_qty_yesterday_close = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_qty_yesterday_close
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_buidt_close_price" Then
                            
                            c_sheet_close_price = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_price") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "vendor_close_price") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_close_local") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_mid_price") <> 0 Then
                                    c_sheet_close_price = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_close_price
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_buidt_ytd_pnl_local_net" Then
                            
                            c_sheet_ytd_pnl = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "ytd_pnl_local_gross") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "vendor_close_price") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_close_local") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_mid_price") <> 0 Then
                                    c_sheet_ytd_pnl = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_ytd_pnl
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_buidt_net_cash_local_with_comm" Then
                            
                            c_sheet_ntcf = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "net_cash_local_with_comm") <> 0 Then
                                    c_sheet_ntcf = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_ntcf
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_tra_cost_total" Then
                            
                            c_sheet_total_comm = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "tra_cost_total") <> 0 Then
                                    c_sheet_total_comm = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_total_comm
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_intraday_commission_local" Then
                            
                            c_sheet_total_comm = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "intraday_commission_local") <> 0 Then
                                    c_sheet_total_comm = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_total_comm
                        
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid" Then
                            
                            c_sheet_aim_account = 0
                            c_sheet_aim_pb = 0
                            c_sheet_product_id = 0
                            
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "identifier") <> 0 Then
                                    c_sheet_product_id = n
                                ElseIf InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "aim_account") <> 0 Then
                                    c_sheet_aim_account = n
                                ElseIf InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "aim_prime_broker") <> 0 Then
                                    c_sheet_aim_pb = n
                                End If
                            Next n
                            
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_aim_account & " & ""_"" & RC" & c_sheet_aim_pb & " & ""_"" & RC" & c_sheet_product_id
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_underlying" Then
                            
                            c_sheet_aim_account = 0
                            c_sheet_aim_pb = 0
                            c_sheet_underlying_id = 0
                            
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "under_product_id") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "identifier") <> 0 Then
                                    c_sheet_underlying_id = n
                                ElseIf InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "aim_account") <> 0 Then
                                    c_sheet_aim_account = n
                                ElseIf InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "aim_prime_broker") <> 0 Then
                                    c_sheet_aim_pb = n
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_aim_account & " & ""_"" & RC" & c_sheet_aim_pb & " & ""_"" & RC" & c_sheet_underlying_id
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_qty_current" Then
                            
                            c_sheet_qty_current = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "qty_current") <> 0 Then
                                    c_sheet_qty_current = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_qty_current
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_qty_yesterday_close" Or Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_close_qty" Then
                            
                            c_sheet_qty_close = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "qty_yesterday_close") <> 0 Then
                                    c_sheet_qty_close = n
                                    Exit For
                                End If
                            Next n
                            
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_qty_close
                        
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_close_price" Or Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_vendor_close_price" Or Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_yesterday_close_local" Or Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_yesterday_mid_price" Then
                            
                            c_sheet_close_price = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_price") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "vendor_close_price") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_close_local") <> 0 Or InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "yesterday_mid_price") <> 0 Then
                                    c_sheet_close_price = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_close_price
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_ytd_pnl_local_net" Then
                            
                            c_sheet_ytd_pnl = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "ytd_pnl_local_gross") <> 0 Then
                                    c_sheet_ytd_pnl = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_ytd_pnl
                        
                        ElseIf Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "CUSTOM_accpbpid_net_cash_local_with_comm" Then
                            
                            c_sheet_ntcf = 0
                            
                            For n = 1 To m - 1
                                If InStr(LCase(Worksheets(tmp_sheet).Cells(l_header_aim_view, n)), "net_cash_local_with_comm") <> 0 Then
                                    c_sheet_ntcf = n
                                    Exit For
                                End If
                            Next n
                            
                            Worksheets(tmp_sheet).Cells(j, m).Value = "=RC" & c_sheet_ntcf
                            
                        End If
                        
                    Else ' il s agit d un champ de la base kronos pictet
                        Worksheets(tmp_sheet).Cells(j, m).Value = Replace("=RTD(""Kronos.RTDServer"";"""";R1C2;R" & l_header_aim_view & "C;RC1)", ";", ",")
                        
                        If m = 2 Then
                            Worksheets(tmp_sheet).Cells(j, m).Calculate
                            
                            If Worksheets(tmp_sheet).Cells(j, m).Value <> "" Then
                                If (j - tmp_last_line_data) * 1.3 > settings_view(i)(1) Then
                                    need_new_passage_through_the_loop = True
                                End If
                            End If
                            
                        End If
                    End If
                End If
            Next m
        End If
        
    Next j
    
    
    If need_new_passage_through_the_loop = True Then
        GoTo new_passage_through_the_loop
        Debug.Print "aim_autocomplete_formula_view: " & "#sheet=" & tmp_sheet & " $new_passage_throug_the_loop"
    Else
        Debug.Print "aim_autocomplete_formula_view: " & "#sheet=" & tmp_sheet & " $empty unecessary formulas"
        For j = tmp_last_line_data + settings_view(i)(1) + 1 To 5000
            Set tmp_range = Worksheets(tmp_sheet).Cells(j, 2)
            
            If tmp_range.HasFormula = False Then
                Exit For
            Else
                For m = 2 To 250
                    If Worksheets(tmp_sheet).Cells(l_header_aim_view, m) = "" Then
                        Exit For
                    Else
                        Worksheets(tmp_sheet).Cells(j, m).Value = ""
                    End If
                Next m
            End If
            
        Next j
    End If
    
bypass_auto_formula:
Next i

End Sub



Public Function aim_get_underyling_spot_ticker(ByVal index_ticker As String) As String

aim_get_underyling_spot_ticker = ""


Dim oBBG As New cls_Bloomberg_Sync
Dim data_bbg As Variant
data_bbg = oBBG.bdp(Array(index_ticker), Array("UNDL_SPOT_TICKER"), output_format.of_vec_without_header)

If Left(data_bbg(0)(0), 1) <> "#" Then
    aim_get_underyling_spot_ticker = "EI09" & data_bbg(0)(0)
End If



End Function


Public Function aim_need_sync_option(Optional ByVal diy As Boolean = False) As Boolean

Debug.Print "aim_need_sync_option: #do-it-yourself=" & diy

aim_need_sync_option = False

If Worksheets("Parametres").Cells(11, 10) <> "" And Worksheets("Parametres").Cells(11, 10) = False Then
    Debug.Print "aim_need_sync_option: $aim_sync_product_with_open already running ! -> Exit"
    Exit Function
End If

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer





Dim update_threshold As Double
update_threshold = 60 'en secondes
    Debug.Print "aim_need_sync_option: update threshold check set to " & update_threshold & " seconds"

Debug.Print "aim_need_sync_option: $aim_autocomplete_formula_view"
Call aim_autocomplete_formula_view


Dim date_last_update As Date


'check si date last update est rempli sinon insere now()-1min
If Worksheets("Parametres").Cells(11, 10) = "" Or IsDate(Worksheets("Parametres").Cells(11, 10)) = False Then
    Debug.Print "aim_need_sync_option: insert date in parametres because not present or broken"
    Worksheets("Parametres").Cells(11, 10) = Now() - 1
End If

date_last_update = Worksheets("Parametres").Cells(11, 10)
    Debug.Print "aim_need_sync_option: get last check date. Result=" & FormatDateTime(date_last_update, vbShortDate) & " " & FormatDateTime(date_last_update, vbLongTime)


If ActiveSheet.name = "Open" And date_last_update < Now - (update_threshold / 86400) Then
    
    Debug.Print "aim_need_sync_option: last check too old. Check if new produts are in aim."
    
    Debug.Print "aim_need_sync_option: Set last_check_date to false to bypass $aim_need_sync_option"
    Worksheets("Parametres").Cells(11, 10) = False 'lock aim_need_sync_option
    
    'un check est lance pour voir s il est necessaire de faire un sync option
    Dim vec_entries_in_views As Variant
    vec_entries_in_views = aim_get_product_in_view(Array(aim_view_code.equities, aim_view_code.Options, aim_view_code.futures))
    
    
    'construction d'un mono vecteur
    k = 0
    Dim vec_product_in_aim() As Variant
    For i = 0 To UBound(vec_entries_in_views, 1)
        If IsArray(vec_entries_in_views(i)) Then
            
            For j = 0 To UBound(vec_entries_in_views(i), 1)
                ReDim Preserve vec_product_in_aim(k)
                vec_product_in_aim(k) = vec_entries_in_views(i)(j)
                k = k + 1
            Next j
            
        End If
    Next i
    
    Dim vec_product_no_in_open As Variant
    vec_product_no_in_open = aim_get_vec_product_not_present_open(vec_product_in_aim)
    
    If IsArray(vec_product_no_in_open) Then
        
        Debug.Print "aim_need_sync_option: New products detected."
        
        aim_need_sync_option = True
        
        If diy = True Then
            Debug.Print "aim_need_sync_option: As do it yourself is set to true, run aim_sync_product."
            Debug.Print "aim_need_sync_option: $aim_sync_product_with_open"
            Call aim_sync_product_with_open
        End If
    Else
        Debug.Print "aim_need_sync_option: Nothing to do, all products are already in Open."
    End If
    
    Debug.Print "aim_need_sync_option: Update last check date. Set to=" & FormatDateTime(Now(), vbLongDate)
    Worksheets("Parametres").Cells(11, 10) = Now()
    
Else
    Debug.Print "aim_need_sync_option: Too quick refresh"
End If



End Function


Public Sub aim_sync_product_with_open()

Debug.Print "aim_sync_product_with_open: #NO INPUT"

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Debug.Print "aim_sync_product_with_open: $aim_autocomplete_formula_view"
Call aim_autocomplete_formula_view

Dim vec_entries_in_views As Variant
vec_entries_in_views = aim_get_product_in_view(Array(aim_view_code.equities, aim_view_code.Options, aim_view_code.futures))

    Dim vec_entries_need_to_be_in_open() As Variant
    k = 0
    For i = 0 To UBound(vec_entries_in_views, 1)
        If IsArray(vec_entries_in_views(i)) Then
            For j = 0 To UBound(vec_entries_in_views(i), 1)
                ReDim Preserve vec_entries_need_to_be_in_open(k)
                vec_entries_need_to_be_in_open(k) = vec_entries_in_views(i)(j)
                k = k + 1
            Next j
        Else
        End If
    Next i



'distinct ticker pour equity DB
Dim vec_distinct_underlying_id_equity() As Variant, vec_distinct_underlying_id_index() As Variant, vec_distinct_future() As Variant, count_future As Integer
    ReDim vec_distinct_underlying_id_equity(0)
    vec_distinct_underlying_id_equity(0) = ""
    
    ReDim vec_distinct_underlying_id_index(0)
    vec_distinct_underlying_id_index(0) = ""
    
    count_future = 0
    


Dim count_underlying_equity As Integer, count_unterlying_index As Integer
count_underlying_equity = 0
count_unterlying_index = 0

Dim data_bbg_produt_without_ticker As Variant

k = 0
p = 0
For i = 0 To UBound(vec_entries_in_views, 1)
    If IsEmpty(vec_entries_in_views(i)) Then
    Else
        For j = 0 To UBound(vec_entries_in_views(i), 1)
            
check_underlying_type:
            
            'repere les fut
            If vec_entries_in_views(i)(j)(3) = aim_instrument_type.future Then
                ReDim Preserve vec_distinct_future(count_future)
                vec_distinct_future(count_future) = vec_entries_in_views(i)(j)
                count_future = count_future + 1
                
'                For m = 0 To UBound(vec_distinct_underlying_id_index, 1)
'                    If vec_entries_in_views(i)(j)(1) = vec_distinct_underlying_id_index(m) Then
'                        Exit For
'                    Else
'                        If m = UBound(vec_distinct_underlying_id_index, 1) Then
'                            If m = 0 And vec_distinct_underlying_id_index(0) = "" Then
'                                vec_distinct_underlying_id_index(count_unterlying_index) = vec_entries_in_views(i)(j)(1) 'underlying
'                                count_unterlying_index = count_unterlying_index + 1
'                            Else
'                                ReDim Preserve vec_distinct_underlying_id_index(count_unterlying_index)
'                                vec_distinct_underlying_id_index(count_unterlying_index) = vec_entries_in_views(i)(j)(1) 'underyling
'                                count_unterlying_index = count_unterlying_index + 1
'                            End If
'                        End If
'                    End If
'                Next m
                
                
                
            End If
            
            If InStr(UCase(vec_entries_in_views(i)(j)(2)), UCase("equity")) <> 0 Or InStr(UCase(vec_entries_in_views(i)(j)(2)), UCase("equitiy")) <> 0 Then
                
                vec_entries_in_views(i)(j)(2) = Replace(UCase(vec_entries_in_views(i)(j)(2)), "EQUITIY", "Equity")
                
                For m = 0 To UBound(vec_distinct_underlying_id_equity, 1)
                    If vec_entries_in_views(i)(j)(1) = vec_distinct_underlying_id_equity(m) Then
                        Exit For
                    Else
                        If m = UBound(vec_distinct_underlying_id_equity, 1) Then
                            If m = 0 And vec_distinct_underlying_id_equity(0) = "" Then
                                vec_distinct_underlying_id_equity(count_underlying_equity) = vec_entries_in_views(i)(j)(1) 'underlying
                                count_underlying_equity = count_underlying_equity + 1
                            Else
                                ReDim Preserve vec_distinct_underlying_id_equity(count_underlying_equity)
                                vec_distinct_underlying_id_equity(count_underlying_equity) = vec_entries_in_views(i)(j)(1) 'underyling
                                count_underlying_equity = count_underlying_equity + 1
                            End If
                        End If
                    End If
                Next m
            ElseIf InStr(UCase(vec_entries_in_views(i)(j)(2)), UCase("index")) <> 0 Then
                
                If vec_entries_in_views(i)(j)(1) = "" Then 'underyling id vide car pam db n a pas la valeur
                    vec_entries_in_views(i)(j)(1) = aim_get_underyling_spot_ticker(vec_entries_in_views(i)(j)(2))
                End If
                
                For m = 0 To UBound(vec_distinct_underlying_id_index, 1)
                    If vec_entries_in_views(i)(j)(1) = vec_distinct_underlying_id_index(m) Then
                        Exit For
                    Else
                        If m = UBound(vec_distinct_underlying_id_index, 1) And Left(vec_entries_in_views(i)(j)(1), 4) = "EI09" Then
                            If m = 0 And vec_distinct_underlying_id_index(0) = "" Then
                                vec_distinct_underlying_id_index(count_unterlying_index) = vec_entries_in_views(i)(j)(1) 'underlying
                                count_unterlying_index = count_unterlying_index + 1
                            Else
                                ReDim Preserve vec_distinct_underlying_id_index(count_unterlying_index)
                                vec_distinct_underlying_id_index(count_unterlying_index) = vec_entries_in_views(i)(j)(1) 'underyling
                                count_unterlying_index = count_unterlying_index + 1
                            End If
                        End If
                    End If
                Next m
            ElseIf vec_entries_in_views(i)(j)(2) = "" Then
                
                'tente un appel bloomberg
                data_bbg_produt_without_ticker = aim_get_bloomberg_data(Array("/buid/" & vec_entries_in_views(i)(j)(0)), Array("PARSEKYABLE_DES"))
                
                If Left(data_bbg_produt_without_ticker(0)(0), 1) <> "#" And (InStr(UCase(data_bbg_produt_without_ticker(0)(0)), "EQUITY") <> 0 Or InStr(UCase(data_bbg_produt_without_ticker(0)(0)), "INDEX") <> 0) Then
                    vec_entries_in_views(i)(j)(2) = data_bbg_produt_without_ticker(0)(0)
                    'repasse dans la boucle
                    GoTo check_underlying_type
                End If
                
            Else
                
                
                
            End If
        Next j
    End If
Next i


Dim vec_entries_alreaday_in_equity_db As Variant
vec_entries_alreaday_in_equity_db = aim_get_product_from_equity_db

Dim vec_new_entry_to_create_equity_db As Variant
vec_new_entry_to_create_equity_db = aim_get_vec_product_not_present_local_db(vec_distinct_underlying_id_equity, vec_entries_alreaday_in_equity_db)


'creation des entrees dans equity database
If count_underlying_equity > 0 Then
    
    If IsArray(vec_new_entry_to_create_equity_db) Then
        If vec_new_entry_to_create_equity_db(0) <> "" Then
            
            Debug.Print "aim_sync_product_with_open: missing product in equity_db, need to create " & UBound(vec_new_entry_to_create_equity_db, 1) + 1 & " entries. $aim_new_entry_equity_db"
            
            Call aim_new_entry_equity_db(vec_new_entry_to_create_equity_db)
        End If
    End If
End If



'creation des index - A DEV une fois que le process de calcul de pnl sera en place
Dim vec_entries_alreaday_in_index_db As Variant
vec_entries_alreaday_in_equity_db = aim_get_product_from_index_db()

Dim vec_new_entry_to_create_index_db As Variant
vec_new_entry_to_create_index_db = aim_get_vec_product_not_present_local_db(vec_distinct_underlying_id_index, vec_entries_alreaday_in_equity_db)
If IsArray(vec_new_entry_to_create_index_db) Then
    
    
    Call aim_new_entry_index_db(vec_new_entry_to_create_index_db)
End If


'ouverture des future
If count_future > 0 Then
    
    Dim vec_future_need_to_be_insert_in_internal_db() As Variant
    Dim count_future_need_to_be_insert_in_internal_db As Integer
    count_future_need_to_be_insert_in_internal_db = 0
    
    Dim vec_future_already_in_internal_db As Variant
    vec_future_already_in_internal_db = aim_mount_data_internal_db_future()
    
    For i = 0 To UBound(vec_distinct_future, 1)
        If IsArray(vec_future_already_in_internal_db) Then
            
            For j = 0 To UBound(vec_future_already_in_internal_db, 1)
                If vec_future_already_in_internal_db(j)(0) = vec_distinct_future(i)(0) Then
                    Exit For
                Else
                    If j = UBound(vec_future_already_in_internal_db, 1) Then
                        ReDim Preserve vec_future_need_to_be_insert_in_internal_db(count_future_need_to_be_insert_in_internal_db)
                        vec_future_need_to_be_insert_in_internal_db(count_future_need_to_be_insert_in_internal_db) = vec_distinct_future(i)(0)
                        count_future_need_to_be_insert_in_internal_db = count_future_need_to_be_insert_in_internal_db + 1
                    End If
                End If
            Next j
            
        Else
            'pas de future, regarde au niveau underlying
            Dim vec_index_already_in_internal_db As Variant
            vec_index_already_in_internal_db = aim_mount_data_internal_db_index()
            
            If IsArray(vec_index_already_in_internal_db) Then
                For j = 0 To UBound(vec_index_already_in_internal_db, 1)
                    If vec_index_already_in_internal_db(j)(0) = vec_distinct_future(i)(1) Then
                        ReDim Preserve vec_future_need_to_be_insert_in_internal_db(count_future_need_to_be_insert_in_internal_db)
                        vec_future_need_to_be_insert_in_internal_db(count_future_need_to_be_insert_in_internal_db) = vec_distinct_future(i)(0)
                        count_future_need_to_be_insert_in_internal_db = count_future_need_to_be_insert_in_internal_db + 1
                        Exit For
                    Else
                        If j = UBound(vec_index_already_in_internal_db, 1) Then
                            MsgBox ("unable to find underlying index: " & vec_distinct_future(i)(1))
                        End If
                    End If
                Next j
            Else
                MsgBox ("No index in index db. Unable to setup futures")
                Exit For
            End If
            
        End If
    Next i

    
    If count_future_need_to_be_insert_in_internal_db > 0 Then
        Debug.Print "aim_sync_product_with_open: missing future in index_db " & UBound(vec_future_need_to_be_insert_in_internal_db, 1) + 1 & " entries. $aim_new_entry_future_db"
        Call aim_new_entry_future_db(vec_future_need_to_be_insert_in_internal_db)
    End If
    
End If


'clean open
Call aim_clean_open_dead_line


'load les positions dans open
Dim vec_entries_to_insert_in_open As Variant
vec_entries_to_insert_in_open = aim_get_vec_product_not_present_open(vec_entries_need_to_be_in_open)


If IsArray(vec_entries_to_insert_in_open) Then
    Debug.Print "aim_sync_product_with_open: insert " & UBound(vec_entries_need_to_be_in_open, 1) + 1 & " entries in open. $aim_insert_new_position_in_open"
    Call aim_insert_new_position_in_open(vec_entries_to_insert_in_open)
End If

'reorder view
Call aim_reorder_open

Call aim_prepare_stats_derivatives_multi_account

Application.Calculation = xlCalculationAutomatic

End Sub


Public Function aim_get_bloomberg_data(ByVal vec_tickers As Variant, ByVal vec_fields As Variant) As Variant

Dim oBBG As New cls_Bloomberg_Sync
aim_get_bloomberg_data = oBBG.bdp(vec_tickers, vec_fields, output_format.of_vec_without_header)

End Function


Public Sub aim_roll_future()

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer


Dim vec_future_actually_in_db()
vec_future_actually_in_db = aim_mount_data_internal_db_future


If IsArray(vec_future_actually_in_db) Then
    
    For i = 0 To UBound(vec_future_actually_in_db(0), 1)
        If vec_future_actually_in_db(0)(i) = "product_id" Then
            dim_vec_fut_product_id = i
        ElseIf vec_future_actually_in_db(0)(i) = "underlying_id" Then
            dim_vec_fut_underlying_id = i
        ElseIf vec_future_actually_in_db(0)(i) = "ticker" Then
            dim_vec_fut_ticker = i
        ElseIf vec_future_actually_in_db(0)(i) = "expiry_date" Then
            
        ElseIf vec_future_actually_in_db(0)(i) = "contract_size" Then
            
        ElseIf vec_future_actually_in_db(0)(i) = "line" Then
            
        End If
    Next i
    
    For i = 0 To UBound(vec_future_actually_in_db, 1)
        'doit-il est desactiver
    Next i

End If

End Sub


Public Sub aim_clean_dead_options_line()

Dim i As Integer, j As Integer, k As Integer

Dim open_last_line_threshold As Integer
open_last_line_threshold = 25

Dim is_last_line As Boolean


For i = 25 To 32000
    
    is_last_line = True
    
    For j = 0 To open_last_line_threshold
        If Worksheets("Open").Cells(i + j, 1) <> "" Then
            is_last_line = False
            Exit For
        End If
    Next j
    
    If is_last_line = True Then
        Exit For
    Else
        
        If Worksheets("Open").Cells(i, 1) <> "" Then
            If IsError(Worksheets("Open").Cells(i, 9)) And Worksheets("Open").Cells(i, 23) < Date Then
                Worksheets("Open").rows(i).Clear
            End If
        End If
    End If
    
Next i


End Sub


Public Sub aim_clean_open_dead_line()

Dim i As Integer, j As Integer, k As Integer

Dim open_last_line_threshold As Integer
open_last_line_threshold = 25

Dim is_last_line As Boolean


For i = 25 To 32000
    
    is_last_line = True
    
    For j = 0 To open_last_line_threshold
        If Worksheets("Open").Cells(i + j, 1) <> "" Then
            is_last_line = False
            Exit For
        End If
    Next j
    
    If is_last_line = True Then
        Exit For
    Else
        If Worksheets("Open").Cells(i, 1) <> "" Then
            
            If IsError(Worksheets("Open").Cells(i, 9)) And IsDate(Worksheets("Open").Cells(i, 9)) Then
                
                If aim_get_working_mode() = aim_mode.rtd_auto_db_pictet Or aim_get_working_mode() = "" Then
                
                    If Worksheets("Open").Cells(i, 9) < Date Then
                        Worksheets("Open").rows(i).Clear
                    Else
                        If aim_get_working_mode = aim_mode.rtd_auto_db_pictet Then
                            Worksheets("Open").rows(i).Clear
                        End If
                    End If
                
                ElseIf aim_get_working_mode() = aim_mode.manual_import_mav_excel_export Then
                    
                    Worksheets("Open").rows(i).Clear
                
                End If
            End If
            
'            'si qty null
'            If Worksheets("Open").Cells(i, 6) = "S" Or Worksheets("Open").Cells(i, 6) = "F" Then
'                If IsError(Worksheets("Open").Cells(i, 29)) = False Then
'                    If Worksheets("Open").Cells(i, 29) = 0 Then
'                        Worksheets("Open").rows(i).Clear
'                    End If
'                End If
'            End If
            
            
            'si equity et qu un derive est deja present autour
            If Worksheets("Open").Cells(i, 6) = "S" Then
                If Worksheets("Open").Cells(i - 1, 97) = Worksheets("Open").Cells(i, 97) Or Worksheets("Open").Cells(i + 1, 97) = Worksheets("Open").Cells(i, 97) Then
                    Worksheets("Open").rows(i).Clear
                End If
            End If
            
        End If
    End If
    
Next i

End Sub


Public Function aim_get_product_from_equity_db() As Variant

Debug.Print "aim_get_product_from_equity_db : #NO INPUT"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer

k = 0
Dim vec_product() As Variant
For i = l_header_internal_db_equity + 2 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_product(k)
        vec_product(k) = Worksheets("Equity_Database").Cells(i, 1).Value
        k = k + 1
    End If
Next i

If k = 0 Then
    aim_get_product_from_equity_db = Empty
    Debug.Print "aim_get_product_from_equity_db :" & "OUTPUT " & "@no entry in equity_db"
Else
    aim_get_product_from_equity_db = vec_product
    Debug.Print "aim_get_product_from_equity_db :" & "OUTPUT " & "@vec_underlying_id of " & UBound(vec_product, 1) + 1 & " entries"
End If

End Function


Public Function aim_get_product_from_index_db() As Variant

Debug.Print "aim_get_product_from_index_db: #NO INPUT"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer

k = 0
Dim vec_product() As Variant
For i = l_header_internal_db_index + 2 To 32000 Step 3
    If Worksheets("Index_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_product(k)
        vec_product(k) = Worksheets("Index_Database").Cells(i, 1).Value
        k = k + 1
    End If
Next i

If k = 0 Then
    aim_get_product_from_index_db = Empty
    Debug.Print "aim_get_product_from_index_db: " & "OUTPUT " & "@no entry in index_db"
Else
    aim_get_product_from_index_db = vec_product
    Debug.Print "aim_get_product_from_index_db :" & "OUTPUT " & "@vec_underlying_id of " & UBound(vec_product, 1) + 1 & " entries"
End If

End Function


Public Function aim_get_product_and_underlying_from_open() As Variant

Debug.Print "aim_get_product_and_underlying_from_open: #NO INPUT"

Dim i As Integer, j As Integer, k As Integer

Dim tmp_qty As Double
Dim tmp_buidt As String
Dim tmp_account As String
Dim tmp_pb As String


k = 0
Dim open_last_line_threshold As Integer, is_last_line As Boolean
    open_last_line_threshold = 10
    Debug.Print "aim_get_product_and_underlying_from_open: $last_line_threshold=" & open_last_line_threshold
Dim vec_product() As Variant
For i = l_header_internal_open + 1 To 32000
    
    is_last_line = True
    
    For j = 0 To open_last_line_threshold
        If Worksheets("Open").Cells(i + j, 1) = "" Then
        Else
            is_last_line = False
            Exit For
        End If
    Next j
    
    
    If is_last_line = True Then
        Debug.Print "aim_get_product_and_underlying_from_open: $reachs last line"
        Exit For
    ElseIf is_last_line = False And Worksheets("Open").Cells(i, 1) <> "" Then
        
        tmp_account = ""
        tmp_buidt = Worksheets("Open").Cells(i, 93)
        
        If tmp_buidt <> "" Then
            tmp_account = Left(tmp_buidt, InStr(tmp_buidt, "_") - 1)
        End If
        
        tmp_pb = Worksheets("Open").Cells(i, 91)
        
        
        tmp_qty = 0
        
        ReDim Preserve vec_product(k)
        
        If Worksheets("Open").Cells(i, 6) = "C" Or Worksheets("Open").Cells(i, 6) = "P" Then
            
            If IsError(Worksheets("Open").Cells(i, 9)) = False Then
                If Worksheets("Open").Cells(i, 9) <> "" Then
                    If IsNumeric(Worksheets("Open").Cells(i, 9)) Then
                        tmp_qty = Worksheets("Open").Cells(i, 9)
                    End If
                End If
            End If
            
            'options - distingué sur equity/index
            If InStr(UCase(Worksheets("Open").Cells(i, c_internal_open_ticker_option)), "EQUITY") <> 0 Then
                vec_product(k) = Array(Worksheets("Open").Cells(i, 2).Value, Worksheets("Open").Cells(i, 1).Value, aim_instrument_type.option_equity, tmp_qty, tmp_account, tmp_pb)
                
                If IsError(Worksheets("Open").Cells(i, 29)) = False Then
                    If IsNumeric(Worksheets("Open").Cells(i, 29)) Then
                        k = k + 1
                        tmp_qty = Worksheets("Open").Cells(i, 29)
                        ReDim Preserve vec_product(k)
                        vec_product(k) = Array(Worksheets("Open").Cells(i, 1).Value, Worksheets("Open").Cells(i, 1).Value, aim_instrument_type.equity, tmp_qty, tmp_account, tmp_pb)
                    End If
                End If
                
            ElseIf InStr(UCase(Worksheets("Open").Cells(i, c_internal_open_ticker_option)), "INDEX") <> 0 Then
                vec_product(k) = Array(Worksheets("Open").Cells(i, 2).Value, Worksheets("Open").Cells(i, 1).Value, aim_instrument_type.option_index, tmp_qty, tmp_account, tmp_pb)
            Else
                
            End If
            
            
        
        ElseIf Worksheets("Open").Cells(i, 6) = "F" Then
            
            If IsError(Worksheets("Open").Cells(i, 29)) = False Then
                If IsNumeric(Worksheets("Open").Cells(i, 29)) Then
                    tmp_qty = Worksheets("Open").Cells(i, 29)
                End If
            End If
            
            vec_product(k) = Array(Worksheets("Open").Cells(i, 2).Value, Worksheets("Open").Cells(i, 1).Value, aim_instrument_type.future, tmp_qty, tmp_account, tmp_pb)
        ElseIf Worksheets("Open").Cells(i, 6) = "S" Then
            
            If IsError(Worksheets("Open").Cells(i, 29)) = False Then
                If IsNumeric(Worksheets("Open").Cells(i, 29)) Then
                    tmp_qty = Worksheets("Open").Cells(i, 29)
                End If
            End If
            
            vec_product(k) = Array(Worksheets("Open").Cells(i, 1).Value, Worksheets("Open").Cells(i, 1).Value, aim_instrument_type.equity, tmp_qty, tmp_account, tmp_pb)
        End If
        
        k = k + 1
            
    End If
Next i

If k = 0 Then
    aim_get_product_and_underlying_from_open = Empty
    Debug.Print "aim_get_product_and_underlying_from_open: " & "OUTPUT " & "@no entry in open"
Else
    aim_get_product_and_underlying_from_open = vec_product
    Debug.Print "aim_get_product_and_underlying_from_open: " & "OUTPUT " & "@vec_product_id of " & UBound(vec_product, 1) + 1 & " entries"
End If

End Function


Public Function aim_get_product_from_open() As Variant 'retourne un vecteur de vecteur (product_id, instrument_type, qty, account, pb account)

Debug.Print "aim_get_product_from_open: #NO INPUT"

Dim i As Integer, j As Integer, k As Integer

Dim seperator_buidt As String
    seperator_buidt = "_"

Dim tmp_qty As Double
Dim tmp_buidt As String
Dim tmp_account As String
Dim tmp_pb As String

k = 0
Dim open_last_line_threshold As Integer, is_last_line As Boolean
    open_last_line_threshold = 50
    Debug.Print "aim_get_product_from_open: $last_line_threshold=" & open_last_line_threshold
Dim vec_product() As Variant
For i = l_header_internal_open + 1 To 32000
    
    is_last_line = True
    
    For j = 0 To open_last_line_threshold
        If Worksheets("Open").Cells(i + j, 1) = "" Then
        Else
            is_last_line = False
            Exit For
        End If
    Next j
    
    
    If is_last_line = True Then
        Debug.Print "aim_get_product_from_open: $reachs last line"
        Exit For
    ElseIf is_last_line = False And Worksheets("Open").Cells(i, 1) <> "" Then
        
        tmp_account = ""
        
        
            tmp_buidt = Worksheets("Open").Cells(i, 93)
        
            If tmp_buidt <> "" Then
                tmp_account = Left(tmp_buidt, InStr(tmp_buidt, seperator_buidt) - 1)
            End If
            
            
        tmp_pb = Worksheets("Open").Cells(i, 91)
            
        
        tmp_qty = 0
        
        ReDim Preserve vec_product(k)
        
        If Worksheets("Open").Cells(i, 6) = "C" Or Worksheets("Open").Cells(i, 6) = "P" Then
            
            If IsError(Worksheets("Open").Cells(i, 9)) = False Then
                If Worksheets("Open").Cells(i, 9) <> "" Then
                    If IsNumeric(Worksheets("Open").Cells(i, 9)) Then
                        tmp_qty = Worksheets("Open").Cells(i, 9)
                    End If
                End If
            End If
            
            'options - distingué sur equity/index
            If InStr(UCase(Worksheets("Open").Cells(i, c_internal_open_ticker_option)), "EQUITY") <> 0 Then
                vec_product(k) = Array(Worksheets("Open").Cells(i, 2).Value, aim_instrument_type.option_equity, tmp_qty, tmp_account, tmp_pb)
                
                If IsError(Worksheets("Open").Cells(i, 29)) = False Then
                    If IsNumeric(Worksheets("Open").Cells(i, 29)) Then
                        k = k + 1
                        tmp_qty = Worksheets("Open").Cells(i, 29)
                        ReDim Preserve vec_product(k)
                        vec_product(k) = Array(Worksheets("Open").Cells(i, 1).Value, aim_instrument_type.equity, tmp_qty, tmp_account, tmp_pb)
                    End If
                End If
                
            ElseIf InStr(UCase(Worksheets("Open").Cells(i, c_internal_open_ticker_option)), "INDEX") <> 0 Then
                vec_product(k) = Array(Worksheets("Open").Cells(i, 2).Value, aim_instrument_type.option_index, tmp_qty, tmp_account, tmp_pb)
            Else
                
            End If
            
            
        
        ElseIf Worksheets("Open").Cells(i, 6) = "F" Then
            
            If IsError(Worksheets("Open").Cells(i, 29)) = False Then
                If IsNumeric(Worksheets("Open").Cells(i, 29)) Then
                    tmp_qty = Worksheets("Open").Cells(i, 29)
                End If
            End If
            
            vec_product(k) = Array(Worksheets("Open").Cells(i, 2).Value, aim_instrument_type.future, tmp_qty, tmp_account, tmp_pb)
        ElseIf Worksheets("Open").Cells(i, 6) = "S" Then
            
            If IsError(Worksheets("Open").Cells(i, 29)) = False Then
                If IsNumeric(Worksheets("Open").Cells(i, 29)) Then
                    tmp_qty = Worksheets("Open").Cells(i, 29)
                End If
            End If
            
            vec_product(k) = Array(Worksheets("Open").Cells(i, 1).Value, aim_instrument_type.equity, tmp_qty, tmp_account, tmp_pb)
        End If
        
        k = k + 1
            
    End If
Next i

If k = 0 Then
    aim_get_product_from_open = Empty
    Debug.Print "aim_get_product_from_open: " & "OUTPUT " & "@no entry in open"
Else
    aim_get_product_from_open = vec_product
    Debug.Print "aim_get_product_from_open: " & "OUTPUT " & "@vec_product_id of " & UBound(vec_product, 1) + 1 & " entries"
End If


End Function



Public Function aim_get_product_in_view(ByVal vec_views As Variant) As Variant

If IsArray(vec_views) Then
Else
    If IsEmpty(vec_views) Then
        aim_get_product_in_view = Empty
        Exit Function
    End If
End If

Debug.Print "aim_get_product_in_view: " & "INPUT " & "#vec_views of " & UBound(vec_views, 1) + 1 & " views"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer

Dim output() As Variant
Dim vec_tmp() As Variant

Dim status_assign_column As Variant

k = 0
For i = 0 To UBound(vec_views, 1)
    
    
    
    Dim column_product_id_concern As Integer, column_underlying_id_concern As Integer, column_ticker_concern As Integer, column_aim_account As Integer, column_aim_prime_broker As Integer, column_qty_current As Integer, column_qty_close As Integer, column_ntcf As Integer
    Dim worksheet_concern As String
    
    worksheet_concern = aim_get_worksheet_xls_name_from_view_code(vec_views(i))
    column_product_id_concern = aim_get_product_column_from_view_code(vec_views(i))
    column_underlying_id_concern = aim_get_underlying_id_column_from_view_code(vec_views(i))
    column_ticker_concern = aim_get_view_column_index_from_header(vec_views(i), "bby_code")
    column_aim_account = aim_get_view_column_index_from_header(vec_views(i), "aim_account")
    column_aim_prime_broker = aim_get_view_column_index_from_header(vec_views(i), "aim_prime_broker")
    
    
    'pour eviter de remonter des lignes inutiles, check que qty current, qty close, ntcf
    column_qty_current = aim_get_view_column_index_from_header(vec_views(i), "qty_current")
    column_qty_close = aim_get_view_column_index_from_header(vec_views(i), "qty_yesterday_close")
    column_ntcf = aim_get_view_column_index_from_header(vec_views(i), "net_cash_local_with_comm")
    
    
    Debug.Print "aim_get_product_in_view: " & "$worksheet=" & worksheet_concern & ", $column_product=" & column_product_id_concern & ", $column_underlying=" & column_underlying_id_concern & ", $column_ticker=" & column_ticker_concern
    
    m = 0
    For j = l_header_aim_view + 1 To 10000
        If Worksheets(worksheet_concern).Cells(j, column_product_id_concern) = "" Then
             Debug.Print "aim_get_product_in_view: $last line reachs"
            Exit For
        Else
            If Worksheets(worksheet_concern).Cells(j, column_qty_current) = 0 And Worksheets(worksheet_concern).Cells(j, column_qty_close) = 0 And Worksheets(worksheet_concern).Cells(j, column_ntcf) = 0 Then
            Else
                ReDim Preserve vec_tmp(m)
                vec_tmp(m) = Array(Worksheets(worksheet_concern).Cells(j, column_product_id_concern).Value, Worksheets(worksheet_concern).Cells(j, column_underlying_id_concern).Value, Worksheets(worksheet_concern).Cells(j, column_ticker_concern).Value, -1, Worksheets(worksheet_concern).Cells(j, column_aim_account).Value, Worksheets(worksheet_concern).Cells(j, column_aim_prime_broker).Value)
                
                If vec_views(i) = aim_view_code.equities Then
                    vec_tmp(m)(3) = aim_instrument_type.equity
                ElseIf vec_views(i) = aim_view_code.futures Then
                    vec_tmp(m)(3) = aim_instrument_type.future
                ElseIf vec_views(i) = aim_view_code.Options Then
                    If InStr(UCase(vec_tmp(m)(2)), "EQUITY") <> 0 Then
                        vec_tmp(m)(3) = aim_instrument_type.option_equity
                    ElseIf InStr(UCase(vec_tmp(m)(2)), "INDEX") <> 0 Then
                        vec_tmp(m)(3) = aim_instrument_type.option_index
                    End If
                End If
                
                m = m + 1
            End If
        End If
    Next j
    
    ReDim Preserve output(i)
    If m = 0 Then
        output(i) = Empty
    Else
        output(i) = vec_tmp
    End If
    
    k = k + 1
    
Next i


If k = 0 Then
    aim_get_product_in_view = Empty
    Debug.Print "aim_get_product_in_view: " & "OUTPUT " & "@no entry in open"
Else
    aim_get_product_in_view = output
    Debug.Print "aim_get_product_in_view: " & "OUTPUT " & "@vec_product_id of " & UBound(output, 1) & " entries"
End If


End Function


Private Function aim_get_vec_product_not_present_open(ByVal vec_open_positions_in_aim As Variant) As Variant

aim_get_vec_product_not_present_open = Empty

If IsEmpty(vec_open_positions_in_aim) Then
    Exit Function
End If

Debug.Print "aim_get_vec_product_not_present_open: " & "INPUTS " & "#vec_open_postions_in_aim of " & UBound(vec_open_positions_in_aim, 1) & " entries"


Dim vec_product_already_in_open As Variant
vec_product_already_in_open = aim_get_product_from_open()
Debug.Print "aim_get_vec_product_not_present_open: $aim_get_product_from_open"

If IsEmpty(vec_product_already_in_open) Then
    aim_get_vec_product_not_present_open = vec_open_positions_in_aim
    Exit Function
Else
    
    Debug.Print "aim_get_vec_product_not_present_open: $matching"
    Dim vec_product_not_in_open() As Variant
    k = 0
    For i = 0 To UBound(vec_open_positions_in_aim, 1)
        If vec_open_positions_in_aim(i)(5) <> "N/A" Then
            For j = 0 To UBound(vec_product_already_in_open, 1)
                If vec_open_positions_in_aim(i)(0) = vec_product_already_in_open(j)(0) And vec_open_positions_in_aim(i)(4) = vec_product_already_in_open(j)(3) And vec_open_positions_in_aim(i)(5) = vec_product_already_in_open(j)(4) Then
                    Exit For
                Else
                    If j = UBound(vec_product_already_in_open, 1) Then
                        ReDim Preserve vec_product_not_in_open(k)
                        vec_product_not_in_open(k) = vec_open_positions_in_aim(i)
                        k = k + 1
                    End If
                End If
            Next j
        End If
    Next i
    
    If k > 0 Then
        aim_get_vec_product_not_present_open = vec_product_not_in_open
        Debug.Print "aim_get_vec_product_not_present_open: " & "OUTPUT " & "@vec_product_not_in_open of " & UBound(vec_product_not_in_open, 1) & " entries"
    Else
        aim_get_vec_product_not_present_open = Empty
        Debug.Print "aim_get_vec_product_not_present_open: " & "OUTPUT " & "@all open positions in AIM are already in open"
    End If
    
End If




End Function


'reception d'un vecteur d'underlying_id ou alors d'un vecteur de vecteur (product_id / underlying_id)
Private Function aim_get_vec_product_not_present_local_db(ByVal vec_required_product As Variant, Optional ByVal vec_internal_product As Variant = Empty) As Variant

Debug.Print "aim_get_vec_product_not_present_local_db: " & "INPUTS " & "#vec_required_products, optional #vec_internal_products_already_setup"

Dim i As Integer, j As Integer, k As Integer, m As Integer

aim_get_vec_product_not_present_local_db = Empty

Dim tmp_vec_vec_required_product() As Variant

If IsArray(vec_required_product) Then
    
    'check vecteur de vecteur (product_id, underyling_id)
    k = 0
    For i = 0 To UBound(vec_required_product, 1)
        ReDim Preserve tmp_vec_vec_required_product(k)
        
        If IsArray(vec_required_product(i)) Then
            tmp_vec_vec_required_product(k) = vec_required_product(i)(1)
        Else
            tmp_vec_vec_required_product(k) = vec_required_product(i)
        End If
        
        k = k + 1
    Next i
    
    vec_required_product = tmp_vec_vec_required_product
    
Else
    If vec_required_product = Empty Then
        Debug.Print "aim_get_vec_product_not_present_local_db: @vec_required_products is empty"
        Exit Function
    End If
End If

If IsArray(vec_internal_product) Then
Else
    If vec_internal_product = Empty Then
        
        Dim tmp_vec_internal_product() As Variant
        Dim vec_product_equity_db As Variant, vec_product_index_db As Variant
        vec_product_equity_db = aim_get_product_from_equity_db
        vec_product_index_db = aim_get_product_from_index_db
        
        k = 0
        If IsArray(vec_product_equity_db) Then
            For i = 0 To UBound(vec_product_equity_db, 1)
                ReDim Preserve tmp_vec_internal_product(k)
                tmp_vec_internal_product(k) = vec_product_equity_db(i)
                k = k + 1
            Next i
        End If
        
        If IsArray(vec_product_index_db) Then
            For i = 0 To UBound(vec_product_index_db, 1)
                ReDim Preserve tmp_vec_internal_product(k)
                tmp_vec_internal_product(k) = vec_product_index_db(i)
                k = k + 1
            Next i
        End If
        
        If k = 0 Then
            aim_get_vec_product_not_present_local_db = vec_required_product
            Debug.Print "aim_get_vec_product_not_present_local_db: @no product in internal db"
            Exit Function
        Else
            vec_internal_product = tmp_vec_internal_product
        End If
        
        
        
    End If
End If


k = 0
Dim vec_diff() As Variant
For i = 0 To UBound(vec_required_product, 1)
    For j = 0 To UBound(vec_internal_product, 1)
        If vec_required_product(i) = vec_internal_product(j) Then
            Exit For
        Else
            If j = UBound(vec_internal_product, 1) Then
                ReDim Preserve vec_diff(k)
                vec_diff(k) = vec_required_product(i)
                k = k + 1
            End If
        End If
    Next j
Next i

If k > 0 Then
    aim_get_vec_product_not_present_local_db = vec_diff
    Debug.Print "aim_get_vec_product_not_present_local_db: " & "OUTPUT" & "@vec_products_not_in_local_db of " & UBound(vec_diff, 1) + 1 & " entries"
Else
    Debug.Print "aim_get_vec_product_not_present_local_db: " & "OUTPUT" & "@all products are already in local db"
End If


End Function


Public Function aim_assign_column(ByVal vec_list_view As Variant) As Variant

Debug.Print "aim_assign_column: " & "INPUT " & "#vec_list_aim_views"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer

Dim output() As Variant
Dim vec_tmp() As Variant

For i = 0 To UBound(vec_list_view, 1)
    
    If vec_list_view(i) = aim_view_code.EOD Then
        
        Debug.Print "aim_assign_column: " & "$view EOD"
        
        k = 0
        
        For j = 1 To 250
            If Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = "" Then
                Exit For
            Else
                If Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_product_id Then
                    c_aim_eod_product_id = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_underlying_id Then
                    c_aim_eod_underlying_id = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_instrument_type Then
                    c_aim_eod_instrument_type = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_ticker Then
                    c_aim_eod_ticker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_currency Then
                    c_aim_eod_currency = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_close_qty Then
                    c_aim_eod_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_close_price Then
                    c_aim_eod_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_local_close_ytd_pnl_gross Then
                    c_aim_eod_local_close_ytd_pnl_gross = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_local_close_commission_broker Then
                    c_aim_eod_local_close_commission_broker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_local_close_commission_total Then
                    c_aim_eod_local_close_commission_total = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_local_close_dividend Then
                    c_aim_eod_local_close_dividend = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_local_close_ytd_pnl_net Then
                    c_aim_eod_local_close_ytd_pnl_net = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_aim_account Then
                    c_aim_eod_aim_account = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_buidt Then
                    c_aim_eod_buidt = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_buidt_underlying Then
                    c_aim_eod_buidt_underlying = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_buidt_close_qty Then
                    c_aim_eod_buidt_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_buidt_close_price Then
                    c_aim_eod_buidt_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_built_local_close_ytd_pnl_net Then
                    c_aim_eod_built_local_close_ytd_pnl_net = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_built_close_commission_total Then
                    c_aim_eod_built_close_commission_total = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_aim_prime_broker Then
                    c_aim_eod_aim_prime_broker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_accpbpid Then
                    c_aim_eod_accpbpid = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_accpbpid_underlying Then
                    c_aim_eod_accpbpid_underlying = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_accpbpid_close_qty Then
                    c_aim_eod_accpbpid_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_accpbpid_close_price Then
                    c_aim_eod_accpbpid_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_eod).Cells(l_header_aim_view, j) = c_header_aim_eod_accpbpid_local_close_ytd_pnl Then
                    c_aim_eod_accpbpid_local_close_ytd_pnl = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_eod).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                
                End If
                
            End If
        Next j
        
        ReDim Preserve output(i)
        output(i) = vec_tmp
        
    ElseIf vec_list_view(i) = aim_view_code.equities Then
        
        Debug.Print "aim_assign_column: " & "$view equities"
        
        k = 0
        
        For j = 1 To 250
            If Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = "" Then
                Exit For
            Else
                If Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_product_id Then
                    c_aim_equities_product_id = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_ticker Then
                    c_aim_equities_ticker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_description Then
                    c_aim_equities_description = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_currency Then
                    c_aim_equities_currency = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_current_qty Then
                    c_aim_equities_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_close_qty Then
                    c_aim_equities_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_close_price Then
                    c_aim_equities_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_local_intraday_net_trading_cash_flow Then
                    c_aim_equities_local_intraday_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_local_intraday_commission Then
                    c_aim_equities_local_intraday_commission = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_local_intraday_dividend Then
                    c_aim_equities_local_intraday_dividend = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_aim_account Then
                    c_aim_equities_aim_account = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_aim_prime_broker Then
                    c_aim_equities_aim_prime_broker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_accpbpid Then
                    c_aim_equities_accpbpid = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_accpbpid_underlying Then
                    c_aim_equities_accpbpid_underlying = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_accpbpid_current_qty Then
                    c_aim_equities_accpbpid_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_accpbpid_close_qty Then
                    c_aim_equities_accpbpid_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_accpbpid_close_price Then
                    c_aim_equities_accpbpid_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_equities).Cells(l_header_aim_view, j) = c_header_aim_equities_accpbpid_local_intraday_net_trading_cash_flow Then
                    c_aim_equities_accpbpid_local_intraday_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_equities).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                End If
                
            End If
        Next j
        
        ReDim Preserve output(i)
        output(i) = vec_tmp
        
    ElseIf vec_list_view(i) = aim_view_code.futures Then
        
        Debug.Print "aim_assign_column: " & "$view futures"
        
        k = 0
        
        For j = 1 To 250
            If Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = "" Then
                Exit For
            Else
                If Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_product_id Then
                    c_aim_futures_product_id = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_underlying_id Then
                    c_aim_futures_underlying_id = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_ticker Then
                    c_aim_futures_ticker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_currency Then
                    c_aim_futures_currency = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_current_qty Then
                    c_aim_futures_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_close_qty Then
                    c_aim_futures_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_close_price Then
                    c_aim_futures_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_local_intraday_net_trading_cash_flow Then
                    c_aim_futures_local_intraday_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_local_intraday_commission Then
                    c_aim_futures_local_intraday_commission = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_aim_account Then
                    c_aim_futures_aim_account = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_buidt Then
                    c_aim_futures_buidt = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_buidt_underlying Then
                    c_aim_futures_buidt_underlying = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_buidt_current_qty Then
                    c_aim_futures_buidt_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_buidt_close_qty Then
                    c_aim_futures_buidt_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_buidt_close_price Then
                    c_aim_futures_buidt_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_buidt_local_intrady_net_trading_cash_flow Then
                    c_aim_futures_buidt_local_intrady_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_buidt_local_intraday_commission Then
                    c_aim_futures_buidt_local_intraday_commission = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_aim_prime_broker Then
                    c_aim_futures_aim_prime_broker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_accpbpid Then
                    c_aim_futures_accpbpid = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_accpbpid_underlying Then
                    c_aim_futures_accpbpid_underlying = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_accpbpid_current_qty Then
                    c_aim_futures_accpbpid_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_accpbpid_close_qty Then
                    c_aim_futures_accpbpid_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_accpbpid_close_price Then
                    c_aim_futures_accpbpid_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, j) = c_header_aim_futures_accpbpid_local_intraday_net_trading_cash_flow Then
                    c_aim_futures_accpbpid_local_intraday_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_futures).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                
                End If
                
            End If
        Next j
        
        ReDim Preserve output(i)
        output(i) = vec_tmp
        
    ElseIf vec_list_view(i) = aim_view_code.Options Then
        
        Debug.Print "aim_assign_column: " & "$view options"
        
        k = 0
        
        For j = 1 To 250
            If Worksheets(aim_view_options).Cells(l_header_aim_view, j) = "" Then
                Exit For
            Else
                If Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_product_id Then
                    c_aim_options_product_id = j
                        
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_underlying_id Then
                    c_aim_options_underlying_id = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_ticker_with_yellow_key Then
                    c_aim_options_ticker_with_yellow_key = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_ticker_without_yellow_key Then
                    c_aim_options_ticker_without_yellow_key = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_investment_class Then
                    c_aim_options_investment_class = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_currency Then
                    c_aim_options_currency = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_current_qty Then
                    c_aim_options_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_close_qty Then
                    c_aim_options_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_close_price_mid Then
                    c_aim_options_close_price_mid = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_local_intraday_net_trading_cash_flow Then
                    c_aim_options_local_intraday_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_local_intraday_commission Then
                    c_aim_options_local_intraday_commission = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_last_transaction_code Then
                    c_aim_options_last_transaction_code = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_aim_account Then
                    c_aim_options_aim_account = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_buidt Then
                    c_aim_options_buidt = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_buidt_underlying Then
                    c_aim_options_buidt_underlying = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_buidt_current_qty Then
                    c_aim_options_buidt_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_buidt_close_qty Then
                    c_aim_options_buidt_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_buidt_close_price Then
                    c_aim_options_buidt_close_price = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_buidt_local_intraday_net_trading_cash_flow Then
                    c_aim_options_buidt_local_intraday_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_buidt_local_intraday_commission Then
                    c_aim_options_buidt_local_intraday_commission = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_aim_prime_broker Then
                    c_aim_options_aim_prime_broker = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_accpbpid Then
                    c_aim_options_accpbpid = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_accpbpid_underlying Then
                    c_aim_options_accpbpid_underlying = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_accpbpid_current_qty Then
                    c_aim_options_accpbpid_current_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_accpbpid_close_qty Then
                    c_aim_options_accpbpid_close_qty = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_accpbpid_close_price_mid Then
                    c_aim_options_accpbpid_close_price_mid = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                
                ElseIf Worksheets(aim_view_options).Cells(l_header_aim_view, j) = c_header_aim_options_accpbpid_local_intraday_net_trading_cash_flow Then
                    c_aim_options_accpbpid_local_intraday_net_trading_cash_flow = j
                    
                    ReDim Preserve vec_tmp(k)
                    vec_tmp(k) = Array(Worksheets(aim_view_options).Cells(l_header_aim_view, j).Value, j)
                    k = k + 1
                    
                End If
                
            End If
        Next j
        
        ReDim Preserve output(i)
        output(i) = vec_tmp
        
    End If
    
Next i

aim_assign_column = output

Debug.Print "aim_assign_column: " & "OUTPUT " & "@vec_view_with_columns"

End Function


Sub aim_new_entry_future_db(ByVal identifier As Variant)

If IsArray(identifier) Then
Else
    If IsEmpty(identifier) Then
        Exit Sub
    End If
End If

Debug.Print "aim_new_entry_future_db: " & "INPUT " & "#vec_product_id_future of " & UBound(identifier, 1) & " entries"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

'check le type
Dim oBBG As New cls_Bloomberg_Sync
Dim bbg_fields As Variant
bbg_fields = Array("TICKER", "PARSEKYABLE_DES", "UNDL_SPOT_TICKER", "LAST_TRADEABLE_DT", "FUT_CONT_SIZE")

Dim dim_bdp_TICKER As Integer, dim_bdp_PARSEKYABLE_DES As Integer, dim_bdp_UNDL_SPOT_TICKER As Integer, _
    dim_bdp_LAST_TRADEABLE_DT As Integer, dim_bdp_FUT_CONT_SIZE As Integer

For i = 0 To UBound(bbg_fields, 1)
    If bbg_fields(i) = "TICKER" Then
        dim_bdp_TICKER = i
    ElseIf bbg_fields(i) = "PARSEKYABLE_DES" Then
        dim_bdp_PARSEKYABLE_DES = i
    ElseIf bbg_fields(i) = "UNDL_SPOT_TICKER" Then
        dim_bdp_UNDL_SPOT_TICKER = i
    ElseIf bbg_fields(i) = "LAST_TRADEABLE_DT" Then
        dim_bdp_LAST_TRADEABLE_DT = i
    ElseIf bbg_fields(i) = "FUT_CONT_SIZE" Then
        dim_bdp_FUT_CONT_SIZE = i
    End If
Next i



Dim vec_ticker_fut() As Variant
For i = 0 To UBound(identifier, 1)
    
    ReDim Preserve vec_ticker_fut(i)
    
    If Left(identifier(i), 6) <> "/buid/" Then
        vec_ticker_fut(i) = "/buid/" & identifier(i)
    Else
        vec_ticker_fut(i) = identifier(i)
    End If
Next i

Dim output_bbg As Variant
Debug.Print "aim_new_entry_future_db: $get_bloomberg_datas"
output_bbg = oBBG.bdp(vec_ticker_fut, bbg_fields, output_format.of_vec_without_header)

Dim count_future As Integer, found_underlying As Boolean

Dim list_future_already_in_db As Variant, list_index_already_in_db As Variant
    Debug.Print "aim_new_entry_future_db: $aim_mount_data_internal_db_future"
    list_future_already_in_db = aim_mount_data_internal_db_future()
    Debug.Print "aim_new_entry_future_db: $aim_mount_data_internal_db_index"
    list_index_already_in_db = aim_mount_data_internal_db_index(Array(1, 2))
    
    Dim dim_future_product_id As Integer, dim_future_underlying_id As Integer, dim_future_ticker As Integer, _
        dim_future_expiry_dt As Integer, dim_future_contract_size As Integer, dim_future_line As Integer
    
    If IsArray(list_future_already_in_db) Then
        For i = 0 To UBound(list_future_already_in_db(0), 1)
            If list_future_already_in_db(0)(i) = "product_id" Then
                dim_future_product_id = i
            ElseIf list_future_already_in_db(0)(i) = "underlying_id" Then
                dim_future_underlying_id = i
            ElseIf list_future_already_in_db(0)(i) = "ticker" Then
                dim_future_ticker = i
            ElseIf list_future_already_in_db(0)(i) = "expiry_date" Then
                dim_future_expiry_dt = i
            ElseIf list_future_already_in_db(0)(i) = "contract_size" Then
                dim_future_contract_size = i
            ElseIf list_future_already_in_db(0)(i) = "line" Then
                dim_future_line = i
            End If
        Next i
    Else
        dim_future_product_id = 0
        dim_future_underlying_id = 1
        dim_future_ticker = 2
        dim_future_expiry_dt = 3
        dim_future_contract_size = 4
        dim_future_line = 5
    End If
    
    
    Dim dim_index_underlying_id As Integer, dim_index_line As Integer
    If IsArray(list_index_already_in_db) Then
        For i = 0 To UBound(list_index_already_in_db(0), 1)
            If list_index_already_in_db(0)(i) = "underlying_id" Then
                dim_index_underlying_id = i
            ElseIf list_index_already_in_db(0)(i) = "line" Then
                dim_index_line = i
            End If
        Next i
    Else
        Debug.Print "aim_new_entry_future_db: @no index in local db"
        MsgBox ("no index in DB")
        Exit Sub
    End If
    
    
    


Dim tmp_vec_future() As Variant
Dim line_underlying_index As Integer
    line_underlying_index = 0

Dim tmp_min_date As Date
Dim min_pos As Integer
Dim tmp_vec As Variant


Dim vec_underyling_need_to_be_refresh_in_open() As Variant
Dim count_underyling_need_to_be_refresh_in_open As Integer
    count_underyling_need_to_be_refresh_in_open = 0

For i = 0 To UBound(identifier, 1)
    
    count_future = 0
    found_underlying = False
    
    If Left(output_bbg(i)(dim_bdp_TICKER), 1) <> "#" Then
        
        
        If IsArray(list_index_already_in_db) Then
            For j = 0 To UBound(list_index_already_in_db, 1)
                If list_index_already_in_db(j)(dim_index_underlying_id) = "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER) Then
                    found_underlying = True
                    line_underlying_index = list_index_already_in_db(j)(dim_index_line)
                    Exit For
                End If
            Next j
        End If
        
        
        If line_underlying_index <> 0 Then
            If IsArray(list_future_already_in_db) Then
                
                For j = 0 To UBound(list_future_already_in_db, 1)
                    If list_future_already_in_db(j)(dim_future_underlying_id) = "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER) Then
                        
                        If list_future_already_in_db(j)(dim_future_expiry_dt) >= Date Then
                            
                            ReDim Preserve tmp_vec_future(count_future)
                            tmp_vec_future(count_future) = list_future_already_in_db(j)
                            count_future = count_future + 1
                            
                        End If
                        
                    Else
                        
                    End If
                Next j
            Else
                
            End If
        End If
        
        
        If found_underlying Then
            
            If count_future >= 2 Then
                Debug.Print "aim_new_entry_future_db: @too many futures already setup for underlying=" & list_index_already_in_db(j)(dim_index_underlying_id)
                MsgBox ("already 2 futures in index db.")
            Else
                'rajoute la nouvelle entree
                ReDim Preserve tmp_vec_future(count_future)
                
                If IsArray(list_future_already_in_db) Then
                    tmp_vec_future(count_future) = list_future_already_in_db(0)
                Else
                    tmp_vec_future(count_future) = Array("", "", "", "", "", "", "", "", "", "", "", "")
                End If
                
                tmp_vec_future(count_future)(dim_future_product_id) = identifier(i)
                tmp_vec_future(count_future)(dim_future_underlying_id) = "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER)
                tmp_vec_future(count_future)(dim_future_ticker) = output_bbg(i)(dim_bdp_TICKER) & " Index"
                tmp_vec_future(count_future)(dim_future_expiry_dt) = output_bbg(i)(dim_bdp_LAST_TRADEABLE_DT)
                tmp_vec_future(count_future)(dim_future_contract_size) = output_bbg(i)(dim_bdp_FUT_CONT_SIZE)
                tmp_vec_future(count_future)(dim_future_line) = line_underlying_index
                
                    
                    If count_underyling_need_to_be_refresh_in_open = 0 Then
                        ReDim Preserve vec_underyling_need_to_be_refresh_in_open(count_underyling_need_to_be_refresh_in_open)
                        vec_underyling_need_to_be_refresh_in_open(count_underyling_need_to_be_refresh_in_open) = "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER)
                        count_underyling_need_to_be_refresh_in_open = count_underyling_need_to_be_refresh_in_open + 1
                    Else
                        For m = 0 To UBound(vec_underyling_need_to_be_refresh_in_open, 1)
                            If "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER) = vec_underyling_need_to_be_refresh_in_open(m) Then
                                Exit For
                            Else
                                If m = UBound(vec_underyling_need_to_be_refresh_in_open, 1) Then
                                    ReDim Preserve vec_underyling_need_to_be_refresh_in_open(count_underyling_need_to_be_refresh_in_open)
                                    vec_underyling_need_to_be_refresh_in_open(count_underyling_need_to_be_refresh_in_open) = "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER)
                                    count_underyling_need_to_be_refresh_in_open = count_underyling_need_to_be_refresh_in_open + 1
                                End If
                            End If
                        Next m
                    End If
                    
                
                'reorder
                If UBound(tmp_vec_future, 1) > 0 Then
                    
                    For m = 0 To UBound(tmp_vec_future, 1)
                        
                        min_pos = m
                        tmp_min_date = tmp_vec_future(m)(dim_future_expiry_dt)
                        
                        For n = m + 1 To UBound(tmp_vec_future, 1)
                            If tmp_vec_future(n)(dim_future_expiry_dt) < tmp_min_date Then
                                min_pos = n
                                tmp_min_date = tmp_vec_future(n)(dim_future_expiry_dt)
                            End If
                        Next n
                        
                        
                        If m <> min_pos Then
                            
                            tmp_vec = tmp_vec_future(m)
                            tmp_vec_future(m) = tmp_vec_future(min_pos)
                            tmp_vec_future(min_pos) = tmp_vec
                            
                        End If
                        
                        
                    Next m
                    
                End If
                
                
                
                'vide
                'Debug.Print "aim_new_entry_future_db: $delete all futures (max2) for underlying " & list_index_already_in_db(j)(dim_index_underlying_id)
                Worksheets("Index_Database").Cells(line_underlying_index, 31) = ""
                Worksheets("Index_Database").Cells(line_underlying_index, 32) = ""
                
                Worksheets("Index_Database").Cells(line_underlying_index, 33) = ""
                Worksheets("Index_Database").Cells(line_underlying_index + 1, 33) = ""
                
                Worksheets("Index_Database").Cells(line_underlying_index, 34) = ""
                Worksheets("Index_Database").Cells(line_underlying_index + 1, 34) = ""
                
                Worksheets("Index_Database").Cells(line_underlying_index, 47) = ""
                Worksheets("Index_Database").Cells(line_underlying_index + 1, 47) = ""
                
                
                
                'insertion dans index db
                'Debug.Print "aim_new_entry_future_db: insert all sorted futures (max2) for underlying " & list_index_already_in_db(j)(dim_index_underlying_id)
                For j = 0 To UBound(tmp_vec_future, 1)
                    Worksheets("Index_Database").Cells(tmp_vec_future(j)(dim_future_line), 31 + j) = tmp_vec_future(j)(dim_future_product_id)
                    Worksheets("Index_Database").Cells(tmp_vec_future(j)(dim_future_line) + j, 33) = tmp_vec_future(j)(dim_future_expiry_dt)
                        Worksheets("Index_Database").Cells(tmp_vec_future(j)(dim_future_line) + j, 33).NumberFormat = "dd.mm.yyyy"
                    Worksheets("Index_Database").Cells(tmp_vec_future(j)(dim_future_line) + j, 34) = tmp_vec_future(j)(dim_future_ticker)
                    Worksheets("Index_Database").Cells(tmp_vec_future(j)(dim_future_line) + j, 47) = tmp_vec_future(j)(dim_future_contract_size)
                Next j
                
                'recharge la liste des futures pour prendre en compte la nouvelle entree
                list_future_already_in_db = aim_mount_data_internal_db_future()
                
            End If
            
        Else
            Debug.Print "aim_new_entry_future_db: @underlying not found in index db. Looking for: " & "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER) & " for future=" & identifier(i)
            MsgBox ("underlying not found in index db: " & "EI09" & output_bbg(i)(dim_bdp_UNDL_SPOT_TICKER))
        End If
        
    End If
bypass_entry_fut:
Next i


'clean des fut open
If count_underyling_need_to_be_refresh_in_open > 0 Then
    
    Dim is_last_line As Boolean
    Dim open_threshold As Integer
    open_threshold = 25
    
    
    For i = l_header_internal_open + 1 To 5000
        
        is_last_line = True
        
        For j = 0 To open_threshold
            If Worksheets("Open").Cells(i + j, 1) <> "" Then
                is_last_line = False
                Exit For
            End If
        Next j
        
        If is_last_line = True Then
            Exit For
        Else
            If Worksheets("Open").Cells(i, 6) = "F" Then
                'est-il dans le fut a mettre a jour ?
                For j = 0 To UBound(vec_underyling_need_to_be_refresh_in_open, 1)
                    If vec_underyling_need_to_be_refresh_in_open(j) = Worksheets("Open").Cells(i, 1) Then
                        Worksheets("Open").rows(i).Clear
                        Exit For
                    End If
                Next j
            End If
        End If
        
    Next i
    
End If


End Sub


Sub aim_new_entry_index_db(ByVal identifier As Variant)

If IsArray(identifier) = True Then
Else
    If IsEmpty(identifier) = True Then
        Exit Sub
    End If
End If

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim oBBG As New cls_Bloomberg_Sync
Dim l_xls As Worksheet

'recupere les donnees pour le matching currency/curerncy code
Dim dim_currency_txt As Integer, dim_currency_code As Integer
dim_currency_txt = 0
dim_currency_code = 1
dim_currency_row = 2
Dim vec_currency() As Variant
k = 0
For i = 14 To 32
    If Worksheets("Parametres").Cells(i, 1) <> "" Then
        ReDim Preserve vec_currency(k)
        vec_currency(k) = Array(UCase(Worksheets("Parametres").Cells(i, 1).Value), Worksheets("Parametres").Cells(i, 5).Value, i)
        k = k + 1
    End If
Next i

Debug.Print "aim_new_entry_index_db: mount data currency from parametres (txt, code, row). Found " & k & " curriences"


'detection de la fin d'index database SANS LE COMPTEUR
Dim l_index_db_last As Integer
For i = l_header_internal_db_index + 2 To 500 Step 3
    If Worksheets("Index_Database").Cells(i, 1) = "" Then
        l_index_db_last = i - 2
        Exit For
    End If
Next i

Debug.Print "aim_new_entry_index_db: detect last line from index_db without counter. $last_line=" & l_index_db_last

Dim output_bbg As Variant
Dim fields_bbg As Variant
fields_bbg = Array("CRNCY", "NAME", "PARSEKYABLE_DES", "PARSEKYABLE_DES_SOURCE", "TICKER", "SECURITY_DES", _
    "MOST_ACTIVE_FUTURE_TICKER", "PX_LAST")

Dim dim_bbg_CRNCY As Integer, dim_bbg_NAME As Integer, dim_bbg_PARSEKYABLE_DES As Integer, dim_bbg_PARSEKYABLE_DES_SOURCE As Integer, _
    dim_bbg_TICKER As Integer, dim_bbg_SECURITY_DES As Integer, dim_bbg_MOST_ACTIVE_FUTURE_TICKER As Integer, _
    dim_bbg_OPT_CHAIN As Integer, dim_bbg_PX_LAST As Integer


For i = 0 To UBound(fields_bbg, 1)
    If fields_bbg(i) = "CRNCY" Then
        dim_bbg_CRNCY = i
    ElseIf fields_bbg(i) = "NAME" Then
        dim_bbg_NAME = i
    ElseIf fields_bbg(i) = "PARSEKYABLE_DES_SOURCE" Then
        dim_bbg_PARSEKYABLE_DES_SOURCE = i
    ElseIf fields_bbg(i) = "MOST_ACTIVE_FUTURE_TICKER" Then
        dim_bbg_MOST_ACTIVE_FUTURE_TICKER = i
    ElseIf fields_bbg(i) = "OPT_CHAIN" Then
        dim_bbg_OPT_CHAIN = i
    ElseIf fields_bbg(i) = "PX_LAST" Then
        dim_bbg_PX_LAST = i
    End If
Next i


'ajustement de l'identifier buid pour l'appel bbg buid
Dim vec_id_buid() As Variant
k = 0
For i = 0 To UBound(identifier, 1)
    
    ReDim Preserve vec_id_buid(i)
    
    If UCase(Left(identifier(i), 6)) <> UCase("/buid/") Then
        vec_id_buid(i) = "/buid/" & identifier(i)
    Else
        vec_id_buid(i) = identifier(i)
    End If
Next i
Debug.Print "aim_new_entry_equity_db: $download bloomberg datas"
output_bbg = oBBG.bdp(vec_id_buid, fields_bbg, output_format.of_vec_without_header)


Dim tmp_strike As Integer
Dim tmp_maturity As Integer
Dim quarter_maturity As Variant
    quarter_maturity = Array(3, 6, 9, 12)

Dim vec_future() As Variant, vec_option() As Variant

For i = 0 To UBound(output_bbg, 1)
    ReDim Preserve vec_future(i)
    vec_future(i) = output_bbg(i)(dim_bbg_MOST_ACTIVE_FUTURE_TICKER) & " Index"
    
    
    For j = 0 To UBound(quarter_maturity, 1)
        If Month(Date) < quarter_maturity(j) Then
            tmp_maturity = quarter_maturity(j)
            Exit For
        End If
    Next j
    
    If output_bbg(i)(dim_bbg_PX_LAST) < 100 Then
        tmp_strike = CInt(10 * Round(output_bbg(i)(dim_bbg_PX_LAST) / 10, 0))
    Else
        tmp_strike = CInt(100 * Round(output_bbg(i)(dim_bbg_PX_LAST) / 100, 0))
    End If
    
    ReDim Preserve vec_option(i)
    vec_option(i) = Left(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), InStr(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), " ") - 1) & " " & tmp_maturity & " P" & tmp_strike & " Index"
Next i



    Dim output_bbg_future As Variant, output_bbg_option As Variant
    Dim flds_bbg_future As Variant, flds_bbg_option As Variant
    Dim dim_bbg_future_FUT_CONT_SIZE As Integer
    Dim dim_bbg_option_OPT_CONT_SIZE As Integer
    
    
    flds_bbg_future = Array("FUT_CONT_SIZE")
    flds_bbg_option = Array("OPT_CONT_SIZE")
    
    For i = 0 To UBound(flds_bbg_future, 1)
        If flds_bbg_future(i) = "FUT_CONT_SIZE" Then
            dim_bbg_future_FUT_CONT_SIZE = i
        End If
    Next i
    
    For i = 0 To UBound(flds_bbg_option, 1)
        If flds_bbg_option(i) = "OPT_CONT_SIZE" Then
            dim_bbg_option_OPT_CONT_SIZE = i
        End If
    Next i
    
output_bbg_future = oBBG.bdp(vec_future, flds_bbg_future, output_format.of_vec_without_header)
output_bbg_option = oBBG.bdp(vec_option, flds_bbg_option, output_format.of_vec_without_header)

    

Dim tmp_currrency_code As Integer

Dim l_col As Integer, l_cols As Integer, l_row_currency As Integer
Dim l_database_row  As Integer
k = 0
Set l_xls = Worksheets("Index_Database")
    
    With l_xls
        
        .Activate
    
        For i = 0 To UBound(output_bbg, 1)
            
            
            
            If Left(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), 1) <> "#" Then 'check valid ticker
                
                For j = 0 To UBound(vec_currency, 1)
                    If UCase(vec_currency(j)(dim_currency_txt)) = output_bbg(i)(dim_bbg_CRNCY) Then
                        tmp_currrency_code = vec_currency(j)(dim_currency_code)
                        l_row_currency = vec_currency(j)(dim_currency_row)
                        Exit For
                    Else
                        If j = UBound(vec_currency, 1) Then
                            Debug.Print "aim_new_entry_index_db: $unable to find currency " & output_bbg(i)(dim_bbg_CRNCY) & " in parametres. Bypass " & identifier(i) & " to next entry"
                            GoTo bypass_new_entry_index
                        End If
                    End If
                Next j
                
                l_database_row = l_index_db_last + (3 * k)
            
                .rows(4).EntireRow.Copy Destination:=.rows(l_database_row + 1)
                .rows(5).EntireRow.Copy Destination:=.rows(l_database_row + 2)
                .rows(6).EntireRow.Copy Destination:=.rows(l_database_row + 3)
                
                Debug.Print "aim_new_entry_index_db: $insert new entry in equity db. Identifier=" & identifier(i) & " - " & output_bbg(i)(dim_bbg_NAME) & " - " & output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE)
                
                'A - underlying id
                .Cells(l_database_row + 2, 1) = identifier(i)
                
                'B - index name
                .Cells(l_database_row + 2, 2) = output_bbg(i)(dim_bbg_NAME)
                
                'C - characteristics
                .Cells(l_database_row + 2, 3) = "FUTURE"
                
                'J - daily premium
                .Cells(l_database_row + 2, 10).Value = "=(RC22-RC152)*Parametres!R" & l_row_currency & "C6*1/1000"
                
                'K - daily StartCurReval
                .Cells(l_database_row + 2, 11).Value = "=(RC40-RC153)*Parametres!R" & l_row_currency & "C6*1/1000"
                
                'M - Result Executed
                .Cells(l_database_row + 2, 14).Value = "=((DSUM(exe,R25C14,R[-1]C1:RC1)/1000)+RC22-RC23)*Parametres!R" & l_row_currency & "C6+RC149"
                
                'AB - Executed
                .Cells(l_database_row + 2, 28).Value = "=(DSUM(AIM_EOD_JSO,""ytd_pnl_local_net"",R[-1]C[147]:RC[147])-RC30)*Parametres!R" & l_row_currency & "C6"
                
                'AC - Daily
                .Cells(l_database_row + 2, 29).Value = "=(-RC42+RC46-RC41+IF(RC31<>"""",(((RC35-RC43)*RC37*RC47)+RC39+RC40)+((RC35-RC44)*RC45*RC47),0)+IF(RC32<>"""",(((R[1]C35-R[1]C43)*R[1]C37*R[1]C47)+R[1]C39+R[1]C40)+((R[1]C35-R[1]C44)*R[1]C45*R[1]C47),0))*Parametres!R" & l_row_currency & "C6"
                
                'AH - inscription d'un ticker de fut pour le pricing des options
                .Cells(l_database_row + 2, 34).Value = output_bbg(i)(dim_bbg_MOST_ACTIVE_FUTURE_TICKER) & " Index"
                
                'CZ - currency code override
                .Cells(l_database_row + 2, 104).Value = "=RC107"
                
                'DC - currency
                .Cells(l_database_row + 2, 107).Value = tmp_currrency_code
                
                'DD - shortname (prend le ticker)
                If InStr(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), " ") <> 0 Then
                    .Cells(l_database_row + 2, 108).Value = Left(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), InStr(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), " ") - 1)
                Else
                    .Cells(l_database_row + 2, 108).Value = output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE)
                End If
                
                'DE - ric option
                If InStr(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), " ") <> 0 Then
                    .Cells(l_database_row + 2, 109).Value = Left(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), InStr(output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), " ") - 1)
                Else
                    .Cells(l_database_row + 2, 109).Value = output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE)
                End If
                
                'DF - bloomberg ticker spot
                .Cells(l_database_row + 2, 110).Value = output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE)
                
                'DH - quotite option
                If IsNumeric(output_bbg_option(i)(dim_bbg_option_OPT_CONT_SIZE)) Then
                    .Cells(l_database_row + 2, 112).Value = output_bbg_option(i)(dim_bbg_option_OPT_CONT_SIZE)
                Else
                    .Cells(l_database_row + 2, 112).Value = CInt(InputBox("Option quotity for " & output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), "option quotity", 100))
                End If
                
                'DI - quotite future
                If IsNumeric(output_bbg_future(i)(dim_bbg_future_FUT_CONT_SIZE)) Then
                    'pour eviter les problemes de delta, demander quand meme confirmation a l utilisateur
                    '.Cells(l_database_row + 2, 113).Value = output_bbg_future(i)(dim_bbg_future_FUT_CONT_SIZE)
                    .Cells(l_database_row + 2, 113).Value = CInt(InputBox("Future quotity for " & output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), "future quotity", output_bbg_future(i)(dim_bbg_future_FUT_CONT_SIZE)))
                Else
                    .Cells(l_database_row + 2, 113).Value = CInt(InputBox("Future quotity for " & output_bbg(i)(dim_bbg_PARSEKYABLE_DES_SOURCE), "future quotity"))
                End If
                
                'DQ - date position system
                .Cells(l_database_row + 2, 121) = Date
                
                
                'FS
                 .Cells(l_database_row + 1, 175) = c_header_aim_eod_underlying_id
                 .Cells(l_database_row + 2, 175).Value = "=RC1"
                
                
                
                'mise en place des formules
                l_cols = 208
        
                For l_col = 5 To l_cols
                    If .Cells(2, l_col).Value = "F" Then l_formula = .Cells(l_database_row + 2, l_col).Value: .Cells(l_database_row + 2, l_col).formula = "=" & Replace(l_formula, ";", ",")
                    If .Cells(3, l_col).Value = "F" Then l_formula = .Cells(l_database_row + 3, l_col).Value: .Cells(l_database_row + 3, l_col).formula = "=" & Replace(l_formula, ";", ",")
                Next l_col
                
                
                
                k = k + 1
            End If
bypass_new_entry_index:
        Next i

    End With

End Sub


Public Sub aim_patch_gpb_entry()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


    correction_factor_net_trading_cash_flow = Array(Array(4, 100))
    correction_factor_ytd_pnl = Array(Array(4, 100))
    correction_factor_commission = Array(Array(4, 100))

Dim l_equity_ccy_code As Integer

For i = 27 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        If Worksheets("Equity_Database").Cells(i, 44) = 4 Then
            
            l_equity_ccy_code = Worksheets("Equity_Database").Cells(i, 44)
            
            'AB - YTD Pnl close
            For m = 0 To UBound(correction_factor_ytd_pnl, 1)
                If l_equity_ccy_code = correction_factor_ytd_pnl(m)(0) Then
                    Worksheets("Equity_Database").Cells(i, 28).Value = "=" & correction_factor_ytd_pnl(m)(1) & "*" & Replace(Worksheets("Equity_Database").Cells(4, 28).Value, ";", ",")
                    Exit For
                End If
            Next m
            
            
            'AC net trading cash flow - check si exception
            For m = 0 To UBound(correction_factor_net_trading_cash_flow, 1)
                If l_equity_ccy_code = correction_factor_net_trading_cash_flow(m)(0) Then
                    Worksheets("Equity_Database").Cells(i, 29).Value = "=" & correction_factor_net_trading_cash_flow(m)(1) & "*" & Replace(Worksheets("Equity_Database").Cells(4, 29).Value, ";", ",")
                    Exit For
                End If
            Next m
            
            
            'AD - 30 - intraday commission & ytd close comm
            For m = 0 To UBound(correction_factor_commission, 1)
                If l_equity_ccy_code = correction_factor_commission(m)(0) Then
                    
                    'code 0 pour eviter de se faire ecraser les formules par set status
                    Worksheets("Equity_Database").Cells(i, 4).Value = 0
                    
                    ' intraday fees equitis
                    Worksheets("Equity_Database").Cells(i, 30).Value = "=" & correction_factor_commission(m)(1) & "*" & Replace(Worksheets("Equity_Database").Cells(4, 30).Value, ";", ",")
                    
                    'BP - intraday (equity + options)
                    Worksheets("Equity_Database").Cells(i, 68).Value = "=" & correction_factor_commission(m)(1) & "*DSUM(AIM_Equities_JSO,""intraday_commission_local"",R[-1]C100:RC100)+" & correction_factor_commission(m)(1) & "*DSUM(AIM_Options_JSO,""intraday_commission_local"",R[-1]C101:RC101)"
                    
                    'BQ - ytd close (equity + options)
                    Worksheets("Equity_Database").Cells(i, 69).Value = "=" & correction_factor_commission(m)(1) & "*" & Replace(Worksheets("Equity_Database").Cells(4, 69).Value, ";", ",")
                    
                End If
            Next m
            
            
            
        End If
    End If
Next i




End Sub


Sub aim_new_entry_equity_db(ByVal identifier As Variant)

Dim override_code As Variant
Dim correction_factor_price_spot As Variant, correction_factor_price_close As Variant, correction_factor_net_trading_cash_flow As Variant, correction_factor_ytd_pnl As Variant, correction_factor_commission As Variant
    
    
    override_code = Array(Array(4, 0))
    correction_factor_price_spot = Array(Array(-4, 100))
    correction_factor_price_close = Array(Array(4, 100))
    correction_factor_net_trading_cash_flow = Array(Array(4, 100))
    correction_factor_ytd_pnl = Array(Array(4, 100))
    correction_factor_commission = Array(Array(4, 100))


If IsArray(identifier) = True Then
Else
    If IsEmpty(identifier) = True Then
        Exit Sub
    End If
End If

Debug.Print "aim_new_entry_equity_db: " & "INPUT " & "#vec_underlying_id of " & UBound(identifier, 1) + 1 & " entries"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim oBBG As New cls_Bloomberg_Sync
Dim l_xls As Worksheet

'recupere les donnees pour le matching currency/curerncy code
Dim dim_currency_txt As Integer, dim_currency_code As Integer
dim_currency_txt = 0
dim_currency_code = 1
dim_currency_row = 2
Dim vec_currency() As Variant
k = 0
For i = 14 To 32
    If Worksheets("Parametres").Cells(i, 1) <> "" Then
        ReDim Preserve vec_currency(k)
        vec_currency(k) = Array(UCase(Worksheets("Parametres").Cells(i, 1).Value), Worksheets("Parametres").Cells(i, 5).Value, i)
        k = k + 1
    End If
Next i

Debug.Print "aim_new_entry_equity_db: mount data currency from parametres (txt, code, row). Found " & k & " curriences"


Dim dim_sector_txt As String, dim_sector_code As String
dim_sector_txt = 0
dim_sector_code = 1
Dim vec_sector() As Variant
k = 0
For i = 17 To 26
    ReDim Preserve vec_sector(k)
    vec_sector(k) = Array(Worksheets("Simulation").Cells(i, 27).Value, Worksheets("Simulation").Cells(i, 28).Value)
    k = k + 1
Next i

Debug.Print "aim_new_entry_equity_db: mount data sector from simulatation (txt, code). Found " & k & " sectors"



'detection de la fin d'equity database SANS LE COMPTEUR
Dim l_equity_db_last As Integer
For i = l_header_internal_db_equity To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        l_equity_db_last = i - 2
        Exit For
    End If
Next i

Debug.Print "aim_new_entry_equity_db: detect last line from equity_db without counter. $last_line=" & l_equity_db_last


Dim output_bbg As Variant
Dim fields_bbg As Variant
fields_bbg = Array("CRNCY", "ID_ISIN", "NAME", "INDUSTRY_GROUP", "INDUSTRY_SECTOR", "GICS_SECTOR_NAME", "GICS_INDUSTRY_GROUP_NAME", "GICS_INDUSTRY_NAME", "GICS_SUB_INDUSTRY_NAME", "COUNTRY", "TICKER_AND_EXCH_CODE", "ID_BB_UNIQUE", "REL_INDEX")

Dim dim_bbg_CRNCY As Integer, dim_bbg_isin As Integer, dim_bbg_NAME As Integer, dim_bbg_indu_group As Integer, _
    dim_bbg_indu_sect As Integer, dim_bbg_gics_sector_name As Integer, dim_bbg_gics_industry_group_name As Integer, _
    dim_bbg_gics_industry_name As Integer, dim_bbg_gics_sub_industry_name As Integer, dim_bbg_gics_country As Integer, _
    dim_bbg_TICKER_AND_EXCH_CODE As Integer, dim_bbg_id_bb_unique As Integer, dim_bbg_rel_index As Integer


For i = 0 To UBound(fields_bbg, 1)
    If fields_bbg(i) = "CRNCY" Then
        dim_bbg_CRNCY = i
    ElseIf fields_bbg(i) = "ID_ISIN" Then
        dim_bbg_isin = i
    ElseIf fields_bbg(i) = "NAME" Then
        dim_bbg_NAME = i
    ElseIf fields_bbg(i) = "INDUSTRY_GROUP" Then
        dim_bbg_indu_group = i
    ElseIf fields_bbg(i) = "INDUSTRY_SECTOR" Then
        dim_bbg_indu_sect = i
    ElseIf fields_bbg(i) = "GICS_SECTOR_NAME" Then
        dim_bbg_gics_sector_name = i
    ElseIf fields_bbg(i) = "GICS_INDUSTRY_GROUP_NAME" Then
        dim_bbg_gics_industry_group_name = i
    ElseIf fields_bbg(i) = "GICS_INDUSTRY_NAME" Then
        dim_bbg_gics_industry_name = i
    ElseIf fields_bbg(i) = "GICS_SUB_INDUSTRY_NAME" Then
        dim_bbg_gics_sub_industry_name = i
    ElseIf fields_bbg(i) = "COUNTRY" Then
        dim_bbg_gics_country = i
    ElseIf fields_bbg(i) = "TICKER_AND_EXCH_CODE" Then
        dim_bbg_TICKER_AND_EXCH_CODE = i
    ElseIf fields_bbg(i) = "ID_BB_UNIQUE" Then
        dim_bbg_id_bb_unique = i
    ElseIf fields_bbg(i) = "REL_INDEX" Then
        dim_bbg_rel_index = i
    End If
Next i


'ajustement de l'identifier buid pour l'appel bbg buid
Dim vec_id_buid() As Variant
k = 0
For i = 0 To UBound(identifier, 1)
    
    ReDim Preserve vec_id_buid(i)
    
    If UCase(Left(identifier(i), 6)) <> UCase("/buid/") Then
        vec_id_buid(i) = "/buid/" & identifier(i)
    Else
        vec_id_buid(i) = identifier(i)
    End If
Next i
Debug.Print "aim_new_entry_equity_db: $download bloomberg datas"
output_bbg = oBBG.bdp(vec_id_buid, fields_bbg, output_format.of_vec_without_header)



Dim l_equity_id As String
Dim l_equity_characteristics As String
Dim l_equity_code As Integer
Dim l_equity_bb_ric As String
Dim l_equity_bbcode As String
Dim l_quotity As Integer

Dim l_equity_name As String
Dim l_equity_short_name As String
Dim l_option_bbcode As String
Dim l_equity_isin As String
Dim l_equity_rel_index As String

Dim l_equity_industry_group As String
Dim l_equity_industry_sector As String

Dim l_equity_type As String

Dim l_equity_ccy_lib As String 'currency txt
Dim l_equity_ccy_code As Integer 'currency code
Dim l_ccy_row As Integer 'currency row dans parametre

For j = 0 To UBound(output_bbg, 1)

    'check si l identifiant est correct
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If output_bbg(j)(0) = "#N/A security" Then
        Debug.Print "aim_new_entry_equity_db: error with product: " & identifier(j) & " seems not to be a valid Bloomberg identifier. Bypass to next entry"
        GoTo bypass_new_entry_equity_db
    End If
    
    l_equity_id = identifier(j)
    l_equity_bbcode = UCase(output_bbg(j)(dim_bbg_TICKER_AND_EXCH_CODE)) & " EQUITY"
    l_equity_bb_ric = UCase(output_bbg(j)(dim_bbg_TICKER_AND_EXCH_CODE))
    
    l_quotity = 100 'default quotity, dans le cadre de AIM sert uniquement pour les options rentrees a la main
    l_option_bbcode = Replace(UCase(output_bbg(j)(dim_bbg_TICKER_AND_EXCH_CODE)), " EQUITY", "")
    
    
        'patch pour certains pays
        Dim patch_mp_option As Variant
        patch_mp_option = Array(Array("GY", "GR"), Array("VX", "SW"))
        
        For m = 0 To UBound(patch_mp_option, 1)
            If InStr(UCase(l_option_bbcode), " " & patch_mp_option(m)(0) & " ") <> 0 Then
                l_option_bbcode = Replace(l_option_bbcode, " " & patch_mp_option(m)(0) & " ", " " & patch_mp_option(m)(1) & " ")
            End If
        Next m
    
    l_equity_type = "STOCK"
    l_equity_ccy_lib = output_bbg(j)(dim_bbg_CRNCY)
    l_equity_isin = output_bbg(j)(dim_bbg_isin)
    l_equity_name = output_bbg(j)(dim_bbg_NAME)
    l_equity_short_name = Trim(Left(l_equity_bb_ric, 4))
    
    l_equity_industry_group = output_bbg(j)(dim_bbg_indu_group)
    l_equity_industry_sector = output_bbg(j)(dim_bbg_indu_sect)
    
    l_equity_rel_index = output_bbg(j)(dim_bbg_rel_index)
    
    For m = 0 To UBound(vec_currency, 1)
        If UCase(vec_currency(m)(dim_currency_txt)) = UCase(output_bbg(j)(dim_bbg_CRNCY)) Then
            l_equity_ccy_code = vec_currency(m)(dim_currency_code)
            l_ccy_row = vec_currency(m)(dim_currency_row)
            Exit For
        Else
            If m = UBound(vec_currency, 1) Then
                Debug.Print "aim_new_entry_equity_db: $problem with currency " & UCase(output_bbg(j)(dim_bbg_CRNCY)) & " not setup in parametres for product=" & UCase(output_bbg(j)(dim_bbg_TICKER_AND_EXCH_CODE)) & " EQUITY"
                MsgBox ("unable to find currency in parameters. Looking for: " & UCase(output_bbg(j)(dim_bbg_CRNCY)) & ". Bypass to next entry")
                GoTo bypass_new_entry_equity_db
            End If
        End If
    Next m
    
    
    'Application.ScreenUpdating = False
    
    ' ANDREASSON LENNART MODIFICATION DU CODE 0 POUR EQUITY_ID
    
    Set l_xls = Worksheets("Equity_Database")
    
    With l_xls
    
        .Activate
        
        l_database_row = l_equity_db_last + (2 * j)
        
        .rows(3).EntireRow.Copy Destination:=.rows(l_database_row + 1)
        .rows(4).EntireRow.Copy Destination:=.rows(l_database_row + 2)
        .rows(l_database_row + 1).EntireRow.Hidden = False
        .rows(l_database_row + 2).EntireRow.Hidden = False
        
        Debug.Print "aim_new_entry_equity_db: $insert new entry in equity db. Identifier=" & identifier(j)
        
        'A - 01 - underyling identifier
        .Cells(l_database_row + 2, 1).Value = identifier(j)
        
        'B - 02 - product name
        .Cells(l_database_row + 2, 2).Value = output_bbg(j)(dim_bbg_NAME)
        
        'C - 03 - characteristics
        .Cells(l_database_row + 2, 3).Value = l_equity_type
        
        'D - 04 - code
        For m = 0 To UBound(override_code, 1)
            If l_equity_ccy_code = override_code(m)(0) Then
                .Cells(l_database_row + 2, 4).Value = override_code(m)(1)
                Exit For
            End If
        Next m
        
        'K - 11 - daily premium
            l_formula = "(RC39-RC95)*Parametres!R" & l_ccy_row & "C6*1/1000"
        .Cells(l_database_row + 2, 11).Value = "=" & Replace(l_formula, ";", ",")
        
        'L - daily startcurreveal
            l_formula = "(RC28-RC96)*Parametres!R" & l_ccy_row & "C6*1/1000"
        .Cells(l_database_row + 2, 12).Value = "=" & Replace(l_formula, ";", ",")
        
        'O
            l_formula = "=((DSUM(exe;R25C15;R[-1]C1:RC1)+RC39-RC40)/1000)*Parametres!R" & l_ccy_row & "C6+RC92"
        .Cells(l_database_row + 2, 15).Value = Replace(l_formula, ";", ",")
        
        
        'Q  - existe-t-il une ancienne entree dans cointrin dont le pnl est deja agrege dans exe ?
        
        
        'U
            'l_formula = "=((RC22-RC27)*RC26+RC29+(RC26-RC25)*RC27-RC30+RC28-RC31+RC41+IF(RC33=0;-RC32;(RC22-RC32)*RC33))*RC50*Parametres!R" & l_ccy_row & "C6"
            'le net trading cash flow de aim est net de comm
            'l_formula = "=((RC22-RC27)*RC26+(RC29+RC30)+(RC26-RC25)*RC27-RC30+RC28-RC31+RC41+IF(RC33=0;-RC32;(RC22-RC32)*RC33))*RC50*Parametres!R" & l_ccy_row & "C6"
            l_formula = "=((RC22-RC27)*RC26+RC29+(RC26-RC25)*RC27+RC28-RC31+RC41+IF(RC33=0;-RC32;(RC22-RC32)*RC33))*RC50*Parametres!R" & l_ccy_row & "C6"
        .Cells(l_database_row + 2, 21).formula = Replace(l_formula, ";", ",")
        
        'V - 22 - spot price
        .Cells(l_database_row + 1, 22).Value = l_equity_short_name & space(1) & "Spot" 'header
        
        For m = 0 To UBound(correction_factor_price_spot, 1)
            If l_equity_ccy_code = correction_factor_price_spot(m)(0) Then
                .Cells(l_database_row + 2, 22).Value = correction_factor_price_spot(m)(1) & "*" & .Cells(4, 22).Value
            End If
        Next m
        
        
        
        'W - header
        .Cells(l_database_row + 1, 23).Value = l_equity_short_name & space(1) & "Beta"
        
        'AA - 27 close price
        .Cells(l_database_row + 1, 27).Value = l_equity_short_name & space(1) & "Close" 'header
        
        For m = 0 To UBound(correction_factor_price_close, 1)
            If l_equity_ccy_code = correction_factor_price_close(m)(0) Then
                .Cells(l_database_row + 2, 27).Value = correction_factor_price_close(m)(1) & "*(" & .Cells(4, 27).Value & ")"
            End If
        Next m
        
        
        
        'AB - YTD Pnl close
        For m = 0 To UBound(correction_factor_ytd_pnl, 1)
            If l_equity_ccy_code = correction_factor_ytd_pnl(m)(0) Then
                '.Cells(l_database_row + 2, 28).Value = correction_factor_ytd_pnl(m)(1) & "*DSUM(AIM_EOD_JSO,""ytd_pnl_local_net"",R[-1]C101:RC101)"
                .Cells(l_database_row + 2, 28).Value = correction_factor_ytd_pnl(m)(1) & "*" & .Cells(4, 28).Value
                Exit For
            End If
        Next m
        
        
        'AC net trading cash flow - check si exception
        For m = 0 To UBound(correction_factor_net_trading_cash_flow, 1)
            If l_equity_ccy_code = correction_factor_net_trading_cash_flow(m)(0) Then
                '.Cells(l_database_row + 2, 29).Value = correction_factor_net_trading_cash_flow(m)(1) & "*DSUM(AIM_Equities_JSO,""net_cash_local_with_comm"",R[-1]C1:RC1)"
                .Cells(l_database_row + 2, 29).Value = correction_factor_net_trading_cash_flow(m)(1) & "*" & .Cells(4, 29).Value
                Exit For
            End If
        Next m
        
        'AD - 30 - intraday commission & ytd close comm
        For m = 0 To UBound(correction_factor_commission, 1)
            If l_equity_ccy_code = correction_factor_commission(m)(0) Then
                
                ' intraday fees equitis
                '.Cells(l_database_row + 2, 30).Value = correction_factor_commission(m)(1) & "*DSUM(AIM_Equities_JSO,""intraday_commission_local"",R[-1]C1:RC1)"
                .Cells(l_database_row + 2, 30).Value = correction_factor_commission(m)(1) & "*" & .Cells(4, 30).Value
                
                'BP - intraday (equity + options)
                .Cells(l_database_row + 2, 68).Value = correction_factor_commission(m)(1) & "*DSUM(AIM_Equities_JSO,""intraday_commission_local"",R[-1]C100:RC100)+" & correction_factor_commission(m)(1) & "*DSUM(AIM_Options_JSO,""intraday_commission_local"",R[-1]C101:RC101)"
                
                'BQ - ytd close (equity + options)
                '.Cells(l_database_row + 2, 69).Value = correction_factor_commission(m)(1) & "*DSUM(AIM_EOD_JSO,""tra_cost_total"",R[-1]C101:RC101)"
                .Cells(l_database_row + 2, 69).Value = correction_factor_commission(m)(1) & "*" & .Cells(4, 69).Value
                
            End If
        Next m
        
        
        'AE
        l_equity_start_pl = 0
        .Cells(l_database_row + 2, 31).Value = l_equity_start_pl
        
        'AN
        l_equity_premium = 0
        .Cells(l_database_row + 2, 40).Value = l_equity_premium
        
        'AO
        
        'AR
        .Cells(l_database_row + 1, 44).Value = "Currency"
        .Cells(l_database_row + 2, 44).Value = l_equity_ccy_code
        
        'AS
        .Cells(l_database_row + 2, 45).Value = Trim(Left(UCase(output_bbg(j)(dim_bbg_TICKER_AND_EXCH_CODE)), 4))
        
        'AT
        .Cells(l_database_row + 2, 46).Value = l_option_bbcode
        
        'AU - Ticker
        .Cells(l_database_row + 2, 47).Value = aim_patch_ticker_marketplace(UCase(output_bbg(j)(dim_bbg_TICKER_AND_EXCH_CODE)) & " EQUITY")
        
        'AV
        .Cells(l_database_row + 2, 48).Value = ""
        
        'AW
        .Cells(l_database_row + 2, 49).Value = l_quotity
        
        'AY
        .Cells(l_database_row + 1, 51).Value = "ISIN" 'header
        .Cells(l_database_row + 2, 51).Value = output_bbg(j)(dim_bbg_isin)
        
        'AZ
        .Cells(l_database_row + 2, 52).Value = Right(Replace(UCase(output_bbg(j)(dim_bbg_TICKER_AND_EXCH_CODE)), " EQUITY", ""), 2)
        
        'BA
        .Cells(l_database_row + 2, 53).Value = output_bbg(j)(dim_bbg_indu_sect)
        
        'BB
        .Cells(l_database_row + 2, 54).Value = output_bbg(j)(dim_bbg_indu_group)
        
        'BC
        '.Cells(l_database_row + 2, 55).Value = get_Beta_Sector_Code(l_equity_industry_sector)
        For m = 0 To UBound(vec_sector, 1)
            If UCase(vec_sector(m)(dim_sector_txt)) = UCase(output_bbg(j)(dim_bbg_indu_sect)) Then
                .Cells(l_database_row + 2, 55).Value = vec_sector(m)(dim_sector_code)
                Exit For
            End If
        Next m
        
        
        'BK
        If Left(output_bbg(j)(dim_bbg_gics_sector_name), 1) <> "#" Then
            .Cells(l_database_row + 2, 63).Value = output_bbg(j)(dim_bbg_gics_sector_name)
        End If
        
        'BL
        If Left(output_bbg(j)(dim_bbg_gics_industry_group_name), 1) <> "#" Then
            .Cells(l_database_row + 2, 64).Value = output_bbg(j)(dim_bbg_gics_industry_group_name)
        End If
        
        'BM
        If Left(output_bbg(j)(dim_bbg_gics_industry_name), 1) <> "#" Then
            .Cells(l_database_row + 2, 65).Value = output_bbg(j)(dim_bbg_gics_industry_name)
        End If
        
        'BN
        If Left(output_bbg(j)(dim_bbg_gics_sub_industry_name), 1) <> "#" Then
            .Cells(l_database_row + 2, 66).Value = output_bbg(j)(dim_bbg_gics_sub_industry_name)
        End If
        
        'BO
        If Left(output_bbg(j)(dim_bbg_gics_country), 1) <> "#" Then
            .Cells(l_database_row + 2, 67).Value = output_bbg(j)(dim_bbg_gics_country)
        End If
        
        'CV
        .Cells(l_database_row + 1, 100).Value = "identifier" 'header
        .Cells(l_database_row + 2, 100).FormulaLocal = "=RC1"
        
        
        'CW
        .Cells(l_database_row + 1, 101).Value = "under_product_id" 'header
        .Cells(l_database_row + 2, 101).FormulaLocal = "=RC1"
        
        
        
        'DG
        .Cells(l_database_row + 2, 111).Value = Date
        
        
        'DU
        .Cells(l_database_row + 2, 125).Value = output_bbg(j)(dim_bbg_rel_index)
        
        
        
        'EH - perso rel_1d
        Dim c_equity_db_daily_pct_change As Integer
        c_equity_db_daily_pct_change = 123
        
        Dim region As Variant
            region = Array(Array("Asia/Pacific", Array("JPY", "HKD", "AUD", "SGD", "TWD", "KRW", "INR", "THB", "CNY"), "HSCEI Index"), Array("Europe", Array("CHF", "EUR", "GBP", "SEK", "NOK", "DKK", "PLN"), "SX5E Index"), Array("America", Array("USD", "CAD", "BRL"), "SPX Index"))
        
        Dim found_rel_index As Boolean
            found_rel_index = False
        For m = 0 To UBound(region, 1)
            For n = 0 To UBound(region(m)(1), 1)
                If region(m)(1)(n) = output_bbg(j)(dim_bbg_CRNCY) Then
                    .Cells(l_database_row + 2, 138).FormulaLocal = "=IF(RC" & c_equity_db_daily_pct_change & "=0;0;100*(RC" & c_equity_db_daily_pct_change & "-(BDP(""" & region(m)(2) & """;""last_price"")/BDP(""" & region(m)(2) & """;""px_yest_close"")-1)))"
                    found_rel_index = True
                    Exit For
                End If
            Next n
        Next m
        
            'si introuvable mettre le SPX
            If found_rel_index = False Then
                .Cells(l_database_row + 2, 138).FormulaLocal = "=IF(RC" & c_equity_db_daily_pct_change & "=0;0;100*(RC" & c_equity_db_daily_pct_change & "-(BDP(""SPX Index"";""last_price"")/BDP(""SPX Index"";""px_yest_close"")-1)))"
            End If
        
        
        'Formattage
        l_cols = 208
        
        For l_col = 5 To l_cols
            If .Cells(2, l_col).Value = "F" Then l_formula = .Cells(l_database_row + 2, l_col).Value: .Cells(l_database_row + 2, l_col).formula = "=" & Replace(l_formula, ";", ",")
        Next l_col
                
    End With
          
bypass_new_entry_equity_db:
Next j

End Sub


Private Function aim_mount_data_internal_db_index(Optional ByVal vec_column As Variant) As Variant

Debug.Print "aim_mount_data_internal_db_index: " & "INPUT " & "OPT #vec_numeric_column_to_get"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim data_index() As Variant
Dim tmp_vec() As Variant

ReDim tmp_vec(2)
    tmp_vec(0) = "underlying_id"
    tmp_vec(1) = "ticker"
    tmp_vec(2) = "line"


'header
If IsMissing(vec_column) = False Then
    For i = 0 To UBound(vec_column)
        ReDim Preserve tmp_vec(2 + 1 + i)
        tmp_vec(2 + 1 + i) = Worksheets("Index_Database").Cells(l_header_internal_db_index, vec_column(i))
    Next i
End If

ReDim Preserve data_index(0)
data_index(0) = tmp_vec

k = 1
For i = l_header_internal_db_index + 2 To 500 Step 3
    If Worksheets("Index_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        
        ReDim Preserve data_index(k)
            
        ReDim Preserve tmp_vec(2)
        tmp_vec(0) = Worksheets("Index_Database").Cells(i, 1).Value
        tmp_vec(1) = Worksheets("Index_Database").Cells(i, 110).Value
        tmp_vec(2) = i
        
        If IsMissing(vec_column) = False Then
            '+ les autres colonnes demandee
            For j = 0 To UBound(vec_column, 1)
                ReDim Preserve tmp_vec(2 + 1 + j)
                tmp_vec(2 + 1 + j) = Worksheets("Index_Database").Cells(i, vec_column(j))
            Next j
        End If
        
        data_index(k) = tmp_vec
        
        k = k + 1
        
    End If
Next i


If k = 1 Then
    aim_mount_data_internal_db_index = Empty
    Debug.Print "aim_mount_data_internal_db_index: " & "OUTPUT " & "@no entry in index db."
Else
    aim_mount_data_internal_db_index = data_index
    Debug.Print "aim_mount_data_internal_db_index: " & "OUTPUT " & "@vec_index_underyling_id of " & UBound(data_index, 1) & " entries"
End If

End Function


Private Function aim_mount_data_indernal_multi_accounts_db_future(Optional ByVal vec_column As Variant) As Variant

Debug.Print "aim_mount_data_indernal_multi_accounts_db_future: " & "INPUT " & "OPT #vec_numeric_column_to_get"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim data_future() As Variant
Dim tmp_vec() As Variant

ReDim tmp_vec(5)
    tmp_vec(0) = "product_id"
    tmp_vec(1) = "underlying_id"
    tmp_vec(2) = "ticker"
    tmp_vec(3) = "expiry_date"
    tmp_vec(4) = "contract_size"
    tmp_vec(5) = "line"

'header
If IsMissing(vec_column) = False Then
    For i = 0 To UBound(vec_column)
        ReDim Preserve tmp_vec(5 + 1 + i)
        tmp_vec(5 + 1 + i) = Worksheets(sheet_index_db_multi_accounts).Cells(l_header_internal_db_index, vec_column(i))
    Next i
End If

ReDim Preserve data_future(0)
data_future(0) = tmp_vec


k = 1
For i = l_header_internal_db_index + 2 To 500 Step 3
    If Worksheets(sheet_index_db_multi_accounts).Cells(i, 1) = "" Then
        Exit For
    Else
        '1ere maturity
        If Worksheets(sheet_index_db_multi_accounts).Cells(i, 31) <> "" And Worksheets(sheet_index_db_multi_accounts).Cells(i, 31) <> 0 And Worksheets(sheet_index_db_multi_accounts).Cells(i, 33) <> "" And Worksheets(sheet_index_db_multi_accounts).Cells(i, 34) <> "" Then
            ReDim Preserve data_future(k)
            
            ReDim Preserve tmp_vec(5)
            tmp_vec(0) = Worksheets(sheet_index_db_multi_accounts).Cells(i, 31).Value
            tmp_vec(1) = Worksheets(sheet_index_db_multi_accounts).Cells(i, 1).Value
            tmp_vec(2) = Worksheets(sheet_index_db_multi_accounts).Cells(i, 34).Value
            tmp_vec(3) = Worksheets(sheet_index_db_multi_accounts).Cells(i, 33).Value
            tmp_vec(4) = Worksheets(sheet_index_db_multi_accounts).Cells(i, 47).Value
            tmp_vec(5) = i
            
            If IsMissing(vec_column) = False Then
                '+ les autres colonnes demandee
                For j = 0 To UBound(vec_column, 1)
                    ReDim Preserve tmp_vec(5 + 1 + j)
                    tmp_vec(5 + 1 + j) = Worksheets(sheet_index_db_multi_accounts).Cells(i, vec_column(j))
                Next j
            End If
            
            data_future(k) = tmp_vec
            
            k = k + 1
        End If
        
        
        '2e maturity
        If Worksheets(sheet_index_db_multi_accounts).Cells(i, 32) <> "" And Worksheets(sheet_index_db_multi_accounts).Cells(i, 32) <> 0 And Worksheets(sheet_index_db_multi_accounts).Cells(i + 1, 33) <> "" And Worksheets(sheet_index_db_multi_accounts).Cells(i + 1, 34) <> "" Then
            ReDim Preserve data_future(k)
            
            ReDim Preserve tmp_vec(5)
            tmp_vec(0) = Worksheets(sheet_index_db_multi_accounts).Cells(i, 32).Value
            tmp_vec(1) = Worksheets(sheet_index_db_multi_accounts).Cells(i, 1).Value
            tmp_vec(2) = Worksheets(sheet_index_db_multi_accounts).Cells(i + 1, 34).Value
            tmp_vec(3) = Worksheets(sheet_index_db_multi_accounts).Cells(i + 1, 33).Value
            tmp_vec(4) = Worksheets(sheet_index_db_multi_accounts).Cells(i + 1, 47).Value
            tmp_vec(5) = i + 1
            
            If IsMissing(vec_column) = False Then
                '+ les autres colonnes demandee
                For j = 0 To UBound(vec_column, 1)
                    ReDim Preserve tmp_vec(5 + 1 + j)
                    tmp_vec(5 + 1 + j) = Worksheets(sheet_index_db_multi_accounts).Cells(i, vec_column(j))
                Next j
            End If
            
            data_future(k) = tmp_vec
            
            k = k + 1
        End If
        
        
    End If
Next i

If k = 1 Then
    aim_mount_data_indernal_multi_accounts_db_future = Empty
    Debug.Print "aim_mount_data_indernal_multi_accounts_db_future: " & "OUTPUT " & "@no entry in future db."
Else
    aim_mount_data_indernal_multi_accounts_db_future = data_future
    Debug.Print "aim_mount_data_indernal_multi_accounts_db_future: " & "OUTPUT " & "@vec_future_id of " & UBound(data_future, 1) & " entries"
End If

End Function


Private Function aim_mount_data_internal_db_future(Optional ByVal vec_column As Variant) As Variant

Debug.Print "aim_mount_data_internal_db_future: " & "INPUT " & "OPT #vec_numeric_column_to_get"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim data_future() As Variant
Dim tmp_vec() As Variant

ReDim tmp_vec(5)
    tmp_vec(0) = "product_id"
    tmp_vec(1) = "underlying_id"
    tmp_vec(2) = "ticker"
    tmp_vec(3) = "expiry_date"
    tmp_vec(4) = "contract_size"
    tmp_vec(5) = "line"

'header
If IsMissing(vec_column) = False Then
    For i = 0 To UBound(vec_column)
        ReDim Preserve tmp_vec(5 + 1 + i)
        tmp_vec(5 + 1 + i) = Worksheets("Index_Database").Cells(l_header_internal_db_index, vec_column(i))
    Next i
End If

ReDim Preserve data_future(0)
data_future(0) = tmp_vec


k = 1
For i = l_header_internal_db_index + 2 To 500 Step 3
    If Worksheets("Index_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        '1ere maturity
        If Worksheets("Index_Database").Cells(i, 31) <> "" And Worksheets("Index_Database").Cells(i, 31) <> 0 And Worksheets("Index_Database").Cells(i, 33) <> "" And Worksheets("Index_Database").Cells(i, 34) <> "" Then
            ReDim Preserve data_future(k)
            
            ReDim Preserve tmp_vec(5)
            tmp_vec(0) = Worksheets("Index_Database").Cells(i, 31).Value
            tmp_vec(1) = Worksheets("Index_Database").Cells(i, 1).Value
            tmp_vec(2) = Worksheets("Index_Database").Cells(i, 34).Value
            tmp_vec(3) = Worksheets("Index_Database").Cells(i, 33).Value
            tmp_vec(4) = Worksheets("Index_Database").Cells(i, 47).Value
            tmp_vec(5) = i
            
            If IsMissing(vec_column) = False Then
                '+ les autres colonnes demandee
                For j = 0 To UBound(vec_column, 1)
                    ReDim Preserve tmp_vec(5 + 1 + j)
                    tmp_vec(5 + 1 + j) = Worksheets("Index_Database").Cells(i, vec_column(j))
                Next j
            End If
            
            data_future(k) = tmp_vec
            
            k = k + 1
        End If
        
        
        '2e maturity
        If Worksheets("Index_Database").Cells(i, 32) <> "" And Worksheets("Index_Database").Cells(i, 32) <> 0 And Worksheets("Index_Database").Cells(i + 1, 33) <> "" And Worksheets("Index_Database").Cells(i + 1, 34) <> "" Then
            ReDim Preserve data_future(k)
            
            ReDim Preserve tmp_vec(5)
            tmp_vec(0) = Worksheets("Index_Database").Cells(i, 32).Value
            tmp_vec(1) = Worksheets("Index_Database").Cells(i, 1).Value
            tmp_vec(2) = Worksheets("Index_Database").Cells(i + 1, 34).Value
            tmp_vec(3) = Worksheets("Index_Database").Cells(i + 1, 33).Value
            tmp_vec(4) = Worksheets("Index_Database").Cells(i + 1, 47).Value
            tmp_vec(5) = i + 1
            
            If IsMissing(vec_column) = False Then
                '+ les autres colonnes demandee
                For j = 0 To UBound(vec_column, 1)
                    ReDim Preserve tmp_vec(5 + 1 + j)
                    tmp_vec(5 + 1 + j) = Worksheets("Index_Database").Cells(i, vec_column(j))
                Next j
            End If
            
            data_future(k) = tmp_vec
            
            k = k + 1
        End If
        
        
    End If
Next i

If k = 1 Then
    aim_mount_data_internal_db_future = Empty
    Debug.Print "aim_mount_data_internal_db_future: " & "OUTPUT " & "@no entry in future db."
Else
    aim_mount_data_internal_db_future = data_future
    Debug.Print "aim_mount_data_internal_db_future: " & "OUTPUT " & "@vec_future_id of " & UBound(data_future, 1) & " entries"
End If



End Function


Private Function aim_mount_data_internal_db_equity(Optional ByVal vec_column As Variant) As Variant

Debug.Print "aim_mount_data_internal_db_equity: " & "INPUT " & "OPT #vec_numeric_column_to_get"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim count_equity_db As Integer
    count_equity_db = 0

Dim data_internal_db_equity() As Variant
Dim tmp_vec() As Variant


ReDim Preserve data_internal_db_equity(0)
                
ReDim tmp_vec(UBound(vec_column, 1) + 1)
    tmp_vec(0) = "line"

If IsMissing(vec_column) = False Then
    For m = 0 To UBound(vec_column, 1)
        ReDim Preserve tmp_vec(m + 1)
        tmp_vec(m + 1) = Worksheets("Equity_Database").Cells(l_header_internal_db_equity, vec_column(m)).Value
    Next m
End If
                
data_internal_db_equity(0) = tmp_vec

q = 1
                
                
For m = l_header_internal_db_equity + 2 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(m, 1) = "" Then
        Exit For
    Else
        ReDim Preserve data_internal_db_equity(q)
        
        ReDim Preserve tmp_vec(0)
        tmp_vec(0) = m
        
        If IsMissing(vec_column) = False Then
            For n = 0 To UBound(vec_column, 1)
                ReDim Preserve tmp_vec(n + 1)
                If IsError(Worksheets("Equity_Database").Cells(m, vec_column(n))) = False Then
                    tmp_vec(n + 1) = Worksheets("Equity_Database").Cells(m, vec_column(n)).Value
                Else
                    tmp_vec(n + 1) = Empty
                End If
            Next n
        End If
        
        data_internal_db_equity(q) = tmp_vec
        
        q = q + 1
        count_equity_db = count_equity_db + 1
    End If
Next m

If count_equity_db = 0 Then
    aim_mount_data_internal_db_equity = Empty
    Debug.Print "aim_mount_data_internal_db_equity: " & "OUTPUT " & "@no entry in equity db."
Else
    aim_mount_data_internal_db_equity = data_internal_db_equity
    Debug.Print "aim_mount_data_internal_db_equity: " & "OUTPUT " & "@vec_underlying_equity_id of " & UBound(data_internal_db_equity, 1) & " entries"
End If

End Function


'reception d'un vecteur de vecteur avec product_id, underlying_id, ticker, instrument_type, account & pb_account
Public Sub aim_insert_new_position_in_open(ByVal vec_product_underlying_ticker_instrument_type_account_and_pb As Variant)

Dim correction_factor_close_price As Variant
    correction_factor_close_price = Array(Array(4, 100))
    
Dim correction_factor_net_trading_cash_flow As Variant
    correction_factor_net_trading_cash_flow = Array(Array(4, 100))


Dim main_aim_account As String
main_aim_account = Worksheets("Parametres").Cells(17, 18).Value

Dim override_vega_formula As Variant
    override_vega_formula = Array(Array("EI09VIX", "=RC18*1000"))

Dim weight_valeur_eur As Variant
    weight_valeur_eur = Array(Array("EI09VIX", 0), Array("EI09SX5ED", 0.7))


If IsArray(vec_product_underlying_ticker_instrument_type_account_and_pb) Then
Else
    If IsEmpty(vec_product_underlying_ticker_instrument_type_account_and_pb) Then
        Debug.Print "aim_insert_new_position_in_open : @no product to open, array is empty"
        Exit Sub
    End If
End If

Debug.Print "aim_insert_new_position_in_open: " & " INPUT " & "#vec_product_underlying_ticker_instrument_type of " & UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1) + 1 & " entries"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim debug_test As Variant

Dim date_tmp As Date

Dim oBBG As New cls_Bloomberg_Sync


Dim count_resize_vec_product As Integer
    count_resize_vec_product = 0
'precheck des des nlles entrees
    
    'eviter d inserer security si option deja presente
    Dim vec_product_already_in_open As Variant
    vec_product_already_in_open = aim_get_product_and_underlying_from_open()
    
    For i = 0 To UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1)
        If IsArray(vec_product_already_in_open) Then
            For j = 0 To UBound(vec_product_already_in_open, 1)
                'check si le produit n est pas deja dans open
                If vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) = vec_product_already_in_open(j)(0) And vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) = vec_product_already_in_open(j)(4) And vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) = vec_product_already_in_open(j)(5) Then
                    vec_product_underlying_ticker_instrument_type_account_and_pb(i) = vec_product_underlying_ticker_instrument_type_account_and_pb(UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1) - count_resize_vec_product)
                    count_resize_vec_product = count_resize_vec_product + 1
                Else
                    'check si equity et qu une option est deja dans open
                    If vec_product_underlying_ticker_instrument_type_account_and_pb(i)(3) = aim_instrument_type.equity Then
                        If vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) = vec_product_already_in_open(j)(1) And vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) = vec_product_already_in_open(j)(5) Then
                            vec_product_underlying_ticker_instrument_type_account_and_pb(i) = vec_product_underlying_ticker_instrument_type_account_and_pb(UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1) - count_resize_vec_product)
                            count_resize_vec_product = count_resize_vec_product + 1
                        End If
                    End If
                End If
            Next j
        End If
        
        
        If vec_product_underlying_ticker_instrument_type_account_and_pb(i)(3) = aim_instrument_type.equity Then
            'check si une option avec le meme sous-jacent ne veut pas etre inseree
            For j = 0 To UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1)
                If i <> j And vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) = vec_product_underlying_ticker_instrument_type_account_and_pb(j)(1) Then
                    vec_product_underlying_ticker_instrument_type_account_and_pb(i) = vec_product_underlying_ticker_instrument_type_account_and_pb(UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1) - count_resize_vec_product)
                    count_resize_vec_product = count_resize_vec_product + 1
                End If
            Next j
        End If
        
    Next i
    
    If count_resize_vec_product > 0 Then
        
        If UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1) + 1 = count_resize_vec_product Then
            Exit Sub
        End If
        
        ReDim Preserve vec_product_underlying_ticker_instrument_type_account_and_pb(UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1) - count_resize_vec_product)
    End If


'mount les sous-jacent afin de recuperer les div, les future contract size
Dim vec_underlying_equity() As Variant, count_underlying_equity As Integer
    ReDim vec_underlying_equity(0)
    vec_underlying_equity(0) = ""
    count_underlying_equity = 0

Dim vec_underlying_index() As Variant, count_underlying_index As Integer
    ReDim vec_underlying_index(0)
    vec_underlying_index(0) = ""
    count_underlying_index = 0
    

k = 0
For i = 0 To UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1)
    If vec_product_underlying_ticker_instrument_type_account_and_pb(i)(3) = aim_instrument_type.option_equity Then
        
        For j = 0 To UBound(vec_underlying_equity, 1)
            If vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) = vec_underlying_equity(j) Then
                Exit For
            Else
                If j = UBound(vec_underlying_equity, 1) Then
                    If j = 0 And vec_underlying_equity(0) = "" Then
                        vec_underlying_equity(0) = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    Else
                        ReDim Preserve vec_underlying_equity(UBound(vec_underlying_equity, 1) + 1)
                         vec_underlying_equity(UBound(vec_underlying_equity, 1)) = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    End If
                    
                    count_underlying_equity = count_underlying_equity + 1
                    
                End If
            End If
        Next j
        
    ElseIf vec_product_underlying_ticker_instrument_type_account_and_pb(i)(3) = aim_instrument_type.option_index Then
        
        
        
    End If
Next i

If count_underlying_equity > 0 Then
    
    Debug.Print "aim_insert_new_position_in_open: $Bloomberg download datas underlying equities for dividends"
    
    'patch des noms
    For i = 0 To UBound(vec_underlying_equity, 1)
        If Left(vec_underlying_equity(i), 6) <> "/buid/" Then
            vec_underlying_equity(i) = "/buid/" & vec_underlying_equity(i)
        End If
    Next i
    
    
    Dim output_bbg_underlying As Variant
    output_bbg_underlying = oBBG.bdp(vec_underlying_equity, Array("DVD_EX_DT", "EQY_DPS", "LAST_DPS_GROSS"))
    
End If




'remonte les donnes de currency
Dim vec_currency() As Variant
    Dim dim_currency_txt As Integer, dim_currency_code As Integer, dim_currency_line As Integer, dim_currency_color As Integer
    dim_currency_txt = 0
    dim_currency_code = 1
    dim_currency_color = 2
    dim_currency_line = 3
    
    
k = 0
For i = 14 To 32
    If Worksheets("Parametres").Cells(i, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_currency(k)
        vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, Worksheets("Parametres").Cells(i, 9).Value, i)
        k = k + 1
    End If
Next i

Debug.Print "aim_insert_new_position_in_open: Get currencies from parametres (txt, code, color). " & k & " & currencies found."


Dim vec_tickers() As Variant
k = 0
For i = 0 To UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1)
    ReDim Preserve vec_tickers(k)
    
    If UCase(Left(vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0), 6)) <> UCase("/buid/") Then
        vec_tickers(k) = "/buid/" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
    Else
        vec_tickers(k) = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
    End If
    
    k = k + 1
Next i


    ' s'assurer que present dans les bases internes respectives
    


'appel API pour toutes les nouvelles entrees
Dim bdp_fields As Variant
    bdp_fields = Array("UNDL_ID_BB_UNIQUE", "UNDL_SPOT_TICKER", "NAME", "CRNCY", "PARSEKYABLE_DES", "TICKER_AND_EXCH_CODE", _
        "ID_ISIN", "OPT_UNDL_ISIN", "OPT_CONT_SIZE", "FUT_CONT_SIZE", "OPT_STRIKE_PX", "OPT_EXPIRE_DT", "OPT_EXER_TYP", _
        "SECURITY_DES", "LAST_TRADEABLE_DT", "OPT_PUT_CALL", "OPT_IDX_DVD")
    
    'detection des dim
    Dim dim_bbg_UNDL_ID_BB_UNIQUE As Integer, dim_bbg_UNDL_SPOT_TICKER As Integer, dim_bbg_NAME As Integer, dim_bbg_CRNCY As Integer, _
        dim_bbg_PARSEKYABLE_DES As Integer, dim_bbg_TICKER_AND_EXCH_CODE As Integer, dim_bbg_ID_ISIN As Integer, _
        dim_bbg_OPT_UNDL_ISIN As Integer, dim_bbg_OPT_CONT_SIZE As Integer, dim_bbg_FUT_CONT_SIZE As Integer, _
        dim_bbg_OPT_STRIKE_PX As Integer, dim_bbg_OPT_EXPIRE_DT As Integer, dim_bbg_OPT_EXER_TYP As Integer, _
        dim_bbg_SECURITY_DES As Integer, dim_bbg_LAST_TRADEABLE_DT As Integer, dim_bbg_OPT_PUT_CALL As Integer, _
        dim_bbg_OPT_IDX_DVD As Integer
    
    For i = 0 To UBound(bdp_fields, 1)
        If bdp_fields(i) = "UNDL_ID_BB_UNIQUE" Then
            dim_bbg_UNDL_ID_BB_UNIQUE = i
        ElseIf bdp_fields(i) = "UNDL_SPOT_TICKER" Then
            dim_bbg_UNDL_SPOT_TICKER = i
        ElseIf bdp_fields(i) = "NAME" Then
            dim_bbg_NAME = i
        ElseIf bdp_fields(i) = "CRNCY" Then
            dim_bbg_CRNCY = i
        ElseIf bdp_fields(i) = "PARSEKYABLE_DES" Then
            dim_bbg_PARSEKYABLE_DES = i
        ElseIf bdp_fields(i) = "TICKER_AND_EXCH_CODE" Then
            dim_bbg_TICKER_AND_EXCH_CODE = i
        ElseIf bdp_fields(i) = "ID_ISIN" Then
            dim_bbg_ID_ISIN = i
        ElseIf bdp_fields(i) = "OPT_UNDL_ISIN" Then
            dim_bbg_OPT_UNDL_ISIN = i
        ElseIf bdp_fields(i) = "OPT_CONT_SIZE" Then
            dim_bbg_OPT_CONT_SIZE = i
        ElseIf bdp_fields(i) = "FUT_CONT_SIZE" Then
            dim_bbg_FUT_CONT_SIZE = i
        ElseIf bdp_fields(i) = "OPT_STRIKE_PX" Then
            dim_bbg_OPT_STRIKE_PX = i
        ElseIf bdp_fields(i) = "OPT_EXPIRE_DT" Then
            dim_bbg_OPT_EXPIRE_DT = i
        ElseIf bdp_fields(i) = "OPT_EXER_TYP" Then
            dim_bbg_OPT_EXER_TYP = i
        ElseIf bdp_fields(i) = "SECURITY_DES" Then
            dim_bbg_SECURITY_DES = i
        ElseIf bdp_fields(i) = "LAST_TRADEABLE_DT" Then
            dim_bbg_LAST_TRADEABLE_DT = i
        ElseIf bdp_fields(i) = "OPT_PUT_CALL" Then
            dim_bbg_OPT_PUT_CALL = i
        ElseIf bdp_fields(i) = "OPT_IDX_DVD" Then
            dim_bbg_OPT_IDX_DVD = i
        End If
    Next i
    
    
Dim output_bdp As Variant
Debug.Print "aim_insert_new_position_in_open: $download datas from Bloomberg for new open entries"
output_bdp = oBBG.bdp(vec_tickers, bdp_fields, output_format.of_vec_without_header)


Dim data_internal_db_equity() As Variant, count_equity_db As Integer
    count_equity_db = 0
Dim data_internal_db_index As Variant, count_index_db As Integer
    count_index_db = 0
Dim data_internal_db_future As Variant, count_future_db As Integer
    count_future_db = 0

Dim l_underlying_ccy_row As Integer

Dim count_code_3_equity_db As Integer
    count_code_3_equity_db = 0

Dim v_0a_underlying_id As String 'underlying Id
Dim v_0b_product_id As String 'option Id
Dim v_0c_product_instrument_type As String 'product instrument type
Dim v_0d_underlying_instrument_type As String 'Underlying type
Dim v_0e_option_category As String 'Option category
Dim v_0f_option_type As String 'Option type (Call/Put)
Dim v_0g_underyling_name As String 'Underlying name -> could be override, so take value in internal DB
Dim v_0h_option_strike As Double 'Option strike
Dim v_0i_nbre_contracts As Variant 'Total, nbre contracts
Dim v_0j_position_insert_date As Variant 'Option date, position insert date in open
Dim v_0k_free_calc As Variant 'Close
Dim v_0o_strategy As Variant 'Strategy
Dim v_0p_option_price As Variant 'Last, option pricing
Dim v_0q_result As Variant 'Result
Dim v_0r_delta As Variant 'Delta (formula)
Dim v_0s_gamma As Variant 'Gamma (formula)
Dim v_0t_vega As Variant 'Vega (formula)
Dim v_0u_theta As Variant 'Theta (formula)
Dim v_0v_future_last As Variant 'future last (formula)
Dim v_0w_expiry_date As Date 'Maturity date
Dim v_0x_rd As Variant 'RD (formula)
Dim v_0y_volatility As Variant 'Vol (formula)
Dim v_0z_ddd As Variant 'DDD (formula)
Dim v_ac_hedge As Variant 'Hedge (formula)
Dim v_ad_nav_pct As Variant 'nav_pct (formula)
Dim v_ae_theo_iv As Variant 'theo_iv (formula)
Dim v_af_valeur_eur As Variant 'valeur eur (formula)
Dim v_ah_dividend_date As Variant 'dividend date (formula)
Dim v_ai_dividend_cash As Variant 'dividend cash (formula)
Dim v_ar_derivative_close_price As Variant 'static value
Dim v_as_derivative_close_position As Variant 'static value - de EOD
Dim v_at_derivative_close_valuation As Variant 'formula
Dim v_au_derivative_realized_pnl As Variant 'formula
Dim v_av_derivative_option_pnl As Variant 'formula
Dim v_aw_derivative_position_close_live_view As Variant 'formula
Dim v_ax_vega_all As Variant 'vega_all (formula)
Dim v_ay_theta_all As Variant 'theta_all (formula)
Dim v_az_valeur_eur As Variant 'valeur eur (formula)
Dim v_cm_aim_prime_broker As String ' aim prime broker
Dim v_cn_aim_account As Variant 'aim account
Dim v_co_buidt As String 'id account + produit
Dim v_cp_buidt_underlying As String 'id account + underlying_id
Dim v_cr_accpbpid As String 'account + pb + product id
Dim v_cs_accpbpid_underlying As String 'account + pb + underlying_id
Dim v_cq_sort_id As Variant 'sort id sans currency : product name, account, expiry (formula)
Dim v_cy_row_internal_db As Integer 'row internal db
Dim v_cz_underlying_ticker As String 'underlying ticker
Dim v_da_product_ticker As String 'product ticker
Dim v_db_characteristic As String 'Characteristic
Dim v_dc_currency_code As Double 'Currency code - could be overrided
Dim v_df_quotity_option As Double 'quotity option
Dim v_dg_quotity_future As Double 'quotity future
Dim v_dh_sector_code As Double 'sector code
Dim v_dk_market As String 'market
Dim v_dl_isin As Variant 'isin
     

Dim l_color_index As Integer
     
Dim tmp_vec() As Variant
Dim tmp_formula As String
Dim l_future_spot As String
Dim currency_color_code As Integer
Dim currency_code As Integer

For i = 0 To UBound(vec_product_underlying_ticker_instrument_type_account_and_pb, 1)
    
    'mise a zero des variables
    v_0a_underlying_id = ""
    v_0b_product_id = ""
    v_0c_product_instrument_type = ""
    v_0d_underlying_instrument = ""
    v_0e_option_category = ""
    v_0f_option_type = ""
    v_0g_underyling_name = ""
    v_0h_option_strike = 0
    v_0i_nbre_contracts = 0
    v_0j_position_insert_date = Date
    v_0k_free_calc = ""
    v_0o_strategy = ""
    v_0p_option_price = ""
    v_0q_result = ""
    v_0r_delta = ""
    v_0s_gamma = ""
    v_0t_vega = ""
    v_0u_theta = ""
    v_0v_future_last = ""
    v_0w_expiry_date = Date
    v_0x_rd = ""
    v_0y_volatility = ""
    v_0z_ddd = ""
    v_ac_hedge = ""
    v_ad_nav_pct = ""
    v_ae_theo_iv = ""
    v_af_valeur_eur = ""
    v_ah_dividend_date = ""
    v_ai_dividend_cash = ""
    v_ar_derivative_close_price = ""
    v_as_derivative_close_position = ""
    v_at_derivative_close_valuation = ""
    v_au_derivative_realized_pnl = ""
    v_av_derivative_option_pnl = ""
    v_aw_derivative_position_close_live_view = ""
    v_ax_vega_all = ""
    v_ay_theta_all = ""
    v_az_valeur_eur = ""
    v_cm_aim_prime_broker = ""
    v_cn_aim_account = ""
    v_co_buidt = ""
    v_cp_buidt_underlying = ""
    v_cr_accpbpid = ""
    v_cs_accpbpid_underlying = ""
    v_cq_sort_id = ""
    v_cy_row_internal_db = 0
    v_cz_underlying_ticker = ""
    v_da_product_ticker = ""
    v_db_characteristic = ""
    v_dc_currency_code = 0
    v_df_quotity_option = 0
    v_dg_quotity_future = 0
    v_dh_sector_code = 0
    v_dk_market = ""
    v_dl_isin = ""
    
    
    
    
    
    For j = 0 To UBound(vec_currency, 1)
        If UCase(vec_currency(j)(dim_currency_txt)) = UCase(output_bdp(i)(dim_bbg_CRNCY)) Then
            currency_code = vec_currency(j)(dim_currency_code)
            l_underlying_ccy_row = vec_currency(j)(dim_currency_line)
            currency_color_code = vec_currency(j)(dim_currency_color)
            Exit For
        Else
            If j = UBound(vec_currency, 1) Then
                Debug.Print "aim_insert_new_position_in_open: $problem with currency, " & UCase(output_bdp(i)(dim_bbg_CRNCY)) & " not setup in parametres. " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2) & " will not be insert in open. Bypass to next entry."
                MsgBox ("currency " & output_bdp(i)(dim_bbg_CRNCY) & " not found in Parametres. Bypass entry " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2))
                GoTo bypass_insert_new_entry_in_open
            End If
        End If
    Next j
    
    
    
    Select Case vec_product_underlying_ticker_instrument_type_account_and_pb(i)(3)
    
        Case aim_instrument_type.equity 'EQUITY
            
            If count_equity_db = 0 Then
                'mount les donnees d'equity database
                data_internal_db_equity = aim_mount_data_internal_db_equity(Array(1, 2, 3, 4, 44, 46, 47, 48, 49, 50, 51, 52, 55, 102, 103))
                
                If IsArray(data_internal_db_equity) Then
                    count_equity_db = UBound(data_internal_db_equity, 1)
                Else
                    Debug.Print "aim_insert_new_position_in_open: @no entry in equity database. Quit"
                    MsgBox ("no entry in equtiy db. ->Quit")
                    Exit Sub
                End If
            End If
            
            
            'localise l'underlying dans equity_db
            For m = 0 To UBound(data_internal_db_equity, 1)
                If data_internal_db_equity(m)(1) = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) Then
                    
                    v_0a_underlying_id = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    v_0b_product_id = 0
                    v_0c_product_instrument_type = "S"
                    v_0d_underlying_instrument_type = "E"
                    v_0e_option_category = 0
                    v_0f_option_type = "S"
                    v_0h_option_strike = 1
                    v_0i_nbre_contracts = ""
                    v_0j_position_insert_date = Date
                    v_0k_free_calc = ""
                    v_0o_strategy = ""
                    v_0p_option_price = 0
                    v_0q_result = 0
                    v_0r_delta = "=RC9*RC26+RC29"
                    v_0s_gamma = 0
                    v_0t_vega = 0
                    v_0u_theta = 0
                    v_0v_future_last = Replace("=VLOOKUP(RC1;equity_database;22;FALSE)", ";", ",")
                    v_0w_expiry_date = Date
                    v_0x_rd = ""
                    v_0y_volatility = ""
                    v_0z_ddd = 0
                        
                        'tmp_formula = "=IF(R[-1]C1=RC1;0;VLOOKUP(RC1;<%SHEET%>;<%COL%>;FALSE))"
                        'tmp_formula = Replace(tmp_formula, "<%SHEET%>", "Equity_Database"): tmp_formula = Replace(tmp_formula, "<%COL%>", "24")
                        tmp_formula = "=IF(R[-1]C97=RC97;0;IF(ISNUMBER(VLOOKUP(RC97;AIM_Equities_DPB;3;FALSE));VLOOKUP(RC97;AIM_Equities_DPB;3;FALSE);0))"
                        
                    v_ac_hedge = Replace(tmp_formula, ";", ",")
                    
                    v_ad_nav_pct = Replace("=IF(RC[-29]=R[-1]C[-29],"""",VLOOKUP(RC[-29],Equity_Database!R25C1:R4000C41,34,FALSE))", ";", ",")
                    
                    v_ae_theo_iv = 0
                    
                        tmp_formula = "=RC18*RC22*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", 1)
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_af_valeur_eur = Replace(tmp_formula, ";", ",")
                    
                    v_ah_dividend_date = ""
                    v_ai_dividend_cash = ""
                    
                    v_ar_derivative_close_price = 0
                    v_as_derivative_close_position = 0
                    v_at_derivative_close_valuation = 0
                    v_au_derivative_realized_pnl = 0
                    v_av_derivative_option_pnl = 0
                    v_aw_derivative_position_close_live_view = 0
                    
                    v_ax_vega_all = 0
                    v_ay_theta_all = 0
                    
                        tmp_formula = "=RC18*RC22*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", 1)
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_az_valeur_eur = tmp_formula
                    
                    v_cm_aim_prime_broker = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5)
                    
                    If data_internal_db_equity(m)(15) <> "" Then
                        v_cn_aim_account = Replace(data_internal_db_equity(m)(15), " ", "")
                    Else
                        v_cn_aim_account = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4)
                    End If
                    
                    
                    v_co_buidt = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    v_cp_buidt_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    
                    v_cq_sort_id = "=RC7 & ""_"" & RC92 & ""_"" & RC23" ' possibilite de rendre static
                    
                    v_cr_accpbpid = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    v_cs_accpbpid_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    
                    
                    v_cz_underlying_ticker = aim_patch_ticker_marketplace(vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)) 'plutot que equity_db car aim dynamique
                    
                    v_da_product_ticker = ""
                    v_df_quotity_option = 1
                    v_dg_quotity_future = 1
                    
                    
                    For n = 0 To UBound(data_internal_db_equity(0), 1)
                        If data_internal_db_equity(0)(n) = "line" Then
                            v_cy_row_internal_db = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "Position Statut" Then
                            If data_internal_db_equity(m)(n) = 3 Or data_internal_db_equity(m)(n) = 13 Then
                                count_code_3_equity_db = count_code_3_equity_db + 1
                            End If
                        ElseIf data_internal_db_equity(0)(n) = "Equities_Name" Then
                            v_0g_underyling_name = data_internal_db_equity(m)(n)
                        'ElseIf data_internal_db_equity(0)(n) = "BLOOMBERG" Then
                            'v_cz_underlying_ticker = data_internal_db_equity(m)(n) 'risque d'etre obsolete
                        ElseIf data_internal_db_equity(0)(n) = "Characteristics" Then
                            v_db_characteristic = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "Sector code" Then
                            v_dh_sector_code = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "Currency code color override" Then 'pour le sort
                            v_dc_currency_code = data_internal_db_equity(m)(n)
                            
                            For p = 0 To UBound(vec_currency, 1)
                                If vec_currency(p)(dim_currency_code) = data_internal_db_equity(m)(n) Then
                                    currency_color_code = vec_currency(p)(dim_currency_color)
                                    Exit For
                                Else
                                    If p = UBound(vec_currency, 1) Then
                                        Debug.Print "aim_insert_new_position_in_open: @currency override problem. Code=" & data_internal_db_equity(m)(n) & ". Bypass to next entry"
                                        GoTo bypass_insert_new_entry_in_open
                                    End If
                                End If
                            Next p
                            
                        ElseIf data_internal_db_equity(0)(n) = "Market" Then
                            v_dk_market = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "ISIN" Then
                            v_dl_isin = data_internal_db_equity(m)(n) 'override par api
                        End If
                    Next n
                    
                    Exit For
                Else
                    If m = UBound(data_internal_db_equity, 1) Then
                        Debug.Print "aim_insert_new_position_in_open: @unable to find underyling=" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) & " for product: " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                        MsgBox ("underlying not present in equity db. Bypass " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2))
                        GoTo bypass_insert_new_entry_in_open
                    End If
                End If
            Next m
        
        Case aim_instrument_type.future 'FUTURE
            
            If count_future_db = 0 Then
                data_internal_db_future = aim_mount_data_internal_db_future(Array(1, 2, 3, 4, 104, 107, 110, 112, 113))
                
                If IsArray(data_internal_db_future) Then
                    count_future_db = UBound(data_internal_db_future, 1)
                Else
                    MsgBox ("no entry in index db. ->Quit")
                    Exit Sub
                End If
                
            End If
            
            'localise le futur dans la list d index db
            For m = 1 To UBound(data_internal_db_future, 1)
                If data_internal_db_future(m)(0) = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) Then 'match sur le contract de fut
                    
                    l_future_spot = "=Index_Database!R" & data_internal_db_future(m)(4) & "C35"
                    
                    v_0a_underlying_id = data_internal_db_future(m)(1)
                    v_0b_product_id = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    v_0c_product_instrument_type = "F"
                    v_0d_underlying_instrument_type = "I"
                    v_0e_option_category = 0
                    v_0f_option_type = "F"
                    v_0h_option_strike = 1
                    v_0i_nbre_contracts = ""
                    v_0j_position_insert_date = Date
                    v_0k_free_calc = ""
                    v_0o_strategy = ""
                    v_0p_option_price = 0
                    v_0q_result = 0
                    v_0r_delta = "=RC9*RC26+RC29"
                    v_0s_gamma = 0
                    v_0t_vega = 0
                        For n = 0 To UBound(override_vega_formula, 1)
                            If override_vega_formula(n)(0) = data_internal_db_future(m)(1) Then
                                v_0t_vega = override_vega_formula(n)(1)
                            End If
                        Next n
                    v_0u_theta = 0
                    v_0v_future_last = "=Index_Database!R" & data_internal_db_future(m)(5) & "C35"
                    v_0w_expiry_date = output_bdp(i)(dim_bbg_LAST_TRADEABLE_DT) 'prend la donnee directement dans l'API
                    v_0x_rd = ""
                    v_0y_volatility = ""
                    v_0z_ddd = 0
                        
                        'tmp_formula = "=Index_Database!R" & data_internal_db_future(m)(5) & "C37+Index_Database!R" & data_internal_db_future(m)(5) & "C45"
                        'tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC93,AIM_Futures_DT,3,FALSE)),VLOOKUP(RC93,AIM_Futures_DT,3,FALSE),0)"
                        tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC96,AIM_Futures_DPB,3,FALSE)),VLOOKUP(RC96,AIM_Futures_DPB,3,FALSE),0)"
                    v_ac_hedge = tmp_formula
                    
                    v_ad_nav_pct = Replace("=IF(RC[-29]=R[-1]C[-29],"""",VLOOKUP(RC[-29],Index_Database!R25C1:R120C200,123,FALSE))", ";", ",")
                    v_ae_theo_iv = 0
                    
                        tmp_formula = "=RC18*RC22*<%WEIGHT%>*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_FUT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                        
                        For p = 0 To UBound(weight_valeur_eur, 1)
                            If weight_valeur_eur(p)(0) = data_internal_db_future(m)(1) Then
                                tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "" & weight_valeur_eur(p)(1))
                                Exit For
                            Else
                                If p = UBound(weight_valeur_eur, 1) Then
                                    tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "1")
                                End If
                            End If
                        Next p
                        
                        
                    v_af_valeur_eur = Replace(tmp_formula, ";", ",")
                    
                    v_ah_dividend_date = ""
                    v_ai_dividend_cash = ""
                    
                    v_ar_derivative_close_price = 0
                    v_as_derivative_close_position = 0
                    v_at_derivative_close_valuation = 0
                    v_au_derivative_realized_pnl = 0
                    v_av_derivative_option_pnl = 0
                    v_aw_derivative_position_close_live_view = 0
                    
                    v_ax_vega_all = 0
                    v_ay_theta_all = 0
                        
                        tmp_formula = "=RC18*RC22*<%QUOTITY%>*<%RATE%>*<%WEIGHT%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_FUT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                        
                        For p = 0 To UBound(weight_valeur_eur, 1)
                            If weight_valeur_eur(p)(0) = data_internal_db_future(m)(1) Then
                                tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "" & weight_valeur_eur(p)(1))
                                Exit For
                            Else
                                If p = UBound(weight_valeur_eur, 1) Then
                                    tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "1")
                                End If
                            End If
                        Next p
                        
                    v_az_valeur_eur = Replace(tmp_formula, ";", ",")
                    
                    v_cm_aim_prime_broker = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5)
                    v_cn_aim_account = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4)
                    
                    v_co_buidt = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    v_cp_buidt_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    
                    v_cq_sort_id = "=RC7 & ""_"" & RC92 & ""_"" & RC23"
                    
                    v_cr_accpbpid = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    v_cs_accpbpid_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    
                    
                    v_cy_row_internal_db = data_internal_db_future(m)(4)
                    v_da_product_ticker = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                    v_df_quotity_option = 1
                    v_dg_quotity_future = output_bdp(i)(dim_bbg_FUT_CONT_SIZE)
                    v_dh_sector_code = 0
                    v_dk_market = ""
                    v_dl_isin = ""
                    
                    
                    For n = 0 To UBound(data_internal_db_future(0), 1)
                        If data_internal_db_future(0)(n) = "Futures_Name" Then
                            v_0g_underyling_name = data_internal_db_future(m)(n)
                        ElseIf data_internal_db_future(0)(n) = "BLOOMBERG" Then
                            v_cz_underlying_ticker = data_internal_db_future(m)(n)
                        ElseIf data_internal_db_future(0)(n) = "Characteristics" Then
                            v_db_characteristic = data_internal_db_future(m)(n)
                        ElseIf data_internal_db_future(0)(n) = "Currency code color override" Then
                            v_dc_currency_code = data_internal_db_future(m)(n)
                            
                            For p = 0 To UBound(vec_currency, 1)
                                If vec_currency(p)(dim_currency_code) = data_internal_db_future(m)(n) Then
                                    currency_color_code = vec_currency(p)(dim_currency_color)
                                    Exit For
                                Else
                                    If p = UBound(vec_currency, 1) Then
                                        Debug.Print "aim_insert_new_position_in_open: @currency override problem. Code=" & data_internal_db_future(m)(n) & ". Bypass " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                                        GoTo bypass_insert_new_entry_in_open
                                    End If
                                End If
                            Next p
                            
                        End If
                    Next n
                    
                    Exit For
                Else
                    If m = UBound(data_internal_db_future, 1) Then
                        Debug.Print "aim_insert_new_position_in_open: @unable to find underyling=" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) & " for product: " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                        MsgBox ("future not present in index db. Bypass")
                        GoTo bypass_insert_new_entry_in_open
                    End If
                End If
            Next m
            
        Case aim_instrument_type.option_equity 'OPTION EQUITY
            
            If count_equity_db = 0 Then
                'mount les donnees d'equity database
                data_internal_db_equity = aim_mount_data_internal_db_equity(Array(1, 2, 3, 4, 44, 46, 47, 48, 49, 50, 51, 52, 55, 102, 103))
                
                If IsArray(data_internal_db_equity) Then
                    count_equity_db = UBound(data_internal_db_equity, 1)
                Else
                    MsgBox ("no entry in equity db. ->Quit")
                    Exit Sub
                End If
            End If
            
            
            For m = 0 To UBound(data_internal_db_equity, 1)
                'If data_internal_db_equity(m)(1) = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) Then
                If data_internal_db_equity(m)(1) = output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE) Then 'utilisation de l underyling recuperer de l api pour eviter le probleme du refresh sequentiel des tables AIM qui peut etre legerement desynchro
                    
                    'v_0a_underlying_id = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) ' pas assez fiable a cause des refresh sequentiel
                    v_0a_underlying_id = output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE)
                    v_0b_product_id = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    v_0c_product_instrument_type = "O"
                    v_0d_underlying_instrument_type = "E"
                    v_0e_option_category = "2"
                    v_0f_option_type = UCase(Left(output_bdp(i)(dim_bbg_OPT_PUT_CALL), 1))
                    v_0h_option_strike = output_bdp(i)(dim_bbg_OPT_STRIKE_PX)
                    v_0i_nbre_contracts = Replace("=VLOOKUP(RC96;AIM_OPTIONS_DPB;3;FALSE)", ";", ",")
                    v_0j_position_insert_date = Date
                    v_0k_free_calc = ""
                    v_0o_strategy = ""
                        
                        tmp_formula = "=IF(RC6=""C"";CALL<%TYPE%>;PUT<%TYPE%>)"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                    v_0p_option_price = Replace(tmp_formula, ";", ",")
                    
                        tmp_formula = "=RC16*RC9*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_0q_result = tmp_formula
                    
                        tmp_formula = "=RC9*RC26*<%QUOTITY%>+RC29"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                    v_0r_delta = tmp_formula
                    
                        tmp_formula = "=RC9*RC27*<%QUOTITY%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                    v_0s_gamma = tmp_formula
                    
                        tmp_formula = "=RC9*RC28*<%QUOTITY%>*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE) / 100)
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_0t_vega = Replace(tmp_formula, ";", ",")
                    
                        tmp_formula = "=IF(RC6=""C"";CALL<%TYPE%>1-CALL<%TYPE%>;PUT<%TYPE%>1-PUT<%TYPE%>)*RC9*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_0u_theta = Replace(tmp_formula, ";", ",")
                    
                    v_0v_future_last = Replace("=VLOOKUP(RC1;equity_database;22;FALSE)", ";", ",")
                    
                    v_0w_expiry_date = output_bdp(i)(dim_bbg_OPT_EXPIRE_DT)
                    
                     
                    
                        tmp_formula = "=IF(RC105<>"""";BDP(RC105" & ";""OPT_IMPLIED_VOLATILITY_MID""))"
                    v_0y_volatility = Replace(tmp_formula, ";", ",")
                    
                        tmp_formula = "=DELTA<%TYPE%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                    v_0z_ddd = Replace(tmp_formula, ";", ",")
                    
                        'tmp_formula = "=IF(R[-1]C1=RC1;0;VLOOKUP(RC1;<%SHEET%>;<%COL%>;FALSE))"
                        'tmp_formula = Replace(tmp_formula, "<%SHEET%>", "Equity_Database")
                        'tmp_formula = Replace(tmp_formula, "<%COL%>", "24")
                        tmp_formula = "=IF(R[-1]C97=RC97;0;IF(ISNUMBER(VLOOKUP(RC97;AIM_Equities_DPB;3;FALSE));VLOOKUP(RC97;AIM_Equities_DPB;3;FALSE);0))"
                    v_ac_hedge = Replace(tmp_formula, ";", ",")
                    
                        tmp_formula = "=IF(RC[-29]=R[-1]C[-29],"""",VLOOKUP(RC[-29],Equity_Database!R25C1:R4000C41,34,FALSE))"
                    v_ad_nav_pct = Replace(tmp_formula, ";", ",")
                    
                    
                        tmp_formula = "=Theo_IV_<%TYPE%>*<%QUOTITY%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", "BIN")
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                    v_ae_theo_iv = Replace(tmp_formula, ";", ",")
                    
                    
                        tmp_formula = "=RC18*RC22*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", 1)
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_af_valeur_eur = Replace(tmp_formula, ";", ",")
                    
                    
                    v_ah_dividend_date = ""
                    v_ai_dividend_cash = ""
                    
                        For n = 1 To UBound(output_bbg_underlying, 1)
                            If output_bbg_underlying(n)(0) = "/buid/" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) Then
                                
                                If IsDate(output_bbg_underlying(n)(1)) Then
                                    If output_bbg_underlying(n)(1) >= Date Then
                                        If IsNumeric(output_bbg_underlying(n)(2)) Then
                                            v_ah_dividend_date = output_bbg_underlying(n)(1)
                                            v_ai_dividend_cash = output_bbg_underlying(n)(2)
                                        End If
                                    End If
                                End If
                                
                                Exit For
                            End If
                        Next n
                        
                    
                    
                        tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC[-42],AIM_EOD_JSO,7,FALSE)),VLOOKUP(RC[-42],AIM_EOD_JSO,7,FALSE),0)"
                        
                        For n = 0 To UBound(correction_factor_close_price, 1)
                            If correction_factor_close_price(n)(0) = currency_code Then
                                tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC[-42],AIM_EOD_JSO,7,FALSE)),<%CORRECTION_FACTOR%>*VLOOKUP(RC[-42],AIM_EOD_JSO,7,FALSE),0)"
                                tmp_formula = Replace(tmp_formula, "<%CORRECTION_FACTOR%>", correction_factor_close_price(n)(1))
                                Exit For
                            End If
                        Next n
                        
                    v_ar_derivative_close_price = tmp_formula
                    
                        tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC96,AIM_Options_DPB,4,FALSE)),VLOOKUP(RC96,AIM_Options_DPB,4,FALSE),0)"
                    v_as_derivative_close_position = tmp_formula
                
                        tmp_formula = "=RC[-2]*RC[-37]*RC[64]"
                    v_at_derivative_close_valuation = tmp_formula
                    
                        'tmp_formula = "=(IF(ISNUMBER(VLOOKUP(RC[-45],AIM_Options_JSO,10,FALSE)),VLOOKUP(RC[-45],AIM_Options_JSO,10,FALSE),0)+((RC9-RC45)*RC110*RC44))"
                        'avec support echeance
                        tmp_formula = "=IF(AND(EX_DATE>RC[-24],RC[-38]=0),0,(IF(ISNUMBER(VLOOKUP(RC96,AIM_Options_DPB,6,FALSE)),VLOOKUP(RC96,AIM_Options_DPB,6,FALSE),0)+((RC9-RC45)*RC110*RC44)))"
                        
                        For n = 0 To UBound(correction_factor_net_trading_cash_flow, 1)
                            If correction_factor_net_trading_cash_flow(n)(0) = currency_code Then
                                tmp_formula = "=IF(AND(EX_DATE>RC[-24],RC[-38]=0),0,(IF(ISNUMBER(VLOOKUP(RC96,AIM_Options_DPB,6,FALSE)),<%CORRECTION_FACTOR%>*VLOOKUP(RC96,AIM_Options_DPB,6,FALSE),0)+((RC9-RC45)*RC110*RC44)))"
                                tmp_formula = Replace(tmp_formula, "<%CORRECTION_FACTOR%>", correction_factor_net_trading_cash_flow(n)(1))
                            End If
                        Next n
                        
                    v_au_derivative_realized_pnl = tmp_formula
                    
                        tmp_formula = "=(RC9*RC16*RC110)-RC46+RC47"
                    v_av_derivative_option_pnl = tmp_formula
                    
                        tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC2,AIM_Options_JSO,8,FALSE)),VLOOKUP(RC2,AIM_Options_JSO,8,FALSE),0)"
                    v_aw_derivative_position_close_live_view = tmp_formula
                    
                    
                    
                        tmp_formula = "=RC9*RC28*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE) / 100)
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_ax_vega_all = tmp_formula
                    
                        tmp_formula = "=IF(RC6=""C"";CALL<%TYPE%>1-CALL<%TYPE%>;PUT<%TYPE%>1-PUT<%TYPE%>)*RC9*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_ay_theta_all = Replace(tmp_formula, ";", ",")
                    
                        tmp_formula = "=RC18*RC22*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_az_valeur_eur = Replace(tmp_formula, ";", ",")
                    
                    v_cm_aim_prime_broker = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5)
                    
                    If data_internal_db_equity(m)(15) <> "" Then
                        v_cn_aim_account = data_internal_db_equity(m)(15)
                    Else
                        v_cn_aim_account = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4)
                    End If
                    
                    v_co_buidt = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    'v_cp_buidt_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    v_cp_buidt_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE)
                    
                    v_cq_sort_id = "=RC7 & ""_"" & RC92 & ""_"" & RC23"
                    
                    v_cr_accpbpid = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    'v_cs_accpbpid_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    v_cs_accpbpid_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE)
                    
                    
                    v_da_product_ticker = output_bdp(i)(dim_bbg_SECURITY_DES) & " Equity"
                    
                    v_df_quotity_option = output_bdp(i)(dim_bbg_OPT_CONT_SIZE)
                    v_dg_quotity_future = 1
                    
                    
                    For n = 0 To UBound(data_internal_db_equity(0), 1)
                        If data_internal_db_equity(0)(n) = "Equities_Name" Then
                            v_0g_underyling_name = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "Position Statut" Then
                            If data_internal_db_equity(m)(n) = 3 Or data_internal_db_equity(m)(n) = 13 Then
                                count_code_3_equity_db = count_code_3_equity_db + 1
                            End If
                        ElseIf data_internal_db_equity(0)(n) = "Currency" Then
                                tmp_formula = "=BDP(""" & aim_get_risk_free_asset_ticker_based_on_currency_and_expiry(data_internal_db_equity(m)(n), CInt(output_bdp(i)(dim_bbg_OPT_EXPIRE_DT) - Date)) & """;""LAST_TRADE"")/100"
                            v_0x_rd = Replace(tmp_formula, ";", ",")
                            
                            
                            
                        ElseIf data_internal_db_equity(0)(n) = "line" Then
                            v_cy_row_internal_db = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "BLOOMBERG" Then
                            v_cz_underlying_ticker = aim_patch_ticker_marketplace(data_internal_db_equity(m)(n))
                        ElseIf data_internal_db_equity(0)(n) = "Characteristics" Then
                            v_db_characteristic = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "Currency code color override" Then
                            v_dc_currency_code = data_internal_db_equity(m)(n)
                            
                            For p = 0 To UBound(vec_currency, 1)
                                If vec_currency(p)(dim_currency_code) = data_internal_db_equity(m)(n) Then
                                    currency_color_code = vec_currency(p)(dim_currency_color)
                                    Exit For
                                Else
                                    If p = UBound(vec_currency, 1) Then
                                        Debug.Print "aim_insert_new_position_in_open: @currency override problem. Code=" & data_internal_db_equity(m)(n) & ". Bypass " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                                        GoTo bypass_insert_new_entry_in_open
                                    End If
                                End If
                            Next p
                            
                        ElseIf data_internal_db_equity(0)(n) = "Sector code" Then
                            v_dh_sector_code = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "Market" Then
                            v_dk_market = data_internal_db_equity(m)(n)
                        ElseIf data_internal_db_equity(0)(n) = "ISIN" Then
                            v_dl_isin = data_internal_db_equity(m)(n)
                        End If
                    Next n
                    
                    Exit For
                Else
                    If m = UBound(data_internal_db_equity, 1) Then
                        Debug.Print "aim_insert_new_position_in_open: @unable to find underyling=" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) & " for product: " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                        MsgBox ("underlying not present in equity db. Bypass")
                        GoTo bypass_insert_new_entry_in_open
                    End If
                End If
            Next m
        
        
        Case aim_instrument_type.option_index 'OPTION INDEX
            
            If count_index_db = 0 Then
                data_internal_db_index = aim_mount_data_internal_db_index(Array(2, 3, 104, 107, 110, 112, 113))
                
                If IsArray(data_internal_db_index) Then
                    count_index_db = UBound(data_internal_db_index, 1)
                Else
                    MsgBox ("no entry in index db. ->Quit")
                    Exit Sub
                End If
                
            End If
            
            
            For m = 0 To UBound(data_internal_db_index, 1)
                
                'If vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) = data_internal_db_index(m)(0) Then
                If output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE) = data_internal_db_index(m)(0) Then
                
                
                    'v_0a_underlying_id = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) ' pas assez fiable a cause des refresh sequentiel
                    v_0a_underlying_id = output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE)
                    v_0b_product_id = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    v_0c_product_instrument_type = "O"
                    v_0d_underlying_instrument_type = "I"
                    v_0e_option_category = "1"
                    v_0f_option_type = UCase(Left(output_bdp(i)(dim_bbg_OPT_PUT_CALL), 1))
                    v_0h_option_strike = output_bdp(i)(dim_bbg_OPT_STRIKE_PX)
                    'v_0i_nbre_contracts = Replace("=VLOOKUP(RC2;AIM_OPTIONS_JSO;7;FALSE)", ";", ",")
                    'v_0i_nbre_contracts = Replace("=VLOOKUP(RC93;AIM_OPTIONS_DT;3;FALSE)", ";", ",")
                    v_0i_nbre_contracts = Replace("=VLOOKUP(RC96;AIM_OPTIONS_DPB;3;FALSE)", ";", ",")
                    v_0j_position_insert_date = Date
                    v_0k_free_calc = ""
                    v_0o_strategy = ""
                    
                        tmp_formula = "=IF(RC6=""C"";CALL<%TYPE%>;PUT<%TYPE%>)"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                    v_0p_option_price = Replace(tmp_formula, ";", ",")
                    
                        tmp_formula = "=RC16*RC9*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_0q_result = tmp_formula
                    
                        tmp_formula = "=RC9*RC28*<%QUOTITY%>*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%RATE%>"
                        
                        For j = 0 To UBound(override_vega_formula, 1)
                            If override_vega_formula(j)(0) = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) Then
                                Debug.Print "aim_insert_new_position_in_open: $found vega_formula override: " & override_vega_formula(j)(1)
                                tmp_formula = override_vega_formula(j)(1)
                            End If
                        Next j
                        
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE) / 100)
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    
                    v_0t_vega = tmp_formula
                    
                    
                        tmp_formula = "=IF(RC6=""C"";CALL<%TYPE%>1-CALL<%TYPE%>;PUT<%TYPE%>1-PUT<%TYPE%>)*RC9*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_0u_theta = Replace(tmp_formula, ";", ",")
                    
                    v_0v_future_last = Replace("=VLOOKUP(RC1;index_database;36;FALSE)", ";", ",")
                    
                    v_0w_expiry_date = output_bdp(i)(dim_bbg_OPT_EXPIRE_DT)
                    
                        tmp_formula = "=IF(RC105<>"""";BDP(RC105;""OPT_IMPLIED_VOLATILITY_MID""))"
                    v_0y_volatility = Replace(tmp_formula, ";", ",")
                    
                        tmp_formula = "=DELTA<%TYPE%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                    v_0z_ddd = Replace(tmp_formula, ";", ",")
                    
                    v_ac_hedge = 0 'une ligne pour chaque future est de toute facon remontee
                    
                        tmp_formula = "=IF(RC[-29]=R[-1]C[-29];"""";VLOOKUP(RC[-29];Index_Database!R25C1:R122C130;123;FALSE))"
                    v_ad_nav_pct = Replace(tmp_formula, ";", ",")
                        
                        tmp_formula = "=Theo_IV_<%TYPE%>*<%QUOTITY%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", "BS")
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", "100")
                    v_ae_theo_iv = tmp_formula
                    
                    
                    v_ah_dividend_date = ""
                    v_ai_dividend_cash = ""
                    If IsDate(output_bdp(i)(dim_bbg_OPT_EXPIRE_DT)) And IsNumeric(output_bdp(i)(dim_bbg_OPT_IDX_DVD)) Then
                        v_ah_dividend_date = output_bdp(i)(dim_bbg_OPT_EXPIRE_DT)
                        'v_ai_dividend_cash = output_bdp(i)(dim_bbg_OPT_IDX_DVD)
                        v_ai_dividend_cash = "=BDP(RC105,""OPT_IDX_DVD"")" 'formule bdp dynamique
                    End If
                    
                    
                        tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC[-42],AIM_EOD_JSO,7,FALSE)),VLOOKUP(RC[-42],AIM_EOD_JSO,7,FALSE),0)"
                    v_ar_derivative_close_price = tmp_formula
                    
                        'tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC[-43],AIM_EOD_JSO,8,FALSE)),VLOOKUP(RC[-43],AIM_EOD_JSO,8,FALSE),0)"
                        'tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC93,AIM_EOD_DT,3,FALSE)),VLOOKUP(RC93,AIM_EOD_DT,3,FALSE),0)"
                        tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC96,AIM_Options_DPB,4,FALSE)),VLOOKUP(RC96,AIM_Options_DPB,4,FALSE),0)"
                    v_as_derivative_close_position = tmp_formula
                
                        tmp_formula = "=RC[-2]*RC[-37]*RC[64]"
                    v_at_derivative_close_valuation = tmp_formula
                    
                        'tmp_formula = "=(IF(ISNUMBER(VLOOKUP(RC[-45],AIM_Options_JSO,10,FALSE)),VLOOKUP(RC[-45],AIM_Options_JSO,10,FALSE),0)+((RC9-RC45)*RC110*RC44))"
                        'tmp_formula = "=(IF(ISNUMBER(VLOOKUP(RC93,AIM_Options_DT,6,FALSE)),VLOOKUP(RC93,AIM_Options_DT,6,FALSE),0)+((RC9-RC45)*RC110*RC44))"
                        tmp_formula = "=IF(AND(EX_DATE>RC[-24],RC[-38]=0),0,(IF(ISNUMBER(VLOOKUP(RC96,AIM_Options_DPB,6,FALSE)),VLOOKUP(RC96,AIM_Options_DPB,6,FALSE),0)+((RC9-RC45)*RC110*RC44)))"
                    v_au_derivative_realized_pnl = tmp_formula
                    
                        tmp_formula = "=(RC9*RC16*RC110)-RC46+RC47"
                    v_av_derivative_option_pnl = tmp_formula
                    
                        'tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC2,AIM_Options_JSO,8,FALSE)),VLOOKUP(RC2,AIM_Options_JSO,8,FALSE),0)"
                        tmp_formula = "=IF(ISNUMBER(VLOOKUP(RC93,AIM_Options_DT,4,FALSE)),VLOOKUP(RC93,AIM_Options_DT,4,FALSE),0)"
                    v_aw_derivative_position_close_live_view = tmp_formula
                    
                        tmp_formula = "=RC9*RC28*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE) / 100)
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_ax_vega_all = tmp_formula
                    
                        tmp_formula = "=IF(RC6=""C"";CALL<%TYPE%>1-CALL<%TYPE%>;PUT<%TYPE%>1-PUT<%TYPE%>)*RC9*<%QUOTITY%>*<%RATE%>"
                        tmp_formula = Replace(tmp_formula, "<%TYPE%>", UCase(Left(output_bdp(i)(dim_bbg_OPT_EXER_TYP), 1)))
                        tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE))
                        tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                    v_ay_theta_all = Replace(tmp_formula, ";", ",")
                    
                    v_cm_aim_prime_broker = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5)
                    v_cn_aim_account = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4)
                    
                    v_co_buidt = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    'v_cp_buidt_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    v_cp_buidt_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE)
                    
                    v_cq_sort_id = "=RC7 & ""_"" & RC92 & ""_"" & RC23"
                    
                    v_cr_accpbpid = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0)
                    'v_cs_accpbpid_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1)
                    v_cs_accpbpid_underlying = vec_product_underlying_ticker_instrument_type_account_and_pb(i)(4) & "_" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(5) & "_" & output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE)
                    
                    
                    v_da_product_ticker = output_bdp(i)(dim_bbg_SECURITY_DES) & " Index"
                    
                    v_df_quotity_option = output_bdp(i)(dim_bbg_OPT_CONT_SIZE)
                    
                    v_dh_sector_code = 0
                    
                    v_dk_market = ""
                    
                    v_dl_isin = ""
                    
                    
                    For n = 0 To UBound(data_internal_db_index(0), 1)
                        If data_internal_db_index(0)(n) = "Futures_Name" Then
                            v_0g_underyling_name = data_internal_db_index(m)(n)
                        ElseIf data_internal_db_index(0)(n) = "line" Then
                            v_cy_row_internal_db = data_internal_db_index(m)(n)
                        ElseIf data_internal_db_index(0)(n) = "Quotite" Then
                            
                                tmp_formula = "=RC9*RC26*<%QUOTITY%>+RC29"
                                tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", output_bdp(i)(dim_bbg_OPT_CONT_SIZE) / data_internal_db_index(m)(n))
                            v_0r_delta = tmp_formula
                            
                            
                                tmp_formula = "=RC9*RC27*<%QUOTITY%>"
                                tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", 1 / data_internal_db_index(m)(n))
                            v_0s_gamma = tmp_formula
                            
                            
                                tmp_formula = "=RC18*RC22*<%WEIGHT%>*IF(Parametres!R14C19=0,1,(1/Parametres!R12C18))*<%QUOTITY%>*<%RATE%>"
                                tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", data_internal_db_index(m)(n))
                                tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                                
                                For p = 0 To UBound(weight_valeur_eur, 1)
                                    If weight_valeur_eur(p)(0) = output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE) Then
                                        Debug.Print "aim_insert_new_position_in_open: $found weighted_valeur_eur_formula override: " & weight_valeur_eur(p)(1)
                                        tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "" & weight_valeur_eur(p)(1))
                                        Exit For
                                    Else
                                        If p = UBound(weight_valeur_eur, 1) Then
                                            tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "1")
                                        End If
                                    End If
                                Next p
                                
                                
                            v_af_valeur_eur = tmp_formula
                            
                            
                                tmp_formula = "=RC18*RC22*<%QUOTITY%>*<%RATE%>*<%WEIGHT%>"
                                tmp_formula = Replace(tmp_formula, "<%QUOTITY%>", 1 / data_internal_db_index(m)(n))
                                tmp_formula = Replace(tmp_formula, "<%RATE%>", "Parametres!R" & l_underlying_ccy_row & "C6")
                                
                                For p = 0 To UBound(weight_valeur_eur, 1)
                                    If weight_valeur_eur(p)(0) = output_bdp(i)(dim_bbg_UNDL_ID_BB_UNIQUE) Then
                                        Debug.Print "aim_insert_new_position_in_open: $found weighted_valeur_eur_formula override: " & weight_valeur_eur(p)(1)
                                        tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "" & weight_valeur_eur(p)(1))
                                        Exit For
                                    Else
                                        If p = UBound(weight_valeur_eur, 1) Then
                                            tmp_formula = Replace(tmp_formula, "<%WEIGHT%>", "1")
                                        End If
                                    End If
                                Next p
                                
                            v_az_valeur_eur = tmp_formula
                            
                            v_dg_quotity_future = data_internal_db_index(m)(n)
                            
                        ElseIf data_internal_db_index(0)(n) = "Currency" Then
                                tmp_formula = "=BDP(""" & aim_get_risk_free_asset_ticker_based_on_currency_and_expiry(data_internal_db_index(m)(n), CInt(output_bdp(i)(dim_bbg_OPT_EXPIRE_DT) - Date)) & """;""LAST_TRADE"")/100"
                            v_0x_rd = Replace(tmp_formula, ";", ",")
                            
                            
                        ElseIf data_internal_db_index(0)(n) = "Currency code color override" Then
                            
                            v_dc_currency_code = data_internal_db_index(m)(n)
                            
                            For p = 0 To UBound(vec_currency, 1)
                                If vec_currency(p)(dim_currency_code) = data_internal_db_index(m)(n) Then
                                    currency_color_code = vec_currency(p)(dim_currency_color)
                                    Exit For
                                Else
                                    If p = UBound(vec_currency, 1) Then
                                        Debug.Print "aim_insert_new_position_in_open: @currency override problem. Code=" & data_internal_db_index(m)(n) & ". Bypass " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                                        GoTo bypass_insert_new_entry_in_open
                                    End If
                                End If
                            Next p
                            
                        ElseIf data_internal_db_index(0)(n) = "BLOOMBERG" Then
                            
                            v_cz_underlying_ticker = data_internal_db_index(m)(n)
                        
                        ElseIf data_internal_db_index(0)(n) = "Characteristics" Then
                            
                            v_db_characteristic = data_internal_db_index(m)(n)
                            
                        End If
                    Next n
                    
                    Exit For
                Else
                    If m = UBound(data_internal_db_index, 1) Then
                        Debug.Print "aim_insert_new_position_in_open: @unable to find underyling=" & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(1) & " for product: " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
                        MsgBox ("underlying not present in index db. Bypass")
                        GoTo bypass_insert_new_entry_in_open
                    End If
                End If
                
            Next m
            
    End Select
    
    Dim l_open_last_row As Integer
    
    Dim empty_line_thresold As Integer
    empty_line_threshold = 20
    Dim is_end_of_open As Boolean
    
    With Worksheets("Open")
        
        Debug.Print "aim_insert_new_position_in_open: $insert_position_in_open for " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(0) & " - " & vec_product_underlying_ticker_instrument_type_account_and_pb(i)(2)
        
        For j = l_header_internal_open To 3200
            is_end_of_open = True
            
            For m = 0 To empty_line_threshold
                If .Cells(j + m, 1) <> "" Then
                    is_end_of_open = False
                    Exit For
                End If
            Next m
            
            If is_end_of_open = True Then
                l_open_last_row = j - 1
                Exit For
            End If
        Next j
        
        
        .Activate
        
        'Copy the line at the end
        .Range("A4:IV4").Copy (.Range("A" & l_open_last_row + 1))
            
            'mise en place des formules
            For j = 1 To 250
                If .Cells(2, j).Value = "F" Then
                    .Cells(l_open_last_row + 1, j).formula = "=" & Replace(.Cells(l_open_last_row + 1, j).Value, ";", ",")
                End If
            Next j
            
        
        .rows(l_open_last_row + 1).EntireRow.Hidden = False
        .rows(l_open_last_row + 1).Select
        
            .Range("H" & l_open_last_row + 1 & ":IV" & l_open_last_row + 1).HorizontalAlignment = xlCenter
            .Range("A" & l_open_last_row + 1 & ":IV" & l_open_last_row + 1).Font.name = "Times New Roman"
            .Range("A" & l_open_last_row + 1 & ":B" & l_open_last_row + 1).Font.size = 7
            .Range("C" & l_open_last_row + 1 & ":IV" & l_open_last_row + 1).Font.size = 10

                'Format cells
                .Cells(l_open_last_row + 1, 8).NumberFormat = "General"
                .Cells(l_open_last_row + 1, 9).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 10).NumberFormat = "dd.mm"
                .Cells(l_open_last_row + 1, 11).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 15).NumberFormat = "#,##0.00"
                .Cells(l_open_last_row + 1, 15).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 16).NumberFormat = "#,##0.00"
                .Cells(l_open_last_row + 1, 16).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 17).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 17).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 18).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 18).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 19).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 19).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 20).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 20).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 21).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 21).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 22).NumberFormat = "#,##0.00"
                .Cells(l_open_last_row + 1, 22).HorizontalAlignment = xlRight
                .Cells(l_open_last_row + 1, 23).NumberFormat = "dd mmm yy"
                .Cells(l_open_last_row + 1, 24).NumberFormat = "#,##0.00%"
                .Cells(l_open_last_row + 1, 25).NumberFormat = "#,##0.00"
                .Cells(l_open_last_row + 1, 29).NumberFormat = "0"
                .Cells(l_open_last_row + 1, 30).NumberFormat = "#,##0.00%"
                .Cells(l_open_last_row + 1, 32).NumberFormat = "#,##0"
                .Cells(l_open_last_row + 1, 33).NumberFormat = "0"
                .Cells(l_open_last_row + 1, 107).NumberFormat = "0.00"
                .Cells(l_open_last_row + 1, 108).NumberFormat = "0.00"
                .Cells(l_open_last_row + 1, 109).NumberFormat = "0.00"
                
        
        'A - underlying id
        .Cells(l_open_last_row + 1, 1).Value = v_0a_underlying_id

        'B - product id
        .Cells(l_open_last_row + 1, 2).Value = v_0b_product_id

        'C - instrument type
        .Cells(l_open_last_row + 1, 3).Value = v_0c_product_instrument_type

        'D - underyling instrument type
        .Cells(l_open_last_row + 1, 4).Value = v_0d_underlying_instrument_type

        'E - option category (on equity / on index)
        .Cells(l_open_last_row + 1, 5).Value = v_0e_option_category

        'F - option type (call / put)
        .Cells(l_open_last_row + 1, 6).Value = v_0f_option_type

        'G - name
        .Cells(l_open_last_row + 1, 7).Value = v_0g_underyling_name

        'H - strike
        .Cells(l_open_last_row + 1, 8).Value = v_0h_option_strike

        'I - nbre contracts
        .Cells(l_open_last_row + 1, 9).formula = v_0i_nbre_contracts

        'J - insert position date
        .Cells(l_open_last_row + 1, 10).Value = v_0j_position_insert_date

        'K - free calc
        .Cells(l_open_last_row + 1, 11).Value = v_0k_free_calc
        
        'O - strategy
        If v_0c_product_instrument_type <> "O" Then 'maintient la formule si option
            .Cells(l_open_last_row + 1, 15).Value = v_0o_strategy
        End If
        
        'P - pricing option
        .Cells(l_open_last_row + 1, 16).Value = v_0p_option_price
        
        'Q - RESULT
        .Cells(l_open_last_row + 1, 17).Value = v_0q_result
        
        'R - DELTA
        .Cells(l_open_last_row + 1, 18).Value = v_0r_delta
        
        'S - GAMMA
        .Cells(l_open_last_row + 1, 19).Value = v_0s_gamma
        
        'T - VEGA+1
        .Cells(l_open_last_row + 1, 20).Value = v_0t_vega
        
        'U - THETA
        .Cells(l_open_last_row + 1, 21).Value = v_0u_theta
        
        'V - underlying price
        .Cells(l_open_last_row + 1, 22).Value = v_0v_future_last
        
        'W - exp_date
        .Cells(l_open_last_row + 1, 23).Value = v_0w_expiry_date
        
        'X - rd
        .Cells(l_open_last_row + 1, 24).Value = v_0x_rd
        
        'Y - volatility
        .Cells(l_open_last_row + 1, 25).Value = v_0y_volatility
        
        'Z - ddd
        .Cells(l_open_last_row + 1, 26).Value = v_0z_ddd
        
        'AC - hedge
        .Cells(l_open_last_row + 1, 29).Value = v_ac_hedge
        
        'AD - Nav pct
        .Cells(l_open_last_row + 1, 30).Value = v_ad_nav_pct
        
        'AE - theo iv
        .Cells(l_open_last_row + 1, 31).Value = v_ae_theo_iv
        
        'AF - valeur eur
        .Cells(l_open_last_row + 1, 32).Value = v_af_valeur_eur
        
        'AH - dividend ex-date
        .Cells(l_open_last_row + 1, 34).Value = v_ah_dividend_date
        
        'AI - dividend
        .Cells(l_open_last_row + 1, 35).Value = v_ai_dividend_cash
        
        'AR - price close
        .Cells(l_open_last_row + 1, 44).Value = v_ar_derivative_close_price
        
        'AS - position close
        .Cells(l_open_last_row + 1, 45).Value = v_as_derivative_close_position
        
        'AT - valorisation close
        .Cells(l_open_last_row + 1, 46).Value = v_at_derivative_close_valuation
        
        'AU - realized_pl
        .Cells(l_open_last_row + 1, 47).Value = v_au_derivative_realized_pnl
        
        'AV - option P&L
        .Cells(l_open_last_row + 1, 48).Value = v_av_derivative_option_pnl
        
        'AW - postion close live view
        .Cells(l_open_last_row + 1, 49).Value = v_aw_derivative_position_close_live_view
        
        'AX - Vega_1%_ALL
        .Cells(l_open_last_row + 1, 50).Value = v_ax_vega_all
        
        'AY - Theta
        .Cells(l_open_last_row + 1, 51).Value = v_ay_theta_all
        
        'AZ - Valeur EUR
        .Cells(l_open_last_row + 1, 52).Value = v_az_valeur_eur
        
        'CM - pb account
        .Cells(l_open_last_row + 1, 91).Value = v_cm_aim_prime_broker
        
        'CN - aim account
        .Cells(l_open_last_row + 1, 92).Value = v_cn_aim_account
        
        'CO - buidt
        .Cells(l_open_last_row + 1, 93).Value = v_co_buidt
        
        'CP - buidt_underyling
        .Cells(l_open_last_row + 1, 94).Value = v_cp_buidt_underlying
        
        'CQ - sort_id
        .Cells(l_open_last_row + 1, 95).Value = v_cq_sort_id
        
        'CR - accpbpid
        .Cells(l_open_last_row + 1, 96).Value = v_cr_accpbpid
        
        'CS - accpbid_underlying
        .Cells(l_open_last_row + 1, 97).Value = v_cs_accpbpid_underlying
        
        'CY - row internal db
        .Cells(l_open_last_row + 1, 103).Value = v_cy_row_internal_db
        
        'CZ - underyling ticker
        .Cells(l_open_last_row + 1, 104).Value = v_cz_underlying_ticker
        
        'DA - product ticker
        .Cells(l_open_last_row + 1, 105).Value = v_da_product_ticker
        
        'DB - underlying characteristics
        .Cells(l_open_last_row + 1, 106).Value = v_db_characteristic
        
        'DC - currency code
        .Cells(l_open_last_row + 1, 107).Value = v_dc_currency_code
        
        'DF - quotity option
        .Cells(l_open_last_row + 1, 110).Value = v_df_quotity_option
        
        'DG - quotity future
        .Cells(l_open_last_row + 1, 111).Value = v_dg_quotity_future
        
        'DH - sector code
        .Cells(l_open_last_row + 1, 112).Value = v_dh_sector_code
        
        'DK - market
        .Cells(l_open_last_row + 1, 115).Value = v_dk_market
        
        'DL - isin
        .Cells(l_open_last_row + 1, 116).Value = v_dl_isin
        
        .rows(l_open_last_row + 1).Interior.ColorIndex = currency_color_code 'coloriage de la ligne grace a la devise ou l'override devise
            .Cells(l_open_last_row + 1, 29).Interior.ColorIndex = 1
            .Cells(l_open_last_row + 1, 29).Font.ColorIndex = 2
            
            If main_aim_account <> "" Then
                If main_aim_account <> v_cn_aim_account Then
                    'passage de la ligne en italic
                    .rows(l_open_last_row + 1).Font.Italic = True
                End If
            End If
        
            
    End With
    
    l_open_last_row = l_open_last_row + 1
    
bypass_insert_new_entry_in_open:

Next i


Dim vb_answer As Variant
If count_code_3_equity_db > 0 Then
    vb_answer = MsgBox("Some entries are desactivated. Run DB Equities->Set Status ?", vbYesNo, "DB Equities->Set Status ?")
    
    If vb_answer = vbYes Then
        Debug.Print "aim_insert_new_position_in_open: $set_Status_Equities_SPEED"
        Call set_Status_Equities_SPEED
    End If
    
End If


End Sub


Public Sub aim_reorder_open(Optional ByVal l_open_last_row As Variant)

Application.Calculation = xlCalculationManual

Worksheets("Open").Activate

Dim open_empty_threshold As Integer
open_empty_threshold = 25

Dim is_last_line As Boolean


If IsMissing(l_open_last_row) Then
    For i = 26 To 32000
        
        is_last_line = True
        
        For j = 0 To open_empty_threshold
            If Worksheets("Open").Cells(i + j, 1) <> "" Then
                is_last_line = False
                Exit For
            End If
        Next j
        
        If is_last_line = True Then
            l_open_last_row = i - 1
            Exit For
        End If
        
    Next i
End If


Worksheets("Open").Range("A25:IV" & l_open_last_row).sort Key1:=Range("DC26"), Order1:=xlAscending, Key2:=Range("CQ26"), Order2:=xlAscending, header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


End Sub


Public Function aim_get_risk_free_asset_ticker_based_on_currency_and_expiry(ByVal currency_code As Integer, day_to_expiry As Integer) As String

aim_get_risk_free_asset_ticker_based_on_currency_and_expiry = "USDR1 CURNCY"

Dim l_ccy_string  As String

If currency_code = 1 Then  'chf
    l_ccy_string = "SF"
ElseIf currency_code = 2 Then  'eur
    l_ccy_string = "EU"
Else
    l_ccy_string = "US"
End If



If day_to_expiry < 45 Then
    l_ccy_string = l_ccy_string & "DRA"
ElseIf day_to_expiry < 75 Then
    l_ccy_string = l_ccy_string & "DRB"
ElseIf day_to_expiry < 105 Then
    l_ccy_string = l_ccy_string & "DRC"
ElseIf day_to_expiry < 205 Then
    l_ccy_string = l_ccy_string & "DRF"
Else
    l_ccy_string = l_ccy_string & "DR1"
End If

aim_get_risk_free_asset_ticker_based_on_currency_and_expiry = l_ccy_string & space(1) & "CURNCY"

End Function


Public Function aim_get_view_kronos_db(ByVal aim_view_code As Integer) As String

aim_get_view_kronos_db = Worksheets(aim_get_worksheet_xls_name_from_view_code(aim_view_code)).Cells(1, 2).Value

End Function


Public Function aim_get_rtd_data(ByVal progID As String, ByVal server As String, ByVal aim_view As String, ByVal aim_view_column As String, ByVal aim_view_id_line As Integer) As Variant

aim_get_rtd_data = WorksheetFunction.RTD(progID, server, aim_view, aim_view_column, CStr(aim_view_id_line))

End Function


Public Sub aim_manual_update_datas_views(ByVal vec_view_and_columns As Variant)

Dim datas_view As Variant

Dim xls_sheet As String

Dim i As Integer, j As Integer
For i = 0 To UBound(vec_view_and_columns, 1)
    datas_view = aim_get_datas_view_rtd(vec_view_and_columns(i)(0), vec_view_and_columns(i)(1), output_format_rtd.vec_without_header)
    Worksheets(aim_get_worksheet_xls_name_from_view_code(vec_view_and_columns(i)(0))).Range("B" & l_header_aim_view + 1 & ":" & xlColumnValue(UBound(datas_view, 2) + 1) & l_header_aim_view + UBound(datas_view, 1)).Value2 = datas_view
Next i

End Sub


Public Function aim_get_datas_view_rtd(ByVal aim_view_code As Integer, ByVal vec_column As Variant, Optional ByVal output_format As Integer = output_format_rtd.range_with_header) As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

aim_get_datas_view_rtd = Empty

Dim tmp_view As String
Dim tmp_value As Variant
Dim is_last_line As Boolean

Dim tmp_view_name As String

Dim vec_data() As Variant
Dim tmp_vec() As Variant


tmp_view_name = aim_get_view_kronos_db(aim_view_code)
k = 0


For i = 0 To UBound(vec_column, 1)
    ReDim Preserve tmp_vec(i)
    tmp_vec(i) = vec_column(i)
Next i

ReDim Preserve vec_data(0)
vec_data(0) = tmp_vec

For i = 1 To 32000
    tmp_value = aim_get_rtd_data("Kronos.RTDServer", "", tmp_view_name, tmp_view_name & "_id", i)
    
    If tmp_value <> "" Then
        'ReDim Preserve tmp_vec(0)
        'tmp_vec(0) = tmp_value
        
        For j = 0 To UBound(vec_column, 1)
            ReDim Preserve tmp_vec(j)
            tmp_vec(j) = aim_get_rtd_data("Kronos.RTDServer", "", tmp_view_name, vec_column(j), i)
        Next j
        
        ReDim Preserve vec_data(i)
        vec_data(i) = tmp_vec
        
    Else
        Exit For
    End If
Next i

Dim matrix_data() As Variant


If output_format = output_format_rtd.vec_with_header Then
    aim_get_datas_view_rtd = vec_data
ElseIf output_format = output_format_rtd.vec_without_header Then
    
    ReDim tmp_vec(UBound(vec_data, 1) - 1)
    
    For i = 1 To UBound(vec_data, 1)
        tmp_vec(i - 1) = vec_data(i)
    Next i
    
    aim_get_datas_view_rtd = tmp_vec
    
ElseIf output_format = output_format_rtd.range_with_header Then
    
    ReDim matrix_data(1 To UBound(vec_data, 1) + 1, 1 To UBound(vec_data(0), 1) + 1)
    
    For i = 0 To UBound(vec_data, 1)
        For j = 0 To UBound(vec_data(i), 1)
            matrix_data(i + 1, j + 1) = vec_data(i)(j)
        Next j
    Next i
    
    aim_get_datas_view_rtd = matrix_data
    
ElseIf output_format = output_format_rtd.range_without_header Then
    
    ReDim matrix_data(1 To UBound(vec_data, 1), 1 To UBound(vec_data(0), 1) + 1)
    
    For i = 1 To UBound(vec_data, 1)
        For j = 0 To UBound(vec_data(i), 1)
            matrix_data(i, j + 1) = vec_data(i)(j)
        Next j
    Next i
    
    aim_get_datas_view_rtd = matrix_data
    
End If


End Function


Public Sub aim_rec_mav()

Dim product_asset_type_to_match() As Variant
    product_asset_type_to_match = Array("Equity", "Equity Option", "Index Future", "Index Option")
    
    

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Application.Calculation = xlManual

Dim oReg As New VBScript_RegExp_55.RegExp
Dim matches As VBScript_RegExp_55.MatchCollection
Dim match As VBScript_RegExp_55.match

oReg.Global = True

Dim dicMAV As New Scripting.Dictionary



dim_ticker = 0
dim_asset_type = 1
dim_id = 2
dim_pos = 3

Dim tmp_vec_pos As Variant


'passe en revue la liste des fichiers ouvert pour voir si l'export ne s y trouve pas
Dim tmp_wrbk As Workbook
oReg.Pattern = "grid[\d]+\.xls"
For Each tmp_wrbk In Workbooks
    
    Set matches = oReg.Execute(tmp_wrbk.name)
    
    For Each match In matches
        
        'fichier trouve
        
        'detection des colonnes
        For i = 2 To 30
            If tmp_wrbk.Worksheets(1).Cells(1, i) = "" Then
                c_export_mav_ticker = 1
                Exit For
            Else
                If tmp_wrbk.Worksheets(1).Cells(1, i) = "Asset Type" Then
                    c_export_mav_asset_type = i
                ElseIf tmp_wrbk.Worksheets(1).Cells(1, i) = "BB Unique Id" Then
                    c_export_mav_id = i
                ElseIf tmp_wrbk.Worksheets(1).Cells(1, i) = "Units" Then
                    c_export_mav_position = i
                End If
            End If
        Next i
        
        k = 0
        For i = 5 To 32000
                
            If tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_ticker) <> "" And tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_asset_type) <> "" And tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_id) <> "" And tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_position) <> "" And IsNumeric(tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_position)) Then
                
                For j = 0 To UBound(product_asset_type_to_match, 1)
                    
                    If UCase(product_asset_type_to_match(j)) = UCase(tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_asset_type).Value) Then
                    
                        If dicMAV.Exists(tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_id).Value) Then
                            tmp_vec_pos = dicMAV.Item(tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_id).Value)
                            tmp_vec_pos(dim_pos) = tmp_vec_pos(dim_pos) + tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_position).Value
                            dicMAV.Item(tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_id).Value) = tmp_vec_pos
                        Else
                            
                            tmp_vec_pos = Array(tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_ticker).Value, tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_asset_type).Value, tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_id).Value, tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_position).Value)
                            dicMAV.Add tmp_wrbk.Worksheets(1).Cells(i, c_export_mav_id).Value, tmp_vec_pos
                            k = k + 1
                        End If
                        
                        Exit For
                        
                    End If
                Next j
                
            End If
                    
        Next i
        
        If k > 0 Then
            
            
            
            Dim tmp_asset_type As String
            Dim dicOpen As New Scripting.Dictionary
            
            'remonte tous les produits d open
            open_threshold = 25
            Dim is_last_line As Boolean
            
            k = 0
            For i = 26 To 6000
                is_last_line = True
                
                For j = 0 To open_threshold
                    If Worksheets("Open").Cells(i + j, 1) <> "" Then
                        is_last_line = False
                        Exit For
                    End If
                Next j
                
                If is_last_line = True Then
                    Exit For
                Else
                    Dim tmp_pos_open As Double
                    If Worksheets("Open").Cells(i, 1) <> "" Then
                        
                        If (Worksheets("Open").Cells(i, 6) = "C" Or Worksheets("Open").Cells(i, 6) = "P") Then
                            
                            If Worksheets("Open").Cells(i, 9) <> "" And IsNumeric(Worksheets("Open").Cells(i, 9)) And Worksheets("Open").Cells(i, 9) <> 0 Then
                            
                                If dicOpen.Exists(Worksheets("Open").Cells(i, 2).Value) Then
                                    tmp_vec_pos = dicOpen.Item(Worksheets("Open").Cells(i, 2).Value)
                                    tmp_vec_pos(dim_pos) = tmp_vec_pos(dim_pos) + Worksheets("Open").Cells(i, 9).Value
                                    dicOpen.Item(Worksheets("Open").Cells(i, 2).Value) = tmp_vec_pos
                                Else
                                    tmp_vec_pos = Array(Worksheets("Open").Cells(i, 105).Value, Worksheets("Open").Cells(i, 4).Value & "Option", Worksheets("Open").Cells(i, 2).Value, Worksheets("Open").Cells(i, 9).Value)
                                    dicOpen.Add Worksheets("Open").Cells(i, 2).Value, tmp_vec_pos
                                    k = k + 1
                                End If
                            
                            End If
                            
                            'check underlying
                            If Worksheets("Open").Cells(i, 29) <> "" And IsNumeric(Worksheets("Open").Cells(i, 29)) And Worksheets("Open").Cells(i, 29) <> 0 Then
                                
                                If dicOpen.Exists(Worksheets("Open").Cells(i, 1).Value) Then
                                    tmp_vec_pos = dicOpen.Item(Worksheets("Open").Cells(i, 1).Value)
                                    tmp_vec_pos(dim_pos) = tmp_pos_open + Worksheets("Open").Cells(i, 29).Value
                                    dicOpen.Item(Worksheets("Open").Cells(i, 1).Value) = tmp_vec_pos
                                Else
                                    tmp_vec_pos = Array(Worksheets("Open").Cells(i, 104).Value, "Equity", Worksheets("Open").Cells(i, 1).Value, Worksheets("Open").Cells(i, 29).Value)
                                    dicOpen.Add Worksheets("Open").Cells(i, 1).Value, tmp_vec_pos
                                    k = k + 1
                                End If
                                
                            End If
                        
                        ElseIf Worksheets("Open").Cells(i, 6) = "S" Then
                            
                            If Worksheets("Open").Cells(i, 29) <> "" And IsNumeric(Worksheets("Open").Cells(i, 29)) And Worksheets("Open").Cells(i, 29) <> 0 Then
                                
                                If dicOpen.Exists(Worksheets("Open").Cells(i, 1).Value) Then
                                    tmp_vec_pos = dicOpen.Item(Worksheets("Open").Cells(i, 1).Value)
                                    tmp_vec_pos(dim_pos) = tmp_pos_open + Worksheets("Open").Cells(i, 29).Value
                                    dicOpen.Item(Worksheets("Open").Cells(i, 1).Value) = tmp_vec_pos
                                Else
                                    tmp_vec_pos = Array(Worksheets("Open").Cells(i, 104).Value, "Equity", Worksheets("Open").Cells(i, 1).Value, Worksheets("Open").Cells(i, 29).Value)
                                    dicOpen.Add Worksheets("Open").Cells(i, 1).Value, tmp_vec_pos
                                    k = k + 1
                                End If
                                
                            End If
                        
                        ElseIf Worksheets("Open").Cells(i, 6) = "F" Then
                            
                            If Worksheets("Open").Cells(i, 29) <> "" And IsNumeric(Worksheets("Open").Cells(i, 29)) And Worksheets("Open").Cells(i, 29) <> 0 Then
                                
                                If dicOpen.Exists(Worksheets("Open").Cells(i, 2).Value) Then
                                    tmp_vec_pos = dicOpen.Item(Worksheets("Open").Cells(i, 2).Value)
                                    tmp_vec_pos(dim_pos) = tmp_pos_open + Worksheets("Open").Cells(i, 29).Value
                                    dicOpen.Item(Worksheets("Open").Cells(i, 2).Value) = tmp_vec_pos
                                Else
                                    tmp_vec_pos = Array(Worksheets("Open").Cells(i, 105).Value, "Future", Worksheets("Open").Cells(i, 2).Value, Worksheets("Open").Cells(i, 29).Value)
                                    dicOpen.Add Worksheets("Open").Cells(i, 2).Value, tmp_vec_pos
                                    k = k + 1
                                End If
                                
                            End If
                            
                        End If
                        
                    End If
                End If
                
            Next i
            
            
            Dim vec_not_in_open()
            Dim vec_in_open_wrong_qty() As Variant
            
            If k > 0 Then
                
                'matching
                
                For Each elem_mav In dicMAV
                    
                    mav_id = dicMAV.Item(elem_mav)(dim_id)
                    mav_pos = dicMAV.Item(elem_mav)(dim_pos)
                    
                    If dicMAV.Item(elem_mav)(dim_pos) <> 0 Then
                        
                        If dicOpen.Exists(elem_mav) Then
                            'check pos
                            If dicOpen.Item(elem_mav)(dim_pos) <> dicMAV.Item(elem_mav)(dim_pos) Then
                                Debug.Print "wrong qty " & Chr(9) & Chr(9) & elem_mav & Chr(9) & dicMAV.Item(elem_mav)(dim_ticker) & Chr(9) & "mav: " & dicMAV.Item(elem_mav)(dim_pos) & Chr(9) & "open: " & dicOpen.Item(elem_mav)(dim_pos)
                            End If
                        Else
                            Debug.Print "not in open " & Chr(9) & elem_mav & Chr(9) & dicMAV.Item(elem_mav)(dim_ticker) & Chr(9) & "mav: " & dicMAV.Item(elem_mav)(dim_pos)
                        End If
                        
                    End If
                    
                Next
                
                
            Else
                MsgBox ("no position in open")
            End If
            
            
        Else
            MsgBox ("no position in aim extraction")
            Exit Sub
        End If
        
        Exit For
    Next
    
Next

End Sub


Public Sub aim_reconciliation_mav()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer


Dim path_aim_mav_report As String, name_aim_mav_report As String
path_aim_mav_report = "Q:\front\Kronos\AIM_Report_MATCH.xls"

    name_aim_mav_report = Right(path_aim_mav_report, InStr(StrReverse(path_aim_mav_report), "\") - 1)


Call open_file(path_aim_mav_report, True)

'repere les colonnes
Dim empty_threshold As Integer, is_column_row_empty As Boolean
empty_threshold = 5

Dim c_aim_mav_report_bbuid As Integer, c_aim_mav_report_unit As Integer, c_aim_mav_is_short As Integer


For i = 1 To 50
    
    is_column_row_empty = True
    
    For j = 0 To empty_threshold
        If Workbooks(name_aim_mav_report).Worksheets(1).Cells(1, i + j) <> "" Then
            is_column_row_empty = False
            Exit For
        End If
    Next j
    
    If is_column_row_empty = True Then
        Exit For
    Else
        If Workbooks(name_aim_mav_report).Worksheets(1).Cells(1, i) = "BB Unique Id" Then
            c_aim_mav_report_bbuid = i
        ElseIf Workbooks(name_aim_mav_report).Worksheets(1).Cells(1, i) = "Units" Then
            c_aim_mav_report_unit = i
        ElseIf Workbooks(name_aim_mav_report).Worksheets(1).Cells(1, i) = "Short Y/N?" Then
            c_aim_mav_is_short = i
        End If
    End If
    
Next i

If c_aim_mav_report_bbuid = 0 Or c_aim_mav_report_unit = 0 Or c_aim_mav_is_short = 0 Then
    MsgBox ("necessary column missing")
    Exit Sub
End If


Dim already_find_top_level As Variant
    already_find_top_level = False
Dim already_find_first_line As Variant
    already_find_first_line = False

Dim vec_mav_id_position() As Variant

k = -1
Dim tmp_qty As Double, tmp_buid As String
For i = 1 To 32000
    is_column_row_empty = True
    
    For j = 0 To empty_threshold
        If Workbooks(name_aim_mav_report).Worksheets(1).Cells(i + j, 1) <> "" Then
            is_column_row_empty = False
            Exit For
        End If
    Next j
    
    If is_column_row_empty = True Then
        Exit For
    Else
        
        If already_find_top_level = False Then
            If Workbooks(name_aim_mav_report).Worksheets(1).Cells(i, 1) = "Top Level" Then
                already_find_top_level = i
            End If
        Else
            
            If Workbooks(name_aim_mav_report).Worksheets(1).Cells(i, c_aim_mav_report_bbuid) <> "" And InStr(Workbooks(name_aim_mav_report).Worksheets(1).Cells(i, c_aim_mav_report_bbuid), "cash") = 0 Then
                
                'check des is short
                
                
                
                tmp_buid = Workbooks(name_aim_mav_report).Worksheets(1).Cells(i, c_aim_mav_report_bbuid)
                
                If Workbooks(name_aim_mav_report).Worksheets(1).Cells(i, c_aim_mav_report_unit) = "" Then
                    tmp_qty = 0
                Else
                    tmp_qty = Workbooks(name_aim_mav_report).Worksheets(1).Cells(i, c_aim_mav_report_unit).Value
                End If
                
                
                
                If k = -1 Then
                    k = k + 1
                    ReDim Preserve vec_mav_id_position(k)
                    vec_mav_id_position(k) = Array(tmp_buid, tmp_qty)
                    k = k + 1
                Else
                    
                    'parcours le vecteur pour voir si le produit n existe pas deja
                    
                    For j = 0 To UBound(vec_mav_id_position, 1)
                        If vec_mav_id_position(j)(0) = tmp_buid Then
                            vec_mav_id_position(j)(1) = vec_mav_id_position(j)(1) + tmp_qty
                            Exit For
                        Else
                            If j = UBound(vec_mav_id_position, 1) Then
                                ReDim Preserve vec_mav_id_position(k)
                                vec_mav_id_position(k) = Array(tmp_buid, tmp_qty)
                                k = k + 1
                            End If
                        End If
                    Next j
                    
                End If
            
            End If
        End If
        
    End If
Next i

Workbooks(name_aim_mav_report).Close False



Dim vec_open_product_type_qty As Variant
vec_open_product_type_qty = aim_get_product_from_open()


'rec avec open

Dim vec_qty_difference() As Variant, count_qty_difference As Integer
    count_qty_difference = 0
Dim vec_missing_open() As Variant, count_missing_open As Integer
    count_missing_open = 0
'Dim vec_missing_mav() As Variant





If k = -1 Or IsArray(vec_open_product_type_qty) = False Then
    MsgBox ("no data in mav report or in open")
Else
    
    
    For i = 0 To UBound(vec_mav_id_position, 1)
        For j = 0 To UBound(vec_open_product_type_qty, 1)
            If vec_mav_id_position(i)(0) = vec_open_product_type_qty(j)(0) Then
                
                If vec_mav_id_position(i)(1) <> vec_open_product_type_qty(j)(2) Then
                    'problem qty
                    ReDim Preserve vec_qty_difference(count_qty_difference)
                    vec_qty_difference(count_qty_difference) = Array(vec_mav_id_position(i)(0), vec_mav_id_position(i)(1), vec_open_product_type_qty(j)(2))
                    count_qty_difference = count_qty_difference + 1
                End If
                
                Exit For
            Else
                If j = UBound(vec_open_product_type_qty, 1) Then
                    'entry missing
                    ReDim Preserve vec_missing_open(count_missing_open)
                    vec_missing_open(count_missing_open) = Array(vec_mav_id_position(i)(0), vec_mav_id_position(i)(1))
                    count_missing_open = count_missing_open + 1
                End If
            End If
        Next j
    Next i
    
    
    
    
    
End If


End Sub


Public Function aim_patch_ticker_marketplace(ByVal bbg_ticker As String) As String

Dim i As Integer

Dim mrkt_place As Variant
mrkt_place = Array(Array("GR", "GY"))

aim_patch_ticker_marketplace = UCase(bbg_ticker)

For i = 0 To UBound(mrkt_place, 1)
    If InStr(aim_patch_ticker_marketplace, " " & mrkt_place(i)(0) & " ") <> 0 Then
        aim_patch_ticker_marketplace = Replace(aim_patch_ticker_marketplace, " " & mrkt_place(i)(0) & " ", " " & mrkt_place(i)(1) & " ")
    End If
Next i


End Function



Public Function aim_get_buidt_underlying_derivaitves_multi_account() As Variant

aim_get_buidt_underlying_derivaitves_multi_account = Empty

Dim vec_buidt_underlying() As Variant

Dim i As Integer, j As Integer, k As Integer

k = 0
For i = l_header_internal_db_index + 2 To 250 Step 3
    If Worksheets(sheet_index_db_multi_accounts).Cells(i, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_buidt_underlying(k)
        vec_buidt_underlying(k) = Array(Worksheets(sheet_index_db_multi_accounts).Cells(i, 1).Value, Worksheets(sheet_index_db_multi_accounts).Cells(i, 2).Value, Worksheets(sheet_index_db_multi_accounts).Cells(i, 3).Value, Worksheets(sheet_index_db_multi_accounts).Cells(i, 4).Value)
        k = k + 1
    End If
Next i

If k > 0 Then
    aim_get_buidt_underlying_derivaitves_multi_account = vec_buidt_underlying
End If

End Function



Public Sub aim_update_futures_index_db_multi_accounts()

Debug.Print "aim_update_futures_index_db_multi_accounts: INPUT nothing"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

'repere les colonnes
Dim c_aim_futures_aim_account As Integer, c_aim_futures_aim_buidt As Integer, c_aim_futures_buidt_underlying As Integer


For i = 1 To 100
    If Worksheets(aim_view_futures).Cells(l_header_aim_view, i) = "aim_account" Then
        c_aim_futures_aim_account = i
    ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, i) = "CUSTOM_buidt" Then
        c_aim_futures_aim_buidt = i
    ElseIf Worksheets(aim_view_futures).Cells(l_header_aim_view, i) = "CUSTOM_buidt_underlying" Then
        c_aim_futures_buidt_underlying = i
        Exit For 'attention break de la loop ici
    End If
Next i


'remonte les produits de aim futures
Dim vec_aim_futures() As Variant
k = 0
For i = l_header_aim_view + 1 To 150
    If Worksheets(aim_view_futures).Cells(i, c_aim_futures_aim_buidt).Value = "" Then
        Exit For
    Else
        ReDim Preserve vec_aim_futures(k)
        vec_aim_futures(k) = Array(Worksheets(aim_view_futures).Cells(i, c_aim_futures_aim_buidt).Value, Worksheets(aim_view_futures).Cells(i, c_aim_futures_buidt_underlying).Value, Worksheets(aim_view_futures).Cells(i, c_aim_futures_aim_account).Value)
        k = k + 1
    End If
Next i


If k > 0 Then
    
    Debug.Print "aim_update_futures_index_db_multi_accounts find " & k & " " & " futures in aim_futures view"
    
    'remonte la liste de ceux deja present dans index_db_multi_accounts
    Dim vec_future_multi_accounts_already_in_db As Variant
    Debug.Print "aim_update_futures_index_db_multi_accounts $aim_mount_data_indernal_multi_accounts_db_future"
    vec_future_multi_accounts_already_in_db = aim_mount_data_indernal_multi_accounts_db_future()
    
    
    'repere ceux a ajouter
    Dim vec_future_need_to_be_insert_multi_account_db() As Variant
    k = 0
    If IsArray(vec_future_multi_accounts_already_in_db) Then
        For i = 0 To UBound(vec_aim_futures, 1)
            For j = 0 To UBound(vec_future_multi_accounts_already_in_db, 1)
                If vec_aim_futures(i)(0) = vec_future_multi_accounts_already_in_db(j)(0) Then
                    Exit For
                Else
                    If j = UBound(vec_future_multi_accounts_already_in_db, 1) Then
                        ReDim Preserve vec_future_need_to_be_insert_multi_account_db(k)
                        vec_future_need_to_be_insert_multi_account_db(k) = vec_aim_futures(i)
                        k = k + 1
                    End If
                End If
            Next j
        Next i
    Else
        vec_future_need_to_be_insert_multi_account_db = vec_aim_futures
        k = UBound(vec_aim_futures, 1) + 1
    End If
    
    Dim vec_ticker_futures() As Variant
    If k > 0 Then 'y a t il des futures a inserer
        
        Debug.Print "aim_update_futures_index_db_multi_accounts find " & k & " futures which are not setup in " & sheet_index_db_multi_accounts
        
        For i = 0 To UBound(vec_future_need_to_be_insert_multi_account_db, 1)
            ReDim Preserve vec_ticker_futures(i)
            vec_ticker_futures(i) = "/buid/" & Replace(vec_future_need_to_be_insert_multi_account_db(i)(0), vec_future_need_to_be_insert_multi_account_db(i)(2) & "_", "")
        Next i
        
        'appel bbg pour les details (ticker, expiry date, quotity)
        Dim data_future_bbg As Variant, field_bbg As Variant
        Dim oBBG As New cls_Bloomberg_Sync
        
        field_bbg = Array("PARSEKYABLE_DES_SOURCE", "FUT_CONT_SIZE", "LAST_TRADEABLE_DT")
        
        For i = 0 To UBound(field_bbg, 1)
            If field_bbg(i) = "PARSEKYABLE_DES_SOURCE" Then
                dim_bbg_TICKER = i
            ElseIf field_bbg(i) = "FUT_CONT_SIZE" Then
                dim_bbg_quotity = i
            ElseIf field_bbg(i) = "LAST_TRADEABLE_DT" Then
                dim_bbg_expiry_date = i
            End If
        Next i
        
        data_future_bbg = oBBG.bdp(vec_ticker_futures, field_bbg, output_format.of_vec_without_header)
        
        
        
        'mise en place des futures
        Dim find_expired_fut As Boolean
        For i = 0 To UBound(vec_future_need_to_be_insert_multi_account_db, 1)
            
            For j = l_header_internal_db_index + 2 To 150 Step 3
                If Worksheets(sheet_index_db_multi_accounts).Cells(j, 1) = "" Then
                    Exit For
                Else
                    If Worksheets(sheet_index_db_multi_accounts).Cells(j, 1) = vec_future_need_to_be_insert_multi_account_db(i)(1) Then
                        
                        'y a - t - il un emplacement de libre ?
find_a_place_for_future:

                        find_expired_fut = False
                        
                        If Worksheets(sheet_index_db_multi_accounts).Cells(j, 31) = "" Then
                            'encore aucun fut de setupe
                            Debug.Print "aim_update_futures_index_db_multi_accounts "
                            
                            If data_future_bbg(i)(dim_bbg_expiry_date) >= Date - 3 Then
                                
                                Worksheets(sheet_index_db_multi_accounts).Cells(j - 1, 31) = "CUSTOM_buidt"
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 31) = vec_future_need_to_be_insert_multi_account_db(i)(0)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 33) = data_future_bbg(i)(dim_bbg_expiry_date)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 34) = data_future_bbg(i)(dim_bbg_TICKER)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 47) = data_future_bbg(i)(dim_bbg_quotity)
                            End If
                        ElseIf Worksheets(sheet_index_db_multi_accounts).Cells(j, 32) = "" Then
                            
                            'faut-il inverser l'ordre des fut ?
                            If data_future_bbg(i)(dim_bbg_expiry_date) < Worksheets(sheet_index_db_multi_accounts).Cells(j, 33) Then
                                
                                If Worksheets(sheet_index_db_multi_accounts).Cells(j, 33) >= Date - 3 Then
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j - 1, 32) = "CUSTOM_buidt"
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j, 32) = Worksheets(sheet_index_db_multi_accounts).Cells(j, 31)
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 33) = Worksheets(sheet_index_db_multi_accounts).Cells(j, 33)
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 34) = Worksheets(sheet_index_db_multi_accounts).Cells(j, 34)
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 47) = Worksheets(sheet_index_db_multi_accounts).Cells(j, 47)
                                End If
                                
                                
                                Worksheets(sheet_index_db_multi_accounts).Cells(j - 1, 31) = "CUSTOM_buidt"
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 31) = vec_future_need_to_be_insert_multi_account_db(i)(0)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 33) = data_future_bbg(i)(dim_bbg_expiry_date)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 34) = data_future_bbg(i)(dim_bbg_TICKER)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 47) = data_future_bbg(i)(dim_bbg_quotity)
                            Else
                                
                                Worksheets(sheet_index_db_multi_accounts).Cells(j - 1, 32) = "CUSTOM_buidt"
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 32) = vec_future_need_to_be_insert_multi_account_db(i)(0)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 33) = data_future_bbg(i)(dim_bbg_expiry_date)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 34) = data_future_bbg(i)(dim_bbg_TICKER)
                                Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 47) = data_future_bbg(i)(dim_bbg_quotity)
                                
                            End If
                            
                        Else
                            'y - a - t - il un fut echu ?
                            If Worksheets(sheet_index_db_multi_accounts).Cells(j, 33) < Date - 3 Then
                                
                                If Worksheets(sheet_index_db_multi_accounts).Cells(j, 32) <> "" Then
                                    'remonter le 2e fut en 1ere position
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j - 1, 31) = "CUSTOM_buidt"
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j, 31) = Worksheets(sheet_index_db_multi_accounts).Cells(j, 32)
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j, 33) = Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 33)
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j, 34) = Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 34)
                                    Worksheets(sheet_index_db_multi_accounts).Cells(j, 47) = Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 47)
                                        
                                        Worksheets(sheet_index_db_multi_accounts).Cells(j - 1, 32) = ""
                                        Worksheets(sheet_index_db_multi_accounts).Cells(j, 32) = ""
                                        Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 33) = ""
                                        Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 34) = ""
                                        Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 47) = ""
                                    
                                    
                                    find_expired_fut = True
                                    
                                End If
                                
                            End If
                            
                            
                            If Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 33) < Date - 3 Then
                                Worksheets(sheet_index_db_multi_accounts).Cells(j - 1, 32) = ""
                                Worksheets(sheet_index_db_multi_accounts).Cells(j, 32) = ""
                                Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 33) = ""
                                Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 34) = ""
                                Worksheets(sheet_index_db_multi_accounts).Cells(j + 1, 47) = ""
                                
                                find_expired_fut = True
                                
                            End If
                            
                            If find_expired_fut = True Then
                                GoTo find_a_place_for_future
                            Else
                                MsgBox ("no place anymore for a new future in " & Worksheets(sheet_index_db_multi_accounts).Cells(j, 4))
                            End If
                            
                        End If
                        
                    End If
                End If
            Next j
        Next i
    Else
        Debug.Print "aim_update_futures_index_db_multi_accounts everything already in " & sheet_index_db_multi_accounts
    End If
        

End If

End Sub


Public Sub aim_prepare_stats_derivatives_multi_account()

Debug.Print "aim_prepare_stats_derivatives_multi_account: INPUT nothing"

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

'une ligne par combi de account_buid_underlying

'list d apres open toutes les combinaisons necessaire
Dim open_threshold As Integer
    open_threshold = 25
    
Dim is_last_line As Boolean

Dim vec_open_buidt_underyling() As Variant
k = 0
For i = l_header_internal_open + 1 To 3500
    
    is_last_line = True
    
    For j = 0 To open_threshold
        If Worksheets("Open").Cells(i + j, 1) <> "" Then
            is_last_line = False
            Exit For
        End If
    Next
    
    If is_last_line = True Then
        Exit For
    Else
        If Worksheets("Open").Cells(i, 1) <> "" Then
            
            'check si derivative sur index
            If Worksheets("Open").Cells(i, 4) = "I" Then
                
                If k = 0 Then
                    ReDim Preserve vec_open_buidt_underyling(k)
                    vec_open_buidt_underyling(k) = Array(Worksheets("Open").Cells(i, 94).Value, Worksheets("Open").Cells(i, 92).Value, Worksheets("Open").Cells(i, 1).Value, Worksheets("Open").Cells(i, 7).Value)
                    k = k + 1
                Else
                    For m = 0 To UBound(vec_open_buidt_underyling, 1)
                        If vec_open_buidt_underyling(m)(0) = Worksheets("Open").Cells(i, 94) Then
                            Exit For
                        Else
                            If m = UBound(vec_open_buidt_underyling, 1) Then
                                ReDim Preserve vec_open_buidt_underyling(k)
                                vec_open_buidt_underyling(k) = Array(Worksheets("Open").Cells(i, 94).Value, Worksheets("Open").Cells(i, 92).Value, Worksheets("Open").Cells(i, 1).Value, Worksheets("Open").Cells(i, 7).Value)
                                k = k + 1
                            End If
                        End If
                    Next m
                End If
                
            End If
            
        End If
    End If
    
    
Next i


'remonte les buidt_underying de la feuille de stat deja present
If k > 0 Then
    
    Debug.Print "aim_prepare_stats_derivatives_multi_account: find " & k & " buidt_underlying_id"
    
    Dim vec_index_db_multi_accounts As Variant
    Debug.Print "aim_prepare_stats_derivatives_multi_account $aim_get_buidt_underlying_derivaitves_multi_account"
    vec_index_db_multi_accounts = aim_get_buidt_underlying_derivaitves_multi_account()
    
    
    'match les manquants
    Dim vec_need_to_be_open_in_index_multi_accounts() As Variant
    k = 0
    
    If IsArray(vec_index_db_multi_accounts) Then
        
        For i = 0 To UBound(vec_open_buidt_underyling, 1)
            For j = 0 To UBound(vec_index_db_multi_accounts, 1)
                If vec_open_buidt_underyling(i)(0) = vec_index_db_multi_accounts(j)(0) Then
                    Exit For
                Else
                    If j = UBound(vec_index_db_multi_accounts, 1) Then
                        ReDim Preserve vec_need_to_be_open_in_index_multi_accounts(k)
                        vec_need_to_be_open_in_index_multi_accounts(k) = vec_open_buidt_underyling(i)
                        k = k + 1
                    End If
                End If
            Next j
        Next i
        
    Else
        vec_need_to_be_open_in_index_multi_accounts = vec_open_buidt_underyling
        k = UBound(vec_open_buidt_underyling, 1) + 1
    End If
    
    
    
    
End If




'insertions des lignes dans la sheet
If k > 0 Then
    
    Debug.Print "aim_prepare_stats_derivatives_multi_account find " & k & " buidt_underyling missing in " & sheet_index_db_multi_accounts
    
    'appel bbg pour connaitre les crncy afin de determiner les currency code
    Dim vec_ticker_index() As Variant
    For i = 0 To UBound(vec_need_to_be_open_in_index_multi_accounts, 1) 'pas forcement efficient car plusieurs fois le meme possible
        ReDim Preserve vec_ticker_index(i)
        vec_ticker_index(i) = "/buid/" & vec_need_to_be_open_in_index_multi_accounts(i)(2)
    Next i
    
    Dim oBBG As New cls_Bloomberg_Sync
    Dim data_bbg As Variant, bbg_field As Variant
    bbg_field = Array("CRNCY")
    data_bbg = oBBG.bdp(vec_ticker_index, bbg_field, output_format.of_vec_without_header)
    
        
    
    Dim vec_currency() As Variant
    k = 0
    For i = 14 To 32
        If Worksheets("Parametres").Cells(i, 1) <> "" Then
            ReDim Preserve vec_currency(k)
            vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, Worksheets("Parametres").Cells(i, 5).Value, i)
            k = k + 1
        End If
    Next i
    
    
    'repere last line in index_db_multi_account
    Dim l_index_db_multi_accounts As Integer
    
    For i = l_header_internal_db_index + 2 To 350 Step 3
        If Worksheets(sheet_index_db_multi_accounts).Cells(i, 1) = "" Then
            l_index_db_multi_accounts = i - 2
            Exit For
        Else
        End If
    Next i
    
    
    For i = 0 To UBound(vec_need_to_be_open_in_index_multi_accounts, 1)
        
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 1, 1) = "CUSTOM_buidt_underlying"
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 1) = vec_need_to_be_open_in_index_multi_accounts(i)(0)
        
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 1, 2) = "aim_account"
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 2) = vec_need_to_be_open_in_index_multi_accounts(i)(1)
        
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 1, 3) = "Identifier"
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 3) = vec_need_to_be_open_in_index_multi_accounts(i)(2)
        
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 1, 4) = "Index_name"
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 4) = vec_need_to_be_open_in_index_multi_accounts(i)(3)
        

        'mise en place des formules, en utilisant la ligne de formule d'index db
        For j = 1 To 250
            If Worksheets("Index_Database").Cells(2, j) = "F" Then
                Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, j).Value = Replace("=" & Worksheets("Index_Database").Cells(5, j), ";", ",")
            End If
            
            
            If Worksheets("Index_Database").Cells(3, j) = "F" Then
                Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 3, j).Value = Replace("=" & Worksheets("Index_Database").Cells(6, j), ";", ",")
            End If
        Next j
        
        
        'remplace formules pour fonctionner par trader sur les formules lies aux futures
        
        For j = 0 To UBound(vec_currency, 1)
            If UCase(data_bbg(i)(0)) = UCase(vec_currency(j)(0)) Then
                
                'daily premium
                Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 10).Value = "=(RC22-RC152)*Parametres!R" & vec_currency(j)(2) & "C6*1/1000"
                
                'daily startcurreval
                Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 11).Value = "=(RC40-RC153)*Parametres!R" & vec_currency(j)(2) & "C6*1/1000"
                
                'result executed
                Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 14).Value = "=(RC22-RC23)*Parametres!R" & vec_currency(j)(2) & "C6+RC149"
                
                ' executed - YTD en close
                Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 28).Value = "=(DSUM(AIM_EOD_DT,""CUSTOM_buidt_ytd_pnl_local_net"",R[-1]C1:RC1)-RC30)*Parametres!R" & vec_currency(j)(2) & "C6"
                
                ' daily
                Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 29).Value = "=(-RC42+RC46-RC41+IF(RC31<>"""",(((RC35-RC43)*RC37*RC47)+RC39+RC40)+((RC35-RC44)*RC45*RC47),0)+IF(RC32<>"""",(((R[1]C35-R[1]C43)*R[1]C37*R[1]C47)+R[1]C39+R[1]C40)+((R[1]C35-R[1]C44)*R[1]C45*R[1]C47),0))*Parametres!R" & vec_currency(j)(2) & "C6"
                
                Exit For
            End If
        Next j
        
        
        'reversal
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 16).Value = "0"
        
        ' premium reversal
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 23).Value = "0"
        
        ' fut reversal
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 30).Value = "0"
        
        ' net position
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 37).Value = "=IF(RC31<>"""",DSUM(AIM_Futures_DT,""CUSTOM_buidt_current_qty"",R[-1]C31:RC31),0)"
            Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 3, 37).Value = "=IF(R[-1]C32<>"""",DSUM(AIM_Futures_DT,""CUSTOM_buidt_current_qty"",R[-2]C[-5]:R[-1]C[-5]),0)"
        
        ' close position
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 38).Value = "=IF(RC31<>"""",DSUM(AIM_Futures_DT,""CUSTOM_buidt_close_qty"",R[-1]C31:RC31),0)"
            Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 3, 38).Value = "=IF(R[-1]C32<>"""",DSUM(AIM_Futures_DT,""CUSTOM_buidt_close_qty"",R[-2]C32:R[-1]C32),0)"
        
        ' net trading cash flow
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 39).Value = "=IF(RC31<>"""",DSUM(AIM_Futures_DT,""CUSTOM_buidt_net_cash_local_with_comm"",R[-1]C31:RC31),0)"
            Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 3, 39).Value = "=IF(R[-1]C32<>"""",DSUM(AIM_Futures_DT,""CUSTOM_buidt_net_cash_local_with_comm"",R[-2]C32:R[-1]C32),0)"
        
        
        'avg rate manual in trades
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 44).Value = 0
            Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 3, 44).Value = 0
            
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 45).Value = 0
            Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 3, 45).Value = 0
        
        
        'intraday comm
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 48).Value = "=DSUM(AIM_Futures_DT,""CUSTOM_intraday_commission_local"",R[-1]C1:RC1)+DSUM(AIM_Options_DT,""CUSTOM_intraday_commission_local"",R[-1]C1:RC1)"
        
        'ytd comm
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 49).Value = "=DSUM(AIM_EOD_DT,""CUSTOM_tra_cost_total"",R[-1]C1:RC1)"
        
        
        'currency
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 107).Value = vec_currency(j)(1)
        
        'account_curreny
        Worksheets(sheet_index_db_multi_accounts).Cells(l_index_db_multi_accounts + 2, 176).Value = "=RC2&""_""&RC107"
        
        
        l_index_db_multi_accounts = l_index_db_multi_accounts + 3
        
    Next i
Else
    
    Debug.Print "aim_prepare_stats_derivatives_multi_account everything already in " & sheet_index_db_multi_accounts
    
End If

Debug.Print "aim_prepare_stats_derivatives_multi_account $aim_update_futures_index_db_multi_accounts"
Call aim_update_futures_index_db_multi_accounts 'insertion / update  des futs / traders

End Sub


Public Sub aim_check_option_underyling()

Dim i As Integer, j As Integer, k As Integer

Application.Calculation = xlCalculationManual

Dim open_threshold As Integer, is_last_line As Boolean
    open_threshold = 25

For i = 26 To 3200
    
    is_last_line = True
    
    For j = 0 To open_threshold
        
        If Worksheets("Open").Cells(i + j, 1) <> "" Then
            is_last_line = False
            Exit For
        End If
        
    Next j
    
    If is_last_line = True Then
        Exit For
    End If
    
    If Worksheets("Open").Cells(i, 1) <> "" And Worksheets("Open").Cells(i, 3) = "O" Then
        If InStr(Worksheets("Open").Cells(i, 97), Worksheets("Open").Cells(i, 1)) = 0 Then
            debug_test = "bug underlying"
        End If
    End If
    
Next i

End Sub


Public Sub aim_switch_view_open_all()

Call aim_account_switch_account_open("")

End Sub


Public Sub aim_switch_view_open_lennart()

Call aim_account_switch_account_open("C6414GSL")

End Sub


Public Sub aim_switch_view_open_julien()

Call aim_account_switch_account_open("C6414GSJ")

End Sub
