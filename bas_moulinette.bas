Attribute VB_Name = "bas_Moulinette"
Public Const db_moulinette As String = "db_redi.sqlt3"

    Public Const t_moulinette_order_xls As String = "t_moulinette_order_xls"
        Public Const f_moulinette_order_xls_id As String = "f_moulinette_order_xls_id"
        Public Const f_moulinette_order_xls_group_id As String = "f_moulinette_order_xls_group_id"
        Public Const f_moulinette_order_xls_ticker As String = "f_moulinette_order_xls_ticker"
        Public Const f_moulinette_order_xls_symbol_redi As String = "f_moulinette_order_xls_symbol_redi"
        Public Const f_moulinette_order_xls_datetime As String = "f_moulinette_order_xls_datetime"
        Public Const f_moulinette_order_xls_side As String = "f_moulinette_order_xls_side"
        Public Const f_moulinette_order_xls_order_qty As String = "f_moulinette_order_xls_order_qty"
        Public Const f_moulinette_order_xls_order_price As String = "f_moulinette_order_xls_order_price"
        Public Const f_moulinette_order_xls_json_tag As String = "f_moulinette_order_xls_json_tag"
    
    Public Const t_moulinette_order_redi As String = "t_moulinette_order_redi"
        Public Const f_moulinette_order_redi_RefNum As String = "f_moulinette_order_redi_RefNum" 'TEXT
        Public Const f_moulinette_order_redi_OrderRefKey As String = "f_moulinette_order_redi_OrderRefKey" 'TEXT
        Public Const f_moulinette_order_redi_Desccription As String = "f_moulinette_order_redi_Desccription"
        Public Const f_moulinette_order_redi_BranchSequence As String = "f_moulinette_order_redi_BranchSequence" 'TEXT
        Public Const f_moulinette_order_redi_datetime As String = "f_moulinette_order_redi_datetime"
        Public Const f_moulinette_order_redi_side As String = "f_moulinette_order_redi_side"
        Public Const f_moulinette_order_redi_symbol As String = "f_moulinette_order_redi_symbol"
        Public Const f_moulinette_order_redi_OrderQty As String = "f_moulinette_order_redi_OrderQty"
        Public Const f_moulinette_order_redi_OrderPrice As String = "f_moulinette_order_redi_OrderPrice"
        Public Const f_moulinette_order_redi_ExecQty As String = "f_moulinette_order_redi_ExecQty"
        Public Const f_moulinette_order_redi_ExecPrice As String = "f_moulinette_order_redi_ExecPrice"
        Public Const f_moulinette_order_redi_PriceType As String = "f_moulinette_order_redi_PriceType"
        Public Const f_moulinette_order_redi_Status As String = "f_moulinette_order_redi_Status"
        Public Const f_moulinette_order_redi_UserID As String = "f_moulinette_order_redi_UserID"
    
    Public Const v_moulinette_aggreg_order_redi As String = "v_moulinette_aggreg_order_redi"
        Public Const f_moulinette_aggreg_order_redi_BranchSequence As String = "BranchSequence"
        Public Const f_moulinette_aggreg_order_redi_symbol As String = "symbol"
        Public Const f_moulinette_aggreg_order_redi_first_datetime As String = "first_datetime"
        Public Const f_moulinette_aggreg_order_redi_order_type As String = "order_type"
        Public Const f_moulinette_aggreg_order_redi_last_status As String = "last_status"
        Public Const f_moulinette_aggreg_order_redi_last_exec_datetime As String = "last_exec_datetime"
        Public Const f_moulinette_aggreg_order_redi_OrderQty As String = "OrderQty"
        Public Const f_moulinette_aggreg_order_redi_OrderPrice As String = "OrderPrice"
        Public Const f_moulinette_aggreg_order_redi_ExecQty As String = "ExecQty"
        Public Const f_moulinette_aggreg_order_redi_NTCF As String = "NTCF"
        Public Const f_moulinette_aggreg_order_redi_AvgExecPrice As String = "AvgExecPrice"
        Public Const f_moulinette_aggreg_order_redi_Commissions As String = "Commissions"
        
        Public Const f_moulinette_stat_ticker_OrderQty As String = "stat_ticker_OrderQty"
        Public Const f_moulinette_stat_ticker_ExecQty As String = "stat_ticker_ExecQty"
        Public Const f_moulinette_stat_ticker_NTCF As String = "stat_ticker_NTCF"
        Public Const f_moulinette_stat_ticker_AVGOrderPrice As String = "stat_ticker_AVGOrderPrice"
        Public Const f_moulinette_stat_ticker_AVGExecPrice As String = "stat_ticker_AVGExecPrice"
        
        
    Public Const t_moulinette_bridge_redi As String = "t_moulinette_bridge_redi"
        Public Const f_moulinette_bridge_redi_internal_id As String = "f_moulinette_bridge_redi_internal_id"
        Public Const f_moulinette_bridge_redi_BranchSequence As String = "f_moulinette_bridge_redi_BranchSequence"
        Public Const f_moulinette_bridge_redi_SymbolRedi As String = "f_moulinette_bridge_redi_SymbolRedi"
        Public Const f_moulinette_bridge_redi_SymbolXLS As String = "f_moulinette_bridge_redi_SymbolXLS"
        
    
    Public Const t_moulinette_static As String = "t_moulinette_static"
        Public Const f_moulinette_static_ticker As String = "f_moulinette_static_ticker"
        Public Const f_moulinette_static_crncy As String = "f_moulinette_static_crncy" 'calcul pnl base
        Public Const f_moulinette_static_fut_cont_size As String = "f_moulinette_static_fut_cont_size" 'calc pnl
        
    
    Public Const t_moulinette_helper As String = "t_moulinette_helper"
        Public Const f_moulinette_helper_text1 As String = "f_moulinette_helper_text1"
        Public Const f_moulinette_helper_text2 As String = "f_moulinette_helper_text2"
        Public Const f_moulinette_helper_text3 As String = "f_moulinette_helper_text3"
        Public Const f_moulinette_helper_numeric1 As String = "f_moulinette_helper_numeric1"
        Public Const f_moulinette_helper_numeric2 As String = "f_moulinette_helper_numeric2"
        Public Const f_moulinette_helper_numeric3 As String = "f_moulinette_helper_numeric3"
    
    Public Const sheet_offline As String = "FORMAT2"
        Public Const c_offline_order_xls_id As Integer = 157
        Public Const c_offline_order_xls_group_id As Integer = 158
        Public Const c_offline_order_xls_ticker As Integer = 159
        Public Const c_offline_order_xls_symbol_redi As Integer = 160
        Public Const c_offline_order_xls_datetime As Integer = 161
        Public Const c_offline_order_xls_side As Integer = 162
        Public Const c_offline_order_xls_order_qty As Integer = 163
        Public Const c_offline_order_xls_order_price As Integer = 164
        Public Const c_offline_order_xls_json_tag As Integer = 165
    
    Public Const sheet_report As String = "Moulinette"
        Public Const l_report_summary_long As Integer = 3
        Public Const l_report_summary_short As Integer = 4
        Public Const l_report_summary_net As Integer = 5
        
        
        Public Const l_report_header As Integer = 10
        Public Const c_report_group_xls_id As Integer = 1
        Public Const c_report_trade_xls_id As Integer = 2
        Public Const c_report_trade_redi_id As Integer = 3
        Public Const c_report_datetime As Integer = 4
        Public Const c_report_ticker As Integer = 5
        Public Const c_report_order_type As Integer = 6
        Public Const c_report_order_tag As Integer = 7
        Public Const c_report_order_status As Integer = 8
        Public Const c_report_order_qty As Integer = 9
        Public Const c_report_order_price As Integer = 10
        Public Const c_report_exec_qty As Integer = 11
        Public Const c_report_exec_avg_price As Integer = 12
        Public Const c_report_bbg_px_last As Integer = 13
        Public Const c_report_bbg_px_high As Integer = 14
        Public Const c_report_bbg_px_low As Integer = 15
        Public Const c_report_nominal_open_usd As Integer = 16
        Public Const c_report_nominal_exec_usd As Integer = 17
        Public Const c_report_pnl_local As Integer = 18
        Public Const c_report_pnl_base As Integer = 19
        Public Const c_report_pnl_with_comm As Integer = 20

Public Const prefix_emsx_trades As String = "EMSX_"
        
        


Public Function moulinette_get_db_complete_path()

moulinette_get_db_complete_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\"))) & db_moulinette

End Function






Private Sub moulinette_manipulate_db()

Dim exec_query As Variant

Dim sql_query As String
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DELETE FROM " & t_moulinette_order_xls)
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DELETE FROM " & t_moulinette_order_xls & " WHERE " & f_moulinette_order_xls_id & "=""2564""")
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DELETE FROM " & t_moulinette_order_redi)
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DROP TABLE " & t_moulinette_order_redi)
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DROP TABLE " & t_moulinette_bridge_redi)
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DROP TABLE " & t_moulinette_static)
table_structure = sqlite3_get_table_structure(moulinette_get_db_complete_path, t_moulinette_static)
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DELETE FROM " & t_moulinette_static)
'exec_query = sqlite3_query(moulinette_get_db_complete_path, "DROP VIEW IF EXISTS " & v_moulinette_aggreg_order_redi)

extract_static = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_static)
extract_xls_order = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_xls)
    extract_xls_order_specfic = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_xls & " WHERE " & f_moulinette_order_xls_ticker & "=""EWZ US EQUITY""")
        'date_tmp = FromJulianDay(CDbl(extract_xls_order_specfic(1)(4)))
extract_redi_order = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_redi)
extract_bridge = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_bridge_redi)
    
    
    
    extract_bridge_order = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_bridge_redi & " WHERE " & f_moulinette_bridge_redi_BranchSequence & " LIKE ""EMSX%""")


    sql_query = "SELECT * FROM " & t_moulinette_order_redi & " WHERE " & f_moulinette_order_redi_BranchSequence & "=""IWG0074"""
    extract_red_order_specific = sqlite3_query(moulinette_get_db_complete_path, sql_query)
    
    

'test aggreg query
'sql_query = "SELECT " & f_moulinette_order_redi_BranchSequence & " AS BranchSequence, " & f_moulinette_order_redi_symbol & " AS symbol, " & " MIN(" & f_moulinette_order_redi_datetime & ") AS first_datetime, AVG(" & f_moulinette_order_redi_OrderQty & ") AS OrderQty, AVG(" & f_moulinette_order_redi_OrderPrice & ") AS OrderPrice, SUM(" & f_moulinette_order_redi_ExecQty & ") AS ExecQty, SUM(" & f_moulinette_order_redi_ExecQty & "*" & f_moulinette_order_redi_ExecPrice & ") AS NTCF, SUM(" & f_moulinette_order_redi_ExecQty & "*" & f_moulinette_order_redi_ExecPrice & ")/SUM(" & f_moulinette_order_redi_ExecQty & ") as AvgExecPrice "
'    sql_query = sql_query & " FROM " & t_moulinette_order_redi
'    sql_query = sql_query & " GROUP BY " & f_moulinette_order_redi_BranchSequence
    

extract_view = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi)



    extract_view_specific = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi & " WHERE BranchSequence LIKE ""EMSX%""")
    'date_tmp = FromJulianDay(CDbl(extract_view_specific(1)(3)))
End Sub


Private Sub test_moulinette_init_db()

Call moulinette_init_db(False)

End Sub


Public Sub moulinette_init_db(Optional ByVal bypass_wash_redi_order As Boolean = False)

Dim sql_query As String
Dim exec_query As Variant

Dim create_db_status As Variant, create_tbl_status As Variant
    create_db_status = sqlite3_create_db(moulinette_get_db_complete_path)

Dim create_table_query As String

If sqlite3_check_if_table_already_exist(moulinette_get_db_complete_path, t_moulinette_order_xls) = False Then
    create_table_query = sqlite3_get_query_create_table(t_moulinette_order_xls, Array(Array(f_moulinette_order_xls_id, "TEXT", ""), Array(f_moulinette_order_xls_group_id, "NUMERIC", ""), Array(f_moulinette_order_xls_ticker, "TEXT", ""), Array(f_moulinette_order_xls_symbol_redi, "TEXT", ""), Array(f_moulinette_order_xls_datetime, "NUMERIC", ""), Array(f_moulinette_order_xls_side, "TEXT", ""), Array(f_moulinette_order_xls_order_qty, "INTEGER", ""), Array(f_moulinette_order_xls_order_price, "REAL", ""), Array(f_moulinette_order_xls_json_tag, "TEXT", "")), Array(Array(f_moulinette_order_xls_id, "ASC")))
    create_tbl_status = sqlite3_create_tables(moulinette_get_db_complete_path, Array(create_table_query))
End If

If sqlite3_check_if_table_already_exist(moulinette_get_db_complete_path, t_moulinette_order_redi) = False Then
    create_table_query = sqlite3_get_query_create_table(t_moulinette_order_redi, Array(Array(f_moulinette_order_redi_RefNum, "TEXT", ""), Array(f_moulinette_order_redi_OrderRefKey, "TEXT", ""), Array(f_moulinette_order_redi_Desccription, "TEXT", ""), Array(f_moulinette_order_redi_BranchSequence, "TEXT", ""), Array(f_moulinette_order_redi_datetime, "NUMERIC", ""), Array(f_moulinette_order_redi_side, "TEXT", ""), Array(f_moulinette_order_redi_symbol, "TEXT", ""), Array(f_moulinette_order_redi_OrderQty, "NUMERIC", ""), Array(f_moulinette_order_redi_OrderPrice, "REAL", ""), Array(f_moulinette_order_redi_ExecQty, "NUMERIC", ""), Array(f_moulinette_order_redi_ExecPrice, "REAL", ""), Array(f_moulinette_order_redi_PriceType, "TEXT", ""), Array(f_moulinette_order_redi_Status, "TEXT", ""), Array(f_moulinette_order_redi_UserID, "TEXT", "")), Array(Array(f_moulinette_order_redi_RefNum, "ASC")))
    create_tbl_status = sqlite3_create_tables(moulinette_get_db_complete_path, Array(create_table_query))
End If


sql_query = "CREATE VIEW IF NOT EXISTS " & v_moulinette_aggreg_order_redi & " AS "
    
    sql_query = sql_query & "SELECT " & f_moulinette_order_redi_BranchSequence & " AS " & f_moulinette_aggreg_order_redi_BranchSequence & ", " & f_moulinette_order_redi_symbol & " AS " & f_moulinette_aggreg_order_redi_symbol & ", " & f_moulinette_order_redi_PriceType & " AS " & f_moulinette_aggreg_order_redi_order_type & ", "
        sql_query = sql_query & " MIN(" & f_moulinette_order_redi_datetime & ") AS " & f_moulinette_aggreg_order_redi_first_datetime & ", AVG(" & f_moulinette_order_redi_OrderQty & ") AS " & f_moulinette_aggreg_order_redi_OrderQty & ", AVG(" & f_moulinette_order_redi_OrderPrice & ") AS " & f_moulinette_aggreg_order_redi_OrderPrice & ", SUM(" & f_moulinette_order_redi_ExecQty & ") AS " & f_moulinette_aggreg_order_redi_ExecQty & ", SUM(" & f_moulinette_order_redi_ExecQty & "*" & f_moulinette_order_redi_ExecPrice & ") AS " & f_moulinette_aggreg_order_redi_NTCF & ", SUM(" & f_moulinette_order_redi_ExecQty & "*" & f_moulinette_order_redi_ExecPrice & ")/SUM(" & f_moulinette_order_redi_ExecQty & ") AS " & f_moulinette_aggreg_order_redi_AvgExecPrice & ", 0.0004*SUM(ABS(" & f_moulinette_order_redi_ExecQty & ")*" & f_moulinette_order_redi_ExecPrice & ") AS " & f_moulinette_aggreg_order_redi_Commissions

            sql_query = sql_query & ", (SELECT " & f_moulinette_order_redi_Status & " FROM " & t_moulinette_order_redi & " t2 WHERE t2." & f_moulinette_order_redi_BranchSequence & "=t." & f_moulinette_order_redi_BranchSequence & " AND t2." & f_moulinette_order_redi_symbol & "=t." & f_moulinette_order_redi_symbol & " ORDER BY t2." & f_moulinette_order_redi_datetime & " DESC LIMIT 1) AS " & f_moulinette_aggreg_order_redi_last_status
            sql_query = sql_query & ", (SELECT " & f_moulinette_order_redi_datetime & " FROM " & t_moulinette_order_redi & " t2 WHERE t2." & f_moulinette_order_redi_BranchSequence & "=t." & f_moulinette_order_redi_BranchSequence & " AND t2." & f_moulinette_order_redi_symbol & "=t." & f_moulinette_order_redi_symbol & " ORDER BY t2." & f_moulinette_order_redi_datetime & " DESC LIMIT 1) AS " & f_moulinette_aggreg_order_redi_last_exec_datetime
            
        sql_query = sql_query & " FROM " & t_moulinette_order_redi
        sql_query = sql_query & " GROUP BY " & f_moulinette_order_redi_BranchSequence & ", " & f_moulinette_order_redi_symbol
exec_query = sqlite3_query(moulinette_get_db_complete_path, sql_query)

    'debug_test = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi)


If sqlite3_check_if_table_already_exist(moulinette_get_db_complete_path, t_moulinette_bridge_redi) = False Then
    create_table_query = sqlite3_get_query_create_table(t_moulinette_bridge_redi, Array(Array(f_moulinette_bridge_redi_internal_id, "TEXT", ""), Array(f_moulinette_bridge_redi_BranchSequence, "TEXT", ""), Array(f_moulinette_bridge_redi_SymbolRedi, "TEXT", ""), Array(f_moulinette_bridge_redi_SymbolXLS, "TEXT", "")), Array(Array(f_moulinette_bridge_redi_internal_id, "ASC"), Array(f_moulinette_bridge_redi_BranchSequence, "ASC"), Array(f_moulinette_bridge_redi_SymbolRedi, "ASC")))
    create_tbl_status = sqlite3_create_tables(moulinette_get_db_complete_path, Array(create_table_query))
End If


If sqlite3_check_if_table_already_exist(moulinette_get_db_complete_path, t_moulinette_static) = False Then
    create_table_query = sqlite3_get_query_create_table(t_moulinette_static, Array(Array(f_moulinette_static_ticker, "TEXT", ""), Array(f_moulinette_static_crncy, "TEXT", ""), Array(f_moulinette_static_fut_cont_size, "NUMERIC", "")), Array(Array(f_moulinette_static_ticker, "ASC")))
    create_tbl_status = sqlite3_create_tables(moulinette_get_db_complete_path, Array(create_table_query))
End If


If sqlite3_check_if_table_already_exist(moulinette_get_db_complete_path, t_moulinette_helper) = False Then
    create_table_query = sqlite3_get_query_create_table(t_moulinette_helper, Array(Array(f_moulinette_helper_text1, "TEXT", ""), Array(f_moulinette_helper_text2, "TEXT", ""), Array(f_moulinette_helper_text3, "TEXT", ""), Array(f_moulinette_helper_numeric1, "NUMERIC", ""), Array(f_moulinette_helper_numeric2, "NUMERIC", ""), Array(f_moulinette_helper_numeric3, "NUMERIC", "")))
    create_tbl_status = sqlite3_create_tables(moulinette_get_db_complete_path, Array(create_table_query))
End If

If bypass_wash_redi_order = False Then
    Call moulinette_wash_db
End If

End Sub


Public Sub moulinette_inject_EMSX_orders()

Dim exec_query As Variant
Dim sql_query As String

Dim l_export_omx_all_header As Integer
    l_export_omx_all_header = 1

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

'passe en revue les fichiers
Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
oReg.Global = True


Dim file_export_omx_all As String

Dim found_file As Boolean
found_file = False

oReg.Pattern = "grid(\d){1,}.*xls"
Dim tmp_wrbk As Workbook
For Each tmp_wrbk In Workbooks
    
    Set matches = oReg.Execute(tmp_wrbk.name)
    
    For Each match In matches
        file_export_omx_all = tmp_wrbk.name
        found_file = True
        Exit For
    Next
    
Next

If found_file = False Then
    'MsgBox ("export file not found !")
    Exit Sub
End If


'debug_test = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_redi)
'wash les vieux ordres / 'wash les ordre emsx
sql_query = "DELETE FROM " & t_moulinette_order_redi & " WHERE " & f_moulinette_order_redi_BranchSequence & " LIKE """ & prefix_emsx_trades & "%"""
exec_query = sqlite3_query(moulinette_get_db_complete_path, sql_query)
'debug_test = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_redi)

'passe en revue les colonnes
Dim c_export_omx_all

For i = 1 To 50
    If Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "" Then
        Exit For
    Else
        If Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Security" Then
            c_export_omx_all_ticker = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Ordered" Or Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Amount" Then
            c_export_omx_all_order_qty = i
        ElseIf UCase(Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i)) = UCase("Filled") Then
             c_export_omx_all_exec_qty = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Average Price" Or Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "AvgPrc" Then
            c_export_omx_all_exec_price = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Limit Price" Then
            c_export_omx_all_order_price = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Trader" Then
            c_export_omx_all_userid = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Order Creation Date/Time" Then
            c_export_omx_all_datetime = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Order Create Date" Then
            c_export_omx_all_order_create_date = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Order Entry Time" Then
            c_export_omx_all_order_entry_time = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Order Number" Or Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Order#" Then
            c_export_omx_all_order_id = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Side" Then
            c_export_omx_all_side = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Status" Or Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Color Status" Then
            c_export_omx_all_status = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Ex" Then
            c_export_omx_all_exchange = i
        ElseIf Workbooks(file_export_omx_all).Worksheets(1).Cells(l_export_omx_all_header, i) = "Stop Price" Then
            c_export_omx_all_stop_price = i
        End If
    End If
Next i



If c_export_omx_all_ticker = 0 Then
    MsgBox ("missing column ""Security"" in OMX ALL/EMSX OR ""Security"" and ""Ex"" in EMSX")
    Exit Sub
End If


If c_export_omx_all_order_qty = 0 Then
    MsgBox ("missing column ""Ordered"" in OMX ALL or ""Amount"" in EMSX")
    Exit Sub
End If

If c_export_omx_all_order_price = 0 Then
    MsgBox ("missing column ""Limit Price"" in OMX ALL")
    Exit Sub
End If


If c_export_omx_all_exec_price = 0 Then
    MsgBox ("missing column ""Average Price"" in OMX ALL or ""AvgPrc"" in EMSX")
    Exit Sub
End If

If c_export_omx_all_exec_qty = 0 Then
    MsgBox ("missing column ""Filled"" in OMX ALL or ""FILLED"" in EMSX")
    Exit Sub
End If


If c_export_omx_all_order_id = 0 Then
    MsgBox ("missing column ""Order Number"" in OMX ALL or ""Order#"" in EMSX")
    Exit Sub
End If


If c_export_omx_all_datetime = 0 And (c_export_omx_all_order_create_date = 0 Or c_export_omx_all_order_entry_time = 0) Then
    MsgBox ("missing column ""Order Creation Date/Time"" in OMX ALL or ""Order Create Date"" and ""Order Entry Time"" in EMSX")
    Exit Sub
End If


If c_export_omx_all_side = 0 Then
    MsgBox ("missing column ""Side"" in OMX ALL / EMSX")
    Exit Sub
End If


If c_export_omx_all_status = 0 Then
    MsgBox ("missing column ""Status"" in OMX ALL or ""Color Status"" in EMSX")
    Exit Sub
End If

If c_export_omx_all_userid = 0 Then
    MsgBox ("missing column ""Status"" in OMX ALL or ""Trader"" in EMSX")
    Exit Sub
End If

If c_export_omx_all_stop_price = 0 Then
    MsgBox ("missing column ""Stop Price"" in EMSX")
    Exit Sub
End If



'recupere tout le le monde
Dim tmp_RefNum As String
Dim tmp_OrderRefKey As String
Dim tmp_Description As String
Dim tmp_BranchSequence As String
Dim tmp_datetime As Date
Dim tmp_side As String
Dim tmp_symbol As String
Dim tmp_OrderQty As Double
Dim tmp_OrderPrice As Double
Dim tmp_ExecQty As Double
Dim tmp_ExecPrice As Double
Dim tmp_PriceType As String
Dim tmp_Status As String
Dim tmp_UserID As String

Dim tmp_date As Date


Dim vec_trade_emsx() As Variant
k = 0
For i = l_export_omx_all_header + 1 To 5000
    If Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_id) = "" Then
        Exit For
    Else
        
        'saute pour l instant les stop
        'If IsNumeric(Replace(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_price), ",", "")) Then
        
            tmp_RefNum = CStr(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_id))
            tmp_OrderRefKey = CStr(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_id))
            tmp_Description = Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_side) & " " & Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_qty) & " " & Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_ticker) & " @ " & Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_price)
            tmp_BranchSequence = prefix_emsx_trades & CStr(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_id))
            
            
            If c_export_omx_all_order_entry_time = 0 And c_export_omx_all_order_create_date = 0 Then
                datetime_year = CInt("20" & Mid(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_datetime), 7, 2))
                datetime_month = CInt(Left(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_datetime), 2))
                datetime_day = CInt(Mid(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_datetime), 4, 2))
                
                datetime_hour = CInt(Mid(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_datetime), 10, 2))
                datetime_minute = CInt(Mid(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_datetime), 13, 2))
                datetime_second = CInt(0)
            Else
                datetime_year = CInt(Left(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_create_date), 4))
                datetime_month = CInt(Mid(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_create_date), 6, 2))
                datetime_day = CInt(Right(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_create_date), 2))
                
                datetime_hour = CInt(Left(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_entry_time), 2))
                datetime_minute = CInt(Mid(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_entry_time), 4, 2))
                datetime_second = CInt(Right(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_entry_time), 2))
            End If
            
            
            tmp_datetime = DateSerial(datetime_year, datetime_month, datetime_day) & " " & TimeSerial(datetime_hour, datetime_minute, datetime_second)
            tmp_date = DateSerial(datetime_year, datetime_month, datetime_day)
            
            tmp_side = Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_side)
            
            If c_export_omx_all_exchange = 0 Then
                tmp_symbol = Replace(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_ticker), " ", ".")
            Else
                tmp_symbol = Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_ticker) & "." & Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_exchange)
            End If
            
            tmp_OrderQty = CDbl(Replace(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_qty), ",", ""))
            
            
            If Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_price) = "ST" Then
                tmp_OrderPrice = CDbl(Replace(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_stop_price), ",", ""))
            Else
                tmp_OrderPrice = CDbl(Replace(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_order_price), ",", ""))
            End If
            
            tmp_ExecQty = CDbl(Replace(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_exec_qty), ",", ""))
            
            If Left(UCase(tmp_side), 1) = "B" Or Left(UCase(tmp_side), 1) = "C" Then
                
            ElseIf Left(tmp_side, 1) = "S" Or Left(tmp_side, 1) = "H" Then
                tmp_OrderQty = -tmp_OrderQty
                tmp_ExecQty = -tmp_ExecQty
            End If
            
            
            tmp_ExecPrice = CDbl(Replace(Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_exec_price), ",", ""))
            
            tmp_PriceType = "LIMIT"
            
            tmp_Status = Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_status)
            
            tmp_UserID = Workbooks(file_export_omx_all).Worksheets(1).Cells(i, c_export_omx_all_userid)
            
            
            
            ReDim Preserve vec_trade_emsx(k)
            vec_trade_emsx(k) = Array(tmp_RefNum, tmp_OrderRefKey, tmp_Description, tmp_BranchSequence, ToJulianDay(tmp_datetime), tmp_side, tmp_symbol, tmp_OrderQty, tmp_OrderPrice, tmp_ExecQty, tmp_ExecPrice, tmp_PriceType, tmp_Status, tmp_UserID)
            k = k + 1
        'End If
    End If
Next i

Workbooks(file_export_omx_all).Close False

If k > 0 Then
    insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_order_redi, vec_trade_emsx, Array(f_moulinette_order_redi_RefNum, f_moulinette_order_redi_OrderRefKey, f_moulinette_order_redi_Desccription, f_moulinette_order_redi_BranchSequence, f_moulinette_order_redi_datetime, f_moulinette_order_redi_side, f_moulinette_order_redi_symbol, f_moulinette_order_redi_OrderQty, f_moulinette_order_redi_OrderPrice, f_moulinette_order_redi_ExecQty, f_moulinette_order_redi_ExecPrice, f_moulinette_order_redi_PriceType, f_moulinette_order_redi_Status, f_moulinette_order_redi_UserID))
    'extract_redi = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_redi & " WHERE " & f_moulinette_order_redi_BranchSequence & " LIKE """ & prefix_emsx_trades & "%""")
    'extract_view = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi & " WHERE " & f_moulinette_aggreg_order_redi_BranchSequence & " LIKE """ & prefix_emsx_trades & "%""")
End If



End Sub




'reception vec : ticker / qty / price / stop (optional) / group_id / json_tag
Public Sub moulinette_inject_EMSX_csv_into_xls_trades(ByVal vec_trades As Variant)

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim tmp_id  As String, tmp_group_id As String, tmp_ticker As String, tmp_json As String, tmp_order_price  As Double, tmp_side As String

Dim vec_moulinette() As Variant


k = 0
If IsArray(vec_trades) Then
    
    dim_ticker = 0
    dim_qty = 1
    dim_price = 2
    dim_stop = 3
    dim_group_id = 4
    dim_json_tag = 5
    
    
    For i = 0 To UBound(vec_trades, 1)
        
        Randomize
        tmp_id = Right(year(Now), 2) & Right("0" & Month(Now), 2) & Right("0" & day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & CInt(1000 * Rnd())
        
        tmp_group_id = vec_trades(i)(dim_group_id)
        
        tmp_json = encode_json_for_DB(vec_trades(i)(dim_json_tag))
        
        tmp_ticker = vec_trades(i)(dim_ticker)
        
        tmp_qty = vec_trades(i)(dim_qty)
        
        tmp_order_price = vec_trades(i)(dim_price)
        
        If IsEmpty(vec_trades(i)(dim_stop)) = False Then
            tmp_order_price = vec_trades(i)(dim_stop)
        Else
            tmp_order_price = vec_trades(i)(dim_price)
        End If
        
        
        tmp_side = get_side_redi_plus_optimize_with_lending(tmp_ticker, tmp_qty, 0, 0, 0, 0, 0)
        
        ReDim Preserve vec_moulinette(i)
        vec_moulinette(i) = Array(tmp_id, tmp_group_id, tmp_ticker, get_symbol_redi_plus(tmp_ticker), ToJulianDay(Now), tmp_side, tmp_qty, tmp_order_price, tmp_json)
        
    Next i
    
    
    Dim db_sqlite_insert_status As Variant
    db_sqlite_insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_order_xls, vec_moulinette, Array(f_moulinette_order_xls_id, f_moulinette_order_xls_group_id, f_moulinette_order_xls_ticker, f_moulinette_order_xls_symbol_redi, f_moulinette_order_xls_datetime, f_moulinette_order_xls_side, f_moulinette_order_xls_order_qty, f_moulinette_order_xls_order_price, f_moulinette_order_xls_json_tag))
    
End If

End Sub


Private Sub moulinette_wash_db()

Dim sql_query As String

Dim exec_query As Variant


sql_query = "DELETE FROM " & t_moulinette_bridge_redi & " WHERE " & f_moulinette_bridge_redi_internal_id & " IN ( "
        
        sql_query = sql_query & "SELECT " & f_moulinette_order_xls_id & " FROM " & t_moulinette_order_xls & " WHERE " & f_moulinette_order_xls_datetime & "<" & ToJulianDay(Date)
        
    sql_query = sql_query & " )"

exec_query = sqlite3_query(moulinette_get_db_complete_path, sql_query)



sql_query = "DELETE FROM " & t_moulinette_order_xls & " WHERE " & f_moulinette_order_xls_datetime & "<" & ToJulianDay(Date)
    exec_query = sqlite3_query(moulinette_get_db_complete_path, sql_query)

'wash systematic de la table redi+
sql_query = "DELETE FROM " & t_moulinette_order_redi & " WHERE " & f_moulinette_order_redi_BranchSequence & " NOT LIKE """ & prefix_emsx_trades & "%"""
exec_query = sqlite3_query(moulinette_get_db_complete_path, sql_query)



End Sub


Public Sub moulinette_inject_redi_orders()

Call moulinette_init_db

Dim i As Long, j As Long, k As Long, m As Long, n As Long

If IsRediReady Then
    
        If ThisWorkbook.OrderQuery Is Nothing Then
            Set ThisWorkbook.OrderQuery = New RediLib.CacheControl
        End If
        
        ThisWorkbook.OrderQuery.UserID = ""
        ThisWorkbook.OrderQuery.Password = ""
        vtable = "Message"
        vwhere = "true"
        
        MessageQuery = ThisWorkbook.OrderQuery.Submit(vtable, vwhere, verr)
        
        ThisWorkbook.OrderQuery.Revoke verr
    
    'get_redi_orders = ThisWorkbook.RediOrders
    Dim extract_msg_table As Variant
    extract_msg_table = ThisWorkbook.RediMsg
    
    If IsEmpty(extract_msg_table) Then
        MsgBox ("error api redi")
        Exit Sub
    End If
    
    
    'on repere les differents ordres
    
    'detect des dim
    For i = 0 To UBound(extract_msg_table(0), 1)
        If extract_msg_table(0)(i) = "RefNum" Then
            dim_redi_RefNum = i
        ElseIf extract_msg_table(0)(i) = "BranchSequence" Then
            dim_redi_BranchSequence = i
        ElseIf extract_msg_table(0)(i) = "OrderRefKey" Then
            dim_redi_OrderRefKey = i
        ElseIf extract_msg_table(0)(i) = "Desc" Then
            dim_redi_Desc = i
        ElseIf extract_msg_table(0)(i) = "datetime" Then
            dim_redi_datetime = i
        ElseIf extract_msg_table(0)(i) = "SideAbrev" Then
            dim_redi_SideAbrev = i
        ElseIf extract_msg_table(0)(i) = "Symbol" Then
            dim_redi_Symbol = i
        ElseIf extract_msg_table(0)(i) = "OrderQty" Then
            dim_redi_OrderQty = i
        ElseIf extract_msg_table(0)(i) = "OrderPrice" Then
            dim_redi_OrderPrice = i
        ElseIf extract_msg_table(0)(i) = "ExecQty" Then
            dim_redi_ExecQty = i
        ElseIf extract_msg_table(0)(i) = "ExecPrice" Then
            dim_redi_ExecPrice = i
        ElseIf extract_msg_table(0)(i) = "Status" Then
            dim_redi_Status = i
        ElseIf extract_msg_table(0)(i) = "PriceType" Then
            dim_redi_PriceType = i
        ElseIf extract_msg_table(0)(i) = "UserID" Then
            dim_redi_UserID = i
        End If
    Next i
    

    k = 0
    Dim vec_redi_exec()
    For i = 1 To UBound(extract_msg_table, 1)
        
        'If extract_msg_table(i)(dim_redi_Status) = "Open" Or extract_msg_table(i)(dim_redi_Status) = "Partial" Or (extract_msg_table(i)(dim_redi_Status) = "Complete" And extract_msg_table(i)(dim_redi_ExecPrice) <> "" And extract_msg_table(i)(dim_redi_ExecPrice) <> 0) Then
        If (extract_msg_table(i)(dim_redi_Status) = "Open" Or extract_msg_table(i)(dim_redi_Status) = "Partial" Or extract_msg_table(i)(dim_redi_Status) = "Complete" Or extract_msg_table(i)(dim_redi_Status) = "Canceled") And extract_msg_table(i)(dim_redi_SideAbrev) <> "Invalid" Then
            
            tmp_RefNum = extract_msg_table(i)(dim_redi_RefNum)
            tmp_OrderRefKey = extract_msg_table(i)(dim_redi_OrderRefKey)
            tmp_Desc = extract_msg_table(i)(dim_redi_Desc)
            tmp_BranchSequence = extract_msg_table(i)(dim_redi_BranchSequence)
            tmp_datetime = extract_msg_table(i)(dim_redi_datetime)
            tmp_SideAbrev = extract_msg_table(i)(dim_redi_SideAbrev)
            tmp_symbol = extract_msg_table(i)(dim_redi_Symbol)
            tmp_OrderQty = extract_msg_table(i)(dim_redi_OrderQty)
            tmp_OrderPrice = extract_msg_table(i)(dim_redi_OrderPrice)
            
            If extract_msg_table(i)(dim_redi_ExecPrice) = "" Or extract_msg_table(i)(dim_redi_ExecPrice) = 0 Then
                tmp_ExecQty = 0
            Else
                tmp_ExecQty = extract_msg_table(i)(dim_redi_ExecQty)
            End If
            
            
            tmp_ExecPrice = extract_msg_table(i)(dim_redi_ExecPrice)
            tmp_PriceType = extract_msg_table(i)(dim_redi_PriceType)
            tmp_Status = extract_msg_table(i)(dim_redi_Status)
            tmp_UserID = extract_msg_table(i)(dim_redi_UserID)
            
            
            'ajustement de certaines var
            If InStr(extract_msg_table(i)(dim_redi_SideAbrev), "B") <> 0 Then
                
            ElseIf InStr(extract_msg_table(i)(dim_redi_SideAbrev), "S") <> 0 Then
                tmp_OrderQty = -tmp_OrderQty
                tmp_ExecQty = -tmp_ExecQty
            End If
            
            If tmp_ExecQty = 0 Then
                tmp_ExecPrice = 0
            End If
            
            
            ReDim Preserve vec_redi_exec(k)
            vec_redi_exec(k) = Array(CStr(tmp_RefNum), tmp_OrderRefKey, tmp_Desc, tmp_BranchSequence, tmp_datetime, tmp_SideAbrev, _
                tmp_symbol, tmp_OrderQty, tmp_OrderPrice, tmp_ExecQty, tmp_ExecPrice, tmp_PriceType, tmp_Status, tmp_UserID)
            k = k + 1
            
        End If
        
    Next i
    
    
    If k > 0 Then
        insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_order_redi, vec_redi_exec, Array(f_moulinette_order_redi_RefNum, f_moulinette_order_redi_OrderRefKey, f_moulinette_order_redi_Desccription, f_moulinette_order_redi_BranchSequence, f_moulinette_order_redi_datetime, f_moulinette_order_redi_side, f_moulinette_order_redi_symbol, f_moulinette_order_redi_OrderQty, f_moulinette_order_redi_OrderPrice, f_moulinette_order_redi_ExecQty, f_moulinette_order_redi_ExecPrice, f_moulinette_order_redi_PriceType, f_moulinette_order_redi_Status, f_moulinette_order_redi_UserID))
        'debug_test = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_redi)
        'debug_test = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi)
    End If
    
    
    'aggreg avec view et match pour remplir bridge
    Call moulinette_match_xls_redi_order
    
End If


End Sub


Private Sub moulinette_match_xls_redi_order()

Dim offline_status As Variant
offline_status = moulinette_wash_offline_xls_store()


Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim sql_query As String

Dim extract_redi_order As Variant, extract_view As Variant
'extract_redi_order = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_redi)
'extract_view = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi)

Dim extract_xls_order_without_bridge As Variant
sql_query = "SELECT " & f_moulinette_order_xls_id & ", " & f_moulinette_order_xls_ticker & ", " & f_moulinette_order_xls_symbol_redi & ", " & f_moulinette_order_xls_datetime & ", " & f_moulinette_order_xls_order_qty & ", " & f_moulinette_order_xls_order_price
    sql_query = sql_query & " FROM " & t_moulinette_order_xls
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & " NOT IN ("
        sql_query = sql_query & "SELECT " & f_moulinette_bridge_redi_internal_id & " FROM " & t_moulinette_bridge_redi
    sql_query = sql_query & ")"
    sql_query = sql_query & " AND " & f_moulinette_order_xls_ticker & " NOT LIKE ""%INDEX"""
    
extract_xls_order_without_bridge = sqlite3_query(moulinette_get_db_complete_path, sql_query)
    

Dim date_xls As Date
Dim date_redi As Date

If UBound(extract_xls_order_without_bridge, 1) > 0 Then
    
    'detect des dim
    For i = 0 To UBound(extract_xls_order_without_bridge(0), 1)
        If extract_xls_order_without_bridge(0)(i) = f_moulinette_order_xls_id Then
            dim_xls_id = i
        ElseIf extract_xls_order_without_bridge(0)(i) = f_moulinette_order_xls_ticker Then
            dim_xls_ticker = i
        ElseIf extract_xls_order_without_bridge(0)(i) = f_moulinette_order_xls_symbol_redi Then
            dim_xls_symbol_redi = i 'version excel
        ElseIf extract_xls_order_without_bridge(0)(i) = f_moulinette_order_xls_datetime Then
            dim_xls_datetime = i
        ElseIf extract_xls_order_without_bridge(0)(i) = f_moulinette_order_xls_order_qty Then
            dim_xls_order_qty = i
        ElseIf extract_xls_order_without_bridge(0)(i) = f_moulinette_order_xls_order_price Then
            dim_xls_order_price = i
        End If
    Next i
    
    
    'tente un matching brutal
    Dim time_limit_match_sec As Double
    Dim extract_brutal As Variant
    k = 0
    m = 0
    Dim vec_bridge_to_add() As Variant
    Dim vec_BranchSequence_already_match() As Variant
        ReDim Preserve vec_BranchSequence_already_match(0)
        vec_BranchSequence_already_match(0) = Array("empty", "empty")
    For i = 1 To UBound(extract_xls_order_without_bridge, 1)
        
        sql_query = "SELECT * "
            sql_query = sql_query & " FROM " & v_moulinette_aggreg_order_redi
            sql_query = sql_query & " WHERE " & f_moulinette_aggreg_order_redi_OrderQty & "=" & extract_xls_order_without_bridge(i)(dim_xls_order_qty)
            'sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_symbol & "=""" & Replace(extract_xls_order_without_bridge(i)(dim_xls_symbol_redi), " ", ".") & """"
            sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_OrderPrice & ">" & CDbl(0.99999999 * extract_xls_order_without_bridge(i)(dim_xls_order_price))
            sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_OrderPrice & "<" & CDbl(1.00000001 * extract_xls_order_without_bridge(i)(dim_xls_order_price))
        
        extract_brutal = sqlite3_query(moulinette_get_db_complete_path, sql_query)
        
        If UBound(extract_brutal) = 0 Then
            Debug.Print "not found in redi queue: " & extract_xls_order_without_bridge(i)(dim_xls_id) & " " & extract_xls_order_without_bridge(i)(dim_xls_ticker) & " " & extract_xls_order_without_bridge(i)(dim_xls_order_qty) & " " & extract_xls_order_without_bridge(i)(dim_xls_order_price)
        Else
            
            'detect des dim
            For j = 0 To UBound(extract_brutal(0), 1)
                If extract_brutal(0)(j) = f_moulinette_aggreg_order_redi_BranchSequence Then
                    dim_view_BranchSequence = j
                ElseIf extract_brutal(0)(j) = f_moulinette_aggreg_order_redi_first_datetime Then
                    dim_view_datetime = j
                ElseIf extract_brutal(0)(j) = f_moulinette_aggreg_order_redi_symbol Then
                    dim_view_symbol = j
                End If
            Next j
            
            date_xls = FromJulianDay(CDbl(extract_xls_order_without_bridge(i)(dim_xls_datetime)))
            
            time_limit_match_sec = 10
            time_limit_match_emsx_sec = 180
                
            If UBound(extract_brutal) = 1 Then
                's'assure que l heure est proche
                date_redi = FromJulianDay(CDbl(extract_brutal(1)(dim_view_datetime)))
                
                If Left(extract_brutal(1)(dim_view_BranchSequence), Len(prefix_emsx_trades)) = prefix_emsx_trades Then
                    
                    If Abs(date_xls - date_redi) < (time_limit_match_emsx_sec / (86400)) Then
                        
                        For j = 0 To UBound(vec_BranchSequence_already_match, 1)
                            If extract_brutal(1)(dim_view_BranchSequence) = vec_BranchSequence_already_match(j)(0) And extract_brutal(1)(dim_view_symbol) = vec_BranchSequence_already_match(j)(1) Then
                                Exit For
                            Else
                                If j = UBound(vec_BranchSequence_already_match, 1) Then
                                    ReDim Preserve vec_bridge_to_add(k)
                                    vec_bridge_to_add(k) = Array(extract_xls_order_without_bridge(i)(dim_xls_id), extract_brutal(1)(dim_view_BranchSequence), extract_brutal(1)(dim_view_symbol), extract_xls_order_without_bridge(i)(dim_xls_symbol_redi))
                                    
                                    ReDim Preserve vec_BranchSequence_already_match(k)
                                    vec_BranchSequence_already_match(k) = Array(extract_brutal(1)(dim_view_BranchSequence), extract_brutal(1)(dim_view_symbol))
                                    k = k + 1
                                    
                                End If
                            End If
                        Next j
                        
                    End If
                    
                Else
                
                    If Abs(date_xls - date_redi) < (time_limit_match_sec / (86400)) Then
                        
                        For j = 0 To UBound(vec_BranchSequence_already_match, 1)
                            If extract_brutal(1)(dim_view_BranchSequence) = vec_BranchSequence_already_match(j)(0) And extract_brutal(1)(dim_view_symbol) = vec_BranchSequence_already_match(j)(1) Then
                                Exit For
                            Else
                                If j = UBound(vec_BranchSequence_already_match, 1) Then
                                    ReDim Preserve vec_bridge_to_add(k)
                                    vec_bridge_to_add(k) = Array(extract_xls_order_without_bridge(i)(dim_xls_id), extract_brutal(1)(dim_view_BranchSequence), extract_brutal(1)(dim_view_symbol), extract_xls_order_without_bridge(i)(dim_xls_symbol_redi))
                                    
                                    ReDim Preserve vec_BranchSequence_already_match(k)
                                    vec_BranchSequence_already_match(k) = Array(extract_brutal(1)(dim_view_BranchSequence), extract_brutal(1)(dim_view_symbol))
                                    k = k + 1
                                    
                                End If
                            End If
                        Next j
                    End If
                End If
            Else
                'necessite filtre plus fin
                
                'passe en vue les differentes entree et prend celui dont le time est le plus serre
                Dim spread_time As Double
                    spread_time = 5
                Dim pos_min_spread As Integer
                    pos_min_spread = -1
                
                For j = 1 To UBound(extract_brutal, 1)
                    For m = 0 To UBound(vec_BranchSequence_already_match, 1)
                        If vec_BranchSequence_already_match(m)(0) = extract_brutal(j)(dim_view_BranchSequence) And vec_BranchSequence_already_match(m)(1) = extract_brutal(j)(dim_view_symbol) Then
                            Exit For
                        Else
                            If m = UBound(vec_BranchSequence_already_match, 1) Then
                                
                                date_redi = FromJulianDay(CDbl(extract_brutal(j)(dim_view_datetime)))
                                
                                'mesure time difference
                                If Abs(date_xls - date_redi) < (time_limit_match_sec / (86400)) Then
                                    
                                    If Abs(date_xls - date_redi) < spread_time Then
                                        spread_time = Abs(date_xls - date_redi)
                                        pos_min_spread = j
                                    End If
                                    
                                End If
                            End If
                        End If
                    Next m
                Next j
                
                If pos_min_spread <> -1 Then
                    
                    ReDim Preserve vec_bridge_to_add(k)
                    vec_bridge_to_add(k) = Array(extract_xls_order_without_bridge(i)(dim_xls_id), extract_brutal(pos_min_spread)(dim_view_BranchSequence), extract_brutal(pos_min_spread)(dim_view_symbol), extract_xls_order_without_bridge(i)(dim_xls_symbol_redi))
                    
                    ReDim Preserve vec_BranchSequence_already_match(k)
                    vec_BranchSequence_already_match(k) = Array(extract_brutal(pos_min_spread)(dim_view_BranchSequence), extract_brutal(pos_min_spread)(dim_view_symbol))
                    k = k + 1
                    
                End If
                
            End If
        End If
        
    Next i
    
    If k > 0 Then
        insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_bridge_redi, vec_bridge_to_add, Array(f_moulinette_bridge_redi_internal_id, f_moulinette_bridge_redi_BranchSequence, f_moulinette_bridge_redi_SymbolRedi, f_moulinette_bridge_redi_SymbolXLS))
    End If
    
End If


End Sub


Private Sub test_moulinette_update_static_data_ticker()

debug_test = moulinette_update_static_data_ticker(Array("fp fp EQUITY", "ESU2 INDEX"))

End Sub




Private Function moulinette_update_static_data_ticker(ByVal vec_ticker As Variant) As Variant

moulinette_update_static_data_ticker = Empty

Dim oBBG As New cls_Bloomberg_Sync

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim sql_query As String
sql_query = "SELECT " & f_moulinette_static_ticker & " FROM " & t_moulinette_static
Dim extract_ticker_static As Variant
extract_ticker_static = sqlite3_query(moulinette_get_db_complete_path, sql_query)

k = 0
Dim vec_ticker_static() As Variant
If UBound(extract_ticker_static, 1) = 0 Then
    'on prend *
    For i = 0 To UBound(vec_ticker, 1)
        ReDim Preserve vec_ticker_static(i)
        vec_ticker_static(i) = vec_ticker(i)
        k = k + 1
    Next i
Else
    
    For i = 0 To UBound(vec_ticker, 1)
        For j = 1 To UBound(extract_ticker_static, 1)
            If UCase(vec_ticker(i)) = UCase(extract_ticker_static(j)(0)) Then
                Exit For
            Else
                If j = UBound(extract_ticker_static, 1) Then
                    ReDim Preserve vec_ticker_static(k)
                    vec_ticker_static(k) = vec_ticker(i)
                    k = k + 1
                End If
            End If
        Next j
    Next i
    
End If

Dim tmp_contract_size As Double
If k > 0 Then
    
    Dim vec_field() As Variant
        vec_field = Array("CRNCY", "FUT_CONT_SIZE")
    
    Dim data_bbg As Variant
    data_bbg = oBBG.bdp(vec_ticker_static, vec_field, output_format.of_vec_without_header)
    
    k = 0
    Dim vec_db_static() As Variant
    For i = 0 To UBound(data_bbg, 1)
        
        If Left(data_bbg(i)(0), 1) <> "#" Then
            
            If Left(data_bbg(i)(1), 1) <> "#" Then
                tmp_contract_size = data_bbg(i)(1)
            Else
                tmp_contract_size = 1
            End If
            
            
            ReDim Preserve vec_db_static(k)
            vec_db_static(k) = Array(UCase(vec_ticker_static(i)), UCase(data_bbg(i)(0)), tmp_contract_size)
            k = k + 1
        Else
            'ticker problem
        End If
        
    Next i
    
    If k > 0 Then
        insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_static, vec_db_static, Array(f_moulinette_static_ticker, f_moulinette_static_crncy, f_moulinette_static_fut_cont_size))
    End If
    
End If

Dim extract_static As Variant
extract_static = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_static & " ORDER BY " & f_moulinette_static_ticker & " ASC")

If UBound(extract_static, 1) = 0 Then
Else
    moulinette_update_static_data_ticker = extract_static
End If

End Function


Public Sub moulinette_report_pnl()

Call moulinette_inject_EMSX_orders 'si fichier exporte
Call moulinette_inject_redi_orders 'matching compris

'extract_xls_order = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_xls)
'extract_redi_order = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_redi)
'extract_view = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi)


Dim limit_color_pnl As Double, color_pnl_above As Integer, color_pnl_below As Integer
    limit_color_pnl = 500
    color_pnl_above = 4
    color_pnl_below = 3
    

Application.Calculation = xlCalculationManual

Dim base_crncy As String
    base_crncy = "USD"

Dim oBBG As New cls_Bloomberg_Sync
Dim oJSON As New JSONLib


Dim sql_query As String

sql_query = "SELECT " & f_moulinette_bridge_redi_internal_id & ", " & f_moulinette_bridge_redi_BranchSequence & ", " & f_moulinette_order_xls_group_id & ", " & f_moulinette_order_xls_ticker & ", " & f_moulinette_order_xls_datetime & ", " & f_moulinette_order_xls_order_qty & ", " & f_moulinette_order_xls_order_price & ", " & f_moulinette_order_xls_json_tag & ", " & f_moulinette_aggreg_order_redi_first_datetime & ", " & f_moulinette_aggreg_order_redi_order_type & ", " & f_moulinette_aggreg_order_redi_OrderQty & ", " & f_moulinette_aggreg_order_redi_OrderPrice & ", " & f_moulinette_aggreg_order_redi_ExecQty & ", " & f_moulinette_aggreg_order_redi_NTCF & ", " & f_moulinette_aggreg_order_redi_AvgExecPrice & ", " & f_moulinette_aggreg_order_redi_last_status & ", " & f_moulinette_aggreg_order_redi_Commissions & ", " & f_moulinette_aggreg_order_redi_last_exec_datetime
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_order_xls_symbol_redi & "=" & f_moulinette_bridge_redi_SymbolXLS
    
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    sql_query = sql_query & " AND UPPER(" & f_moulinette_order_xls_ticker & ") NOT LIKE ""%INDEX"""
    
    sql_query = sql_query & " ORDER BY " & f_moulinette_order_xls_ticker & " ASC, " & f_moulinette_order_xls_group_id & " ASC, " & f_moulinette_order_xls_datetime & " ASC"

Dim extract_trades As Variant
extract_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)

If UBound(extract_trades, 1) = 0 Then
    MsgBox ("no trades or problem db, -> Exit")
    Exit Sub
End If
    
    For i = 0 To UBound(extract_trades(0), 1)
        If extract_trades(0)(i) = f_moulinette_bridge_redi_internal_id Then
            dim_trade_xls_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_bridge_redi_BranchSequence Then
            dim_trade_redi_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_group_id Then
            dim_trade_group_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_ticker Then
            dim_trade_ticker = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_datetime Then
            dim_trade_datetime = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_order_type Then
            dim_trade_order_type = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_last_status Then
            dim_trade_last_status = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_order_qty Then
            dim_trade_order_qty = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_order_price Then
            dim_trade_order_price = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_json_tag Then
            dim_trade_json_tag = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_ExecQty Then
            dim_trade_exec_qty = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_AvgExecPrice Then
            dim_trade_exec_price = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_Commissions Then
            dim_trade_comm = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_last_exec_datetime Then
            dim_trade_last_exec_datetime = i
        End If
    Next i




' stat pour summary ticker
sql_query = "SELECT " & f_moulinette_order_xls_ticker & ", SUM(" & f_moulinette_aggreg_order_redi_OrderQty & ") AS " & f_moulinette_stat_ticker_OrderQty & ", SUM(" & f_moulinette_aggreg_order_redi_ExecQty & ") AS " & f_moulinette_stat_ticker_ExecQty & ", SUM(" & f_moulinette_aggreg_order_redi_ExecQty & "*" & f_moulinette_aggreg_order_redi_AvgExecPrice & ") AS " & f_moulinette_stat_ticker_NTCF & ", SUM(" & f_moulinette_aggreg_order_redi_OrderQty & "*" & f_moulinette_aggreg_order_redi_OrderPrice & ")/SUM(" & f_moulinette_aggreg_order_redi_OrderQty & ")" & " AS " & f_moulinette_stat_ticker_AVGOrderPrice & ", SUM(" & f_moulinette_aggreg_order_redi_ExecQty & "*" & f_moulinette_aggreg_order_redi_AvgExecPrice & ")/SUM(" & f_moulinette_aggreg_order_redi_ExecQty & ")" & " AS " & f_moulinette_stat_ticker_AVGExecPrice
    sql_query = sql_query & " FROM " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi & ", " & t_moulinette_order_xls
    sql_query = sql_query & " WHERE " & f_moulinette_aggreg_order_redi_BranchSequence & "=" & f_moulinette_bridge_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_symbol & "=" & f_moulinette_bridge_redi_SymbolRedi
    sql_query = sql_query & " AND " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolXLS & "=" & f_moulinette_order_xls_symbol_redi
    sql_query = sql_query & " GROUP BY " & f_moulinette_order_xls_ticker

Dim extract_stat_ticker As Variant
extract_stat_ticker = sqlite3_query(moulinette_get_db_complete_path, sql_query)
    
    For i = 0 To UBound(extract_stat_ticker(0), 1)
        If extract_stat_ticker(0)(i) = f_moulinette_order_xls_ticker Then
            dim_stat_ticker_ticker = i
        ElseIf extract_stat_ticker(0)(i) = f_moulinette_stat_ticker_OrderQty Then
            dim_stat_ticker_order_qty = i
        ElseIf extract_stat_ticker(0)(i) = f_moulinette_stat_ticker_ExecQty Then
            dim_stat_ticker_exec_qty = i
        ElseIf extract_stat_ticker(0)(i) = f_moulinette_stat_ticker_NTCF Then
            dim_stat_ticker_ntcf = i
        ElseIf extract_stat_ticker(0)(i) = f_moulinette_stat_ticker_AVGOrderPrice Then
            dim_stat_ticker_avg_order_price = i
        ElseIf extract_stat_ticker(0)(i) = f_moulinette_stat_ticker_AVGExecPrice Then
            dim_stat_ticker_avg_exec_price = i
        End If
    Next i

    
    

'second appel pour les tickers
sql_query = "SELECT DISTINCT " & f_moulinette_order_xls_ticker
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id

Dim extract_distinct_ticker As Variant
extract_distinct_ticker = sqlite3_query(moulinette_get_db_complete_path, sql_query)


Dim vec_ticker() As Variant
For i = 1 To UBound(extract_distinct_ticker, 1)
    ReDim Preserve vec_ticker(i - 1)
    vec_ticker(i - 1) = extract_distinct_ticker(i)(0)
Next i


Dim data_bbg As Variant
data_bbg = oBBG.bdp(vec_ticker, Array("px_last"), output_format.of_vec_without_header)


Dim extract_static As Variant
extract_static = moulinette_update_static_data_ticker(vec_ticker)

If IsEmpty(extract_static) Then
    extract_static = Array(Array("", ""))
Else
    For i = 0 To UBound(extract_static(0), 1)
        If extract_static(0)(i) = f_moulinette_static_ticker Then
            dim_static_ticker = i
        ElseIf extract_static(0)(i) = f_moulinette_static_crncy Then
            dim_static_crncy = i
        ElseIf extract_static(0)(i) = f_moulinette_static_fut_cont_size Then
            dim_static_contract_size = i
        End If
    Next i
End If



'on imprime le report
Application.ReferenceStyle = xlA1
Worksheets(sheet_report).Cells.Clear

Worksheets(sheet_report).Cells.ClearOutline


Dim color_group_summary As Integer, color_ticker_summary As Integer
    color_group_summary = 34
    color_ticker_summary = 33

'mise en place header
Worksheets(sheet_report).Cells(l_report_header, c_report_group_xls_id) = "group"
Worksheets(sheet_report).Cells(l_report_header, c_report_trade_xls_id) = "xls id"
Worksheets(sheet_report).Cells(l_report_header, c_report_trade_redi_id) = "redi id"
Worksheets(sheet_report).Cells(l_report_header, c_report_datetime) = "datetime"
Worksheets(sheet_report).Cells(l_report_header, c_report_ticker) = "ticker"
Worksheets(sheet_report).Cells(l_report_header, c_report_bbg_px_last) = "LAST_PRICE"

Worksheets(sheet_report).Cells(l_report_header, c_report_bbg_px_high) = "daily high"
Worksheets(sheet_report).Cells(l_report_header, c_report_bbg_px_low) = "daily low"

Worksheets(sheet_report).Cells(l_report_header, c_report_order_type) = "ord type"
Worksheets(sheet_report).Cells(l_report_header, c_report_order_tag) = "tag"
Worksheets(sheet_report).Cells(l_report_header, c_report_order_status) = "status"
Worksheets(sheet_report).Cells(l_report_header, c_report_order_qty) = "ord qty"
Worksheets(sheet_report).Cells(l_report_header, c_report_order_price) = "ord price"
Worksheets(sheet_report).Cells(l_report_header, c_report_exec_qty) = "exec qty"
Worksheets(sheet_report).Cells(l_report_header, c_report_exec_avg_price) = "exec avg price"
Worksheets(sheet_report).Cells(l_report_header, c_report_nominal_open_usd) = "nom. open"
Worksheets(sheet_report).Cells(l_report_header, c_report_nominal_exec_usd) = "nom. exec"
Worksheets(sheet_report).Cells(l_report_header, c_report_pnl_local) = "pnl local"
Worksheets(sheet_report).Cells(l_report_header, c_report_pnl_base) = "pnl base"
Worksheets(sheet_report).Cells(l_report_header, c_report_pnl_with_comm) = "pnl with comm"

    For i = 1 To c_report_pnl_with_comm
        Worksheets(sheet_report).Cells(l_report_header, i).Font.Bold = True
    Next i


Dim vec_currency() As Variant
k = 0
For i = 14 To 50
    If Worksheets("Parametres").Cells(i, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_currency(k)
        vec_currency(k) = Array(Worksheets("Parametres").Cells(i, 1).Value, i, Worksheets("Parametres").Cells(i, 6).Value)
        k = k + 1
    End If
Next i


Dim last_line_formula As Integer

k = l_report_header + 1
Dim last_group_id As Double, last_ticker As String
    last_group_id = -1
    last_ticker = ""
Dim l_ticker_summary As Integer, l_group_summary As Integer

Dim vec_line_group() As Variant

Dim outline_ticker() As Variant
    Dim count_outline_ticker As Integer
    count_outline_ticker = 0
Dim outline_group() As Variant
    Dim count_outline_group As Integer
    count_outline_group = 0

Dim colOrderTag As Collection, ElemOrderTag As Variant

Dim vec_line_aggreg_ticker() As Variant
    Dim count_line_aggreg_ticker As Integer
    count_line_aggreg_ticker = 0



Dim pnl_long As Double, pnl_short As Double
pnl_long = 0
pnl_short = 0

Dim nom_open_long As Double, nom_open_short As Double
Dim nom_exec_long As Double, nom_exec_short As Double

nom_open_long = 0
nom_open_short = 0
nom_exec_long = 0
nom_exec_short = 0


For i = 1 To UBound(extract_trades, 1)
    
    
    'stat - ticker
    If last_ticker <> extract_trades(i)(dim_trade_ticker) Then
        
        ReDim Preserve vec_line_aggreg_ticker(count_line_aggreg_ticker)
        vec_line_aggreg_ticker(count_line_aggreg_ticker) = k
        count_line_aggreg_ticker = count_line_aggreg_ticker + 1
        
        If k <> l_report_header + 1 Then
            
            m = 0
            For j = k - 1 To l_ticker_summary + 1 Step -1
                
                If Worksheets(sheet_report).Cells(j, c_report_ticker) <> "" And Worksheets(sheet_report).Cells(j, c_report_ticker) <> extract_trades(i)(dim_trade_ticker) Then
                    Exit For
                Else
                    If Worksheets(sheet_report).Cells(j, c_report_group_xls_id) <> "" Then
                        ReDim Preserve vec_line_group(m)
                        vec_line_group(m) = j
                        m = m + 1
                    End If
                End If
                
            Next j
            
            
            If m > 0 Then
                
                tmp_formula_order_qty = ""
                tmp_formula_exec_qty = ""
                tmp_formula_nominal_open = ""
                tmp_formula_nominal_exec = ""
                tmp_formula_pnl_local = ""
                tmp_formula_pnl_base = ""
                tmp_formula_pnl_with_comm = ""
                For j = 0 To UBound(vec_line_group, 1)
                    If j = 0 Then
                        
                    Else
                        tmp_formula_order_qty = tmp_formula_order_qty & "+"
                        tmp_formula_exec_qty = tmp_formula_exec_qty & "+"
                        tmp_formula_nominal_open = tmp_formula_nominal_open & "+"
                        tmp_formula_nominal_exec = tmp_formula_nominal_exec & "+"
                        tmp_formula_pnl_local = tmp_formula_pnl_local & "+"
                        tmp_formula_pnl_base = tmp_formula_pnl_base & "+"
                        tmp_formula_pnl_with_comm = tmp_formula_pnl_with_comm & "+"
                    End If
                    
                    tmp_formula_order_qty = tmp_formula_order_qty & xlColumnValue(c_report_order_qty) & vec_line_group(j)
                    tmp_formula_exec_qty = tmp_formula_exec_qty & xlColumnValue(c_report_exec_qty) & vec_line_group(j)
                    tmp_formula_nominal_open = tmp_formula_nominal_open & xlColumnValue(c_report_nominal_open_usd) & vec_line_group(j)
                    tmp_formula_nominal_exec = tmp_formula_nominal_exec & xlColumnValue(c_report_nominal_exec_usd) & vec_line_group(j)
                    tmp_formula_pnl_local = tmp_formula_pnl_local & xlColumnValue(c_report_pnl_local) & vec_line_group(j)
                    tmp_formula_pnl_base = tmp_formula_pnl_base & xlColumnValue(c_report_pnl_base) & vec_line_group(j)
                    tmp_formula_pnl_with_comm = tmp_formula_pnl_with_comm & xlColumnValue(c_report_pnl_with_comm) & vec_line_group(j)
                Next j
                
                
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_order_qty).FormulaLocal = "=" & tmp_formula_order_qty
                    Worksheets(sheet_report).Cells(l_ticker_summary, c_report_order_qty).NumberFormat = "#,##0"
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_exec_qty).FormulaLocal = "=" & tmp_formula_exec_qty
                    Worksheets(sheet_report).Cells(l_ticker_summary, c_report_exec_qty).NumberFormat = "#,##0"
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_open_usd).FormulaLocal = "=" & tmp_formula_nominal_open
                    Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_open_usd).NumberFormat = "#,##0"
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_exec_usd).FormulaLocal = "=" & tmp_formula_nominal_exec
                    Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_exec_usd).NumberFormat = "#,##0"
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_local).FormulaLocal = "=" & tmp_formula_pnl_local
                    Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_local).NumberFormat = "#,##0"
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_base).FormulaLocal = "=" & tmp_formula_pnl_base
                    Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_base).NumberFormat = "#,##0"
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_with_comm).FormulaLocal = "=" & tmp_formula_pnl_with_comm
                    Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_with_comm).NumberFormat = "#,##0"
                
                
                ReDim Preserve outline_ticker(count_outline_ticker)
                outline_ticker(count_outline_ticker) = l_ticker_summary + 1 & ":" & k - 1
                count_outline_ticker = count_outline_ticker + 1
                
            End If
            
        End If
        
        
        'mise en place ligne aggreg
        Worksheets(sheet_report).Cells(k, c_report_ticker) = extract_trades(i)(dim_trade_ticker)
        Worksheets(sheet_report).Cells(k, c_report_bbg_px_last).FormulaLocal = "=BDP(" & xlColumnValue(c_report_ticker) & k & ";""LAST_PRICE"")"
            Worksheets(sheet_report).Cells(k, c_report_bbg_px_last).NumberFormat = "#,##0.00"
        
        Worksheets(sheet_report).Cells(k, c_report_bbg_px_high).FormulaLocal = "=BDP(" & xlColumnValue(c_report_ticker) & k & ";""PX_HIGH"")"
            Worksheets(sheet_report).Cells(k, c_report_bbg_px_high).NumberFormat = "#,##0.00"
        
        Worksheets(sheet_report).Cells(k, c_report_bbg_px_low).FormulaLocal = "=BDP(" & xlColumnValue(c_report_ticker) & k & ";""PX_LOW"")"
            Worksheets(sheet_report).Cells(k, c_report_bbg_px_low).NumberFormat = "#,##0.00"
        
        
        For j = 1 To UBound(extract_stat_ticker, 1)
            If extract_stat_ticker(j)(dim_stat_ticker_ticker) = extract_trades(i)(dim_trade_ticker) Then
                Worksheets(sheet_report).Cells(k, c_report_order_price) = extract_stat_ticker(j)(dim_stat_ticker_avg_order_price)
                    Worksheets(sheet_report).Cells(k, c_report_order_price).NumberFormat = "#,##0.00"
                Worksheets(sheet_report).Cells(k, c_report_exec_avg_price) = extract_stat_ticker(j)(dim_stat_ticker_avg_exec_price)
                    Worksheets(sheet_report).Cells(k, c_report_exec_avg_price).NumberFormat = "#,##0.00"
                Exit For
            End If
        Next j
        
        
        
        For j = 1 To c_report_pnl_with_comm
            Worksheets(sheet_report).Cells(k, j).Interior.ColorIndex = color_ticker_summary
            Worksheets(sheet_report).Cells(k, j).Font.Bold = True
        Next j
        
        Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions.Delete
        Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=limit_color_pnl
            Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions(1).Interior.ColorIndex = color_pnl_above
        Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=-limit_color_pnl
            Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions(2).Interior.ColorIndex = color_pnl_below
        
        
        l_ticker_summary = k
        k = k + 1
        
    End If
    
    
    'stats - group
    If last_group_id <> extract_trades(i)(dim_trade_group_id) Then
        
        If k <> l_report_header + 2 Then ' a cause de l aggreg ticker
            
            If last_ticker <> extract_trades(i)(dim_trade_ticker) Then
                last_line_formula = k - 2
            Else
                last_line_formula = k - 1
            End If
            
            Worksheets(sheet_report).Cells(l_group_summary, c_report_order_qty).FormulaLocal = "=SUM(" & xlColumnValue(c_report_order_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_order_qty) & last_line_formula & ")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_order_qty).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_qty).FormulaLocal = "=SUM(" & xlColumnValue(c_report_exec_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_exec_qty) & last_line_formula & ")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_qty).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_order_price).FormulaLocal = "=IF(" & xlColumnValue(c_report_order_qty) & l_group_summary & "<>0;SUMPRODUCT(" & xlColumnValue(c_report_order_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_order_qty) & last_line_formula & ";" & xlColumnValue(c_report_order_price) & l_group_summary + 1 & ":" & xlColumnValue(c_report_order_price) & last_line_formula & ")/" & xlColumnValue(c_report_order_qty) & l_group_summary & ";"""")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_order_price).NumberFormat = "#,##0.00"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_avg_price).FormulaLocal = "=IF(" & xlColumnValue(c_report_exec_qty) & l_group_summary & "<>0;SUMPRODUCT(" & xlColumnValue(c_report_exec_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_exec_qty) & last_line_formula & ";" & xlColumnValue(c_report_exec_avg_price) & l_group_summary + 1 & ":" & xlColumnValue(c_report_exec_avg_price) & last_line_formula & ")/" & xlColumnValue(c_report_exec_qty) & l_group_summary & ";"""")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_avg_price).NumberFormat = "#,##0.00"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_open_usd).FormulaLocal = "=-SUMIF(" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & last_line_formula & ";""<0"";" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & last_line_formula & ")+SUMIF(" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & last_line_formula & ";"">=0"";" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & last_line_formula & ")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_open_usd).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_exec_usd).FormulaLocal = "=-SUMIF(" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & last_line_formula & ";""<0"";" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & last_line_formula & ")+SUMIF(" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & last_line_formula & ";"">=0"";" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & last_line_formula & ")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_exec_usd).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_local).FormulaLocal = "=SUM(" & xlColumnValue(c_report_pnl_local) & l_group_summary + 1 & ":" & xlColumnValue(c_report_pnl_local) & last_line_formula & ")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_local).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_base).FormulaLocal = "=SUM(" & xlColumnValue(c_report_pnl_base) & l_group_summary + 1 & ":" & xlColumnValue(c_report_pnl_base) & last_line_formula & ")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_base).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_with_comm).FormulaLocal = "=SUM(" & xlColumnValue(c_report_pnl_with_comm) & l_group_summary + 1 & ":" & xlColumnValue(c_report_pnl_with_comm) & last_line_formula & ")"
                Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_with_comm).NumberFormat = "#,##0"
            
            
            'group
            ReDim Preserve outline_group(count_outline_group)
            outline_group(count_outline_group) = l_group_summary + 1 & ":" & last_line_formula
            count_outline_group = count_outline_group + 1
            
        End If
        
        
        'mise en place ligne aggreg
        Worksheets(sheet_report).Cells(k, c_report_group_xls_id) = extract_trades(i)(dim_trade_group_id)
        Worksheets(sheet_report).Cells(k, c_report_datetime) = FromJulianDay(CDbl(extract_trades(i)(dim_trade_datetime)))
        Worksheets(sheet_report).Cells(k, c_report_bbg_px_last).FormulaLocal = "=" & xlColumnValue(c_report_bbg_px_last) & l_ticker_summary
        
        For j = 1 To c_report_pnl_with_comm
            Worksheets(sheet_report).Cells(k, j).Interior.ColorIndex = color_group_summary
            Worksheets(sheet_report).Cells(k, j).Font.Bold = True
        Next j
        
        
         Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions.Delete
        Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=limit_color_pnl
            Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions(1).Interior.ColorIndex = color_pnl_above
        Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=-limit_color_pnl
            Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormatConditions(2).Interior.ColorIndex = color_pnl_below
        
        l_group_summary = k
        k = k + 1
    End If
    
    
    
    ' ################################################### trade ###################################################
    
    Worksheets(sheet_report).Cells(k, c_report_trade_xls_id) = extract_trades(i)(dim_trade_xls_id)
    Worksheets(sheet_report).Cells(k, c_report_trade_redi_id) = extract_trades(i)(dim_trade_redi_id)
    Worksheets(sheet_report).Cells(k, c_report_datetime) = FromJulianDay(CDbl(extract_trades(i)(dim_trade_last_exec_datetime)))
    Worksheets(sheet_report).Cells(k, c_report_order_type) = extract_trades(i)(dim_trade_order_type)
    Worksheets(sheet_report).Cells(k, c_report_order_status) = extract_trades(i)(dim_trade_last_status)
    
    
    'rank high / low
    If extract_trades(i)(dim_trade_exec_qty) > 0 Then
        'pct du low
        Worksheets(sheet_report).Cells(k, c_report_bbg_px_low).FormulaLocal = "=1-((" & xlColumnValue(c_report_exec_avg_price) & k & "-" & xlColumnValue(c_report_bbg_px_low) & l_ticker_summary & ")/(" & xlColumnValue(c_report_bbg_px_high) & l_ticker_summary & "-" & xlColumnValue(c_report_bbg_px_low) & l_ticker_summary & "))"
            Worksheets(sheet_report).Cells(k, c_report_bbg_px_low).NumberFormat = "0%"
    ElseIf extract_trades(i)(dim_trade_exec_qty) < 0 Then
        'pct du high
        Worksheets(sheet_report).Cells(k, c_report_bbg_px_high).FormulaLocal = "=((" & xlColumnValue(c_report_exec_avg_price) & k & "-" & xlColumnValue(c_report_bbg_px_low) & l_ticker_summary & ")/(" & xlColumnValue(c_report_bbg_px_high) & l_ticker_summary & "-" & xlColumnValue(c_report_bbg_px_low) & l_ticker_summary & "))"
            Worksheets(sheet_report).Cells(k, c_report_bbg_px_high).NumberFormat = "0%"
    End If
    
    If IsNull(extract_trades(i)(dim_trade_json_tag)) Then
        
    Else
        Set colOrderTag = oJSON.parse(CStr(decode_json_from_DB(extract_trades(i)(dim_trade_json_tag))))
        
        If colOrderTag Is Nothing Then
        Else
            p = 0
            For Each ElemOrderTag In colOrderTag
                If p = 0 Then
                    Worksheets(sheet_report).Cells(k, c_report_order_tag).Value = ""
                Else
                    Worksheets(sheet_report).Cells(k, c_report_order_tag).Value = Worksheets(sheet_report).Cells(k, c_report_order_tag).Value & " "
                End If
                
                Worksheets(sheet_report).Cells(k, c_report_order_tag).Value = Worksheets(sheet_report).Cells(k, c_report_order_tag).Value & ElemOrderTag
                p = p + 1
            Next
            
        End If
        
    End If
    
    
    Worksheets(sheet_report).Cells(k, c_report_order_qty) = extract_trades(i)(dim_trade_order_qty)
        Worksheets(sheet_report).Cells(k, c_report_order_qty).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(k, c_report_order_price) = extract_trades(i)(dim_trade_order_price)
    Worksheets(sheet_report).Cells(k, c_report_exec_qty) = extract_trades(i)(dim_trade_exec_qty)
        Worksheets(sheet_report).Cells(k, c_report_exec_qty).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(k, c_report_exec_avg_price) = extract_trades(i)(dim_trade_exec_price)
        Worksheets(sheet_report).Cells(k, c_report_exec_avg_price).NumberFormat = "#,##0.00"
    Worksheets(sheet_report).Cells(k, c_report_bbg_px_last).FormulaLocal = "=" & xlColumnValue(c_report_bbg_px_last) & l_group_summary
        Worksheets(sheet_report).Cells(k, c_report_bbg_px_last).NumberFormat = "#,##0.00"
        
    For j = 1 To UBound(extract_static, 1)
        If UCase(extract_static(j)(dim_static_ticker)) = UCase(extract_trades(i)(dim_trade_ticker)) Then
            
            For m = 0 To UBound(vec_currency, 1)
                If UCase(vec_currency(m)(0)) = UCase(extract_static(j)(dim_static_crncy)) Then
                    
                    'nominal usd open
                    Worksheets(sheet_report).Cells(k, c_report_nominal_open_usd).FormulaLocal = "=" & extract_static(j)(dim_static_contract_size) & "*" & xlColumnValue(c_report_order_qty) & k & "*" & xlColumnValue(c_report_order_price) & k & "*Parametres!F" & vec_currency(m)(1)
                        Worksheets(sheet_report).Cells(k, c_report_nominal_open_usd).NumberFormat = "#,##0"
                        
                        
                        If extract_trades(i)(dim_trade_order_qty) < 0 Then
                            nom_open_short = nom_open_short + extract_trades(i)(dim_trade_order_qty) * extract_trades(i)(dim_trade_order_price) * vec_currency(m)(2)
                        ElseIf extract_trades(i)(dim_trade_order_qty) > 0 Then
                            nom_open_long = nom_open_long + extract_trades(i)(dim_trade_order_qty) * extract_trades(i)(dim_trade_order_price) * vec_currency(m)(2)
                        End If
                        
                    
                    'nominal usd exec
                    Worksheets(sheet_report).Cells(k, c_report_nominal_exec_usd).FormulaLocal = "=" & extract_static(j)(dim_static_contract_size) & "*" & xlColumnValue(c_report_exec_qty) & k & "*" & xlColumnValue(c_report_exec_avg_price) & k & "*Parametres!F" & vec_currency(m)(1)
                        Worksheets(sheet_report).Cells(k, c_report_nominal_exec_usd).NumberFormat = "#,##0"
                        
                        If extract_trades(i)(dim_trade_exec_qty) < 0 Then
                            nom_exec_short = nom_exec_short + extract_trades(i)(dim_trade_exec_qty) * extract_trades(i)(dim_trade_exec_price) * vec_currency(m)(2)
                        ElseIf extract_trades(i)(dim_trade_exec_qty) > 0 Then
                            nom_exec_long = nom_exec_long + extract_trades(i)(dim_trade_exec_qty) * extract_trades(i)(dim_trade_exec_price) * vec_currency(m)(2)
                        End If
                        
                        
                    'pnl local with contract size
                    Worksheets(sheet_report).Cells(k, c_report_pnl_local).FormulaLocal = "=" & extract_static(j)(dim_static_contract_size) & "*" & xlColumnValue(c_report_exec_qty) & k & "*(" & xlColumnValue(c_report_bbg_px_last) & k & "-" & xlColumnValue(c_report_exec_avg_price) & k & ")"
                        Worksheets(sheet_report).Cells(k, c_report_pnl_local).NumberFormat = "#,##0"
                    
                    'pnl base
                    Worksheets(sheet_report).Cells(k, c_report_pnl_base).FormulaLocal = "=" & xlColumnValue(c_report_pnl_local) & k & "*Parametres!F" & vec_currency(m)(1)
                        Worksheets(sheet_report).Cells(k, c_report_pnl_base).NumberFormat = "#,##0"
                    
                    'pnl base + comm
                    Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).FormulaLocal = "=(" & xlColumnValue(c_report_pnl_local) & k & "-" & Round(Abs(extract_trades(i)(dim_trade_comm)), 4) & ")*Parametres!F" & vec_currency(m)(1)
                        Worksheets(sheet_report).Cells(k, c_report_pnl_with_comm).NumberFormat = "#,##0"
                    
                    
                    For n = 0 To UBound(vec_ticker, 1)
                        If vec_ticker(n) = extract_trades(i)(dim_trade_ticker) Then
                            
                            If extract_trades(i)(dim_trade_exec_qty) < 0 Then
                                pnl_short = pnl_short + vec_currency(m)(2) * ((extract_static(j)(dim_static_contract_size) * extract_trades(i)(dim_trade_exec_qty) * (data_bbg(n)(0) - extract_trades(i)(dim_trade_exec_price))) - Abs(extract_trades(i)(dim_trade_comm)))
                            ElseIf extract_trades(i)(dim_trade_exec_qty) > 0 Then
                                pnl_long = pnl_long + vec_currency(m)(2) * ((extract_static(j)(dim_static_contract_size) * extract_trades(i)(dim_trade_exec_qty) * (data_bbg(n)(0) - extract_trades(i)(dim_trade_exec_price))) - Abs(extract_trades(i)(dim_trade_comm)))
                            End If
                            
                            Exit For
                        End If
                    Next n
                    
                    
                    Exit For
                End If
            Next m
            
            Exit For
        End If
    Next j
    
    For j = c_report_group_xls_id To c_report_pnl_with_comm
        Worksheets(sheet_report).Cells(k, j).FormatConditions.Delete
        Worksheets(sheet_report).Cells(k, j).FormatConditions.Add type:=xlExpression, Formula1:="=$" & xlColumnValue(c_report_pnl_with_comm) & "$" & k & "<-500"
        Worksheets(sheet_report).Cells(k, j).FormatConditions(1).Interior.ColorIndex = color_pnl_below
        Worksheets(sheet_report).Cells(k, j).FormatConditions.Add type:=xlExpression, Formula1:="=$" & xlColumnValue(c_report_pnl_with_comm) & "$" & k & ">500"
        Worksheets(sheet_report).Cells(k, j).FormatConditions(2).Interior.ColorIndex = color_pnl_above
    Next j
    
    
    ' ##############################################################################################################
    
    'mise en place calcul stat agregg (sum etc.)
    If i = UBound(extract_trades, 1) Then
        
        
        'stat - ticker
        m = 0
        For j = k - 1 To l_ticker_summary + 1 Step -1
            
            If Worksheets(sheet_report).Cells(j, c_report_ticker) <> "" And Worksheets(sheet_report).Cells(j, c_report_ticker) <> extract_trades(i)(dim_trade_ticker) Then
                Exit For
            Else
                If Worksheets(sheet_report).Cells(j, c_report_group_xls_id) <> "" Then
                    ReDim Preserve vec_line_group(m)
                    vec_line_group(m) = j
                    m = m + 1
                End If
            End If
            
        Next j
        
        
        If m > 0 Then
            
            tmp_formula_order_qty = ""
            tmp_formula_exec_qty = ""
            tmp_formula_nominal_open = ""
            tmp_formula_nominal_exec = ""
            tmp_formula_pnl_local = ""
            tmp_formula_pnl_base = ""
            tmp_formula_pnl_with_comm = ""
            For j = 0 To UBound(vec_line_group, 1)
                If j = 0 Then
                    
                Else
                    tmp_formula_order_qty = tmp_formula_order_qty & "+"
                    tmp_formula_exec_qty = tmp_formula_exec_qty & "+"
                    tmp_formula_nominal_open = tmp_formula_nominal_open & "+"
                    tmp_formula_nominal_exec = tmp_formula_nominal_exec & "+"
                    tmp_formula_pnl_local = tmp_formula_pnl_local & "+"
                    tmp_formula_pnl_base = tmp_formula_pnl_base & "+"
                    tmp_formula_pnl_with_comm = tmp_formula_pnl_with_comm & "+"
                End If
                
                tmp_formula_order_qty = tmp_formula_order_qty & xlColumnValue(c_report_order_qty) & vec_line_group(j)
                tmp_formula_exec_qty = tmp_formula_exec_qty & xlColumnValue(c_report_exec_qty) & vec_line_group(j)
                tmp_formula_nominal_open = tmp_formula_nominal_open & xlColumnValue(c_report_nominal_open_usd) & vec_line_group(j)
                tmp_formula_nominal_exec = tmp_formula_nominal_exec & xlColumnValue(c_report_nominal_exec_usd) & vec_line_group(j)
                tmp_formula_pnl_local = tmp_formula_pnl_local & xlColumnValue(c_report_pnl_local) & vec_line_group(j)
                tmp_formula_pnl_base = tmp_formula_pnl_base & xlColumnValue(c_report_pnl_base) & vec_line_group(j)
                tmp_formula_pnl_with_comm = tmp_formula_pnl_with_comm & xlColumnValue(c_report_pnl_with_comm) & vec_line_group(j)
                
            Next j
            
            Worksheets(sheet_report).Cells(l_ticker_summary, c_report_order_qty).FormulaLocal = "=" & tmp_formula_order_qty
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_order_qty).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_ticker_summary, c_report_exec_qty).FormulaLocal = "=" & tmp_formula_exec_qty
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_exec_qty).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_open_usd).FormulaLocal = "=" & tmp_formula_nominal_open
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_open_usd).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_exec_usd).FormulaLocal = "=" & tmp_formula_nominal_exec
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_nominal_exec_usd).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_local).FormulaLocal = "=" & tmp_formula_pnl_local
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_local).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_base).FormulaLocal = "=" & tmp_formula_pnl_base
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_base).NumberFormat = "#,##0"
            Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_with_comm).FormulaLocal = "=" & tmp_formula_pnl_with_comm
                Worksheets(sheet_report).Cells(l_ticker_summary, c_report_pnl_with_comm).NumberFormat = "#,##0"
            
        End If
        
        
        ReDim Preserve outline_ticker(count_outline_ticker)
        outline_ticker(count_outline_ticker) = l_ticker_summary + 1 & ":" & k
        count_outline_ticker = count_outline_ticker + 1
        
        
        
        
        
        'stat - group
        Worksheets(sheet_report).Cells(l_group_summary, c_report_order_qty).FormulaLocal = "=SUM(" & xlColumnValue(c_report_order_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_order_qty) & k & ")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_order_qty).NumberFormat = "#,##0"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_qty).FormulaLocal = "=SUM(" & xlColumnValue(c_report_exec_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_exec_qty) & k & ")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_qty).NumberFormat = "#,##0"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_order_price).FormulaLocal = "=IF(" & xlColumnValue(c_report_order_qty) & l_group_summary & "<>0;SUMPRODUCT(" & xlColumnValue(c_report_order_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_order_qty) & k & ";" & xlColumnValue(c_report_order_price) & l_group_summary + 1 & ":" & xlColumnValue(c_report_order_price) & k & "/" & xlColumnValue(c_report_order_qty) & l_group_summary & ");"""")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_order_price).NumberFormat = "#,##0.00"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_avg_price).FormulaLocal = "=IF(" & xlColumnValue(c_report_exec_qty) & l_group_summary & "<>0;SUMPRODUCT(" & xlColumnValue(c_report_exec_qty) & l_group_summary + 1 & ":" & xlColumnValue(c_report_exec_qty) & k & ";" & xlColumnValue(c_report_exec_avg_price) & l_group_summary + 1 & ":" & xlColumnValue(c_report_exec_avg_price) & k & "/" & xlColumnValue(c_report_exec_qty) & l_group_summary & ");"""")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_avg_price).NumberFormat = "#,##0.00"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_open_usd).FormulaLocal = "=-SUMIF(" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & k & ";""<0"";" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & k & ")+SUMIF(" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & k & ";"">=0"";" & xlColumnValue(c_report_nominal_open_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_open_usd) & k & ")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_open_usd).NumberFormat = "#,##0"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_exec_usd).FormulaLocal = "=-SUMIF(" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & k & ";""<0"";" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & k & ")+SUMIF(" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & k & ";"">=0"";" & xlColumnValue(c_report_nominal_exec_usd) & l_group_summary + 1 & ":" & xlColumnValue(c_report_nominal_exec_usd) & k & ")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_nominal_exec_usd).NumberFormat = "#,##0"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_local).FormulaLocal = "=SUM(" & xlColumnValue(c_report_pnl_local) & l_group_summary + 1 & ":" & xlColumnValue(c_report_pnl_local) & k & ")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_qty).NumberFormat = "#,##0"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_base).FormulaLocal = "=SUM(" & xlColumnValue(c_report_pnl_base) & l_group_summary + 1 & ":" & xlColumnValue(c_report_pnl_base) & k & ")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_exec_qty).NumberFormat = "#,##0"
        Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_with_comm).FormulaLocal = "=SUM(" & xlColumnValue(c_report_pnl_with_comm) & l_group_summary + 1 & ":" & xlColumnValue(c_report_pnl_with_comm) & k & ")"
            Worksheets(sheet_report).Cells(l_group_summary, c_report_pnl_with_comm).NumberFormat = "#,##0"
        
            'group
            ReDim Preserve outline_group(count_outline_group)
            outline_group(count_outline_group) = l_group_summary + 1 & ":" & k
            count_outline_group = count_outline_group + 1
        
    End If
    
    
    Worksheets(sheet_report).Cells(k, c_report_pnl_base).FormatConditions.Delete
    Worksheets(sheet_report).Cells(k, c_report_pnl_base).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=limit_color_pnl
        Worksheets(sheet_report).Cells(k, c_report_pnl_base).FormatConditions(1).Interior.ColorIndex = color_pnl_above
    Worksheets(sheet_report).Cells(k, c_report_pnl_base).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=-limit_color_pnl
        Worksheets(sheet_report).Cells(k, c_report_pnl_base).FormatConditions(2).Interior.ColorIndex = color_pnl_below
    
    last_ticker = extract_trades(i)(dim_trade_ticker)
    last_group_id = extract_trades(i)(dim_trade_group_id)
    
    k = k + 1
    
Next i

Dim report_last_line As Integer
    report_last_line = k - 1


'mise en place outline
If count_outline_group > 0 Then
    For i = 0 To UBound(outline_group, 1)
        Worksheets(sheet_report).rows(outline_group(i)).rows.Group
    Next i
End If


If count_outline_ticker > 0 Then
    For i = 0 To UBound(outline_ticker, 1)
        Worksheets(sheet_report).rows(outline_ticker(i)).rows.Group
    Next i
End If


Dim formula_pnl_sum As String, formula_pnl_max As String, formula_pnl_min As String, formula_pnl_vec As String
    formula_nominal_open = ""
    formula_nominal_exec = ""
    formula_pnl_sum = ""
    formula_pnl_max = ""
    formula_pnl_min = ""
    formula_pnl_vec = ""


If count_line_aggreg_ticker > 0 Then
    
    For i = 0 To UBound(vec_line_aggreg_ticker, 1)
        If i = 0 Then
        Else
            formula_pnl_vec = formula_pnl_vec & ";"
            formula_nominal_open = formula_nominal_open & ";"
            formula_nominal_exec = formula_nominal_exec & ";"
        End If
        
        formula_pnl_vec = formula_pnl_vec & xlColumnValue(c_report_pnl_with_comm) & vec_line_aggreg_ticker(i)
        formula_nominal_open = formula_nominal_open & xlColumnValue(c_report_nominal_open_usd) & vec_line_aggreg_ticker(i)
        formula_nominal_exec = formula_nominal_exec & xlColumnValue(c_report_nominal_exec_usd) & vec_line_aggreg_ticker(i)
        
    Next i
    
    
    'header
    Worksheets(sheet_report).Cells(l_report_summary_long, c_report_bbg_px_last) = "LONG"
    Worksheets(sheet_report).Cells(l_report_summary_short, c_report_bbg_px_last) = "SHORT"
    Worksheets(sheet_report).Cells(l_report_summary_net, c_report_bbg_px_last) = "NET"
    
    Worksheets(sheet_report).Cells(l_report_summary_long, c_report_nominal_open_usd) = nom_open_long
        Worksheets(sheet_report).Cells(l_report_summary_long, c_report_nominal_open_usd).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(l_report_summary_long, c_report_nominal_exec_usd) = nom_exec_long
        Worksheets(sheet_report).Cells(l_report_summary_long, c_report_nominal_exec_usd).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(l_report_summary_long, c_report_pnl_with_comm) = pnl_long
        Worksheets(sheet_report).Cells(l_report_summary_long, c_report_pnl_with_comm).NumberFormat = "#,##0"
        
            Worksheets(sheet_report).Cells(l_report_summary_long, c_report_pnl_with_comm).FormatConditions.Delete
            Worksheets(sheet_report).Cells(l_report_summary_long, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=limit_color_pnl
                Worksheets(sheet_report).Cells(l_report_summary_long, c_report_pnl_with_comm).FormatConditions(1).Interior.ColorIndex = color_pnl_above
            Worksheets(sheet_report).Cells(l_report_summary_long, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=-limit_color_pnl
                Worksheets(sheet_report).Cells(l_report_summary_long, c_report_pnl_with_comm).FormatConditions(2).Interior.ColorIndex = color_pnl_below
        
    
    Worksheets(sheet_report).Cells(l_report_summary_short, c_report_nominal_open_usd) = nom_open_short
        Worksheets(sheet_report).Cells(l_report_summary_short, c_report_nominal_open_usd).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(l_report_summary_short, c_report_nominal_exec_usd) = nom_exec_short
        Worksheets(sheet_report).Cells(l_report_summary_short, c_report_nominal_exec_usd).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(l_report_summary_short, c_report_pnl_with_comm) = pnl_short
        Worksheets(sheet_report).Cells(l_report_summary_short, c_report_pnl_with_comm).NumberFormat = "#,##0"
    
            Worksheets(sheet_report).Cells(l_report_summary_short, c_report_pnl_with_comm).FormatConditions.Delete
            Worksheets(sheet_report).Cells(l_report_summary_short, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=limit_color_pnl
                Worksheets(sheet_report).Cells(l_report_summary_short, c_report_pnl_with_comm).FormatConditions(1).Interior.ColorIndex = color_pnl_above
            Worksheets(sheet_report).Cells(l_report_summary_short, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=-limit_color_pnl
                Worksheets(sheet_report).Cells(l_report_summary_short, c_report_pnl_with_comm).FormatConditions(2).Interior.ColorIndex = color_pnl_below
    
    
    
    
    Worksheets(sheet_report).Cells(l_report_summary_net, c_report_nominal_open_usd) = nom_open_long + Abs(nom_open_short)
        Worksheets(sheet_report).Cells(l_report_summary_net, c_report_nominal_open_usd).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(l_report_summary_net, c_report_nominal_exec_usd) = nom_exec_long + Abs(nom_exec_short)
        Worksheets(sheet_report).Cells(l_report_summary_net, c_report_nominal_exec_usd).NumberFormat = "#,##0"
    Worksheets(sheet_report).Cells(l_report_summary_net, c_report_pnl_with_comm).FormulaLocal = "=SUM(" & xlColumnValue(c_report_pnl_with_comm) & l_report_header + 1 & ":" & xlColumnValue(c_report_pnl_with_comm) & report_last_line & ")/3"
        Worksheets(sheet_report).Cells(l_report_summary_net, c_report_pnl_with_comm).NumberFormat = "#,##0"
        
            Worksheets(sheet_report).Cells(l_report_summary_net, c_report_pnl_with_comm).FormatConditions.Delete
            Worksheets(sheet_report).Cells(l_report_summary_net, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=limit_color_pnl
                Worksheets(sheet_report).Cells(l_report_summary_net, c_report_pnl_with_comm).FormatConditions(1).Interior.ColorIndex = color_pnl_above
            Worksheets(sheet_report).Cells(l_report_summary_net, c_report_pnl_with_comm).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=-limit_color_pnl
                Worksheets(sheet_report).Cells(l_report_summary_net, c_report_pnl_with_comm).FormatConditions(2).Interior.ColorIndex = color_pnl_below
                
    
    
    Worksheets(sheet_report).Outline.ShowLevels RowLevels:=1
    
End If



Application.Calculation = xlCalculationAutomatic

End Sub


Public Function moulinette_cancel_ticker(ByVal ticker As String) As Variant

ticker = UCase(ticker)

Dim oBBG As New cls_Bloomberg_Sync
Dim oJSON As New JSONLib

Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


'repere les trades a canceler
sql_query = "SELECT " & f_moulinette_aggreg_order_redi_BranchSequence & ", " & f_moulinette_aggreg_order_redi_OrderQty & ", " & f_moulinette_aggreg_order_redi_ExecQty
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_order_xls_ticker & "=""" & ticker & """"
    sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_OrderQty & "<>" & f_moulinette_aggreg_order_redi_ExecQty
    ' ##############################################################################################

Dim extract_trades As Variant
extract_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)


'repere les trades a canceler
If UBound(extract_trades, 1) = 0 Then
    
Else
    For i = 1 To UBound(extract_trades, 1)
        Call moulinette_cancel_trade(extract_trades(i)(0))
    Next i
    
    Call moulinette_report_pnl
End If



End Function


Public Function moulinette_close_ticker(ByVal ticker As String) As Variant

ticker = UCase(ticker)

Dim oBBG As New cls_Bloomberg_Sync
Dim oJSON As New JSONLib

Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


'extract_view = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & v_moulinette_aggreg_order_redi)
'extract_bridge = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_bridge_redi)
'extract_xls = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_xls & " WHERE " & f_moulinette_order_xls_ticker & "=""" & ticker & """")


sql_query = "SELECT " & f_moulinette_order_xls_ticker & ", SUM(" & f_moulinette_aggreg_order_redi_ExecQty & ") AS QtyExecTicker"
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_order_xls_ticker & "=""" & ticker & """"
    ' ##############################################################################################


Dim extract_aggreg_group As Variant
extract_aggreg_group = sqlite3_query(moulinette_get_db_complete_path, sql_query)

Dim qty_to_close As Double
If UBound(extract_aggreg_group, 1) = 0 Then
    MsgBox ("id not found. -> Exit")
    moulinette_close_ticker = False
    Exit Function
Else
    qty_to_close = -extract_aggreg_group(1)(1)
End If



'repere les trades a canceler
sql_query = "SELECT " & f_moulinette_aggreg_order_redi_BranchSequence & ", " & f_moulinette_aggreg_order_redi_OrderQty & ", " & f_moulinette_aggreg_order_redi_ExecQty
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_order_xls_ticker & "=""" & ticker & """"
    sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_OrderQty & "<>" & f_moulinette_aggreg_order_redi_ExecQty
    ' ##############################################################################################

Dim extract_trades As Variant
extract_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)

Dim need_update_report As Boolean
    need_update_report = False

'repere les trades a canceler
If UBound(extract_trades, 1) = 0 Then
    
Else
    For i = 1 To UBound(extract_trades, 1)
        Call moulinette_cancel_trade(extract_trades(i)(0))
    Next i
    
    need_update_report = True
End If



If extract_aggreg_group(1)(1) <> 0 Then

    Dim data_bbg As Variant
    data_bbg = oBBG.bdp(Array(extract_aggreg_group(1)(0)), Array("PX_LAST", "PX_BID", "PX_ASK"), output_format.of_vec_without_header)
    
    If IsNumeric(data_bbg(0)(0)) Then
        
        Dim status_trade As Variant
        status_trade = universal_trades_r_plus(Array(Array(extract_aggreg_group(1)(0), qty_to_close, data_bbg(0)(0), Empty, CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0)), oJSON.toString(Array("*** auto close ticker ***")))))
    
        need_update_report = True
    Else
        MsgBox ("error price bbg")
        moulinette_close_ticker = False
        Exit Function
    End If

End If

If need_update_report = True Then
    Call moulinette_report_pnl
End If


End Function



Public Sub moulinette_btn_close_ticker()

Dim i As Integer
Dim L_ticker As Integer

If ActiveSheet.name = sheet_report Then
    
    If Cells(ActiveCell.row, c_report_ticker).Value <> "" Then
        L_ticker = ActiveCell.row
        
run_fn:
        Call moulinette_close_ticker(CStr(Cells(L_ticker, c_report_ticker).Value))
    Else
        For i = ActiveCell.row - 1 To l_report_header + 1 Step -1
            If Cells(i, c_report_ticker).Value <> "" Then
                L_ticker = i
                GoTo run_fn
            End If
        Next i
    End If
End If

End Sub


Public Sub moulinette_btn_cancel_ticker()

Dim i As Integer
Dim L_ticker As Integer

If ActiveSheet.name = sheet_report Then
    
    If Cells(ActiveCell.row, c_report_ticker).Value <> "" Then
        L_ticker = ActiveCell.row
        
run_fn:
        Call moulinette_cancel_ticker(CStr(Cells(L_ticker, c_report_ticker).Value))
    Else
        For i = ActiveCell.row - 1 To l_report_header + 1 Step -1
            If Cells(i, c_report_ticker).Value <> "" Then
                L_ticker = i
                GoTo run_fn
            End If
        Next i
    End If
End If

End Sub


Public Sub moulinette_btn_cancel_group()

Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


Dim group_summary_row As Integer, group_start_row As Integer, group_end_row As Integer
group_summary_row = -3
group_end_row = -2
group_start_row = -1


If ActiveSheet.name = sheet_report Then
    
    If Cells(ActiveCell.row, c_report_group_xls_id).Value <> "" Then
        
        group_summary_row = ActiveCell.row
        group_start_row = ActiveCell.row + 1
        
find_last_row_group:
        For i = group_start_row To 10000
            
            If Cells(i, c_report_trade_redi_id) = "" Then
                group_end_row = i - 1
                Exit For
            End If
            
        Next i
        
        If group_end_row >= group_start_row Then
            
            Call moulinette_cancel_group(Cells(group_summary_row, c_report_group_xls_id).Value)
            
        End If
        
    Else
        If Cells(ActiveCell.row, c_report_trade_redi_id).Value <> "" Then
            
            'on remonte jusqu'a tomber sur le group sauf si summary ticker
            If Cells(ActiveCell.row, c_report_ticker) <> "" Then
            Else
                For i = ActiveCell.row - 1 To 2 Step -1
                    If Cells(i, c_report_trade_redi_id) = "" Then
                        group_summary_row = i
                        group_start_row = i + 1
                        GoTo find_last_row_group
                    End If
                Next i
            End If
        End If
    End If
    
End If

End Sub



Public Sub moulinette_btn_close_group()

Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


Dim group_summary_row As Integer, group_start_row As Integer, group_end_row As Integer
group_summary_row = -3
group_end_row = -2
group_start_row = -1


If ActiveSheet.name = sheet_report Then
    
    If Cells(ActiveCell.row, c_report_group_xls_id).Value <> "" Then
        
        group_summary_row = ActiveCell.row
        group_start_row = ActiveCell.row + 1
        
find_last_row_group:
        For i = group_start_row To 10000
            
            If Cells(i, c_report_trade_redi_id) = "" Then
                group_end_row = i - 1
                Exit For
            End If
            
        Next i
        
        If group_end_row >= group_start_row Then
            
            Call moulinette_close_group(Cells(group_summary_row, c_report_group_xls_id).Value)
            
        End If
        
    Else
        If Cells(ActiveCell.row, c_report_trade_redi_id).Value <> "" Then
            
            'on remonte jusqu'a tomber sur le group
            
            If Cells(ActiveCell.row, c_report_ticker) <> "" Then
            Else
                For i = ActiveCell.row - 1 To 2 Step -1
                    If Cells(i, c_report_trade_redi_id) = "" Then
                        group_summary_row = i
                        group_start_row = i + 1
                        GoTo find_last_row_group
                    End If
                Next i
            End If
        End If
    End If
    
End If

End Sub




Public Sub test_moulinette_close_group()

debug_test = moulinette_close_group(15211274)

End Sub


Public Function moulinette_close_group(ByVal group_id As Double) As Variant

Dim oBBG As New cls_Bloomberg_Sync
Dim oJSON As New JSONLib

Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


sql_query = "SELECT " & f_moulinette_order_xls_ticker & ", SUM(" & f_moulinette_aggreg_order_redi_ExecQty & ") AS QtyExecGroup"
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_order_xls_group_id & "=" & group_id & ""
    ' ##############################################################################################


Dim extract_aggreg_group As Variant
extract_aggreg_group = sqlite3_query(moulinette_get_db_complete_path, sql_query)

Dim qty_to_close As Double
If UBound(extract_aggreg_group, 1) = 0 Then
    MsgBox ("id not found. -> Exit")
    moulinette_close_group = False
    Exit Function
Else
    qty_to_close = -extract_aggreg_group(1)(1)
End If



'repere la qty a cloturer
sql_query = "SELECT " & f_moulinette_aggreg_order_redi_BranchSequence & ", " & f_moulinette_aggreg_order_redi_OrderQty & ", " & f_moulinette_aggreg_order_redi_ExecQty
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_order_xls_group_id & "=" & group_id & ""
    sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_OrderQty & "<>" & f_moulinette_aggreg_order_redi_ExecQty
    ' ##############################################################################################

Dim extract_trades As Variant
extract_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)

Dim need_update_report As Boolean
    need_update_report = False

'repere les trades a canceler
If UBound(extract_trades, 1) = 0 Then
    
Else
    For i = 1 To UBound(extract_trades, 1)
        Call moulinette_cancel_trade(extract_trades(i)(0))
    Next i
    
    need_update_report = True
    
End If



If extract_aggreg_group(1)(1) <> 0 Then

    Dim data_bbg As Variant
    data_bbg = oBBG.bdp(Array(extract_aggreg_group(1)(0)), Array("PX_LAST", "PX_BID", "PX_ASK"), output_format.of_vec_without_header)
    
    If IsNumeric(data_bbg(0)(0)) Then
        
        Dim status_trade As Variant
        status_trade = universal_trades_r_plus(Array(Array(extract_aggreg_group(1)(0), qty_to_close, data_bbg(0)(0), Empty, group_id, oJSON.toString(Array("*** auto close group ***")))))
        
        need_update_report = True
        
    Else
        MsgBox ("error price bbg")
        moulinette_close_group = False
        Exit Function
    End If

End If


If need_update_report = True Then
    Call moulinette_report_pnl
End If

End Function


Public Function moulinette_cancel_group(ByVal group_id As Double) As Variant


Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


'repere les trades a canceler
sql_query = "SELECT " & f_moulinette_aggreg_order_redi_BranchSequence & ", " & f_moulinette_aggreg_order_redi_OrderQty & ", " & f_moulinette_aggreg_order_redi_ExecQty
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_order_xls_group_id & "=" & group_id & ""
    sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_OrderQty & "<>" & f_moulinette_aggreg_order_redi_ExecQty
    ' ##############################################################################################

Dim extract_trades As Variant
extract_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)


'repere les trades a canceler
If UBound(extract_trades, 1) = 0 Then
    
Else
    For i = 1 To UBound(extract_trades, 1)
        Call moulinette_cancel_trade(extract_trades(i)(0))
    Next i
    
    Call moulinette_report_pnl
End If


End Function





Public Sub moulinette_btn_close_trade()

If ActiveSheet.name = sheet_report Then
    
    If Cells(ActiveCell.row, c_report_trade_redi_id).Value <> "" Then
        
        Call moulinette_close_trade(Cells(ActiveCell.row, c_report_trade_redi_id).Value)
        
        Call moulinette_report_pnl
        
    End If
    
End If

End Sub



Public Sub moulinette_btn_cancel_trade()

If ActiveSheet.name = sheet_report Then
    
    If Cells(ActiveCell.row, c_report_trade_redi_id).Value <> "" Then
        
        Call moulinette_cancel_trade(Cells(ActiveCell.row, c_report_trade_redi_id).Value)
        
        Call moulinette_report_pnl
        
    End If
    
End If

End Sub


Public Function moulinette_cancel_trade(ByVal redi_trade_id As String) As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim sql_query As String


sql_query = "SELECT " & f_moulinette_bridge_redi_internal_id & ", " & f_moulinette_bridge_redi_BranchSequence & ", " & f_moulinette_order_xls_group_id & ", " & f_moulinette_order_xls_ticker & ", " & f_moulinette_order_xls_datetime & ", " & f_moulinette_order_xls_order_qty & ", " & f_moulinette_order_xls_order_price & ", " & f_moulinette_order_xls_json_tag & ", " & f_moulinette_aggreg_order_redi_first_datetime & ", " & f_moulinette_aggreg_order_redi_order_type & ", " & f_moulinette_aggreg_order_redi_OrderQty & ", " & f_moulinette_aggreg_order_redi_OrderPrice & ", " & f_moulinette_aggreg_order_redi_ExecQty & ", " & f_moulinette_aggreg_order_redi_NTCF & ", " & f_moulinette_aggreg_order_redi_AvgExecPrice
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=""" & redi_trade_id & """"
    ' ##############################################################################################

Dim extract_trades As Variant
extract_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)

If UBound(extract_trades, 1) = 0 Then
    MsgBox ("no trades or problem db, -> Exit")
    moulinette_cancel_trade = False
    Exit Function
End If
    
    For i = 0 To UBound(extract_trades(0), 1)
        If extract_trades(0)(i) = f_moulinette_bridge_redi_internal_id Then
            dim_trade_xls_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_bridge_redi_BranchSequence Then
            dim_trade_redi_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_group_id Then
            dim_trade_group_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_ticker Then
            dim_trade_ticker = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_datetime Then
            dim_trade_datetime = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_order_type Then
            dim_trade_order_type = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_order_qty Then
            dim_trade_order_qty = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_order_price Then
            dim_trade_order_price = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_json_tag Then
            dim_trade_json_tag = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_ExecQty Then
            dim_trade_exec_qty = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_AvgExecPrice Then
            dim_trade_exec_price = i
        End If
    Next i


If extract_trades(1)(dim_trade_order_qty) = extract_trades(1)(dim_trade_exec_qty) Then
    MsgBox ("nothing to do, order completed.")
    moulinette_cancel_trade = False
    Exit Function
Else
    
    
    If ThisWorkbook.OrderQuery Is Nothing Then
        Set ThisWorkbook.OrderQuery = New RediLib.CacheControl
    End If
    
    Dim error_cancel As Variant
    
    ThisWorkbook.OrderQuery.UserID = ""
    ThisWorkbook.OrderQuery.Password = ""
    vtable = "Message"
    vwhere = "true"
    
    MessageQuery = ThisWorkbook.OrderQuery.Submit(vtable, vwhere, verr)
    moulinette_cancel_trade = ThisWorkbook.OrderQuery.CancelByBranchSequence(redi_trade_id, error_cancel)
    
End If


End Function



Public Function moulinette_close_trade(ByVal redi_trade_id As String) As Variant

Dim oBBG As New cls_Bloomberg_Sync
Dim oJSON As New JSONLib

'Call moulinette_inject_redi_orders 'matching compris

Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

'remonte la pos net associe au trade

sql_query = "SELECT " & f_moulinette_bridge_redi_internal_id & ", " & f_moulinette_bridge_redi_BranchSequence & ", " & f_moulinette_order_xls_group_id & ", " & f_moulinette_order_xls_ticker & ", " & f_moulinette_order_xls_datetime & ", " & f_moulinette_order_xls_order_qty & ", " & f_moulinette_order_xls_order_price & ", " & f_moulinette_order_xls_json_tag & ", " & f_moulinette_aggreg_order_redi_first_datetime & ", " & f_moulinette_aggreg_order_redi_order_type & ", " & f_moulinette_aggreg_order_redi_OrderQty & ", " & f_moulinette_aggreg_order_redi_OrderPrice & ", " & f_moulinette_aggreg_order_redi_ExecQty & ", " & f_moulinette_aggreg_order_redi_NTCF & ", " & f_moulinette_aggreg_order_redi_AvgExecPrice
    sql_query = sql_query & " FROM " & t_moulinette_order_xls & ", " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi
    sql_query = sql_query & " WHERE " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=" & f_moulinette_aggreg_order_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolRedi & "=" & f_moulinette_aggreg_order_redi_symbol
    
    ' ##############################################################################################
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_BranchSequence & "=""" & redi_trade_id & """"
    ' ##############################################################################################

Dim extract_trades As Variant
extract_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)

If UBound(extract_trades, 1) = 0 Then
    MsgBox ("no trades or problem db, -> Exit")
    moulinette_close_trade = False
    Exit Function
End If
    
    For i = 0 To UBound(extract_trades(0), 1)
        If extract_trades(0)(i) = f_moulinette_bridge_redi_internal_id Then
            dim_trade_xls_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_bridge_redi_BranchSequence Then
            dim_trade_redi_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_group_id Then
            dim_trade_group_id = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_ticker Then
            dim_trade_ticker = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_datetime Then
            dim_trade_datetime = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_order_type Then
            dim_trade_order_type = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_order_qty Then
            dim_trade_order_qty = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_order_price Then
            dim_trade_order_price = i
        ElseIf extract_trades(0)(i) = f_moulinette_order_xls_json_tag Then
            dim_trade_json_tag = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_ExecQty Then
            dim_trade_exec_qty = i
        ElseIf extract_trades(0)(i) = f_moulinette_aggreg_order_redi_AvgExecPrice Then
            dim_trade_exec_price = i
        End If
    Next i


Dim tmp_qty_to_trade_to_close_trade As Double, tmp_group_id As Double, trade_close_price As Double
If extract_trades(1)(dim_trade_exec_qty) = 0 Then
'    moulinette_close_trade = False
'    MsgBox ("position not opened. Nothing to do")
'    Exit Function
Else
    tmp_qty_to_trade_to_close_trade = -extract_trades(1)(dim_trade_exec_qty)
    tmp_group_id = extract_trades(1)(dim_trade_group_id)
    
    Dim data_bbg As Variant
    data_bbg = oBBG.bdp(Array(extract_trades(1)(dim_trade_ticker)), Array("PX_LAST", "PX_BID", "PX_ASK"), output_format.of_vec_without_header)
    
    
    If IsNumeric(data_bbg(0)(0)) Then
        trade_close_price = data_bbg(0)(0)
    Else
        MsgBox ("prob price bloomberg")
        moulinette_close_trade = False
        Exit Function
    End If
    
End If


'place un trade au marche

If extract_trades(1)(dim_trade_order_qty) > extract_trades(1)(dim_trade_exec_qty) Then
    'cancel le solde du trade
    cancel_status = moulinette_cancel_trade(redi_trade_id)
End If


Dim status_trade As Variant
status_trade = universal_trades_r_plus(Array(Array(extract_trades(1)(dim_trade_ticker), tmp_qty_to_trade_to_close_trade, trade_close_price, Empty, tmp_group_id, oJSON.toString(Array("*** auto close trade ***")))))

End Function


Public Function moulinette_wash_offline_xls_store() As Variant

moulinette_wash_offline_xls_store = Empty

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim sql_query As String

Call moulinette_init_db(True)

Application.Calculation = xlCalculationManual


Dim last_line_store As Integer
    last_line_store = 0
    
    
Dim vec_offline_orders() As Variant
Dim tmp_row() As Variant
k = 0
For i = 1 To 15000
    
    If Workbooks("Kronos.xls").Worksheets(sheet_offline).Cells(i, c_offline_order_xls_id) = "" Then
        last_line_store = i - 1
        
        If last_line_store <= 0 Then
            last_line_store = 1
        End If
        
        Exit For
    Else
        
        m = 0
        For j = c_offline_order_xls_id To c_offline_order_xls_json_tag
            ReDim Preserve tmp_row(m)
            
            If Workbooks("Kronos.xls").Worksheets(sheet_offline).Cells(i, j) = "" Then
                tmp_row(m) = Empty
            Else
                tmp_row(m) = Workbooks("Kronos.xls").Worksheets(sheet_offline).Cells(i, j)
            End If
            
            m = m + 1
        Next j
        
        ReDim Preserve vec_offline_orders(k)
        vec_offline_orders(k) = tmp_row
        k = k + 1
        
    End If
    
Next i


'on repasse a l envers pour virer les lignes
Dim vec_final_xls_orders() As Variant, vec_final_xls_orders_id() As Variant

If k > 0 Then
    k = 0
    For i = 0 To UBound(vec_offline_orders, 1)
        
        If FromJulianDay(CDbl(vec_offline_orders(i)(c_offline_order_xls_datetime - c_offline_order_xls_id))) < Date Then
            'continue
            Debug.Print "$moulinette_wash_offline_xls_store " & vec_offline_orders(i)(0)
        Else
            ReDim Preserve vec_final_xls_orders(k)
            ReDim Preserve vec_final_xls_orders_id(k)
            
            vec_final_xls_orders(k) = vec_offline_orders(i)
            vec_final_xls_orders_id(k) = Array(CStr(vec_offline_orders(i)(0)))
            
            k = k + 1
        End If
        
    Next i
End If


'on complete avec redi queue
Dim extract_xls_queue As Variant
extract_xls_queue = sqlite3_query(moulinette_get_db_complete_path, "SELECT " & f_moulinette_order_xls_id & ", " & f_moulinette_order_xls_group_id & ", " & f_moulinette_order_xls_ticker & ", " & f_moulinette_order_xls_symbol_redi & ", " & f_moulinette_order_xls_datetime & ", " & f_moulinette_order_xls_side & ", " & f_moulinette_order_xls_order_qty & ", " & f_moulinette_order_xls_order_price & ", " & f_moulinette_order_xls_json_tag & " FROM " & t_moulinette_order_xls)
m = 0
If UBound(extract_xls_queue, 1) > 0 Then

    For i = 1 To UBound(extract_xls_queue, 1)
        
        If k = 0 Then
            ReDim Preserve vec_final_xls_orders(m)
            ReDim Preserve vec_final_xls_orders_id(m)
    
            vec_final_xls_orders(m) = extract_xls_queue(i)
            vec_final_xls_orders_id(m) = Array(CStr(extract_xls_queue(i)(0)))
            k = k + 1
            m = m + 1
        Else
        
            For j = 0 To UBound(vec_final_xls_orders, 1)
                If CStr(vec_final_xls_orders(j)(0)) = extract_xls_queue(i)(0) Then
                    Exit For
                Else
                    If j = UBound(vec_final_xls_orders, 1) Then
                        'dans db mais pas dans xls, on ajoute l entree
                        ReDim Preserve vec_final_xls_orders(UBound(vec_final_xls_orders, 1) + 1)
                        ReDim Preserve vec_final_xls_orders_id(UBound(vec_final_xls_orders_id, 1) + 1)
    
                        vec_final_xls_orders(UBound(vec_final_xls_orders, 1)) = extract_xls_queue(i)
                        vec_final_xls_orders_id(UBound(vec_final_xls_orders_id, 1)) = Array(CStr(extract_xls_queue(i)(0)))
                        k = k + 1
                    End If
                End If
            Next j
        End If
    Next i


End If





'wash excel area
For i = 1 To last_line_store
    For j = c_offline_order_xls_id To c_offline_order_xls_json_tag
        Workbooks("Kronos.xls").Worksheets(sheet_offline).Cells(i, j) = ""
    Next j
Next i


If k > 0 Then
    
    For i = 0 To UBound(vec_final_xls_orders, 1)
        For j = 0 To UBound(vec_final_xls_orders(i), 1)
            Workbooks("Kronos.xls").Worksheets(sheet_offline).Cells(i + 1, c_offline_order_xls_id + j) = vec_final_xls_orders(i)(j)
        Next j
    Next i
    
    
    moulinette_wash_offline_xls_store = vec_final_xls_orders
    
    'on inject dans le helper et on repere les inconnus
    exec_query = sqlite3_query(moulinette_get_db_complete_path, "DELETE FROM " & t_moulinette_helper)
    insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_helper, vec_final_xls_orders_id, Array(f_moulinette_helper_text1))
    
    
    sql_query = "SELECT " & f_moulinette_helper_text1
        sql_query = sql_query & " FROM " & t_moulinette_helper
        sql_query = sql_query & " WHERE " & f_moulinette_helper_text1 & " NOT IN ( "
            sql_query = sql_query & "SELECT " & f_moulinette_order_xls_id & " FROM " & t_moulinette_order_xls
        sql_query = sql_query & " )"
    
    Dim extract_xls_trades As Variant
    extract_xls_trades = sqlite3_query(moulinette_get_db_complete_path, sql_query)
    
    If UBound(extract_xls_trades, 1) = 0 Then
        
    Else
        Dim vec_completed_order() As Variant
        k = 0
        For i = 1 To UBound(extract_xls_trades, 1)
            
            'match sur order
            
            For j = 0 To UBound(vec_final_xls_orders, 1)
                If CStr(extract_xls_trades(i)(0)) = CStr(vec_final_xls_orders(j)(0)) Then
                    ReDim Preserve vec_completed_order(k)
                    vec_completed_order(k) = vec_final_xls_orders(j)
                    k = k + 1
                    Exit For
                End If
            Next j
            
        Next i
        
        
        If k > 0 Then
            insert_status = sqlite3_insert_with_transaction(moulinette_get_db_complete_path, t_moulinette_order_xls, vec_completed_order, Array(f_moulinette_order_xls_id, f_moulinette_order_xls_group_id, f_moulinette_order_xls_ticker, f_moulinette_order_xls_symbol_redi, f_moulinette_order_xls_datetime, f_moulinette_order_xls_side, f_moulinette_order_xls_order_qty, f_moulinette_order_xls_order_price, f_moulinette_order_xls_json_tag))
        End If
        
    End If
    
Else
    moulinette_wash_offline_xls_store = Empty
End If


End Function



'order_qty
'exec_qty
'exec_avg_price
'nominal_gross_local_open
'nominal_gross_local_exec
Public Function moulinette(ByVal ticker As String, ByVal field As String) As Variant

Dim sql_query As String
Dim i As Integer, j As Integer, k As Integer

' stat pour summary ticker
sql_query = "SELECT " & f_moulinette_order_xls_ticker & ", SUM(" & f_moulinette_aggreg_order_redi_OrderQty & ") AS order_qty, SUM(" & f_moulinette_aggreg_order_redi_ExecQty & ") AS exec_qty, SUM(" & f_moulinette_aggreg_order_redi_ExecQty & "*" & f_moulinette_aggreg_order_redi_AvgExecPrice & ") AS ntcf, SUM(" & f_moulinette_aggreg_order_redi_OrderQty & "*" & f_moulinette_aggreg_order_redi_OrderPrice & ")/SUM(" & f_moulinette_aggreg_order_redi_OrderQty & ")" & " AS " & f_moulinette_stat_ticker_AVGOrderPrice & ", SUM(" & f_moulinette_aggreg_order_redi_ExecQty & "*" & f_moulinette_aggreg_order_redi_AvgExecPrice & ")/SUM(" & f_moulinette_aggreg_order_redi_ExecQty & ")" & " AS exec_avg_price"
    sql_query = sql_query & " FROM " & v_moulinette_aggreg_order_redi & ", " & t_moulinette_bridge_redi & ", " & t_moulinette_order_xls
    sql_query = sql_query & " WHERE " & f_moulinette_aggreg_order_redi_BranchSequence & "=" & f_moulinette_bridge_redi_BranchSequence
    sql_query = sql_query & " AND " & f_moulinette_aggreg_order_redi_symbol & "=" & f_moulinette_bridge_redi_SymbolRedi
    sql_query = sql_query & " AND " & f_moulinette_order_xls_id & "=" & f_moulinette_bridge_redi_internal_id
    sql_query = sql_query & " AND " & f_moulinette_bridge_redi_SymbolXLS & "=" & f_moulinette_order_xls_symbol_redi
    
    ' #######################################################################################
    sql_query = sql_query & " AND " & f_moulinette_order_xls_ticker & "=""" & ticker & """"
     ' #######################################################################################
    
    sql_query = sql_query & " GROUP BY " & f_moulinette_order_xls_ticker

Dim extract_stat_ticker As Variant
extract_stat_ticker = sqlite3_query(moulinette_get_db_complete_path, sql_query)


If UBound(extract_stat_ticker, 1) = 0 Then
    moulinette = "#N/A Ticker"
    Exit Function
Else
    For j = 0 To UBound(extract_stat_ticker(0), 1)
        
        If UCase(extract_stat_ticker(0)(j)) = UCase(Replace(field, " ", "_")) Then
            moulinette = extract_stat_ticker(1)(j)
            Exit Function
        Else
            If j = UBound(extract_stat_ticker(0), 1) Then
                moulinette = "#N/A field"
                Exit Function
            End If
        End If
    Next j
End If


End Function



Private Sub store_in_excel_orders_from_db()

Application.Calculation = xlCalculationManual

extract_xls_order = sqlite3_query(moulinette_get_db_complete_path, "SELECT * FROM " & t_moulinette_order_xls)

For i = 1 To UBound(extract_xls_order, 1)
    For j = 0 To UBound(extract_xls_order(i), 1)
        Workbooks("Kronos.xls").Worksheets(sheet_offline).Cells(i, c_offline_order_xls_id + j) = extract_xls_order(i)(j)
    Next j
Next i


End Sub


' symbol / qty
Private Sub generate_ticket(ByVal vec_ticket)

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer


Dim oRediTicket As New RediLib.Ticket

dim_symbol = 0
dim_qty = 1



With oRediTicket
    
    Dim err_data As Variant
    
    For i = 0 To UBound(vec_ticket, 1)
        
        .symbol = get_symbol_redi_plus(vec_ticket(i)(dim_symbol))
        
        If vec_ticket(i)(dim_qty) < 0 Then
            .side = "sell"
        Else
            .side = "buy"
        End If
        
        .quantity = Abs(vec_ticket(i)(dim_qty))
        
        retValueOrder = .Submit(myerr)
        
    Next i
    
End With


End Sub


Sub get_redi_all()

Application.Calculation = xlCalculationManual

Worksheets("log_msg").Cells.Clear

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim vtable As Variant
Dim vwhere As Variant
Dim verr As Variant

If IsRediReady Then
    
        If ThisWorkbook.OrderQuery Is Nothing Then
            Set ThisWorkbook.OrderQuery = New RediLib.CacheControl
        End If
        
        ThisWorkbook.OrderQuery.UserID = ""
        ThisWorkbook.OrderQuery.Password = ""
        vtable = "Message"
        vwhere = "true"
        
        MessageQuery = ThisWorkbook.OrderQuery.Submit(vtable, vwhere, verr)
        
        ThisWorkbook.OrderQuery.Revoke verr
    
    'get_redi_orders = ThisWorkbook.RediOrders
    Dim extract_msg_table As Variant
    extract_msg_table = ThisWorkbook.RediMsg
    
    
    
    For i = 0 To UBound(extract_msg_table, 1)
        For j = 0 To UBound(extract_msg_table(i), 1)
            Worksheets("log_msg").Cells(i + 1, j + 1) = extract_msg_table(i)(j)
        Next j
    Next i
    
    
    Worksheets("log_msg").rows(1).AutoFilter
    
End If

End Sub


Public Sub test_universal_trade_big_order_with_cancel()

Dim oJSON As New JSONLib

Dim oBBG As New cls_Bloomberg_Sync

Dim data_bbg As Variant
data_bbg = oBBG.bdp(Array("UBSN VX EQUITY", "FP FP EQUITY", "LOGN VX EQUITY"), Array("px_last"), output_format.of_vec_without_header)


Dim debug_test As Variant
    
    Randomize
    tmp_group = CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0))
debug_test = universal_trades_r_plus(Array(Array("FP FP EQUITY", 12000, Round(data_bbg(1)(0), 2), Empty, tmp_group, oJSON.toString(Array("base")))))

End Sub


Public Sub test_universal_trades_r_plus()

Call generate_ticket(Array(Array("UBSN VX EQUITY", 150000), Array("UBSN VX EQUITY", -150000), Array("FP FP EQUITY", 150000), Array("FP FP EQUITY", -150000), Array("LOGN VX EQUITY", 150000), Array("LOGN VX EQUITY", -150000)))

Dim oJSON As New JSONLib

Dim oBBG As New cls_Bloomberg_Sync

Dim data_bbg As Variant
data_bbg = oBBG.bdp(Array("UBSN VX EQUITY", "FP FP EQUITY", "LOGN VX EQUITY"), Array("px_last"), output_format.of_vec_without_header)


Dim debug_test As Variant
    
    Randomize
    tmp_group = CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0))
debug_test = universal_trades_r_plus(Array(Array("UBSN VX EQUITY", 100, Round(0.999 * data_bbg(0)(0), 2), Empty, tmp_group, oJSON.toString(Array("base"))), Array("UBSN VX EQUITY", -200, Round(1.001 * data_bbg(0)(0), 2), Empty, tmp_group, oJSON.toString(Array("blub"))), Array("UBSN VX EQUITY", -150, Round(1.001 * data_bbg(0)(0), 2), Round(1.001 * data_bbg(0)(0), 2), tmp_group, oJSON.toString(Array("stop")))))
    Randomize
    tmp_group = CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0))
debug_test = universal_trades_r_plus(Array(Array("UBSN VX EQUITY", 100, Round(0.999 * data_bbg(0)(0), 2), Empty, tmp_group, oJSON.toString(Array("base"))), Array("UBSN VX EQUITY", -200, Round(1.001 * data_bbg(0)(0), 2), Empty, tmp_group, oJSON.toString(Array("blub"))), Array("UBSN VX EQUITY", -150, Round(1.001 * data_bbg(0)(0), 2), Round(1.001 * data_bbg(0)(0), 2), tmp_group + 5, oJSON.toString(Array("base")))))
    Randomize
    tmp_group = CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0))
debug_test = universal_trades_r_plus(Array(Array("FP FP EQUITY", 50, Round(0.999 * data_bbg(1)(0), 2), Empty, tmp_group, oJSON.toString(Array("base"))), Array("FP FP EQUITY", -200, Round(1.001 * data_bbg(1)(0), 2), Empty, tmp_group, oJSON.toString(Array("blub"))), Array("FP FP EQUITY", -150, Round(1.001 * data_bbg(1)(0), 2), Round(1.001 * data_bbg(1)(0), 2), tmp_group, oJSON.toString(Array("stop")))))
    Randomize
    tmp_group = CDbl(Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & Round(100 * Rnd(), 0))
debug_test = universal_trades_r_plus(Array(Array("LOGN VX EQUITY", 150, Round(0.999 * data_bbg(2)(0), 2), Empty, tmp_group, oJSON.toString(Array("base"))), Array("LOGN VX EQUITY", -1250, Round(1.001 * data_bbg(2)(0), 2), Empty, tmp_group, oJSON.toString(Array("stop")))))

End Sub



