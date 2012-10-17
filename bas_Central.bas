Attribute VB_Name = "bas_Central"
Public Const db_central As String = "db_central.sqlt3"
    Public Const t_central_rank As String = "t_central_rank"
        Public Const f_central_rank_id As String = "f_central_rank_id"
        Public Const f_central_rank_bbg_field As String = "f_central_rank_bbg_field"
        Public Const f_central_rank_optional_name As String = "f_central_rank_optional_name"
        Public Const f_central_rank_order As String = "f_central_rank_order"
        Public Const f_central_rank_rank_if_not_available As String = "f_central_rank_rank_if_not_available"
        Public Const f_central_rank_weight As String = "f_central_rank_weight"
    
    Public Const t_central_data_bbg As String = "t_central_data_bbg"
        Public Const f_central_data_bbg_ticker As String = "f_central_data_bbg_ticker"
    
    Public Const t_central_monitor_field As String = "t_central_monitor_field"
        Public Const f_central_monitor_field_bbg_id As String = "f_central_monitor_field_bbg_id"
        Public Const f_central_monitor_field_db_id As String = "f_central_monitor_field_db_id"
        Public Const f_central_monitor_field_last_update_date As String = "f_central_monitor_field_last_update_date"
    
    Public Const t_central_store_rank As String = "t_central_store_rank"
        Public Const f_central_store_rank_ticker As String = "f_central_store_rank_ticker"
        Public Const f_central_store_rank_import_date As String = "f_central_store_rank_import_date"
        Public Const f_central_store_rank_id As String = "f_central_store_rank_id"
        Public Const f_central_store_rank_value As String = "f_central_store_rank_value"
    
    Public Const t_central_helper As String = "t_central_helper"
        
        Public Const central_helper_nbre_text_fields As Integer = 20
        Public Const central_helper_nbre_numeric_fields As Integer = 20
        
        Public Const f_central_helper_text1 As String = "f_central_helper_text1"
        Public Const f_central_helper_text2 As String = "f_central_helper_text2"
        Public Const f_central_helper_text3 As String = "f_central_helper_text3"
        Public Const f_central_helper_numeric1 As String = "f_central_helper_numeric1"
        Public Const f_central_helper_numeric2 As String = "f_central_helper_numeric2"
        Public Const f_central_helper_numeric3 As String = "f_central_helper_numeric3"


Public Enum central_order_rank
    small_is_best = 0
    big_is_best = 1
End Enum


Public Function central_get_db_fullpath() As String

central_get_db_fullpath = ActiveWorkbook.path & "\" & db_central

End Function


Private Sub central_manip_db()

Dim sql_query As String
'sql_query = "DROP TABLE " & t_central_data_bbg
'exec_query = sqlite3_query(central_get_db_fullpath, sql_query)
'
'sql_query = "DROP TABLE " & t_central_monitor_field
'exec_query = sqlite3_query(central_get_db_fullpath, sql_query)
'
'sql_query = "DROP TABLE " & t_central_helper
'exec_query = sqlite3_query(central_get_db_fullpath, sql_query)

debug_test = sqlite3_query(central_get_db_fullpath, "SELECT * FROM " & t_central_data_bbg)
exec_query = sqlite3_query(central_get_db_fullpath, "UPDATE " & t_central_data_bbg & " SET bbg_PE_RATIO=1")
exec_query = sqlite3_query(central_get_db_fullpath, "UPDATE " & t_central_data_bbg & " SET bbg_PX_TO_BOOK_RATIO=2")
debug_test = sqlite3_query(central_get_db_fullpath, "SELECT * FROM " & t_central_data_bbg)


End Sub


Private Sub central_init_db()

Dim i As Integer, j As Integer, k As Integer

Dim sql_query As String
Dim create_table_query As String
Dim exec_query As Variant

Dim init_db_status As Variant
init_db_status = sqlite3_create_db(central_get_db_fullpath)

Dim create_table_status As Variant

If sqlite3_check_if_table_already_exist(central_get_db_fullpath, t_central_rank) = False Then
    create_table_query = sqlite3_get_query_create_table(t_central_rank, Array(Array(f_central_rank_id, "TEXT", ""), Array(f_central_rank_bbg_field, "TEXT", ""), Array(f_central_rank_optional_name, "TEXT", ""), Array(f_central_rank_order, "INTEGER", ""), Array(f_central_rank_rank_if_not_available, "REAL", ""), Array(f_central_rank_weight, "REAL", "")), Array(Array(f_central_rank_id, "ASC"), Array(f_central_rank_bbg_field, "ASC")))
    create_table_status = sqlite3_create_tables(central_get_db_fullpath, Array(create_table_query))
End If


If sqlite3_check_if_table_already_exist(central_get_db_fullpath, t_central_data_bbg) = False Then
    create_table_query = sqlite3_get_query_create_table(t_central_data_bbg, Array(Array(f_central_data_bbg_ticker, "TEXT", "")), Array(Array(f_central_data_bbg_ticker, "ASC")))
    create_table_status = sqlite3_create_tables(central_get_db_fullpath, Array(create_table_query))
End If


If sqlite3_check_if_table_already_exist(central_get_db_fullpath, t_central_monitor_field) = False Then
    create_table_query = sqlite3_get_query_create_table(t_central_monitor_field, Array(Array(f_central_monitor_field_bbg_id, "TEXT", ""), Array(f_central_monitor_field_db_id, "TEXT", ""), Array(f_central_monitor_field_last_update_date, "NUMERIC", "")), Array(Array(f_central_monitor_field_bbg_id, "ASC")))
    create_table_status = sqlite3_create_tables(central_get_db_fullpath, Array(create_table_query))
End If


If sqlite3_check_if_table_already_exist(central_get_db_fullpath, t_central_store_rank) = False Then
    create_table_query = sqlite3_get_query_create_table(t_central_store_rank, Array(Array(f_central_store_rank_ticker, "TEXT", ""), Array(f_central_store_rank_import_date, "NUMERIC", ""), Array(f_central_store_rank_id, "TEXT", ""), Array(f_central_store_rank_value, "REAL", "")), Array(Array(f_central_store_rank_ticker, "ASC"), Array(f_central_store_rank_import_date, "DESC"), Array(f_central_store_rank_id, "ASC")))
    create_table_status = sqlite3_create_tables(central_get_db_fullpath, Array(create_table_query))
End If


If sqlite3_check_if_table_already_exist(central_get_db_fullpath, t_central_helper) = False Then
    
    Dim vec_fields()
    k = 0
    For i = 1 To central_helper_nbre_text_fields
        ReDim Preserve vec_fields(k)
        vec_fields(k) = Array("f_central_helper_text" & i, "TEXT", "")
        k = k + 1
    Next i
    
    For i = 1 To central_helper_nbre_numeric_fields
        ReDim Preserve vec_fields(k)
        vec_fields(k) = Array("f_central_helper_numeric" & i, "NUMERIC", "")
        k = k + 1
    Next i
    
    create_table_query = sqlite3_get_query_create_table(t_central_helper, vec_fields)
    create_table_status = sqlite3_create_tables(central_get_db_fullpath, Array(create_table_query))
End If

'wash vielleries

'sql_query = "DELETE FROM " & t_central_data_bbg & " WHERE " & f_central_data_bbg_import_date & "<" & ToJulianDay(Date)
'exec_query = sqlite3_query(central_get_db_fullpath, sql_query)

End Sub


Public Function central_get_ticker_rank() As Variant

'central_get_ticker_rank = Array("AAPL US EQUITY", "GOOG US EQUITY", "FP FP EQUITY")
sql_query = "SELECT Ticker FROM t_custom_rank ORDER BY Ticker ASC"

Dim extract_rank As Variant
extract_rank = central_query_on_ranksqlt3(sql_query)

Dim vec_ticker() As Variant

For i = 1 To UBound(extract_rank, 1)
    ReDim Preserve vec_ticker(i - 1)
    vec_ticker(i - 1) = extract_rank(i)(0)
Next i

central_get_ticker_rank = vec_ticker

End Function


Private Function central_get_potential_path_rank() As Variant

Dim db_rank As String
    db_rank = "rank.sqlt3"

central_get_potential_path_rank = Array("q:\front\stouff\rsse\" & db_rank, ActiveWorkbook.path & "\" & db_rank)

End Function


Private Function central_get_path_rank() As String

Dim tmp_paths As Variant
tmp_paths = central_get_potential_path_rank


For i = 0 To UBound(tmp_paths, 1)
    If exist_file(tmp_paths(i)) Then
        central_get_path_rank = tmp_paths(i)
        Exit Function
    End If
Next i

End Function


Private Function central_query_on_ranksqlt3(ByVal sql_query As String) As Variant

central_query_on_ranksqlt3 = sqlite3_query(central_get_path_rank, sql_query)

End Function


Private Sub test_central_create_custom_rank()

debug_test = central_create_custom_rank("test_value", Array(Array("PE_RATIO", "txt_PE_RATIO", central_order_rank.small_is_best, 50, 75), Array("PX_TO_BOOK_RATIO", "txt_PX_TO_BOOK_RATIO", central_order_rank.big_is_best, 45, 25)))

End Sub


'vec_fields: 0 bbg field | 1 optional name | 2 order | 3 rank if not available | 4 weight
Public Function central_create_custom_rank(ByVal rank_name As String, ByVal vec_fields As Variant) As Variant

Call central_init_db

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
Dim sql_query As String
Dim exec_query As Variant

'check if already avaialbe
sql_query = "SELECT DISTINCT " & f_central_rank_id & " FROM " & t_central_rank & " WHERE " & f_central_rank_id & "=""" & rank_name & """"
Dim extract_distinct_rank_name As Variant
extract_distinct_rank_name = sqlite3_query(central_get_db_fullpath, sql_query)

If UBound(extract_distinct_rank_name, 1) > 0 Then
    vbanswer = MsgBox("Rank already existing in the DB. Erase with this new set up ?", vbYesNo)
    
    If vbanswer = vbYes Then
        sql_query = "DELETE FROM " & t_central_rank & " WHERE " & f_central_rank_id & "=""" & rank_name & """"
        exec_query = sqlite3_query(central_get_db_fullpath, sql_query)
        'debug_test = sqlite3_query(central_get_db_fullpath, "SELECT DISTINCT " & f_central_rank_id & " FROM " & t_central_rank & " WHERE " & f_central_rank_id & "=""" & rank_name & """")
    Else
        MsgBox ("-> Exit.")
        Exit Function
    End If
    
End If

Dim output_for_db() As Variant
Dim tmp_row() As Variant
k = 0
For i = 0 To UBound(vec_fields, 1)
    ReDim Preserve tmp_row(0)
    tmp_row(0) = rank_name
    
    For j = 0 To UBound(vec_fields(i), 1)
        ReDim Preserve tmp_row(j + 1)
        tmp_row(j + 1) = vec_fields(i)(j)
    Next j
    
    ReDim Preserve output_for_db(i)
    output_for_db(i) = tmp_row
    k = k + 1
Next i

If k > 0 Then
    insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_rank, output_for_db, Array(f_central_rank_id, f_central_rank_bbg_field, f_central_rank_optional_name, f_central_rank_order, f_central_rank_rank_if_not_available, f_central_rank_weight))
End If

central_create_custom_rank = sqlite3_query(central_get_db_fullpath, "SELECT " & f_central_rank_bbg_field & ", " & f_central_rank_optional_name & ", " & f_central_rank_order & ", " & f_central_rank_rank_if_not_available & ", " & f_central_rank_weight & " FROM " & t_central_rank & " WHERE " & f_central_rank_id & "=""" & rank_name & """")

End Function


Public Function central_get_compatible_sql_field_name(ByVal bbg_field As String) As String

bbg_field = Replace(bbg_field, "%", "pct")
bbg_field = Replace(bbg_field, ".", "_")
bbg_field = Replace(bbg_field, " ", "_")

central_get_compatible_sql_field_name = "bbg_" & UCase(bbg_field)

End Function


Private Sub test_central_load_rank()

debug_test = central_load_rank("test_value", central_get_ticker_rank)

End Sub


Public Function central_load_rank(ByVal rank_name As String, ByVal vec_ticker As Variant) As Variant

Dim vec_alert As Variant
vec_alert = Array(3, 22, 6, 19, 36, 35, 43, 4) 'small is worst

Call central_init_db

Dim oBBG As New cls_Bloomberg_Sync
Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim sql_query As String

Dim tmp_row() As Variant, tmp_column() As Variant


'construct interval vec_alert
For i = 0 To UBound(vec_alert, 1)
    vec_alert(i) = Array(i * (100 / (UBound(vec_alert, 1) + 1)), (i + 1) * (100 / (UBound(vec_alert, 1) + 1)), vec_alert(i))
Next i


'control que rank existe bien
sql_query = "SELECT DISTINCT " & f_central_rank_id & " FROM " & t_central_rank & " WHERE " & f_central_rank_id & "=""" & rank_name & """"
Dim extract_check_rank As Variant
extract_check_rank = sqlite3_query(central_get_db_fullpath, sql_query)

If UBound(extract_check_rank, 1) = 0 Then
    MsgBox ("Problem with DB, rank: " & rank_name & " not found")
    Exit Function
End If

'charge le rank
sql_query = "SELECT * FROM " & t_central_rank & " WHERE " & f_central_rank_id & "=""" & rank_name & """"
Dim extract_rank_composition As Variant
extract_rank_composition = sqlite3_query(central_get_db_fullpath, sql_query)

For i = 0 To UBound(extract_rank_composition(0), 1)
    If extract_rank_composition(0)(i) = f_central_rank_bbg_field Then
        dim_rank_bbg_field = i
    ElseIf extract_rank_composition(0)(i) = f_central_rank_optional_name Then
        dim_rank_optional_name = i
    ElseIf extract_rank_composition(0)(i) = f_central_rank_order Then
        dim_rank_order = i
    ElseIf extract_rank_composition(0)(i) = f_central_rank_rank_if_not_available Then
        dim_rank_rank_if_not_available = i
    ElseIf extract_rank_composition(0)(i) = f_central_rank_weight Then
        dim_rank_weight = i
    End If
Next i


'normalisation des poids sur 1
Dim sum_weight As Double
sum_weight = 0
For i = 1 To UBound(extract_rank_composition, 1)
    sum_weight = sum_weight + extract_rank_composition(i)(dim_rank_weight)
Next i

    For i = 1 To UBound(extract_rank_composition, 1)
        extract_rank_composition(i)(dim_rank_weight) = extract_rank_composition(i)(dim_rank_weight) / sum_weight
    Next i




'charge les champs bbg
Dim vec_bbg_fields() As Variant
For i = 1 To UBound(extract_rank_composition, 1)
    ReDim Preserve vec_bbg_fields(i - 1)
    vec_bbg_fields(i - 1) = extract_rank_composition(i)(dim_rank_bbg_field)
Next i

    'inject dans le helper
    Dim data_helper_fields() As Variant
    For i = 0 To UBound(vec_bbg_fields, 1)
        ReDim Preserve data_helper_fields(i)
        data_helper_fields(i) = Array(vec_bbg_fields(i))
    Next i
    
    exec_status = sqlite3_query(central_get_db_fullpath, "DELETE FROM " & t_central_helper)
    insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_helper, data_helper_fields, Array(f_central_helper_text1))
    'debug_test = sqlite3_query(central_get_db_fullpath, "SELECT " & f_central_helper_text1 & " FROM " & t_central_helper)
    

'update table if missing fields
Dim extract_current_field_bbg_data As Variant
extract_current_field_bbg_data = sqlite3_get_table_structure(central_get_db_fullpath, t_central_data_bbg)

k = 0
For i = 0 To UBound(vec_bbg_fields, 1)
    For j = 1 To UBound(extract_current_field_bbg_data, 1)
        If central_get_compatible_sql_field_name(vec_bbg_fields(i)) = extract_current_field_bbg_data(j)(1) Then
            Exit For
        Else
            If j = UBound(extract_current_field_bbg_data, 1) Then
                'missing field
                sql_query = "ALTER TABLE " & t_central_data_bbg & " ADD COLUMN " & central_get_compatible_sql_field_name(vec_bbg_fields(i)) & " NUMERIC"
                exec_query = sqlite3_query(central_get_db_fullpath, sql_query)
                
                'ajoute egalement dans le bridge
                Dim data_field_bridge() As Variant
                ReDim Preserve data_field_bridge(k)
                data_field_bridge(k) = Array(vec_bbg_fields(i), central_get_compatible_sql_field_name(vec_bbg_fields(i)), ToJulianDay(Date - 1))
                k = k + 1
            End If
        End If
    Next j
Next i

If k > 0 Then
    insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_monitor_field, data_field_bridge, Array(f_central_monitor_field_bbg_id, f_central_monitor_field_db_id, f_central_monitor_field_last_update_date))
    'debug_test = sqlite3_query(central_get_db_fullpath, "SELECT * FROM " & t_central_monitor_field)
End If


'charge le bridge de field
Dim extract_bridge_field As Variant
extract_bridge_field = sqlite3_query(central_get_db_fullpath, "SELECT " & f_central_monitor_field_bbg_id & ", " & f_central_monitor_field_db_id & ", " & f_central_monitor_field_last_update_date & " FROM " & t_central_monitor_field)

    For i = 0 To UBound(extract_bridge_field(0), 1)
        If extract_bridge_field(0)(i) = f_central_monitor_field_bbg_id Then
            dim_bridge_bbg = i
        ElseIf extract_bridge_field(0)(i) = f_central_monitor_field_db_id Then
            dim_bridge_db = i
        ElseIf extract_bridge_field(0)(i) = f_central_monitor_field_last_update_date Then
            dim_bridge_last_update_date = i
        End If
    Next i
    
    'tranforme en date les date sqlite
    For i = 1 To UBound(extract_bridge_field, 1)
        extract_bridge_field(i)(dim_bridge_last_update_date) = FromJulianDay(CDbl(extract_bridge_field(i)(dim_bridge_last_update_date)))
    Next i
    


'repere les tickers manquants et complete * field bbg
Dim data_helper_ticker() As Variant
k = 0
For i = 0 To UBound(vec_ticker, 1)
    ReDim Preserve data_helper_ticker(i)
    data_helper_ticker(i) = Array(vec_ticker(i))
Next i

    exec_status = sqlite3_query(central_get_db_fullpath, "DELETE FROM " & t_central_helper)
    insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_helper, data_helper_ticker, Array(f_central_helper_text1))
    
    sql_query = "SELECT " & f_central_helper_text1 & " FROM " & t_central_helper & " WHERE " & f_central_helper_text1 & " NOT IN ("
            sql_query = sql_query & "SELECT " & f_central_data_bbg_ticker & " FROM " & t_central_data_bbg
        sql_query = sql_query & ")"
    Dim extract_missing_ticker As Variant
    extract_missing_ticker = sqlite3_query(central_get_db_fullpath, sql_query)
    
    If UBound(extract_missing_ticker, 1) > 0 Then
        
        Dim vec_new_ticker() As Variant
        For i = 1 To UBound(extract_missing_ticker, 1)
            ReDim Preserve vec_new_ticker(i - 1)
            vec_new_ticker(i - 1) = extract_missing_ticker(i)(0)
        Next i
        
        
        'repere les champs bbg grace au champs de la structure de la table
        extract_current_field_bbg_data = sqlite3_get_table_structure(central_get_db_fullpath, t_central_data_bbg)
        
        Dim vec_field_for_new_ticker() As Variant
        k = 0
        For i = 1 To UBound(extract_current_field_bbg_data, 1)
            If extract_current_field_bbg_data(i)(1) <> f_central_data_bbg_ticker And extract_current_field_bbg_data(i)(1) <> f_central_data_bbg_import_date Then
                
                For j = 1 To UBound(extract_bridge_field, 1)
                    If extract_current_field_bbg_data(i)(1) = extract_bridge_field(j)(dim_bridge_db) Then
                        ReDim Preserve vec_field_for_new_ticker(k)
                        vec_field_for_new_ticker(k) = extract_bridge_field(j)(dim_bridge_bbg)
                        k = k + 1
                        Exit For
                    End If
                Next j
                
            End If
        Next i
        
        If k > 0 Then
            
            Dim data_bbg_new_ticker As Variant
            data_bbg_new_ticker = oBBG.bdp(vec_new_ticker, vec_field_for_new_ticker, output_format.of_vec_without_header)
            
            'insertion des datas pour les nouveaux tickers
            Dim data_db_new_ticker() As Variant
            For i = 0 To UBound(vec_new_ticker, 1)
                
                ReDim Preserve tmp_row(0)
                tmp_row(0) = vec_new_ticker(i)
                
                For j = 0 To UBound(vec_field_for_new_ticker, 1)
                    ReDim Preserve tmp_row(j + 1)
                    
                    If IsNumeric(data_bbg_new_ticker(i)(j)) Then
                        tmp_row(j + 1) = data_bbg_new_ticker(i)(j)
                    Else
                        tmp_row(j + 1) = Empty
                    End If
                    
                Next j
                
                ReDim Preserve data_db_new_ticker(i)
                data_db_new_ticker(i) = tmp_row
                
            Next i
            
            Dim field_db_new_ticker()
            ReDim Preserve field_db_new_ticker(0)
            field_db_new_ticker(0) = f_central_data_bbg_ticker
            For i = 0 To UBound(vec_field_for_new_ticker, 1)
                ReDim Preserve field_db_new_ticker(i + 1)
                field_db_new_ticker(i + 1) = central_get_compatible_sql_field_name(vec_field_for_new_ticker(i))
            Next i
            
            insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_data_bbg, data_db_new_ticker, field_db_new_ticker)
            
            debug_test = sqlite3_query(central_get_db_fullpath, "SELECT * FROM " & t_central_data_bbg & " ORDER BY " & f_central_data_bbg_ticker & " ASC")
        End If
        
    End If
    
    
    
'mise a jour des champs ne datant pas du jour
Dim vec_bbg_field_need_update() As Variant
k = 0
For i = 0 To UBound(vec_bbg_fields, 1)
    
    For j = 1 To UBound(extract_bridge_field, 1)
        If vec_bbg_fields(i) = extract_bridge_field(j)(dim_bridge_bbg) Then
            If extract_bridge_field(j)(dim_bridge_last_update_date) < Date Then
                ReDim Preserve vec_bbg_field_need_update(k)
                vec_bbg_field_need_update(k) = vec_bbg_fields(i)
                k = k + 1
            End If
        End If
    Next j
    
Next i


If k > 0 Then
    
    Dim vec_ticker_need_update() As Variant
    
    'remonte * tickers de bbg_data
    sql_query = "SELECT " & f_central_data_bbg_ticker & " FROM " & t_central_data_bbg
    Dim extract_ticker_data_bbg As Variant
    extract_ticker_data_bbg = sqlite3_query(central_get_db_fullpath, sql_query)
    
    For i = 1 To UBound(extract_ticker_data_bbg, 1)
        ReDim Preserve vec_ticker_need_update(i - 1)
        vec_ticker_need_update(i - 1) = extract_ticker_data_bbg(i)(0)
    Next i
    
    
    data_bbg_need_update = oBBG.bdp(vec_ticker_need_update, vec_bbg_field_need_update, output_format.of_vec_without_header)
    
    
    
    Dim tmp_queries_for_one_ticker As String
    Dim vec_sql_queries() As Variant
    k = 0
    For i = 0 To UBound(vec_ticker_need_update, 1)
        
        m = 0
        tmp_queries_for_one_ticker = ""
        For j = 0 To UBound(vec_bbg_field_need_update, 1)
            
            If IsNumeric(data_bbg_need_update(i)(j)) Then
                
                If m = 0 Then
                    tmp_queries_for_one_ticker = "UPDATE " & t_central_data_bbg & " SET "
                Else
                    tmp_queries_for_one_ticker = tmp_queries_for_one_ticker & ", "
                End If
                
                tmp_queries_for_one_ticker = tmp_queries_for_one_ticker & central_get_compatible_sql_field_name(vec_bbg_field_need_update(j)) & "=" & data_bbg_need_update(i)(j)
                m = m + 1
            End If
            
        Next j
        
        If m > 0 Then
            
            tmp_queries_for_one_ticker = tmp_queries_for_one_ticker & " WHERE " & f_central_data_bbg_ticker & "=""" & vec_ticker_need_update(i) & """"
            
            ReDim Preserve vec_sql_queries(k)
            vec_sql_queries(k) = tmp_queries_for_one_ticker
            k = k + 1
        End If
        
    Next i
    
    If k > 0 Then
        db_data_bbg_new_state = central_update_db_data_bbg(vec_sql_queries)
        
        'update de la date des champs
        exec_status = sqlite3_query(central_get_db_fullpath, "DELETE FROM " & t_central_helper)
        insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_helper, data_helper_fields, Array(f_central_helper_text1))
        'debug_test = sqlite3_query(central_get_db_fullpath, "SELECT " & f_central_helper_text1 & " FROM " & t_central_helper)
        
        sql_query = "UPDATE " & t_central_monitor_field & " SET " & f_central_monitor_field_last_update_date & "=" & ToJulianDay(Date) & " WHERE " & f_central_monitor_field_bbg_id & " IN ("
                sql_query = sql_query & "SELECT " & f_central_helper_text1 & " FROM " & t_central_helper
            sql_query = sql_query & ")"
        exec_query = sqlite3_query(central_get_db_fullpath, sql_query)
        'debug_test = sqlite3_query(central_get_db_fullpath, "SELECT " & f_central_monitor_field_bbg_id & ", " & f_central_monitor_field_db_id & ", date(" & f_central_monitor_field_last_update_date & ") FROM " & t_central_monitor_field)
        
        
    End If
    
    
    
End If




'extraction des donnees necessaire a la matrix de rank
sql_query = "SELECT " & f_central_data_bbg_ticker
For i = 0 To UBound(vec_bbg_fields, 1)
    sql_query = sql_query & ", " & central_get_compatible_sql_field_name(vec_bbg_fields(i))
Next i
    
    'rajoute dummy field qui contiendra le ranking final
    sql_query = sql_query & ", " & f_central_data_bbg_ticker & " AS final_weighted_rank"
    
    sql_query = sql_query & " FROM " & t_central_data_bbg
    
    sql_query = sql_query & " WHERE " & f_central_data_bbg_ticker & " IN (SELECT " & f_central_helper_text1 & " FROM " & t_central_helper & ")"
    
    sql_query = sql_query & " ORDER BY " & f_central_data_bbg_ticker & " ASC"


Dim extract_data_for_ranking As Variant
extract_data_for_ranking = sqlite3_query(central_get_db_fullpath, sql_query)


Dim data_ranked As Variant
data_ranked = extract_data_for_ranking

Dim min_max_value As Double
Dim min_max_pos As Long

For i = 1 To UBound(extract_data_for_ranking(0), 1) - 1 'saute final rank
    
    For m = 1 To UBound(extract_rank_composition, 1)
        
        If extract_rank_composition(m)(dim_rank_bbg_field) = vec_bbg_fields(i - 1) Then
    
            k = 0
            For j = 1 To UBound(extract_data_for_ranking, 1)
                
                If IsNull(extract_data_for_ranking(j)(i)) = False Then
                    ReDim Preserve tmp_column(k)
                    tmp_column(k) = extract_data_for_ranking(j)(i)
                    k = k + 1
                Else
                    data_ranked(j)(i) = extract_rank_composition(m)(dim_rank_rank_if_not_available)
                End If
                
            Next j
            
            If k > 0 Then
                
                'on sort le vecteur
                For p = 0 To UBound(tmp_column, 1)
                    
                    min_max_pos = p
                    min_max_value = tmp_column(p)
                    
                    For q = p + 1 To UBound(tmp_column, 1)
                        
                        If extract_rank_composition(m)(dim_rank_order) = central_order_rank.big_is_best Then
                            If tmp_column(q) < min_max_value Then
                                min_max_value = tmp_column(q)
                                min_max_pos = q
                            End If
                        ElseIf extract_rank_composition(m)(dim_rank_order) = central_order_rank.small_is_best Then
                            If tmp_column(q) > min_max_value Then
                                min_max_value = tmp_column(q)
                                min_max_pos = q
                            End If
                        End If
                        
                    Next q
                    
                    
                    If min_max_pos <> p Then
                        min_max_value = tmp_column(p)
                        tmp_column(p) = tmp_column(min_max_pos)
                        tmp_column(min_max_pos) = min_max_value
                    End If
                    
                Next p
                
                
                'redonne a chaque titre sa note
                For j = 1 To UBound(extract_data_for_ranking, 1)
                    For p = 0 To UBound(tmp_column, 1)
                        If extract_data_for_ranking(j)(i) = tmp_column(p) Then
                            data_ranked(j)(i) = p * (100 / (UBound(extract_data_for_ranking, 1) - 1))
                            'pas d exit for si meme donnee plus loin
                        End If
                    Next p
                Next j
                
                
            Else
                
            End If
            
            Exit For
        End If
    Next m
    
Next i


'calc final rank
Dim final_rank_ticker As Double

For j = 1 To UBound(extract_data_for_ranking, 1) 'boucle ticker
    
    final_rank_ticker = 0
    For p = 1 To UBound(extract_data_for_ranking(0), 1) - 1 'boucle column with data ranked
        For m = 1 To UBound(extract_rank_composition, 1)
            If extract_rank_composition(m)(dim_rank_bbg_field) = vec_bbg_fields(p - 1) Then
                final_rank_ticker = final_rank_ticker + extract_rank_composition(m)(dim_rank_weight) * data_ranked(j)(p)
                data_ranked(j)(p) = Round(data_ranked(j)(p), 0)
                Exit For
            End If
        Next m
    Next p
    
    data_ranked(j)(UBound(data_ranked(j), 1)) = Round(final_rank_ticker, 0)
    
Next j


'store result in db (api call)
Dim store_data_rank() As Variant

For i = 1 To UBound(data_ranked, 1)
    ReDim Preserve store_data_rank(i - 1)
    store_data_rank(i - 1) = Array(data_ranked(i)(0), ToJulianDay(Date), rank_name, data_ranked(i)(UBound(data_ranked(i), 1)))
Next i
    
    exec_query = "DELETE FROM " & t_central_store_rank & " WHERE " & f_central_store_rank_id & "=""" & rank_name & """"
    insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_store_rank, store_data_rank, Array(f_central_store_rank_ticker, f_central_store_rank_import_date, f_central_store_rank_id, f_central_store_rank_value))
    'debug_test = sqlite3_query(central_get_db_fullpath, "SELECT * FROM " & t_central_store_rank)



'merge data with rank.sqlt3
Dim field_rank() As Variant
    field_rank = Array(Array("NAME", "TEXT"), Array("CRNCY", "TEXT"), Array("GICS_SECTOR_NAME", "TEXT"), Array("GICS_INDUSTRY_NAME", "TEXT"), Array("Rank_EPS_4w_chg_curr_yr", "NUMERIC"), Array("Rank_EPS_4w_chg_nxt_yr", "NUMERIC"), Array("Rank_MoneyFlow", "NUMERIC"), Array("Rank_GEO_GROWTH_5YR_EPS", "NUMERIC"), Array("Rank_R2_5YR_EPS", "NUMERIC"), Array("Rank_EPS", "NUMERIC"))

sql_query = "SELECT Ticker"
    For i = 0 To UBound(field_rank, 1)
        sql_query = sql_query & ", " & field_rank(i)(0)
    Next i
    
    sql_query = sql_query & " FROM t_custom_rank"
    sql_query = sql_query & " ORDER BY Ticker ASC"
Dim extract_rank As Variant
extract_rank = central_query_on_ranksqlt3(sql_query)

'match avec sample rank

Dim complete_data_helper() As Variant
For i = 1 To UBound(data_ranked, 1)
    
    'ticker
    ReDim Preserve tmp_row(0)
    tmp_row(0) = data_ranked(i)(0)
    
    'new rank
    ReDim Preserve tmp_row(1)
    tmp_row(1) = data_ranked(i)(UBound(data_ranked(i), 1))
    
    k = 2
    For j = 1 To UBound(extract_rank, 1)
        If data_ranked(i)(0) = extract_rank(j)(0) Then
            
            'append data
            For m = 1 To UBound(extract_rank(j), 1)
                ReDim Preserve tmp_row(k)
                tmp_row(k) = extract_rank(j)(m)
                k = k + 1
            Next m
            
            Exit For
        End If
    Next j
    
    ReDim Preserve complete_data_helper(i - 1)
    complete_data_helper(i - 1) = tmp_row
    
Next i


    Dim count_helper_field_text As Integer, count_helper_field_numeric As Integer, count_helper_field As Integer
    count_helper_field = 0
    count_helper_field_text = 0
    count_helper_field_numeric = 0
    
    Dim field_helper_complete_data() As Variant
        count_helper_field_text = count_helper_field_text + 1
        ReDim Preserve field_helper_complete_data(count_helper_field)
        field_helper_complete_data(count_helper_field) = "f_central_helper_text1"
        count_helper_field = count_helper_field + 1
        
        
        ReDim Preserve field_helper_complete_data(count_helper_field)
        count_helper_field_numeric = count_helper_field_numeric + 1
        field_helper_complete_data(count_helper_field) = "f_central_helper_numeric1"
        count_helper_field = count_helper_field + 1
        
        'append les fields de rank
        For i = 0 To UBound(field_rank, 1)
            ReDim Preserve field_helper_complete_data(count_helper_field)
            If LCase(field_rank(i)(1)) = "text" Then
                count_helper_field_text = count_helper_field_text + 1
                field_helper_complete_data(count_helper_field) = "f_central_helper_" & LCase(field_rank(i)(1)) & count_helper_field_text
            ElseIf LCase(field_rank(i)(1)) = "numeric" Then
                count_helper_field_numeric = count_helper_field_numeric + 1
                field_helper_complete_data(count_helper_field) = "f_central_helper_" & LCase(field_rank(i)(1)) & count_helper_field_numeric
            End If
            
            count_helper_field = count_helper_field + 1
            
        Next i
        
        If count_helper_field_numeric > central_helper_nbre_numeric_fields Or count_helper_field_text > central_helper_nbre_text_fields Then
            MsgBox ("too much rank fields !")
        End If
        
        
        exec_query = sqlite3_query(central_get_db_fullpath, "DELETE FROM " & t_central_helper)
        insert_status = sqlite3_insert_with_transaction(central_get_db_fullpath, t_central_helper, complete_data_helper, field_helper_complete_data)
        'debug_test = sqlite3_query(central_get_db_fullpath, "SELECT * FROM " & t_central_helper & " ORDER BY f_central_helper_text1 ASC")
        
        'super query de selection depuis helper
        sql_query = "SELECT " & f_central_helper_text1 & " AS Ticker, " & f_central_helper_numeric1 & " AS " & Mid(central_get_compatible_sql_field_name(rank_name), 5)
            
            'append field rank
            count_helper_field_text = 1
            count_helper_field_numeric = 1
            
            For i = 0 To UBound(field_rank, 1)
                If LCase(field_rank(i)(1)) = "text" Then
                    count_helper_field_text = count_helper_field_text + 1
                    sql_query = sql_query & ", " & "f_central_helper_" & LCase(field_rank(i)(1)) & count_helper_field_text & " AS " & field_rank(i)(0)
                ElseIf LCase(field_rank(i)(1)) = "numeric" Then
                    count_helper_field_numeric = count_helper_field_numeric + 1
                    sql_query = sql_query & ", " & "f_central_helper_" & LCase(field_rank(i)(1)) & count_helper_field_numeric & " AS " & field_rank(i)(0)
                End If
                
            Next i
            
        sql_query = sql_query & " FROM " & t_central_helper
        sql_query = sql_query & " ORDER BY " & f_central_helper_text1 & " ASC"
        
        Dim final_merge_extract_with_helper As Variant
        final_merge_extract_with_helper = sqlite3_query(central_get_db_fullpath, sql_query)
    
'print excel report
Dim tmp_wrbk As Workbook

Set tmp_wrbk = Application.Workbooks.Add

tmp_wrbk.Worksheets(1).name = "raw_bbg"
tmp_wrbk.Worksheets(2).name = "rank"
tmp_wrbk.Worksheets(3).name = "report"

Application.Calculation = xlCalculationManual

For i = 0 To UBound(extract_data_for_ranking, 1)
    For j = 0 To UBound(extract_data_for_ranking(i), 1) - 1
        tmp_wrbk.Worksheets(1).Cells(i + 1, j + 1) = extract_data_for_ranking(i)(j)
    Next j
Next i


For i = 0 To UBound(data_ranked, 1)
    For j = 0 To UBound(data_ranked(i), 1)
        tmp_wrbk.Worksheets(2).Cells(i + 1, j + 1) = data_ranked(i)(j)
    Next j
Next i


For i = 0 To UBound(final_merge_extract_with_helper, 1)
    For j = 0 To UBound(final_merge_extract_with_helper(i), 1)
        tmp_wrbk.Worksheets(3).Cells(i + 1, j + 1) = final_merge_extract_with_helper(i)(j)
        
        If IsNumeric(final_merge_extract_with_helper(i)(j)) Then
            For m = 0 To UBound(vec_alert, 1)
                If final_merge_extract_with_helper(i)(j) >= vec_alert(m)(0) And final_merge_extract_with_helper(i)(j) <= vec_alert(m)(1) Then
                    tmp_wrbk.Worksheets(3).Cells(i + 1, j + 1).Interior.ColorIndex = vec_alert(m)(2)
                End If
            Next m
        End If
        
    Next j
Next i

tmp_wrbk.Worksheets(3).rows(1).AutoFilter
tmp_wrbk.Worksheets(3).Activate


End Function


Private Function central_update_db_data_bbg(ByVal vec_queries As Variant) As Variant

Dim i As Long, j As Long, k As Long


Dim initReturn As Variant
Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long


initReturn = SQLite3Initialize()
retValue = SQLite3Open(central_get_db_fullpath, dbHandle)
retValue = SQLite3Finalize(stmHandle)

SQLite3PrepareV2 dbHandle, "BEGIN TRANSACTION", stmHandle
SQLite3Step stmHandle
SQLite3Finalize stmHandle


For i = 0 To UBound(vec_queries, 1)
    SQLite3PrepareV2 dbHandle, vec_queries(i), stmHandle
    retValue = SQLite3Step(stmHandle)
    retValue = SQLite3Reset(stmHandle)
Next i

SQLite3Finalize stmHandle
SQLite3PrepareV2 dbHandle, "COMMIT TRANSACTION", stmHandle
retValue = SQLite3Step(stmHandle)
SQLite3Finalize stmHandle

SQLite3Close dbHandle

central_update_db_data_bbg = sqlite3_query(central_get_db_fullpath, "SELECT * FROM " & t_central_data_bbg & " ORDER BY " & f_central_data_bbg_ticker & " ASC")

End Function
