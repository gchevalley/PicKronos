Attribute VB_Name = "bas_Chart"
Public Sub load_Chart_universal_live(ByVal Chart As String, ByVal view As String)


'greg chevalley
' nouvelle fonction avec support des filtres colonne EH sheet parameters

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim test_debug As Variant, debug_test As Variant

'charge la config des graphs
Dim chars_config(10) As Variant

' 0 - chart / 1 - array(critere) / ( 2 - parameters_line_limit ) / 3 - array(vec1_column) / 4 - array(vec2_column) / 5 - array(vec_sector_column) / 6 - array (vec -> destination) / 7 - sheet_chart
chars_config(0) = Array("P%L", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(17, 132), Array(17, 133)), Array("Short Name", Array("Daily Change", "Nav Daily")), Array("Short Name", Array("Theta_ALL", "Theta Nav")), Array("Short Name", "Sector code"), Array(Array(1, 2), Array("", 3), Array("", 26)), "Chart Daily P&L")
chars_config(3) = Array("P%L Daily", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(17, 132), Array(17, 133)), Array("Short Name", Array("Daily Change", "Nav Daily")), Array("Short Name", Array("Theta_ALL", "Theta Nav")), Array("Short Name", "Sector code"), Array(Array(1, 2), Array("", 3), Array("", 26)), "Chart Daily P&L")
chars_config(5) = Array("P%L Weekly", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(17, 132), Array(17, 133)), Array("Short Name", Array("weekly p&l", "Nav Daily")), Array("Short Name", Array("Theta_ALL", "Theta Nav")), Array("Short Name", "Sector code"), Array(Array(1, 2), Array("", 3), Array("", 26)), "Chart Daily P&L")
chars_config(4) = Array("P%L YTD", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(17, 132), Array(17, 133)), Array("Short Name", Array("Result Total", "Nav Daily")), Array("Short Name", Array("Theta_ALL", "Theta Nav")), Array("Short Name", "Sector code"), Array(Array(1, 2), Array("", 3), Array("", 26)), "Chart Daily P&L")


chars_config(1) = Array("Positions", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(14, 132), Array(14, 133)), Array("Short Name", Array("Valeur_Euro", "Nav Position")), Array("Short Name", Array("Result Total", "Nav Daily")), Array("Short Name", "Sector code"), Array(Array(27, 28), Array("", 29), Array("", 52)), "Chart Positions")
chars_config(2) = Array("Vega", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(20, 132), Array(20, 133)), Array("Short Name", Array("Vega_1%_ALL", "Nav Vol")), Array("", Array("", "")), Array("Short Name", "Sector code"), Array(Array(53, 54), Array("", ""), Array("", 78)), "Chart Volatilities")

chars_config(6) = Array("Rel Perf", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(14, 132), Array(14, 133)), Array("Short Name", Array("Valeur_Euro", "Nav Position")), Array("Short Name", Array("Result Total", "Nav Daily")), Array("Short Name", "Sector code"), Array(Array(27, 28), Array("", 29), Array("", 52)), "Chart Perf", Array(Array("bbg", "rel_1d", 0)))

'repère la config concernée par l'appel
Dim config_chart As Integer
For i = 0 To UBound(chars_config, 1)
    If chars_config(i)(0) = Chart Then
        config_chart = i
        Exit For
    End If
Next i

Dim tmp_colonne_criteria_1 As String, tmp_colonne_criteria_2 As String
'selon le mode (book/nav) arrange la colonne
If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
    
    'colonne criteria
    tmp_colonne_criteria_1 = chars_config(config_chart)(1)(0)(0)
    tmp_colonne_criteria_2 = chars_config(config_chart)(1)(0)(1)
    
    chars_config(config_chart)(1)(0) = tmp_colonne_criteria_1
    chars_config(config_chart)(1)(1) = tmp_colonne_criteria_2
    
    'colonne a grapher
    chars_config(config_chart)(3)(1) = chars_config(config_chart)(3)(1)(0)
    chars_config(config_chart)(4)(1) = chars_config(config_chart)(4)(1)(0)
ElseIf UCase(view) = "NAV" Then
    'colonne criteria
    tmp_colonne_criteria_1 = chars_config(config_chart)(1)(1)(0)
    tmp_colonne_criteria_2 = chars_config(config_chart)(1)(1)(1)
    
    chars_config(config_chart)(1)(0) = tmp_colonne_criteria_1
    chars_config(config_chart)(1)(1) = tmp_colonne_criteria_2
    
    'colonne a grapher
    chars_config(config_chart)(3)(1) = chars_config(config_chart)(3)(1)(1)
    chars_config(config_chart)(4)(1) = chars_config(config_chart)(4)(1)(1)
Else
    'default = book
    chars_config(config_chart)(3)(1) = chars_config(config_chart)(3)(1)(0)
    chars_config(config_chart)(4)(1) = chars_config(config_chart)(4)(1)(0)
End If


Dim oJSON As New JSONLib
Dim oTags As Collection, oTag As Collection


Dim l_rows As Long, l_row As Long

Dim l_val_c1
Dim l_val_c2

Dim l_val_1
Dim l_val_2

Dim l_array_1()
Dim l_array_2()
Dim l_array_sectors()
Dim l_array_line()
Dim l_array_index As Long

Dim l_xls_sheet As Worksheet
Dim l_xls_chart As Chart
Dim l_xls_series As Series

Dim l_chart_rows As Long, l_chart_row As Long

Dim date_tmp As Date, date_today As Date
date_today = Date


Dim c_parameters_chart_name As Integer, c_parameters_chart_criteria_1 As Integer, c_parameters_criteria_2 As Integer, _
    l_parameters_chart_first_line As Integer
    
    l_parameters_chart_first_line = 12
    c_parameters_chart_name = 131
    c_parameters_chart_criteria_1 = 132
    c_parameters_criteria_2 = 133

Dim limit_chart_criteria_1 As Double, limit_chart_criteria_2 As Double

Dim l_chart_database_header As Integer
l_chart_database_header = 12

'repere les limites
Dim l_parameters_chart_config_line As Integer
For i = l_parameters_chart_first_line To 100
    If Worksheets("Parametres").Cells(i, c_parameters_chart_name) = "" And Worksheets("Parametres").Cells(i + 1, c_parameters_chart_name) = "" And Worksheets("Parametres").Cells(i + 2, c_parameters_chart_name) = "" Then
        Exit For
    Else
        If InStr(Chart, " ") <> 0 Then
            If InStr(Worksheets("Parametres").Cells(i, c_parameters_chart_name), Left(Chart, InStr(Chart, " ") - 1)) <> 0 Then
                l_parameters_chart_config_line = i
                Exit For
            End If
        Else
            If InStr(Worksheets("Parametres").Cells(i, c_parameters_chart_name), Chart) <> 0 Then
                l_parameters_chart_config_line = i
                Exit For
            End If
        End If
    End If
Next i






If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
    limit_chart_criteria_1 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_chart_criteria_1).Value
    limit_chart_criteria_2 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_criteria_2).Value
ElseIf UCase(view) = "NAV" Then
    limit_chart_criteria_1 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 2, c_parameters_chart_criteria_1).Value
    limit_chart_criteria_2 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 2, c_parameters_criteria_2).Value
    
    'limit_chart_criteria_1 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_chart_criteria_1).Value
    'limit_chart_criteria_2 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_criteria_2).Value
Else
    Exit Sub
End If



Dim l_equity_db_first_line As Integer
    l_equity_db_first_line = 27
    
    Dim line_step As Integer
    Dim c_equity_db_name As Integer, l_header_equity_db As Integer, c_equity_db_last_column As Integer, _
        c_equity_db_criteria_1 As Integer, c_equity_db_criteria_2 As Integer
        
        c_equity_db_name = 2
        Dim c_equity_db_criteria() As Variant
            ReDim c_equity_db_criteria(UBound(chars_config(config_chart)(1), 1))
        
        Dim c_equity_db_vect_1() As Variant
            ReDim c_equity_db_vect_1(UBound(chars_config(config_chart)(3), 1))
        
        Dim c_equity_db_vect_2() As Variant
            ReDim c_equity_db_vect_2(UBound(chars_config(config_chart)(4), 1))
        
        Dim c_equity_db_vect_3() As Variant
            ReDim c_equity_db_vect_3(UBound(chars_config(config_chart)(5), 1))
        
        c_equity_db_criteria_1 = 0
        c_equity_db_criteria_2 = 0
        l_header_equity_db = 25
        line_step = 2
        
        c_equity_db_last_column = 250

Dim l_index_db_first_line As Integer
    l_index_db_first_line = 27
    
    
    Dim l_header_index_db As Integer, c_index_db_last_column As Integer, c_index_db_criteria_1 As Integer, _
        c_index_db_criteria_2 As Integer
    
    c_equity_db_name = 2
    Dim c_index_db_criteria() As Variant
        ReDim c_index_db_criteria(UBound(chars_config(config_chart)(1), 1))
    
    
    Dim c_index_db_vect_1() As Variant
            ReDim c_index_db_vect_1(UBound(chars_config(config_chart)(3), 1))
        
        Dim c_index_db_vect_2() As Variant
            ReDim c_index_db_vect_2(UBound(chars_config(config_chart)(4), 1))
        
        Dim c_index_db_vect_3() As Variant
            ReDim c_index_db_vect_3(UBound(chars_config(config_chart)(5), 1))
    
    
        
    c_index_db_criteria_1 = 0
    c_index_db_criteria_2 = 0
    l_header_index_db = 25
    line_step = 2
    
    c_index_db_last_column = 250



'repère les colonnes dans equity_db & index_db pour les limites
For i = 1 To c_equity_db_last_column
    For j = 0 To UBound(chars_config(config_chart)(1), 1)
        If UCase(chars_config(config_chart)(1)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
            c_equity_db_criteria(j) = i
        End If
    Next j
    
    'vector criteria 1
    For j = 0 To UBound(chars_config(config_chart)(3), 1)
        If chars_config(config_chart)(3)(j) <> "" Then
            If UCase(chars_config(config_chart)(3)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
                c_equity_db_vect_1(j) = i
            End If
        Else
            c_equity_db_vect_1(j) = 0
        End If
    Next j
    
    'vec criteria 2
    For j = 0 To UBound(chars_config(config_chart)(4), 1)
        If chars_config(config_chart)(4)(j) <> "" Then
            If UCase(chars_config(config_chart)(4)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
                c_equity_db_vect_2(j) = i
            End If
        Else
            c_equity_db_vect_2(j) = 0
        End If
    Next j
    
    'vect sector
    For j = 0 To UBound(chars_config(config_chart)(5), 1)
        If chars_config(config_chart)(5)(j) <> "" Then
            If UCase(chars_config(config_chart)(5)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
                c_equity_db_vect_3(j) = i
            End If
        Else
            c_equity_db_vect_3(j) = 0
        End If
    Next j
    
Next i


For i = 1 To c_index_db_last_column
    For j = 0 To UBound(chars_config(config_chart)(1), 1)
        If UCase(chars_config(config_chart)(1)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
            c_index_db_criteria(j) = i
        End If
    Next j
    
    
    
    'vector criteria 1
    For j = 0 To UBound(chars_config(config_chart)(3), 1)
        If chars_config(config_chart)(3)(j) <> "" Then
            If UCase(chars_config(config_chart)(3)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
                c_index_db_vect_1(j) = i
            End If
        Else
            c_index_db_vect_1(j) = 0
        End If
    Next j
    
    'vec criteria 2
    For j = 0 To UBound(chars_config(config_chart)(4), 1)
        If chars_config(config_chart)(4)(j) <> "" Then
            If UCase(chars_config(config_chart)(4)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
                c_index_db_vect_2(j) = i
            End If
        Else
            c_index_db_vect_2(j) = 0
        End If
    Next j
    
    'vect sector
    For j = 0 To UBound(chars_config(config_chart)(5), 1)
        If chars_config(config_chart)(5)(j) <> "" Then
            If UCase(chars_config(config_chart)(5)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
                c_index_db_vect_3(j) = i
            End If
        Else
            c_index_db_vect_3(j) = 0
        End If
    Next j
    
    
Next i




'chargement des filtres (pour commencer uniquement ceux du type colonne)
'filtre
Dim c_filter_name As Integer
    c_filter_name = 138

Dim c_filter_criteria As Integer
    c_filter_criteria = 139

Dim l_filter_first_line As Integer
    l_filter_first_line = 11
        
        
Dim activate_filter As Boolean

If IsNumeric(Worksheets("Parametres").Cells(l_filter_first_line, c_filter_criteria)) And Worksheets("Parametres").Cells(l_filter_first_line, c_filter_criteria) = 1 Then
    activate_filter = True
Else
    activate_filter = False
End If



If activate_filter = True Then
    Application.Calculation = xlCalculationManual
    
    'repere la derniere ligne
    Dim l_filter_last_line As Integer
    For i = l_filter_first_line To 5000
        If Worksheets("Parametres").Cells(i, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 1, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 2, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 3, c_filter_name) = "" Then
            l_filter_last_line = i - 1
            Exit For
        End If
    Next i
    
    
    'repere les filters colonnes activés
    Dim filter() As Variant
    Dim filter_type() As Variant
    Dim filter_type_criteria() As Variant 'double, int, txt
    Dim filter_compare_type() As Variant ' <, > =, limit, abs
    Dim filter_column_equity_db() As Variant
    Dim filter_column_equity_idx() As Variant
    Dim filter_criteria() As Variant
    
    Dim count_ticker As Integer
        count_ticker = 0
    Dim list_tickers() As Variant
    Dim bbg_fld() As Variant
    Dim bbg_fld_custom() As Variant
    
    Dim vec_array_1() As Variant
    Dim vec_array_2() As Variant
    Dim vec_array_sector() As Variant
    Dim vec_array_line() As Variant
    
    Dim nbre_filters_column As Integer, nbre_filters_custom As Integer, nbre_filters_api As Integer, _
        nbre_filters_api_custom As Integer
    
    
    Dim region As Variant
    region = Array(Array("Asia/Pacific", Array("JPY", "HKD", "AUD", "SGD", "TWD", "KRW", "INR", "THB")), Array("Europe", Array("CHF", "EUR", "GBP", "SEK", "NOK", "DKK", "PLN")), Array("America", Array("USD", "CAD", "BRL")))
    
    Dim currency_code() As Variant
    k = 0
    For j = 14 To 32
        If Worksheets("Parametres").Cells(j, 1) <> "" Then
            ReDim Preserve currency_code(k)
            currency_code(k) = Array(Left(Worksheets("Parametres").Cells(j, 1).Value, 3), Worksheets("Parametres").Cells(j, 5).Value)
            k = k + 1
        End If
    Next j
    
    
    j = 0
    nbre_filters_column = 0
    nbre_filters_custom = 0
    nbre_filters_api = 0
    nbre_filters_api_custom = 0
    
    For i = l_filter_first_line To l_filter_last_line
        If InStr(Worksheets("Parametres").Cells(i, c_filter_name), "col_") <> 0 And Worksheets("Parametres").Cells(i + 1, c_filter_criteria) <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "col_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "col"
            
            ReDim Preserve filter_type_criteria(j)
            If IsNumeric(Worksheets("Parametres").Cells(i + 2, c_filter_criteria)) Then
                filter_type_criteria(j) = "num"
            Else
                filter_type_criteria(j) = "str"
            End If
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            ReDim Preserve filter_compare_type(j)
            filter_compare_type(j) = Worksheets("Parametres").Cells(i + 2, c_filter_name)
            
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            
            j = j + 1
            nbre_filters_column = nbre_filters_column + 1
        
        ElseIf InStr(Worksheets("Parametres").Cells(i, c_filter_name), "custom_") <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "custom_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "custom"
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 1, c_filter_criteria)
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            If filter_criteria(j) <> 0 And UCase(Worksheets("Parametres").Cells(i + 2, c_filter_name)) = "CRITERIA" Then
                filter_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            End If
            
            j = j + 1
            nbre_filters_custom = nbre_filters_custom + 1
            
        ElseIf InStr(Worksheets("Parametres").Cells(i, c_filter_name), "api_") <> 0 And Worksheets("Parametres").Cells(i + 1, c_filter_criteria) <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "api_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "api"
            
            ReDim Preserve filter_type_criteria(j)
            filter_type_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 3, c_filter_criteria)
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            ReDim Preserve bbg_fld(nbre_filters_api)
            bbg_fld(nbre_filters_api) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "api_", "")
            
            j = j + 1
            nbre_filters_api = nbre_filters_api + 1
            
        End If
    Next i
    
    Dim nbre_filters As Integer
        nbre_filters = j - 1 + 1
    
    
    'remonte les colonnes concernees d'equity & index database
    For i = 0 To UBound(filter)
        For j = 1 To c_equity_db_last_column
            If UCase(Replace(filter(i), "_", " ")) = UCase(Replace(Worksheets("Equity_Database").Cells(l_header_equity_db, j), "_", " ")) Then
                filter_column_equity_db(i) = j
                GoTo next_filter_column_equity_database
            End If
        Next j
next_filter_column_equity_database:
    Next i
    
    For i = l_equity_db_first_line To 5000 Step line_step
        Dim l_equity_db_last_line As Integer
        If Worksheets("Equity_Database").Cells(i, c_equity_db_name) = "" And Worksheets("Equity_Database").Cells(i + 1 * line_step, c_equity_db_name) = "" And Worksheets("Equity_Database").Cells(i + 2 * line_step, c_equity_db_name) = "" Then
            l_equity_db_last_line = i - line_step
            Exit For
        End If
    Next i
    
    
    'remonte les colonnes conernees d'index database
    For i = 0 To UBound(filter)
        For j = 1 To c_index_db_last_column
            If UCase(Replace(filter(i), "_", " ")) = UCase(Replace(Worksheets("Index_Database").Cells(l_header_index_db, j), "_", " ")) Then
                filter_column_equity_idx(i) = j
                GoTo next_filter_column_index_database
            End If
        Next j
next_filter_column_index_database:
    Next i
    
    
    
    Dim c_index_db_name As Integer
    c_index_db_name = 2
    line_step = 3
    For i = l_equity_db_first_line To 5000 Step line_step
        Dim l_index_db_last_line As Integer
        If Worksheets("Index_Database").Cells(i, c_index_db_name) = "" And Worksheets("Index_Database").Cells(i + 1 * line_step, c_index_db_name) = "" And Worksheets("Index_Database").Cells(i + 2 * line_step, c_index_db_name) = "" Then
            l_index_db_last_line = i - line_step
            Exit For
        End If
    Next i
    
    
    
    'LANCE LA RECHERCHE DE TITRES
    
    'capte les titres qui rentre dans les criteres d'equity database
    Dim tmp_equity_filter_ok As Boolean
    
    count_ticker = 0
    For i = l_equity_db_first_line To l_equity_db_last_line Step 2
        
        On Error GoTo next_entry_equity_db:
        
        Dim limit_equity_value_criteria_1 As Variant, limit_equity_value_criteria_2 As Variant
        
        
        limit_equity_value_criteria_1 = Worksheets("Equity_Database").Cells(i, c_equity_db_criteria(0))
        limit_equity_value_criteria_2 = Worksheets("Equity_Database").Cells(i, c_equity_db_criteria(1))
        
        If IsNumeric(limit_equity_value_criteria_1) And IsNumeric(limit_equity_value_criteria_2) Then
            If Abs(limit_equity_value_criteria_1) >= limit_chart_criteria_1 And Abs(limit_equity_value_criteria_2) >= limit_chart_criteria_2 Then
                
                'passe dans les filtres de colonnes
                tmp_equity_filter_ok = True
                For j = 0 To UBound(filter)
                    
                    If filter_type(j) = "col" Then
                    
                        ' la colonne existe dans equity DB
                        If tmp_equity_filter_ok = True And filter_column_equity_db(j) <> False And IsNumeric(filter_column_equity_db(j)) Then
                            
                            ' passe le check critere ou non
                            If filter_compare_type(j) = "=" And filter_type_criteria(j) = "str" And UCase(filter_criteria(j)) = UCase(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) Then
                                
                            ElseIf filter_compare_type(j) = "=" And filter_type_criteria(j) = "num" Then
                                debug_test = Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))
                                If filter_criteria(j) = Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j)) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            ElseIf filter_compare_type(j) = "limit" And filter_type_criteria(j) = "num" Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) And filter_criteria(j) <= Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j)) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            ElseIf filter_compare_type(j) = "limit_abs" And filter_type_criteria(j) = "num" Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) And filter_criteria(j) <= Abs(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            Else
                                tmp_equity_filter_ok = False
                            End If
                            
                        Else
                            tmp_equity_filter_ok = False
                        End If
                
                
                    ElseIf filter_type(j) = "custom" Then
                    
                        ' #################################################################################################################
                        ' ############################ passe dans les filtres custom ######################################################
                        ' #################################################################################################################
                        If UCase(filter(j)) = "EQUITY" Then
                            If filter_criteria(j) = 0 Then
                                tmp_equity_filter_ok = False
                            End If
                        ElseIf UCase(filter(j)) = "SHORT" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 5)) And Worksheets("Equity_Database").Cells(i, 5) < 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "LONG" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 5)) And Worksheets("Equity_Database").Cells(i, 5) > 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "STOCK ONLY" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 9)) And Worksheets("Equity_Database").Cells(i, 9) = 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "LENDING" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 26)) And Worksheets("Equity_Database").Cells(i, 26) = 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "EXPECTED_REPORT_DT" Then
                            If filter_criteria(j) > 0 Then
                                nbre_filters_api_custom = nbre_filters_api_custom + 1
                            End If
                        ElseIf UCase(filter(j)) = "DVD_EX_DT" Then
                            If filter_criteria(j) > 0 Then
                                nbre_filters_api_custom = nbre_filters_api_custom + 1
                            End If
                        ElseIf UCase(filter(j)) = "REGION" Then
                            For k = 0 To UBound(region, 1)
                                If region(k)(0) = filter_criteria(j) Then
                                    For m = 0 To UBound(currency_code)
                                        If Worksheets("Equity_Database").Cells(i, 44) = currency_code(m)(1) Then
                                            For n = 0 To UBound(region(k)(1))
                                                If currency_code(m)(0) = region(k)(1)(n) Then
                                                    GoTo bypass_filter_region_check
                                                Else
                                                    If n = UBound(region(k)(1)) Then
                                                        tmp_equity_filter_ok = False
                                                    End If
                                                End If
                                            Next n
                                        End If
                                    Next m
                                End If
                            Next k
bypass_filter_region_check:
                        ElseIf UCase(filter(j)) = "TAG" Then
                        
                            If filter_criteria(j) <> 0 Then
                                
                                If tmp_equity_filter_ok = True Then
                                    tmp_equity_filter_ok = False
                                    
                                    If Worksheets("Equity_Database").Cells(i, 137) <> "" Then
                                        
                                        Set oTags = oJSON.parse(Worksheets("Equity_Database").Cells(i, 137))
        
                                        For Each oTag In oTags
                                            
                                            If oTag.Item(2) = "OPEN" And oTag.Item(3) = filter_criteria(j) Then
                                                tmp_equity_filter_ok = True
                                                Exit For
                                            End If
                                        Next
                                        
                                    End If
                                End If
                            
                            End If
                        End If
                    
                    End If
                        
                Next j
                

                'le titre doit etre pris en compte
                If tmp_equity_filter_ok = True Then
                    
                    If nbre_filters_api = 0 And nbre_filters_api_custom = 0 Then
                    
                        ReDim Preserve l_array_1(l_array_index)
                        ReDim Preserve l_array_2(l_array_index)
                        ReDim Preserve l_array_sectors(l_array_index)
                        ReDim Preserve l_array_line(l_array_index)
                        
                        If c_equity_db_vect_1(1) <> 0 Then
                            l_array_1(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(1)).Value) 'short name + daily change
                        End If
                        
                        If c_equity_db_vect_2(1) <> 0 Then
                            l_array_2(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(1)).Value) 'short name + theta
                        End If
                        
                        l_array_sectors(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(1)).Value) 'short name + beta_sector_CODE
                        l_array_line(l_array_index) = Array("Equity_Database", i)
                        l_array_index = l_array_index + 1
                    
                    Else
                        
                        ReDim Preserve list_tickers(count_ticker)
                        list_tickers(count_ticker) = Worksheets("Equity_Database").Cells(i, 47)
                        
                        
                        ReDim Preserve vec_array_1(count_ticker)
                            vec_array_1(count_ticker) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(1)).Value) 'short name + daily change
                        
                        ReDim Preserve vec_array_2(count_ticker)
                            vec_array_2(count_ticker) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(1)).Value) 'short name + theta
                        
                        
                        ReDim Preserve vec_array_sector(count_ticker)
                            vec_array_sector(count_ticker) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(1)).Value) 'short name + beta_sector_CODE
                        
                        ReDim Preserve vec_array_line(count_ticker)
                            vec_array_line(count_ticker) = Array("Equity_Database", i)
                        
                        count_ticker = count_ticker + 1
                        
                    End If
                    
                End If
                
            End If
        End If
next_entry_equity_db:
    Next i
    
    
    On Error GoTo 0
    
    Dim pass_filter_bbg As Boolean
    
    If count_ticker > 0 Then
        
        If nbre_filters_api > 0 Then
            Dim output_bbg As Variant
            output_bbg = bbg_multi_tickers_and_multi_fields(list_tickers, bbg_fld)
        End If
        
        If nbre_filters_api_custom > 0 Then
            Dim output_bbg_custom As Variant
            bbg_fld_custom = Array("EXPECTED_REPORT_DT", "DVD_EX_DT")
            
            output_bbg_custom = bbg_multi_tickers_and_multi_fields(list_tickers, bbg_fld_custom)
        End If
            
                            
            For i = 0 To UBound(list_tickers, 1)
            
                pass_filter_bbg = True
                
                If nbre_filters_api > 0 Then
                    For j = 0 To UBound(bbg_fld, 1)
                        For k = 0 To UBound(filter, 1)
                            If filter_type(k) = "api" And UCase(bbg_fld(j)) = UCase(filter(k)) Then
                                
                                If Left(output_bbg(i, j), 1) <> "#" Then
                                    If filter_type_criteria(k) = "num" Then
                                        If output_bbg(i, j) >= filter_criteria(k) Then
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    ElseIf filter_type_criteria(k) = "str" Then
                                        If output_bbg(i, j) = filter_criteria(k) Then
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    End If
                                Else
                                    pass_filter_bbg = False
                                End If
                                
                            End If
                        Next k
                    Next j
                End If
                
                
                
                If pass_filter_bbg = True And nbre_filters_api_custom > 0 Then
                    
                    For j = 0 To UBound(bbg_fld_custom, 1)
                        For k = 0 To UBound(filter, 1)
                            If filter_type(k) = "custom" And UCase(bbg_fld_custom(j)) = UCase(filter(k)) Then
                                If Left(output_bbg_custom(i, j), 1) <> "#" Then
                                    
                                    'code custom
                                    If UCase(filter(k)) = "EXPECTED_REPORT_DT" Then
                                        date_tmp = Mid(output_bbg_custom(i, j), 4, 2) & "." & Left(output_bbg_custom(i, j), 2) & "." & Right(output_bbg_custom(i, j), 4)
                                        
                                        test_debug = date_tmp - Date
                                        
                                        If date_tmp - Date <= filter_criteria(k) And date_tmp > Date Then
                                            
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    ElseIf UCase(filter(k)) = "DVD_EX_DT" Then
                                        date_tmp = Mid(output_bbg_custom(i, j), 4, 2) & "." & Left(output_bbg_custom(i, j), 2) & "." & Right(output_bbg_custom(i, j), 4)
                                        
                                        test_debug = date_tmp - Date
                                        
                                        If date_tmp - Date <= filter_criteria(k) And date_tmp > Date Then
                                            
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    End If
                                    
                                End If
                            End If
                        Next k
                    Next j
                    
                End If
                
                
                
                If pass_filter_bbg = True Then
                    
                    ReDim Preserve l_array_1(l_array_index)
                    ReDim Preserve l_array_2(l_array_index)
                    ReDim Preserve l_array_sectors(l_array_index)
                    ReDim Preserve l_array_line(l_array_index)
                            
                    l_array_1(l_array_index) = vec_array_1(i)
                    l_array_2(l_array_index) = vec_array_2(i)
                    l_array_sectors(l_array_index) = vec_array_sector(i)
                    l_array_line(l_array_index) = vec_array_line(i)
                    l_array_index = l_array_index + 1
                    
                End If
                
            Next i
            
            
            
        
    End If
    
    
    'capte les titres qui rentre dans les criteres d'index database
    limit_equity_value_criteria_1 = 0
    limit_equity_value_criteria_2 = 0
    For i = l_index_db_first_line To l_index_db_last_line Step 3

        On Error GoTo next_entry_index_db:
        
        If IsEmpty(c_index_db_criteria(0)) = True Or IsEmpty(c_index_db_criteria(1)) = True Then
            GoTo out_of_index_db
        Else
        
            limit_equity_value_criteria_1 = Worksheets("Index_Database").Cells(i, c_index_db_criteria(0))
            limit_equity_value_criteria_2 = Worksheets("Index_Database").Cells(i, c_index_db_criteria(1))
    
            If IsNumeric(limit_equity_value_criteria_1) And IsNumeric(limit_equity_value_criteria_2) Then
                If Abs(limit_equity_value_criteria_1) >= limit_chart_criteria_1 And Abs(limit_equity_value_criteria_2) >= limit_chart_criteria_2 Then
    
                    'passe dans les filtres de colonnes
                    tmp_equity_filter_ok = True
                    For j = 0 To UBound(filter)
                        If filter_type(j) = "col" Then
                            ' la colonne existe dans equity DB
                            If tmp_equity_filter_ok = True And filter_column_equity_idx(j) <> False And IsNumeric(filter_column_equity_idx(j)) Then
    
                                ' passe le check critere ou non
                                If filter_compare_type(j) = "=" And filter_type_criteria(j) = "str" And UCase(filter_criteria(j)) = UCase(Worksheets("Index_Database").Cells(i, filter_column_equity_idx(j))) Then
    
                                ElseIf filter_compare_type(j) = "=" And filter_type_criteria(j) = "num" Then
                                    If filter_criteria(j) = Worksheets("Index_Database").Cells(i, filter_column_equity_idx(j)) Then
    
                                    End If
                                ElseIf filter_compare_type(j) = "limit" And filter_type_criteria(j) = "num" Then
                                    If IsNumeric(Worksheets("Index_Database").Cells(i, filter_column_equity_idx(j))) And filter_criteria(j) <= Worksheets("Index_Database").Cells(i, filter_column_equity_idx(j)) Then
    
                                    End If
                                ElseIf filter_compare_type(j) = "limit_abs" And filter_type_criteria(j) = "num" Then
                                    If IsNumeric(Worksheets("Index_Database").Cells(i, filter_column_equity_idx(j))) And filter_criteria(j) <= Abs(Worksheets("Index_Database").Cells(i, filter_column_equity_idx(j))) Then
    
                                    End If
                                Else
                                    tmp_equity_filter_ok = False
                                End If
    
                            Else
                                tmp_equity_filter_ok = False
                            End If
                        ElseIf filter_type(j) = "custom" Then
                            If UCase(filter(j)) = "INDEX" Then
                                If filter_criteria(j) = 0 Then
                                    tmp_equity_filter_ok = False
                                End If
                            ElseIf UCase(filter(j)) = "REGION" Then
                                
                                For k = 0 To UBound(region, 1)
                                    If region(k)(0) = filter_criteria(j) Then
                                        For m = 0 To UBound(currency_code)
                                            If Worksheets("Index_Database").Cells(i, 107) = currency_code(m)(1) Then
                                                For n = 0 To UBound(region(k)(1))
                                                    If currency_code(m)(0) = region(k)(1)(n) Then
                                                    
                                                    Else
                                                        If n = UBound(region(k)(1)) Then
                                                            tmp_equity_filter_ok = False
                                                        End If
                                                    End If
                                                Next n
                                            End If
                                        Next m
                                    End If
                                Next k
                                
                            End If
                        ElseIf filter_type(j) = "api" Then
                            
                        End If
                    Next j
                    
                    
                    'le titre doit etre pris en compte
                    If tmp_equity_filter_ok = True Then
    
                        If nbre_filters_api = 0 And (IsEmpty(c_index_db_vect_1(1)) = False) Then
                            ReDim Preserve l_array_1(l_array_index)
                            ReDim Preserve l_array_2(l_array_index)
                            ReDim Preserve l_array_sectors(l_array_index)
                            ReDim Preserve l_array_line(l_array_index)
                            
                            If c_index_db_vect_1(1) <> 0 Then
                                l_array_1(l_array_index) = Array(Worksheets("Index_Database").Cells(i, c_index_db_vect_1(0)).Value, Worksheets("Index_Database").Cells(i, c_index_db_vect_1(1)).Value) 'short name + daily change
                            End If
                            
                            If c_index_db_vect_2(1) <> 0 Then
                                l_array_2(l_array_index) = Array(Worksheets("Index_Database").Cells(i, c_index_db_vect_2(0)).Value, Worksheets("Index_Database").Cells(i, c_index_db_vect_2(1)).Value) 'short name + theta
                            End If
                            
                            l_array_sectors(l_array_index) = Array(Worksheets("Index_Database").Cells(i, 109).Value, 6)
                            
                            l_array_line(l_array_index) = Array("Index_Database", i)
                            
                            l_array_index = l_array_index + 1
    
                        Else
                            'le ticker doit encore passer les test bbg
    
                        End If
    
                    End If
                End If
            End If
        End If
next_entry_index_db:
    Next i
out_of_index_db:
    On Error GoTo 0

' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Else 'fonction habituelle
    
    
    With Worksheets("Equity_Database")
        l_rows = .Cells(1, g_col_a).Value 'end row
        .Range("A26:IV" & l_rows).rows.Hidden = False
                              
        For l_row = 26 To l_rows Step 2
            
            limit_equity_value_criteria_1 = .Cells(l_row + 1, c_equity_db_criteria(0)).Value
            limit_equity_value_criteria_2 = .Cells(l_row + 1, c_equity_db_criteria(1)).Value

            If IsNumeric(limit_equity_value_criteria_1) And IsNumeric(limit_equity_value_criteria_2) Then
                If Abs(limit_equity_value_criteria_1) >= limit_chart_criteria_1 And Abs(limit_equity_value_criteria_2) >= limit_chart_criteria_2 Then
                    
                    ReDim Preserve l_array_1(l_array_index)
                    ReDim Preserve l_array_2(l_array_index)
                    ReDim Preserve l_array_sectors(l_array_index)
                    
                    If c_equity_db_vect_1(1) <> 0 Then
                        l_array_1(l_array_index) = Array(.Cells(l_row + 1, 45).Value, .Cells(l_row + 1, c_equity_db_vect_1(1)).Value) 'short name + daily change
                    End If
                    
                    If c_equity_db_vect_2(1) <> 0 Then
                        l_array_2(l_array_index) = Array(.Cells(l_row + 1, 45).Value, .Cells(l_row + 1, c_equity_db_vect_2(1)).Value) 'short name + theta
                    End If
                    
                    l_array_sectors(l_array_index) = Array(.Cells(l_row + 1, 45).Value, .Cells(l_row + 1, c_equity_db_vect_3(1)).Value) 'short name + beta_sector_CODE
                    l_array_index = l_array_index + 1

                End If
            End If
        Next l_row
    End With
    
    
    With Worksheets("Index_Database")
        l_rows = .Cells(1, g_col_a).Value
        .Range("A26:IV" & l_rows).rows.Hidden = False
        
        For l_row = 26 To l_rows Step 3
            limit_equity_value_criteria_1 = .Cells(l_row + 1, c_equity_db_criteria(0)).Value
            limit_equity_value_criteria_2 = .Cells(l_row + 1, c_equity_db_criteria(1)).Value
            
            If IsNumeric(limit_equity_value_criteria_1) And IsNumeric(limit_equity_value_criteria_2) Then
                If Abs(limit_equity_value_criteria_1) >= limit_chart_criteria_1 Or Abs(limit_equity_value_criteria_2) >= limit_chart_criteria_2 Then
                    
                    ReDim Preserve l_array_1(l_array_index)
                    ReDim Preserve l_array_2(l_array_index)
                    ReDim Preserve l_array_sectors(l_array_index)
                    
                    If c_index_db_vect_1(1) <> 0 Then
                        l_array_1(l_array_index) = Array(.Cells(l_row + 1, 109).Value, .Cells(l_row + 1, c_index_db_vect_1(1)))
                    End If
                    
                    If c_index_db_vect_2(1) <> 0 Then
                        l_array_2(l_array_index) = Array(.Cells(l_row + 1, 109).Value, .Cells(l_row + 1, c_index_db_vect_2(1)))
                    End If
                    
                    l_array_sectors(l_array_index) = Array(.Cells(l_row + 1, 109).Value, 6)
                    l_array_index = l_array_index + 1
                    
                End If
            End If
        Next l_row
    End With
    
End If


        
        
        
        ' construction du graph avec la methode precedente
        l_rows = l_array_index
                         
        Set l_xls_sheet = Worksheets("Chart_Database")
         
         
        For i = 0 To UBound(chars_config(config_chart)(6), 1)
            For j = 0 To UBound(chars_config(config_chart)(6)(i), 1)
                If chars_config(config_chart)(6)(i)(j) <> "" Then
                    'clean de la column
                    For k = 13 To 32000
                        If IsError(Worksheets("Chart_Database").Cells(k, chars_config(config_chart)(6)(i)(j))) = True Then
                            Worksheets("Chart_Database").Cells(k, chars_config(config_chart)(6)(i)(j)) = ""
                        Else
                            If Worksheets("Chart_Database").Cells(k, chars_config(config_chart)(6)(i)(j)) = "" And Worksheets("Chart_Database").Cells(k + 1, chars_config(config_chart)(6)(i)(j)) = "" And Worksheets("Chart_Database").Cells(k + 2, chars_config(config_chart)(6)(i)(j)) = "" Then
                                Exit For
                            Else
                                Worksheets("Chart_Database").Cells(k, chars_config(config_chart)(6)(i)(j)) = ""
                            End If
                        End If
                    Next k
                End If
            Next j
        Next i
        
        For l_row = 0 To l_rows - 1 Step 1
        
            For i = 0 To UBound(chars_config(config_chart)(6), 1)
                For j = 0 To UBound(chars_config(config_chart)(6)(i), 1)
                    If chars_config(config_chart)(6)(i)(j) <> "" Then
                        If i = 0 Then
                            
                            'Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = l_array_1(l_row)(j)
                            
                            If j = 0 Then
                                'header valeur codée en dur pas besoin d'un live
                                Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = l_array_1(l_row)(j)
                            Else
                                If Left(UCase(l_array_line(l_row)(0)), 5) = "INDEX" Then
                                    Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = "=" & l_array_line(l_row)(0) & "!R" & l_array_line(l_row)(1) & "C" & c_index_db_vect_1(1)
                                ElseIf Left(UCase(l_array_line(l_row)(0)), 6) = "EQUITY" Then
                                    Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = "=" & l_array_line(l_row)(0) & "!R" & l_array_line(l_row)(1) & "C" & c_equity_db_vect_1(1)
                                End If
                            End If
                        ElseIf i = 1 Then
                            
                            'Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = l_array_2(l_row)(j)
                            
                            If j = 0 Then
                                'header valeur codée en dur pas besoin d'un live
                                Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = l_array_2(l_row)(j)
                            Else
                                If Left(UCase(l_array_line(l_row)(0)), 5) = "INDEX" Then
                                    Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = "=" & l_array_line(l_row)(0) & "!R" & l_array_line(l_row)(1) & "C" & c_index_db_vect_2(1)
                                ElseIf Left(UCase(l_array_line(l_row)(0)), 6) = "EQUITY" Then
                                    Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = "=" & l_array_line(l_row)(0) & "!R" & l_array_line(l_row)(1) & "C" & c_equity_db_vect_2(1)
                                End If
                            End If
                            
                        ElseIf i = 2 Then
                            'Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = l_array_sectors(l_row)(j)
                            
                            If j = 0 Then
                                'header valeur codée en dur pas besoin d'un live
                                Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = l_array_2(l_row)(j)
                            Else
                                If Left(UCase(l_array_line(l_row)(0)), 5) = "INDEX" Then
                                    Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)) = 6
                                ElseIf Left(UCase(l_array_line(l_row)(0)), 6) = "EQUITY" Then
                                    Worksheets("Chart_Database").Cells(l_row + 13, chars_config(config_chart)(6)(i)(j)).Value = "=" & l_array_line(l_row)(0) & "!R" & l_array_line(l_row)(1) & "C" & c_equity_db_vect_3(1)
                                End If
                            End If
                        End If
                    End If
                Next j
            Next i
            
        Next l_row

    
    
    Dim count_serie_chart As Integer
    count_serie_chart = 0
    
    Application.ScreenUpdating = False
    
    Set l_xls_chart = Charts(chars_config(config_chart)(7))
    With l_xls_chart
        .ChartArea.ClearContents
        .Activate
        
        l_chart_rows = .SeriesCollection.count
        
        For l_chart_row = 1 To l_chart_rows Step 1
            .SeriesCollection(1).Delete
        Next l_chart_row
              
        
        
        If chars_config(config_chart)(6)(0)(1) <> "" Then
            
            Set l_xls_series = .SeriesCollection.NewSeries
            count_serie_chart = count_serie_chart + 1
            
            l_xls_series.name = l_xls_sheet.Range(xlColumnValue(chars_config(config_chart)(6)(0)(1)) & l_chart_database_header)
            
            l_xls_series.Values = l_xls_sheet.Range(xlColumnValue(chars_config(config_chart)(6)(0)(1)) & l_chart_database_header + 1 & ":" & xlColumnValue(chars_config(config_chart)(6)(0)(1)) & l_rows + l_chart_database_header)
            l_xls_series.XValues = l_xls_sheet.Range(xlColumnValue(chars_config(config_chart)(6)(0)(0)) & l_chart_database_header + 1 & ":" & xlColumnValue(chars_config(config_chart)(6)(0)(0)) & l_rows + l_chart_database_header)
        
            For l_row = 0 To l_rows - 1 Step 1
                l_xls_series.Points(l_row + 1).Interior.Pattern = xlSolid
                l_xls_series.Points(l_row + 1).Interior.ColorIndex = format_Chart_ColorIndex(l_xls_sheet.Cells(l_row + l_chart_database_header + 1, chars_config(config_chart)(6)(2)(1)).Value)
            Next l_row
            
            If get_excel_version = excel_version.excel_2010 Then
                ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
            End If
            
            ActiveChart.SeriesCollection(count_serie_chart).DataLabels.Select
            
            If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
                Selection.NumberFormat = "#,##0"
            ElseIf UCase(view) = "NAV" Then
                Selection.NumberFormat = "0.000%"
            End If
            
            
        
        End If
        
        Set l_xls_series = Nothing
        
        
        If chars_config(config_chart)(6)(1)(1) <> "" Then
        
            Set l_xls_series = .SeriesCollection.NewSeries
            count_serie_chart = count_serie_chart + 1
            
            l_xls_series.name = l_xls_sheet.Range(xlColumnValue(chars_config(config_chart)(6)(1)(1)) & l_chart_database_header)
            
            l_xls_series.Values = l_xls_sheet.Range(xlColumnValue(chars_config(config_chart)(6)(1)(1)) & l_chart_database_header + 1 & ":" & xlColumnValue(chars_config(config_chart)(6)(1)(1)) & l_rows + l_chart_database_header)
            l_xls_series.XValues = l_xls_sheet.Range(xlColumnValue(chars_config(config_chart)(6)(0)(0)) & l_chart_database_header + 1 & ":" & xlColumnValue(chars_config(config_chart)(6)(0)(0)) & l_rows + l_chart_database_header)
            
            
            If get_excel_version() = excel_version.excel_2010 Then
                ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
            End If
            
            ActiveChart.SeriesCollection(count_serie_chart).DataLabels.Select
            
            
            If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
                Selection.NumberFormat = "#,##0"
            ElseIf UCase(view) = "NAV" Then
                Selection.NumberFormat = "0.000%"
            End If
            
            ActiveChart.PlotArea.Select
        
        End If
        
    End With
    
    Set l_xls_series = Nothing
    Set l_xls_chart = Nothing
    Set l_xls_sheet = Nothing
    
    Application.ScreenUpdating = True
    
    Application.Calculation = xlCalculationAutomatic

End Sub



Public Sub load_Chart_Sector_universal(ByVal view As String)

Dim debug_test As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim l_rows As Long, l_row As Long
Dim L_value
Dim l_formula As String

Dim l_col_0
Dim l_col_00

Dim l_col_1
Dim l_col_2
Dim l_col_3
Dim l_col_sector

Dim l_val_c1
Dim l_val_c2
Dim l_val_1
Dim l_val_2

Dim l_array_x1()
Dim l_array_y1()
Dim l_array_y2()
Dim l_array_y3()
Dim l_array_y5()
Dim l_array_z1()
Dim l_array_index As Long

Dim l_xls_sheet As Worksheet
Dim l_xls_chart As Chart
Dim l_xls_series As Series

Dim l_chart_rows As Long, l_chart_row As Long



Application.Calculation = xlCalculationManual



Dim c_equity_db_sector As Integer, c_equity_db_industry As Integer, c_equity_db_sector_code As Integer
Dim c_equity_db_valeur As Integer, c_equity_db_daily_chg As Integer


c_equity_db_sector = 53
c_equity_db_industry = 54
c_equity_db_sector_code = 55

If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
    c_equity_db_valeur = 5
    c_equity_db_daily_chg = 10
    
    l_val_c1 = Worksheets("Parametres").Cells(23, 132).Value
    l_val_c2 = Worksheets("Parametres").Cells(23, 133).Value
ElseIf UCase(view) = "NAV" Then
    c_equity_db_valeur = 34
    c_equity_db_daily_chg = 35
    
    l_val_c1 = Worksheets("Parametres").Cells(24, 132).Value
    l_val_c2 = Worksheets("Parametres").Cells(24, 133).Value
End If








'preparation de la matrix
Dim vec_sector_code() As Variant
Dim vec_sector_name() As Variant
Dim vec_industry_name() As Variant
Dim vec_industry_sector_code() As Variant


Dim count_sector As Integer, count_industry As Integer
    count_sector = 0
    count_industry = 0
    
    ReDim vec_sector_code(count_sector)
    ReDim vec_sector_name(count_sector)
    ReDim vec_industry_name(count_industry)
    ReDim vec_industry_sector_code(count_industry)


Dim l_equity_db_header As Integer
l_equity_db_header = 25
For i = l_equity_db_header + 2 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" And Worksheets("Equity_Database").Cells(i + 2, 1) = "" And Worksheets("Equity_Database").Cells(i + 4, 1) = "" Then
        Exit For
    Else
        For j = 0 To UBound(vec_sector_name, 1)
            If Worksheets("Equity_Database").Cells(i, c_equity_db_sector) = vec_sector_name(j) Then
                Exit For
            Else
                If j = UBound(vec_sector_name, 1) Then
                    ReDim Preserve vec_sector_name(count_sector)
                    ReDim Preserve vec_sector_code(count_sector)
                    
                    vec_sector_name(count_sector) = Worksheets("Equity_Database").Cells(i, c_equity_db_sector)
                    vec_sector_code(count_sector) = Worksheets("Equity_Database").Cells(i, c_equity_db_sector_code)
                    
                    count_sector = count_sector + 1
                End If
            End If
        Next j
        
        For j = 0 To UBound(vec_industry_name, 1)
            If Worksheets("Equity_Database").Cells(i, c_equity_db_industry) = vec_industry_name(j) Then
                Exit For
            Else
                If j = UBound(vec_industry_name, 1) Then
                    ReDim Preserve vec_industry_name(count_industry)
                    ReDim Preserve vec_industry_sector_code(count_industry)
                    
                    vec_industry_name(count_industry) = Worksheets("Equity_Database").Cells(i, c_equity_db_industry)
                    vec_industry_sector_code(count_industry) = Worksheets("Equity_Database").Cells(i, c_equity_db_sector_code)
                    
                    count_industry = count_industry + 1
                End If
            End If
        Next j
        
    End If
Next i


'(sort des codes)



'construction de la matrix
Dim matrix_sector() As Variant
ReDim matrix_sector(UBound(vec_sector_name, 1))

Dim matrix_industry() As Variant
ReDim matrix_industry(UBound(vec_industry_name, 1))

Dim dim_name As Integer, dim_code As Integer, dim_valeur As Integer, dim_daily_chg As Integer, dim_valeur_long As Integer, _
    dim_valeur_short As Integer, dim_daily_chg_long As Integer, dim_daily_chg_short As Integer



dim_name = 0
dim_code = 1
dim_valeur = 2
dim_daily_chg = 3
dim_valeur_long = 4
dim_valeur_short = 5
dim_daily_chg_long = 6
dim_daily_chg_short = 7

For i = 0 To UBound(vec_sector_name, 1)
    matrix_sector(i) = Array(vec_sector_name(i), vec_sector_code(i), 0, 0, 0, 0, 0, 0)
Next i

For i = 0 To UBound(vec_industry_name, 1)
    matrix_industry(i) = Array(vec_industry_name(i), vec_industry_sector_code(i), 0, 0, 0, 0, 0, 0)
Next i


' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Dim date_tmp As Date, date_today As Date

date_today = Date

'chargement des filtres (pour commencer uniquement ceux du type colonne)
Dim test_debug As Variant
Dim c_filter_name As Integer
    c_filter_name = 138

Dim c_filter_criteria As Integer
    c_filter_criteria = 139

Dim l_filter_first_line As Integer
    l_filter_first_line = 11
        
        
Dim activate_filter As Boolean

If IsNumeric(Worksheets("Parametres").Cells(l_filter_first_line, c_filter_criteria)) And Worksheets("Parametres").Cells(l_filter_first_line, c_filter_criteria) = 1 Then
    activate_filter = True
Else
    activate_filter = False
End If



If activate_filter = True Then
    Application.Calculation = xlCalculationManual
    
    'repere la derniere ligne
    Dim l_filter_last_line As Integer
    For i = l_filter_first_line To 5000
        If Worksheets("Parametres").Cells(i, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 1, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 2, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 3, c_filter_name) = "" Then
            l_filter_last_line = i - 1
            Exit For
        End If
    Next i
    
    
    'repere les filters colonnes activés
    Dim filter() As Variant
    Dim filter_type() As Variant
    Dim filter_type_criteria() As Variant 'double, int, txt
    Dim filter_compare_type() As Variant ' <, > =, limit, abs
    Dim filter_column_equity_db() As Variant
    Dim filter_column_equity_idx() As Variant
    Dim filter_criteria() As Variant
    
    Dim count_ticker As Integer
        count_ticker = 0
    Dim list_tickers() As Variant
    Dim bbg_fld() As Variant
    Dim bbg_fld_custom() As Variant
    
    Dim vec_array_1() As Variant
    Dim vec_array_2() As Variant
    Dim vec_array_sector() As Variant
    
    Dim vec_ticker_sector() As Variant
    Dim vec_ticker_industry() As Variant
    
    Dim vec_valeur_eur() As Variant
    Dim vec_daily_chg() As Variant
    
    Dim nbre_filters_column As Integer, nbre_filters_custom As Integer, nbre_filters_api As Integer, _
        nbre_filters_api_custom As Integer
    
    
    Dim region As Variant
    region = Array(Array("Asia/Pacific", Array("JPY", "HKD", "AUD", "SGD", "TWD", "KRW", "INR", "THB")), Array("Europe", Array("CHF", "EUR", "GBP", "SEK", "NOK", "DKK", "PLN")), Array("America", Array("USD", "CAD", "BRL")))
    
    Dim currency_code() As Variant
    k = 0
    For j = 14 To 32
        If Worksheets("Parametres").Cells(j, 1) <> "" Then
            ReDim Preserve currency_code(k)
            currency_code(k) = Array(Left(Worksheets("Parametres").Cells(j, 1).Value, 3), Worksheets("Parametres").Cells(j, 5).Value)
            k = k + 1
        End If
    Next j
    
    j = 0
    nbre_filters_column = 0
    nbre_filters_custom = 0
    nbre_filters_api = 0
    nbre_filters_api_custom = 0
    
    For i = l_filter_first_line To l_filter_last_line
        If InStr(Worksheets("Parametres").Cells(i, c_filter_name), "col_") <> 0 And Worksheets("Parametres").Cells(i + 1, c_filter_criteria) <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "col_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "col"
            
            ReDim Preserve filter_type_criteria(j)
            If IsNumeric(Worksheets("Parametres").Cells(i + 2, c_filter_criteria)) Then
                filter_type_criteria(j) = "num"
            Else
                filter_type_criteria(j) = "str"
            End If
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            ReDim Preserve filter_compare_type(j)
            filter_compare_type(j) = Worksheets("Parametres").Cells(i + 2, c_filter_name)
            
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            
            j = j + 1
            nbre_filters_column = nbre_filters_column + 1
        
        ElseIf InStr(Worksheets("Parametres").Cells(i, c_filter_name), "custom_") <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "custom_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "custom"
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 1, c_filter_criteria)
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            If filter_criteria(j) <> 0 And UCase(Worksheets("Parametres").Cells(i + 2, c_filter_name)) = "CRITERIA" Then
                filter_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            End If
            
            j = j + 1
            nbre_filters_custom = nbre_filters_custom + 1
            
        ElseIf InStr(Worksheets("Parametres").Cells(i, c_filter_name), "api_") <> 0 And Worksheets("Parametres").Cells(i + 1, c_filter_criteria) <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "api_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "api"
            
            ReDim Preserve filter_type_criteria(j)
            filter_type_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 3, c_filter_criteria)
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            ReDim Preserve bbg_fld(nbre_filters_api)
            bbg_fld(nbre_filters_api) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "api_", "")
            
            
            nbre_filters_api = nbre_filters_api + 1
            
        End If
    Next i
    
    Dim nbre_filters As Integer
        nbre_filters = j - 1 + 1
    
    
    'remonte les colonnes concernees d'equity
    Dim l_equity_db_first_line As Integer
        l_equity_db_first_line = 27
    
    Dim line_step As Integer
    Dim c_equity_db_name As Integer, l_header_equity_db As Integer, c_equity_db_last_column As Integer
        c_equity_db_name = 2
        l_header_equity_db = 25
        line_step = 2
        
        c_equity_db_last_column = 250
    
    For i = 0 To UBound(filter)
        For j = 1 To c_equity_db_last_column
            If UCase(Replace(filter(i), "_", " ")) = UCase(Replace(Worksheets("Equity_Database").Cells(l_header_equity_db, j), "_", " ")) Then
                filter_column_equity_db(i) = j
                GoTo next_filter_column_equity_database
            End If
        Next j
next_filter_column_equity_database:
    Next i
    
    For i = l_equity_db_first_line To 5000 Step line_step
        Dim l_equity_db_last_line As Integer
        If Worksheets("Equity_Database").Cells(i, c_equity_db_name) = "" And Worksheets("Equity_Database").Cells(i + 1 * line_step, c_equity_db_name) = "" And Worksheets("Equity_Database").Cells(i + 2 * line_step, c_equity_db_name) = "" Then
            l_equity_db_last_line = i - line_step
            Exit For
        End If
    Next i
    

    
    'lance la recherche de titres
    
    
    'capte les titres qui rentre dans les criteres d'equity database
    Dim tmp_equity_filter_ok As Boolean
    
    count_ticker = 0
    For i = l_equity_db_first_line To l_equity_db_last_line Step 2
        
        On Error GoTo next_entry_equity_db:
        
        Dim tmp_equity_valeur_eur As Variant, tmp_equity_daily_change As Variant
        
        
        If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
            tmp_equity_valeur_eur = Worksheets("Equity_Database").Cells(i, g_col_e)
            tmp_equity_daily_change = Worksheets("Equity_Database").Cells(i, g_col_j)
        ElseIf UCase(view) = "NAV" Then
            tmp_equity_valeur_eur = Worksheets("Equity_Database").Cells(i, g_col_ah)
            tmp_equity_daily_change = Worksheets("Equity_Database").Cells(i, g_col_ai)
        End If
        
        If IsNumeric(tmp_equity_valeur_eur) And IsNumeric(tmp_equity_daily_change) Then
                
                'passe dans les filtres de colonnes
                tmp_equity_filter_ok = True
                For j = 0 To UBound(filter)
                    
                    If filter_type(j) = "col" Then
                    
                        ' la colonne existe dans equity DB
                        If tmp_equity_filter_ok = True And filter_column_equity_db(j) <> False And IsNumeric(filter_column_equity_db(j)) Then
                            
                            ' passe le check critere ou non
                            If filter_compare_type(j) = "=" And filter_type_criteria(j) = "str" And UCase(filter_criteria(j)) = UCase(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) Then
                                
                            ElseIf filter_compare_type(j) = "=" And filter_type_criteria(j) = "num" Then
                                If filter_criteria(j) = Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j)) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            ElseIf filter_compare_type(j) = "limit" And filter_type_criteria(j) = "num" Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) And filter_criteria(j) <= Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j)) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            ElseIf filter_compare_type(j) = "limit_abs" And filter_type_criteria(j) = "num" Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) And filter_criteria(j) <= Abs(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            Else
                                tmp_equity_filter_ok = False
                            End If
                            
                        Else
                            tmp_equity_filter_ok = False
                        End If
                
                
                    ElseIf filter_type(j) = "custom" Then
                    
                        ' #################################################################################################################
                        ' ############################ passe dans les filtres custom ######################################################
                        ' #################################################################################################################
                        If UCase(filter(j)) = "EQUITY" Then
                            If filter_criteria(j) = 0 Then
                                tmp_equity_filter_ok = False
                            End If
                        ElseIf UCase(filter(j)) = "SHORT" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 5)) And Worksheets("Equity_Database").Cells(i, 5) < 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "LONG" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 5)) And Worksheets("Equity_Database").Cells(i, 5) > 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "STOCK ONLY" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 9)) And Worksheets("Equity_Database").Cells(i, 9) = 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "LENDING" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 26)) And Worksheets("Equity_Database").Cells(i, 26) = 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "EXPECTED_REPORT_DT" Then
                            If filter_criteria(j) > 0 Then
                                nbre_filters_api_custom = nbre_filters_api_custom + 1
                            End If
                        ElseIf UCase(filter(j)) = "DVD_EX_DT" Then
                            If filter_criteria(j) > 0 Then
                                nbre_filters_api_custom = nbre_filters_api_custom + 1
                            End If
                        
                        ElseIf UCase(filter(j)) = "REGION" Then
                            For k = 0 To UBound(region, 1)
                                If region(k)(0) = filter_criteria(j) Then
                                    For m = 0 To UBound(currency_code)
                                        If Worksheets("Equity_Database").Cells(i, 44) = currency_code(m)(1) Then
                                            For n = 0 To UBound(region(k)(1))
                                                If currency_code(m)(0) = region(k)(1)(n) Then
                                                    GoTo bypass_filter_region_check
                                                Else
                                                    If n = UBound(region(k)(1)) Then
                                                        tmp_equity_filter_ok = False
                                                    End If
                                                End If
                                            Next n
                                        End If
                                    Next m
                                End If
                            Next k
bypass_filter_region_check:
                        ElseIf UCase(filter(j)) = "TAG" Then
                            
                        End If
                    
                    End If
                        
                Next j
                

                'le titre doit etre pris en compte
                If tmp_equity_filter_ok = True Then
                    
                    If nbre_filters_api = 0 And nbre_filters_api_custom = 0 Then
                    
                        'insère les données dans les matrix sector / industry
                        For k = 0 To UBound(matrix_sector, 1)
                            If matrix_sector(k)(dim_name) = Worksheets("Equity_Database").Cells(i, c_equity_db_sector) Then
                                
                                matrix_sector(k)(dim_valeur) = matrix_sector(k)(dim_valeur) + Worksheets("Equity_Database").Cells(i, c_equity_db_valeur)
                                matrix_sector(k)(dim_daily_chg) = matrix_sector(k)(dim_daily_chg) + Worksheets("Equity_Database").Cells(i, c_equity_db_daily_chg)
                                
                                If Worksheets("Equity_Database").Cells(i, c_equity_db_valeur) >= 0 Then
                                    'long
                                    matrix_sector(k)(dim_valeur_long) = matrix_sector(k)(dim_valeur_long) + Worksheets("Equity_Database").Cells(i, c_equity_db_valeur)
                                    matrix_sector(k)(dim_daily_chg_long) = matrix_sector(k)(dim_daily_chg_long) + Worksheets("Equity_Database").Cells(i, c_equity_db_daily_chg)
                                Else
                                    'short
                                    matrix_sector(k)(dim_valeur_short) = matrix_sector(k)(dim_valeur_short) + Worksheets("Equity_Database").Cells(i, c_equity_db_valeur)
                                    matrix_sector(k)(dim_daily_chg_short) = matrix_sector(k)(dim_daily_chg_short) + Worksheets("Equity_Database").Cells(i, c_equity_db_daily_chg)
                                End If
                                
                                Exit For
                            End If
                        Next k
                        
                        
                        For k = 0 To UBound(matrix_industry, 1)
                            If matrix_industry(k)(dim_name) = Worksheets("Equity_Database").Cells(i, c_equity_db_industry) Then
                                
                                matrix_industry(k)(dim_valeur) = matrix_industry(k)(dim_valeur) + Worksheets("Equity_Database").Cells(i, c_equity_db_valeur)
                                matrix_industry(k)(dim_daily_chg) = matrix_industry(k)(dim_daily_chg) + Worksheets("Equity_Database").Cells(i, c_equity_db_daily_chg)
                                
                                If Worksheets("Equity_Database").Cells(i, c_equity_db_valeur) >= 0 Then
                                    'long
                                    matrix_industry(k)(dim_valeur_long) = matrix_industry(k)(dim_valeur_long) + Worksheets("Equity_Database").Cells(i, c_equity_db_valeur)
                                    matrix_industry(k)(dim_daily_chg_long) = matrix_industry(k)(dim_daily_chg_long) + Worksheets("Equity_Database").Cells(i, c_equity_db_daily_chg)
                                Else
                                    'short
                                    matrix_industry(k)(dim_valeur_short) = matrix_industry(k)(dim_valeur_short) + Worksheets("Equity_Database").Cells(i, c_equity_db_valeur)
                                    matrix_industry(k)(dim_daily_chg_short) = matrix_industry(k)(dim_daily_chg_short) + Worksheets("Equity_Database").Cells(i, c_equity_db_daily_chg)
                                End If
                                
                                Exit For
                            End If
                        Next k
                        
                    
                    Else
                        
                        ReDim Preserve list_tickers(count_ticker)
                        list_tickers(count_ticker) = Worksheets("Equity_Database").Cells(i, g_col_au)
                        
                        ReDim Preserve vec_ticker_sector(count_ticker)
                        vec_ticker_sector(count_ticker) = Worksheets("Equity_Database").Cells(i, c_equity_db_sector)
                        
                        ReDim Preserve vec_ticker_industry(count_ticker)
                        vec_ticker_industry(count_ticker) = Worksheets("Equity_Database").Cells(i, c_equity_db_industry)
                        
                        ReDim Preserve vec_valeur_eur(count_ticker)
                        vec_valeur_eur(count_ticker) = Worksheets("Equity_Database").Cells(i, c_equity_db_valeur)
                        
                        ReDim Preserve vec_daily_chg(count_ticker)
                        vec_daily_chg(count_ticker) = Worksheets("Equity_Database").Cells(i, c_equity_db_daily_chg)
                        
                        count_ticker = count_ticker + 1
                        
                    End If
                    
                End If
                
        End If
next_entry_equity_db:
    Next i
    
    
    On Error GoTo 0
    
    Dim pass_filter_bbg As Boolean
    
    If count_ticker > 0 Then
        
        If nbre_filters_api > 0 Then
            Dim output_bbg As Variant
            output_bbg = bbg_multi_tickers_and_multi_fields(list_tickers, bbg_fld)
        End If
        
        If nbre_filters_api_custom > 0 Then
            Dim output_bbg_custom As Variant
            bbg_fld_custom = Array("EXPECTED_REPORT_DT", "DVD_EX_DT")
            
            output_bbg_custom = bbg_multi_tickers_and_multi_fields(list_tickers, bbg_fld_custom)
        End If
            
                            
            For i = 0 To UBound(list_tickers, 1)
            
                pass_filter_bbg = True
                
                If nbre_filters_api > 0 Then
                    For j = 0 To UBound(bbg_fld, 1)
                        For k = 0 To UBound(filter, 1)
                            If filter_type(k) = "api" And UCase(bbg_fld(j)) = UCase(filter(k)) Then
                                
                                If Left(output_bbg(i, j), 1) <> "#" Then
                                    If filter_type_criteria(k) = "num" Then
                                        If output_bbg(i, j) >= filter_criteria(k) Then
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    ElseIf filter_type_criteria(k) = "str" Then
                                        If output_bbg(i, j) = filter_criteria(k) Then
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    End If
                                Else
                                    pass_filter_bbg = False
                                End If
                                
                            End If
                        Next k
                    Next j
                End If
                
                
                
                If pass_filter_bbg = True And nbre_filters_api_custom > 0 Then
                    
                    For j = 0 To UBound(bbg_fld_custom, 1)
                        For k = 0 To UBound(filter, 1)
                            If filter_type(k) = "custom" And UCase(bbg_fld_custom(j)) = UCase(filter(k)) Then
                                If Left(output_bbg_custom(i, j), 1) <> "#" Then
                                    
                                    'code custom
                                    If UCase(filter(k)) = "EXPECTED_REPORT_DT" Then
                                        date_tmp = Mid(output_bbg_custom(i, j), 4, 2) & "." & Left(output_bbg_custom(i, j), 2) & "." & Right(output_bbg_custom(i, j), 4)
                                        
                                        test_debug = date_tmp - Date
                                        
                                        If date_tmp - Date <= filter_criteria(k) And date_tmp > Date Then
                                            
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    ElseIf UCase(filter(k)) = "DVD_EX_DT" Then
                                        date_tmp = Mid(output_bbg_custom(i, j), 4, 2) & "." & Left(output_bbg_custom(i, j), 2) & "." & Right(output_bbg_custom(i, j), 4)
                                        
                                        test_debug = date_tmp - Date
                                        
                                        If date_tmp - Date <= filter_criteria(k) And date_tmp > Date Then
                                            
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    End If
                                    
                                End If
                            End If
                        Next k
                    Next j
                    
                End If
                
                
                
                If pass_filter_bbg = True Then
                    
                    'a entrer dans les 2 matrix sector / industry
                    For k = 0 To UBound(matrix_sector, 1)
                        If matrix_sector(k)(dim_name) = vec_ticker_sector(i) Then
                            
                            matrix_sector(k)(dim_valeur) = matrix_sector(k)(dim_valeur) + vec_valeur_eur(i)
                            matrix_sector(k)(dim_daily_chg) = matrix_sector(k)(dim_daily_chg) + vec_daily_chg(i)
                            
                            If vec_valeur_eur(i) >= 0 Then
                                'long
                                matrix_sector(k)(dim_valeur_long) = matrix_sector(k)(dim_valeur_long) + vec_valeur_eur(i)
                                matrix_sector(k)(dim_daily_chg_long) = matrix_sector(k)(dim_daily_chg_long) + vec_daily_chg(i)
                            Else
                                'short
                                matrix_sector(k)(dim_valeur_short) = matrix_sector(k)(dim_valeur_short) + vec_valeur_eur(i)
                                matrix_sector(k)(dim_daily_chg_short) = matrix_sector(k)(dim_daily_chg_short) + vec_daily_chg(i)
                            End If
                            
                            Exit For
                        End If
                    Next k
                    
                    
                    For k = 0 To UBound(matrix_industry, 1)
                        If matrix_industry(k)(dim_name) = vec_ticker_industry(i) Then
                            
                            matrix_industry(k)(dim_valeur) = matrix_industry(k)(dim_valeur) + vec_valeur_eur(i)
                            matrix_industry(k)(dim_daily_chg) = matrix_industry(k)(dim_daily_chg) + vec_daily_chg(i)
                            
                            If vec_valeur_eur(i) >= 0 Then
                                'long
                                matrix_industry(k)(dim_valeur_long) = matrix_industry(k)(dim_valeur_long) + vec_valeur_eur(i)
                                matrix_industry(k)(dim_daily_chg_long) = matrix_industry(k)(dim_daily_chg_long) + vec_daily_chg(i)
                            Else
                                'short
                                matrix_industry(k)(dim_valeur_short) = matrix_industry(k)(dim_valeur_short) + vec_valeur_eur(i)
                                matrix_industry(k)(dim_daily_chg_short) = matrix_industry(k)(dim_daily_chg_short) + vec_daily_chg(i)
                            End If
                            
                            Exit For
                        End If
                    Next k
                    
                    
                    
                End If
                
            Next i
            
            
            
        
    End If
    
    
    On Error GoTo 0
    
    
    'prepare les vecteurs du graphique
    For i = 0 To UBound(matrix_sector, 1)
        l_val_1 = matrix_sector(i)(dim_valeur)
        l_val_2 = matrix_sector(i)(dim_daily_chg)
        
        If Abs(l_val_1) >= l_val_c1 And Abs(l_val_2) >= l_val_c2 Then
            ReDim Preserve l_array_x1(l_array_index)
            ReDim Preserve l_array_y1(l_array_index)
            ReDim Preserve l_array_y2(l_array_index)
            ReDim Preserve l_array_y3(l_array_index)
            ReDim Preserve l_array_y5(l_array_index)
            ReDim Preserve l_array_z1(l_array_index)
            
            l_array_x1(l_array_index) = matrix_sector(i)(dim_name) 'sector/industry
            l_array_y1(l_array_index) = matrix_sector(i)(dim_valeur) 'valeur
            l_array_y2(l_array_index) = matrix_sector(i)(dim_daily_chg) 'daily chg
            l_array_y3(l_array_index) = matrix_sector(i)(dim_valeur_long)  'long valeur eur
            l_array_y5(l_array_index) = matrix_sector(i)(dim_valeur_short)  'short valeur eur
            l_array_z1(l_array_index) = matrix_sector(i)(dim_code) 'code
            
            l_array_index = l_array_index + 1
        End If
    Next i
    
    For i = 0 To UBound(matrix_industry, 1)
        l_val_1 = matrix_industry(i)(dim_valeur)
        l_val_2 = matrix_industry(i)(dim_daily_chg)
        
        If Abs(l_val_1) >= l_val_c1 And Abs(l_val_2) >= l_val_c2 Then
            ReDim Preserve l_array_x1(l_array_index)
            ReDim Preserve l_array_y1(l_array_index)
            ReDim Preserve l_array_y2(l_array_index)
            ReDim Preserve l_array_y3(l_array_index)
            ReDim Preserve l_array_y5(l_array_index)
            ReDim Preserve l_array_z1(l_array_index)
            
            l_array_x1(l_array_index) = matrix_industry(i)(dim_name) 'sector/industry
            l_array_y1(l_array_index) = matrix_industry(i)(dim_valeur) 'valeur
            l_array_y2(l_array_index) = matrix_industry(i)(dim_daily_chg) 'daily chg
            l_array_y3(l_array_index) = matrix_industry(i)(dim_valeur_long)  'long valeur eur
            l_array_y5(l_array_index) = matrix_industry(i)(dim_valeur_short)  'short valeur eur
            l_array_z1(l_array_index) = matrix_industry(i)(dim_code) 'code
            
            l_array_index = l_array_index + 1
        End If
    Next i

Else 'fonction habituelle

    With Worksheets("Simulation")
        l_rows = .Cells(15, g_col_aa).Value
        .Range("AA17:IV" & l_rows).rows.Hidden = False
        
        'ANDREASSON LENNART
        For l_row = 17 To l_rows Step 1
            l_val_1 = .Cells(l_row, g_col_ac).Value
            l_val_2 = .Cells(l_row, g_col_ae).Value
            If IsNumeric(l_val_1) And IsNumeric(l_val_2) Then
                If Abs(l_val_1) >= l_val_c1 And Abs(l_val_2) >= l_val_c2 Then
                    ReDim Preserve l_array_x1(l_array_index)
                    ReDim Preserve l_array_y1(l_array_index)
                    ReDim Preserve l_array_y2(l_array_index)
                    ReDim Preserve l_array_y3(l_array_index)
                    ReDim Preserve l_array_y5(l_array_index)
                    ReDim Preserve l_array_z1(l_array_index)
                    
                    l_array_x1(l_array_index) = .Cells(l_row, g_col_aa).Value ' sector / indsutry
                    l_array_y1(l_array_index) = "Simulation!R" & l_row & "C" & g_col_ac & "" 'valeur_eur
                    l_array_y2(l_array_index) = "Simulation!R" & l_row & "C" & g_col_ae & "" 'daily chg
                    l_array_y3(l_array_index) = "Simulation!R" & l_row & "C" & g_col_af & "" 'long valeur eur
                    l_array_y5(l_array_index) = "Simulation!R" & l_row & "C" & g_col_ag & "" 'short valeur eur
                    l_array_z1(l_array_index) = .Cells(l_row, g_col_ab).Value 'sector code
                    l_array_index = l_array_index + 1
                End If
            End If
        Next l_row
    End With
    
End If
        
        
        
' construction du graph avec la methode standard
l_rows = l_array_index

Set l_xls_sheet = Worksheets("Chart_Database")
    With l_xls_sheet
        .Range("EA13:EZ2000").ClearContents
        
        For l_row = 0 To l_rows - 1 Step 1
            .Cells(l_row + 13, g_col_ea).Value = l_array_x1(l_row)
            .Cells(l_row + 13, g_col_eb).Value = "=" & l_array_y1(l_row)
            .Cells(l_row + 13, g_col_ec).Value = "=" & l_array_y2(l_row)
            .Cells(l_row + 13, g_col_ed).Value = "=" & l_array_y3(l_row)
            .Cells(l_row + 13, g_col_ef).Value = "=" & l_array_y5(l_row)
            .Cells(l_row + 13, g_col_ez).Value = l_array_z1(l_row)
        Next l_row
    End With


Set l_xls_chart = Charts("Chart Sectors")
Application.ScreenUpdating = False
With l_xls_chart
    .Activate
    .ChartArea.ClearContents
    l_chart_rows = .SeriesCollection.count
    
    For l_chart_row = 1 To l_chart_rows Step 1
        .SeriesCollection(1).Delete
    Next l_chart_row
    
    Set l_xls_series = .SeriesCollection.NewSeries
    l_xls_series.name = l_xls_sheet.Range("EB12")
    l_xls_series.XValues = l_xls_sheet.Range("EA13:EA" & l_rows + 12 & "")
    l_xls_series.Values = l_xls_sheet.Range("EB13:EB" & l_rows + 12 & "")
    
    If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
        l_xls_series.DataLabels.NumberFormat = "#,##0.00"
    ElseIf UCase(view) = "NAV" Then
        l_xls_series.DataLabels.NumberFormat = "0.00%"
    End If
    
    Set l_xls_series = Nothing
    
    
    
    Set l_xls_series = .SeriesCollection.NewSeries
    l_xls_series.name = l_xls_sheet.Range("EC12")
    l_xls_series.Values = l_xls_sheet.Range("EC13:EC" & l_rows + 12 & "")
    Set l_xls_series = Nothing
    
    
    
    Set l_xls_series = .SeriesCollection.NewSeries
    l_xls_series.name = l_xls_sheet.Range("ED12")
    l_xls_series.Values = l_xls_sheet.Range("ED13:ED" & l_rows + 12 & "")
    
    For l_row = 0 To l_rows - 1 Step 1
        l_xls_series.Points(l_row + 1).Interior.Pattern = xlSolid
        l_xls_series.Points(l_row + 1).Interior.ColorIndex = format_Chart_ColorIndex(l_xls_sheet.Cells(l_row + 13, g_col_ez).Value)
    Next l_row
    
    If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
        l_xls_series.DataLabels.NumberFormat = "#,##0.00"
    ElseIf UCase(view) = "NAV" Then
        l_xls_series.DataLabels.NumberFormat = "0.00%"
    End If
    
    Set l_xls_series = Nothing
    
    
    
    Set l_xls_series = .SeriesCollection.NewSeries
    l_xls_series.name = l_xls_sheet.Range("EF12")
    l_xls_series.Values = l_xls_sheet.Range("EF13:EF" & l_rows + 12 & "")
          
    For l_row = 0 To l_rows - 1 Step 1
        l_xls_series.Points(l_row + 1).Interior.Pattern = xlSolid
        l_xls_series.Points(l_row + 1).Interior.ColorIndex = format_Chart_ColorIndex(l_xls_sheet.Cells(l_row + 13, g_col_ez).Value)
    Next l_row
    
    
    Set l_xls_series = Nothing
    
End With

Set l_xls_chart = Nothing
Set l_xls_sheet = Nothing

Application.ScreenUpdating = True

End Sub


Public Sub load_Chart_Rel_Perf_live(ByVal Chart As String, ByVal view As String)


'greg chevalley
' nouvelle fonction avec support des filtres colonne EH sheet parameters

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim test_debug As Variant, debug_test As Variant

'charge la config des graphs
Dim chars_config(10) As Variant

' 0 - chart / 1 - array(critere) / ( 2 - parameters_line_limit ) / 3 - array(vec1_column) / 4 - array(vec2_column) / 5 - array(vec_sector_column) / 6 - array (vec -> destination) / 7 - sheet_chart
chars_config(0) = Array("Rel Perf", Array(Array("Valeur_Euro", "Daily Change"), Array("Nav Position", "Nav Daily")), Array(Array(14, 132), Array(14, 133)), Array("Short Name", Array("Valeur_Euro", "Nav Position")), Array("Short Name", Array("Result Total", "Nav Daily")), Array("Short Name", "Sector"), Array(Array(27, 28), Array("", 29), Array("", 52)), "Chart Perf", Array(Array("bbg", "rel_1d", 0)))

'repère la config concernée par l'appel
Dim config_chart As Integer
For i = 0 To UBound(chars_config, 1)
    If chars_config(i)(0) = Chart Then
        config_chart = i
        Exit For
    End If
Next i

Dim tmp_colonne_criteria_1 As String, tmp_colonne_criteria_2 As String
'selon le mode (book/nav) arrange la colonne
If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
    
    'colonne criteria
    tmp_colonne_criteria_1 = chars_config(config_chart)(1)(0)(0)
    tmp_colonne_criteria_2 = chars_config(config_chart)(1)(0)(1)
    
    chars_config(config_chart)(1)(0) = tmp_colonne_criteria_1
    chars_config(config_chart)(1)(1) = tmp_colonne_criteria_2
    
    'colonne a grapher
    chars_config(config_chart)(3)(1) = chars_config(config_chart)(3)(1)(0)
    chars_config(config_chart)(4)(1) = chars_config(config_chart)(4)(1)(0)
ElseIf UCase(view) = "NAV" Then
    'colonne criteria
    tmp_colonne_criteria_1 = chars_config(config_chart)(1)(1)(0)
    tmp_colonne_criteria_2 = chars_config(config_chart)(1)(1)(1)
    
    chars_config(config_chart)(1)(0) = tmp_colonne_criteria_1
    chars_config(config_chart)(1)(1) = tmp_colonne_criteria_2
    
    'colonne a grapher
    chars_config(config_chart)(3)(1) = chars_config(config_chart)(3)(1)(1)
    chars_config(config_chart)(4)(1) = chars_config(config_chart)(4)(1)(1)
Else
    'default = book
    chars_config(config_chart)(3)(1) = chars_config(config_chart)(3)(1)(0)
    chars_config(config_chart)(4)(1) = chars_config(config_chart)(4)(1)(0)
End If


Dim oJSON As New JSONLib
Dim oTags As Collection, oTag As Collection


Dim l_rows As Long, l_row As Long

Dim l_val_c1
Dim l_val_c2

Dim l_val_1
Dim l_val_2

Dim l_array_valeur_eur()
Dim l_array_daily_pnl()
Dim l_array_nav_pos()
Dim l_array_nav_daily()
Dim l_array_sectors()
Dim l_array_rel_perf()
Dim l_array_rel_index()
Dim l_array_line()
Dim l_array_index As Long

Dim l_xls_sheet As Worksheet
Dim l_xls_chart As Chart
Dim l_xls_series As Series

Dim l_chart_rows As Long, l_chart_row As Long

Dim date_tmp As Date, date_today As Date
date_today = Date


Dim c_parameters_chart_name As Integer, c_parameters_chart_criteria_1 As Integer, c_parameters_criteria_2 As Integer, _
    l_parameters_chart_first_line As Integer
    
    l_parameters_chart_first_line = 12
    c_parameters_chart_name = 131
    c_parameters_chart_criteria_1 = 132
    c_parameters_criteria_2 = 133

Dim limit_chart_criteria_1 As Double, limit_chart_criteria_2 As Double

Dim l_chart_database_header As Integer
l_chart_database_header = 12

'repere les limites
Dim l_parameters_chart_config_line As Integer
For i = l_parameters_chart_first_line To 100
    If Worksheets("Parametres").Cells(i, c_parameters_chart_name) = "" And Worksheets("Parametres").Cells(i + 1, c_parameters_chart_name) = "" And Worksheets("Parametres").Cells(i + 2, c_parameters_chart_name) = "" Then
        Exit For
    Else
        If InStr(Chart, " ") <> 0 Then
            If InStr(Worksheets("Parametres").Cells(i, c_parameters_chart_name), Left(Chart, InStr(Chart, " ") - 1)) <> 0 Then
                l_parameters_chart_config_line = i
                Exit For
            End If
        Else
            If InStr(Worksheets("Parametres").Cells(i, c_parameters_chart_name), Chart) <> 0 Then
                l_parameters_chart_config_line = i
                Exit For
            End If
        End If
    End If
Next i






If UCase(view) = "BOOK" Or UCase(view) = "NOMINAL" Then
    limit_chart_criteria_1 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_chart_criteria_1).Value
    limit_chart_criteria_2 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_criteria_2).Value
ElseIf UCase(view) = "NAV" Then
    limit_chart_criteria_1 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 2, c_parameters_chart_criteria_1).Value
    limit_chart_criteria_2 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 2, c_parameters_criteria_2).Value
    
    'limit_chart_criteria_1 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_chart_criteria_1).Value
    'limit_chart_criteria_2 = Worksheets("Parametres").Cells(l_parameters_chart_config_line + 1, c_parameters_criteria_2).Value
Else
    Exit Sub
End If



Dim l_equity_db_first_line As Integer
    l_equity_db_first_line = 27
    
    Dim line_step As Integer
    Dim c_equity_db_name As Integer, l_header_equity_db As Integer, c_equity_db_last_column As Integer, _
        c_equity_db_criteria_1 As Integer, c_equity_db_criteria_2 As Integer
        
        c_equity_db_name = 2
        Dim c_equity_db_criteria() As Variant
            ReDim c_equity_db_criteria(UBound(chars_config(config_chart)(1), 1))
        
        Dim c_equity_db_vect_1() As Variant
            ReDim c_equity_db_vect_1(UBound(chars_config(config_chart)(3), 1))
        
        Dim c_equity_db_vect_2() As Variant
            ReDim c_equity_db_vect_2(UBound(chars_config(config_chart)(4), 1))
        
        Dim c_equity_db_vect_3() As Variant
            ReDim c_equity_db_vect_3(UBound(chars_config(config_chart)(5), 1))
        
        c_equity_db_criteria_1 = 0
        c_equity_db_criteria_2 = 0
        l_header_equity_db = 25
        line_step = 2
        
        c_equity_db_last_column = 250

Dim l_index_db_first_line As Integer
    l_index_db_first_line = 27
    
    
    Dim l_header_index_db As Integer, c_index_db_last_column As Integer, c_index_db_criteria_1 As Integer, _
        c_index_db_criteria_2 As Integer
    
    c_equity_db_name = 2
    Dim c_index_db_criteria() As Variant
        ReDim c_index_db_criteria(UBound(chars_config(config_chart)(1), 1))
    
    
    Dim c_index_db_vect_1() As Variant
            ReDim c_index_db_vect_1(UBound(chars_config(config_chart)(3), 1))
        
        Dim c_index_db_vect_2() As Variant
            ReDim c_index_db_vect_2(UBound(chars_config(config_chart)(4), 1))
        
        Dim c_index_db_vect_3() As Variant
            ReDim c_index_db_vect_3(UBound(chars_config(config_chart)(5), 1))
    
    
        
    c_index_db_criteria_1 = 0
    c_index_db_criteria_2 = 0
    l_header_index_db = 25
    line_step = 2
    
    c_index_db_last_column = 250



'repère les colonnes dans equity_db & index_db pour les limites
For i = 1 To c_equity_db_last_column
    For j = 0 To UBound(chars_config(config_chart)(1), 1)
        If UCase(chars_config(config_chart)(1)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
            c_equity_db_criteria(j) = i
        End If
    Next j
    
    'vector criteria 1
    For j = 0 To UBound(chars_config(config_chart)(3), 1)
        If chars_config(config_chart)(3)(j) <> "" Then
            If UCase(chars_config(config_chart)(3)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
                c_equity_db_vect_1(j) = i
            End If
        Else
            c_equity_db_vect_1(j) = 0
        End If
    Next j
    
    'vec criteria 2
    For j = 0 To UBound(chars_config(config_chart)(4), 1)
        If chars_config(config_chart)(4)(j) <> "" Then
            If UCase(chars_config(config_chart)(4)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
                c_equity_db_vect_2(j) = i
            End If
        Else
            c_equity_db_vect_2(j) = 0
        End If
    Next j
    
    'vect sector
    For j = 0 To UBound(chars_config(config_chart)(5), 1)
        If chars_config(config_chart)(5)(j) <> "" Then
            If UCase(chars_config(config_chart)(5)(j)) = UCase(Worksheets("Equity_Database").Cells(l_header_equity_db, i)) Then
                c_equity_db_vect_3(j) = i
            End If
        Else
            c_equity_db_vect_3(j) = 0
        End If
    Next j
    
Next i


For i = 1 To c_index_db_last_column
    For j = 0 To UBound(chars_config(config_chart)(1), 1)
        If UCase(chars_config(config_chart)(1)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
            c_index_db_criteria(j) = i
        End If
    Next j
    
    
    
    'vector criteria 1
    For j = 0 To UBound(chars_config(config_chart)(3), 1)
        If chars_config(config_chart)(3)(j) <> "" Then
            If UCase(chars_config(config_chart)(3)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
                c_index_db_vect_1(j) = i
            End If
        Else
            c_index_db_vect_1(j) = 0
        End If
    Next j
    
    'vec criteria 2
    For j = 0 To UBound(chars_config(config_chart)(4), 1)
        If chars_config(config_chart)(4)(j) <> "" Then
            If UCase(chars_config(config_chart)(4)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
                c_index_db_vect_2(j) = i
            End If
        Else
            c_index_db_vect_2(j) = 0
        End If
    Next j
    
    'vect sector
    For j = 0 To UBound(chars_config(config_chart)(5), 1)
        If chars_config(config_chart)(5)(j) <> "" Then
            If UCase(chars_config(config_chart)(5)(j)) = UCase(Worksheets("Index_Database").Cells(l_header_index_db, i)) Then
                c_index_db_vect_3(j) = i
            End If
        Else
            c_index_db_vect_3(j) = 0
        End If
    Next j
    
    
Next i




'chargement des filtres (pour commencer uniquement ceux du type colonne)
'filtre
Dim c_filter_name As Integer
    c_filter_name = 138

Dim c_filter_criteria As Integer
    c_filter_criteria = 139

Dim l_filter_first_line As Integer
    l_filter_first_line = 11
        
        
Dim activate_filter As Boolean

If IsNumeric(Worksheets("Parametres").Cells(l_filter_first_line, c_filter_criteria)) And Worksheets("Parametres").Cells(l_filter_first_line, c_filter_criteria) = 1 Then
    activate_filter = True
Else
    activate_filter = False
End If



If activate_filter = True Then
    Application.Calculation = xlCalculationManual
    
    'repere la derniere ligne
    Dim l_filter_last_line As Integer
    For i = l_filter_first_line To 5000
        If Worksheets("Parametres").Cells(i, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 1, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 2, c_filter_name) = "" And Worksheets("Parametres").Cells(i + 3, c_filter_name) = "" Then
            l_filter_last_line = i - 1
            Exit For
        End If
    Next i
    
    
    'repere les filters colonnes activés
    Dim filter() As Variant
    Dim filter_type() As Variant
    Dim filter_type_criteria() As Variant 'double, int, txt
    Dim filter_compare_type() As Variant ' <, > =, limit, abs
    Dim filter_column_equity_db() As Variant
    Dim filter_column_equity_idx() As Variant
    Dim filter_criteria() As Variant
    
    Dim count_ticker As Integer
        count_ticker = 0
    Dim list_tickers() As Variant
    Dim bbg_fld() As Variant
    Dim bbg_fld_custom() As Variant
    
    Dim vec_array_1() As Variant
    Dim vec_array_2() As Variant
    Dim vec_array_3() As Variant
    Dim vec_array_4() As Variant
    Dim vec_array_5() As Variant
    Dim vec_array_6() As Variant
    
    Dim vec_array_all() As Variant
    
    Dim vec_array_sector() As Variant
    Dim vec_array_line() As Variant
    
    Dim nbre_filters_column As Integer, nbre_filters_custom As Integer, nbre_filters_api As Integer, _
        nbre_filters_api_custom As Integer
    
    
    Dim region As Variant
    region = Array(Array("Asia/Pacific", Array("JPY", "HKD", "AUD", "SGD", "TWD", "KRW", "INR", "THB")), Array("Europe", Array("CHF", "EUR", "GBP", "SEK", "NOK", "DKK", "PLN")), Array("America", Array("USD", "CAD", "BRL")))
    
    Dim currency_code() As Variant
    k = 0
    For j = 14 To 32
        If Worksheets("Parametres").Cells(j, 1) <> "" Then
            ReDim Preserve currency_code(k)
            currency_code(k) = Array(Left(Worksheets("Parametres").Cells(j, 1).Value, 3), Worksheets("Parametres").Cells(j, 5).Value)
            k = k + 1
        End If
    Next j
    
    
    j = 0
    nbre_filters_column = 0
    nbre_filters_custom = 0
    nbre_filters_api = 0
    nbre_filters_api_custom = 0
    
    For i = l_filter_first_line To l_filter_last_line
        If InStr(Worksheets("Parametres").Cells(i, c_filter_name), "col_") <> 0 And Worksheets("Parametres").Cells(i + 1, c_filter_criteria) <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "col_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "col"
            
            ReDim Preserve filter_type_criteria(j)
            If IsNumeric(Worksheets("Parametres").Cells(i + 2, c_filter_criteria)) Then
                filter_type_criteria(j) = "num"
            Else
                filter_type_criteria(j) = "str"
            End If
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            ReDim Preserve filter_compare_type(j)
            filter_compare_type(j) = Worksheets("Parametres").Cells(i + 2, c_filter_name)
            
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            
            j = j + 1
            nbre_filters_column = nbre_filters_column + 1
        
        ElseIf InStr(Worksheets("Parametres").Cells(i, c_filter_name), "custom_") <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "custom_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "custom"
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 1, c_filter_criteria)
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            If filter_criteria(j) <> 0 And UCase(Worksheets("Parametres").Cells(i + 2, c_filter_name)) = "CRITERIA" Then
                filter_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            End If
            
            j = j + 1
            nbre_filters_custom = nbre_filters_custom + 1
            
        ElseIf InStr(Worksheets("Parametres").Cells(i, c_filter_name), "api_") <> 0 And Worksheets("Parametres").Cells(i + 1, c_filter_criteria) <> 0 Then
            
            ReDim Preserve filter(j)
            filter(j) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "api_", "")
            
            ReDim Preserve filter_type(j)
            filter_type(j) = "api"
            
            ReDim Preserve filter_type_criteria(j)
            filter_type_criteria(j) = Worksheets("Parametres").Cells(i + 2, c_filter_criteria)
            
            ReDim Preserve filter_criteria(j)
            filter_criteria(j) = Worksheets("Parametres").Cells(i + 3, c_filter_criteria)
            
            ReDim Preserve filter_column_equity_db(j)
            filter_column_equity_db(j) = False
            
            ReDim Preserve filter_column_equity_idx(j)
            filter_column_equity_idx(j) = False
            
            ReDim Preserve bbg_fld(nbre_filters_api)
            bbg_fld(nbre_filters_api) = Replace(Worksheets("Parametres").Cells(i, c_filter_name), "api_", "")
            
            j = j + 1
            nbre_filters_api = nbre_filters_api + 1
            
        End If
    Next i
    
    Dim nbre_filters As Integer
        nbre_filters = j - 1 + 1
    
    
    'remonte les colonnes concernees d'equity & index database
    For i = 0 To UBound(filter)
        For j = 1 To c_equity_db_last_column
            If UCase(Replace(filter(i), "_", " ")) = UCase(Replace(Worksheets("Equity_Database").Cells(l_header_equity_db, j), "_", " ")) Then
                filter_column_equity_db(i) = j
                GoTo next_filter_column_equity_database
            End If
        Next j
next_filter_column_equity_database:
    Next i
    
    For i = l_equity_db_first_line To 5000 Step line_step
        Dim l_equity_db_last_line As Integer
        If Worksheets("Equity_Database").Cells(i, c_equity_db_name) = "" And Worksheets("Equity_Database").Cells(i + 1 * line_step, c_equity_db_name) = "" And Worksheets("Equity_Database").Cells(i + 2 * line_step, c_equity_db_name) = "" Then
            l_equity_db_last_line = i - line_step
            Exit For
        End If
    Next i
    
    
    'remonte les colonnes conernees d'index database
    For i = 0 To UBound(filter)
        For j = 1 To c_index_db_last_column
            If UCase(Replace(filter(i), "_", " ")) = UCase(Replace(Worksheets("Index_Database").Cells(l_header_index_db, j), "_", " ")) Then
                filter_column_equity_idx(i) = j
                GoTo next_filter_column_index_database
            End If
        Next j
next_filter_column_index_database:
    Next i
    
    
    
    Dim c_index_db_name As Integer
    c_index_db_name = 2
    line_step = 3
    For i = l_equity_db_first_line To 5000 Step line_step
        Dim l_index_db_last_line As Integer
        If Worksheets("Index_Database").Cells(i, c_index_db_name) = "" And Worksheets("Index_Database").Cells(i + 1 * line_step, c_index_db_name) = "" And Worksheets("Index_Database").Cells(i + 2 * line_step, c_index_db_name) = "" Then
            l_index_db_last_line = i - line_step
            Exit For
        End If
    Next i
    
    
    
    'LANCE LA RECHERCHE DE TITRES
    
    'capte les titres qui rentre dans les criteres d'equity database
    Dim tmp_equity_filter_ok As Boolean
    
    Dim vec_ticker() As Variant
        Dim count_all_tickers As Integer
        count_all_tickers = 0
    
    count_ticker = 0
    For i = l_equity_db_first_line To l_equity_db_last_line Step 2
        
        On Error GoTo next_entry_equity_db:
        
        Dim limit_equity_value_criteria_1 As Variant, limit_equity_value_criteria_2 As Variant
        
        
        limit_equity_value_criteria_1 = Worksheets("Equity_Database").Cells(i, c_equity_db_criteria(0))
        limit_equity_value_criteria_2 = Worksheets("Equity_Database").Cells(i, c_equity_db_criteria(1))
        
        If IsNumeric(limit_equity_value_criteria_1) And IsNumeric(limit_equity_value_criteria_2) Then
            If Abs(limit_equity_value_criteria_1) >= limit_chart_criteria_1 And Abs(limit_equity_value_criteria_2) >= limit_chart_criteria_2 Then
                
                'passe dans les filtres de colonnes
                tmp_equity_filter_ok = True
                For j = 0 To UBound(filter)
                    
                    If filter_type(j) = "col" Then
                    
                        ' la colonne existe dans equity DB
                        If tmp_equity_filter_ok = True And filter_column_equity_db(j) <> False And IsNumeric(filter_column_equity_db(j)) Then
                            
                            ' passe le check critere ou non
                            If filter_compare_type(j) = "=" And filter_type_criteria(j) = "str" And UCase(filter_criteria(j)) = UCase(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) Then
                                
                            ElseIf filter_compare_type(j) = "=" And filter_type_criteria(j) = "num" Then
                                debug_test = Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))
                                If filter_criteria(j) = Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j)) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            ElseIf filter_compare_type(j) = "limit" And filter_type_criteria(j) = "num" Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) And filter_criteria(j) <= Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j)) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            ElseIf filter_compare_type(j) = "limit_abs" And filter_type_criteria(j) = "num" Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) And filter_criteria(j) <= Abs(Worksheets("Equity_Database").Cells(i, filter_column_equity_db(j))) Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            Else
                                tmp_equity_filter_ok = False
                            End If
                            
                        Else
                            tmp_equity_filter_ok = False
                        End If
                
                
                    ElseIf filter_type(j) = "custom" Then
                    
                        ' #################################################################################################################
                        ' ############################ passe dans les filtres custom ######################################################
                        ' #################################################################################################################
                        If UCase(filter(j)) = "EQUITY" Then
                            If filter_criteria(j) = 0 Then
                                tmp_equity_filter_ok = False
                            End If
                        ElseIf UCase(filter(j)) = "SHORT" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 5)) And Worksheets("Equity_Database").Cells(i, 5) < 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "LONG" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 5)) And Worksheets("Equity_Database").Cells(i, 5) > 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "STOCK ONLY" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 9)) And Worksheets("Equity_Database").Cells(i, 9) = 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "LENDING" Then
                            If filter_criteria(j) = 1 Then
                                If IsNumeric(Worksheets("Equity_Database").Cells(i, 26)) And Worksheets("Equity_Database").Cells(i, 26) = 0 Then
                                Else
                                    tmp_equity_filter_ok = False
                                End If
                            End If
                        ElseIf UCase(filter(j)) = "EXPECTED_REPORT_DT" Then
                            If filter_criteria(j) > 0 Then
                                nbre_filters_api_custom = nbre_filters_api_custom + 1
                            End If
                        ElseIf UCase(filter(j)) = "DVD_EX_DT" Then
                            If filter_criteria(j) > 0 Then
                                nbre_filters_api_custom = nbre_filters_api_custom + 1
                            End If
                        ElseIf UCase(filter(j)) = "REGION" Then
                            For k = 0 To UBound(region, 1)
                                If region(k)(0) = filter_criteria(j) Then
                                    For m = 0 To UBound(currency_code)
                                        If Worksheets("Equity_Database").Cells(i, 44) = currency_code(m)(1) Then
                                            For n = 0 To UBound(region(k)(1))
                                                If currency_code(m)(0) = region(k)(1)(n) Then
                                                    GoTo bypass_filter_region_check
                                                Else
                                                    If n = UBound(region(k)(1)) Then
                                                        tmp_equity_filter_ok = False
                                                    End If
                                                End If
                                            Next n
                                        End If
                                    Next m
                                End If
                            Next k
bypass_filter_region_check:
                        ElseIf UCase(filter(j)) = "TAG" Then
                        
                            If filter_criteria(j) <> 0 Then
                                
                                If tmp_equity_filter_ok = True Then
                                    tmp_equity_filter_ok = False
                                    
                                    If Worksheets("Equity_Database").Cells(i, 137) <> "" Then
                                        
                                        Set oTags = oJSON.parse(Worksheets("Equity_Database").Cells(i, 137))
        
                                        For Each oTag In oTags
                                            
                                            If oTag.Item(2) = "OPEN" And oTag.Item(3) = filter_criteria(j) Then
                                                tmp_equity_filter_ok = True
                                                Exit For
                                            End If
                                        Next
                                        
                                    End If
                                End If
                            
                            End If
                        End If
                    
                    End If
                        
                Next j
                

                'le titre doit etre pris en compte
                If tmp_equity_filter_ok = True Then
                    
                    If nbre_filters_api = 0 And nbre_filters_api_custom = 0 Then
                        
                        ReDim Preserve l_array_valeur_eur(l_array_index)
                        ReDim Preserve l_array_daily_pnl(l_array_index)
                        ReDim Preserve l_array_nav_pos(l_array_index)
                        ReDim Preserve l_array_nav_daily(l_array_index)
                        ReDim Preserve l_array_sectors(l_array_index)
                        ReDim Preserve l_array_rel_perf(l_array_index)
                            l_array_rel_perf(l_array_index) = 0
                        ReDim Preserve l_array_rel_index(l_array_index)
                            l_array_rel_index(l_array_index) = "SPX INDEX"
                        
                        ReDim Preserve l_array_line(l_array_index)
                        
                        l_array_valeur_eur(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, 46).Value, Worksheets("Equity_Database").Cells(i, 5).Value) 'short name + valeur eur
                        l_array_daily_pnl(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, 46).Value, Worksheets("Equity_Database").Cells(i, 13).Value) 'short name + daily pnl
                        
                        If c_equity_db_vect_1(1) <> 0 Then
                            l_array_nav_pos(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(1)).Value) 'short name + daily change
                        End If
                        
                        If c_equity_db_vect_2(1) <> 0 Then
                            l_array_nav_daily(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(1)).Value) 'short name + theta
                        End If
                        
                        l_array_sectors(l_array_index) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(1)).Value) 'short name + beta_sector_CODE
                        l_array_line(l_array_index) = Array("Equity_Database", i)
                        
                        
                        ReDim Preserve vec_ticker(l_array_index)
                        vec_ticker(l_array_index) = Worksheets("Equity_Database").Cells(i, 47)
                        
                        l_array_index = l_array_index + 1
                    
                    Else
                        
                        ReDim Preserve list_tickers(count_ticker)
                        list_tickers(count_ticker) = Worksheets("Equity_Database").Cells(i, 47)
                        
                        ReDim Preserve vec_array_all(count_ticker)
                            vec_array_all(count_ticker) = Array(Array(Worksheets("Equity_Database").Cells(i, 46).Value, Worksheets("Equity_Database").Cells(i, 5).Value), Array(Worksheets("Equity_Database").Cells(i, 46).Value, Worksheets("Equity_Database").Cells(i, 13).Value), Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_1(1)).Value), Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_2(1)).Value))
                        
                        
                        ReDim Preserve vec_array_sector(count_ticker)
                            vec_array_sector(count_ticker) = Array(Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(0)).Value, Worksheets("Equity_Database").Cells(i, c_equity_db_vect_3(1)).Value) 'short name + beta_sector_CODE
                        
                        ReDim Preserve vec_array_line(count_ticker)
                            vec_array_line(count_ticker) = Array("Equity_Database", i)
                        
                        count_ticker = count_ticker + 1
                        
                    End If
                    
                End If
                
            End If
        End If
next_entry_equity_db:
    Next i
    
    
    On Error GoTo 0
    
    Dim pass_filter_bbg As Boolean
    
    If count_ticker > 0 Then
        
        If nbre_filters_api > 0 Then
            Dim output_bbg As Variant
            output_bbg = bbg_multi_tickers_and_multi_fields(list_tickers, bbg_fld)
        End If
        
        If nbre_filters_api_custom > 0 Then
            Dim output_bbg_custom As Variant
            bbg_fld_custom = Array("EXPECTED_REPORT_DT", "DVD_EX_DT")
            
            output_bbg_custom = bbg_multi_tickers_and_multi_fields(list_tickers, bbg_fld_custom)
        End If
            
                            
            For i = 0 To UBound(list_tickers, 1)
            
                pass_filter_bbg = True
                
                If nbre_filters_api > 0 Then
                    For j = 0 To UBound(bbg_fld, 1)
                        For k = 0 To UBound(filter, 1)
                            If filter_type(k) = "api" And UCase(bbg_fld(j)) = UCase(filter(k)) Then
                                
                                If Left(output_bbg(i, j), 1) <> "#" Then
                                    If filter_type_criteria(k) = "num" Then
                                        If output_bbg(i, j) >= filter_criteria(k) Then
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    ElseIf filter_type_criteria(k) = "str" Then
                                        If output_bbg(i, j) = filter_criteria(k) Then
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    End If
                                Else
                                    pass_filter_bbg = False
                                End If
                                
                            End If
                        Next k
                    Next j
                End If
                
                
                
                If pass_filter_bbg = True And nbre_filters_api_custom > 0 Then
                    
                    For j = 0 To UBound(bbg_fld_custom, 1)
                        For k = 0 To UBound(filter, 1)
                            If filter_type(k) = "custom" And UCase(bbg_fld_custom(j)) = UCase(filter(k)) Then
                                If Left(output_bbg_custom(i, j), 1) <> "#" Then
                                    
                                    'code custom
                                    If UCase(filter(k)) = "EXPECTED_REPORT_DT" Then
                                        date_tmp = Mid(output_bbg_custom(i, j), 4, 2) & "." & Left(output_bbg_custom(i, j), 2) & "." & Right(output_bbg_custom(i, j), 4)
                                        
                                        test_debug = date_tmp - Date
                                        
                                        If date_tmp - Date <= filter_criteria(k) And date_tmp > Date Then
                                            
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    ElseIf UCase(filter(k)) = "DVD_EX_DT" Then
                                        date_tmp = Mid(output_bbg_custom(i, j), 4, 2) & "." & Left(output_bbg_custom(i, j), 2) & "." & Right(output_bbg_custom(i, j), 4)
                                        
                                        test_debug = date_tmp - Date
                                        
                                        If date_tmp - Date <= filter_criteria(k) And date_tmp > Date Then
                                            
                                        Else
                                            pass_filter_bbg = False
                                        End If
                                    End If
                                    
                                End If
                            End If
                        Next k
                    Next j
                    
                End If
                
                
                
                If pass_filter_bbg = True Then
                    
                    ReDim Preserve l_array_valeur_eur(l_array_index)
                    ReDim Preserve l_array_daily_pnl(l_array_index)
                    ReDim Preserve l_array_nav_pos(l_array_index)
                    ReDim Preserve l_array_nav_daily(l_array_index)
                    ReDim Preserve l_array_rel_perf(l_array_index)
                    ReDim Preserve l_array_rel_perf(l_array_index)
                        l_array_rel_perf(l_array_index) = 0
                    ReDim Preserve l_array_rel_index(l_array_index)
                        l_array_rel_index(l_array_index) = "SPX INDEX"
                    ReDim Preserve l_array_line(l_array_index)
                        
                    
                    l_array_valeur_eur(l_array_index) = vec_array_all(i)(0)
                    l_array_daily_pnl(l_array_index) = vec_array_all(i)(1)
                    l_array_nav_pos(l_array_index) = vec_array_all(i)(2)
                    l_array_nav_daily(l_array_index) = vec_array_all(i)(3)
                    
                    l_array_sectors(l_array_index) = vec_array_sector(i)
                    l_array_line(l_array_index) = vec_array_line(i)
                    
                    ReDim Preserve vec_ticker(l_array_index)
                    vec_ticker(l_array_index) = list_tickers(i)
                    
                    l_array_index = l_array_index + 1
                    
                End If
                
            Next i
        
    End If
    

' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    
End If


    'existe-t-il encore une optional 3e serie ?
    If UBound(chars_config(config_chart), 1) > 7 Then
        
        Dim output_bbg_serie As Variant
        output_bbg_serie = bbg_multi_tickers_and_multi_fields(vec_ticker, Array("CHG_PCT_1D", "rel_index", "rel_1d"))
        
        For i = 0 To UBound(output_bbg_serie, 1)
            If Left(output_bbg_serie(i, 1), 1) <> "#" Then
                l_array_rel_index(i) = output_bbg_serie(i, 1) & " INDEX"
            End If
            
            If Left(output_bbg_serie(i, 2), 1) <> "#" And IsNumeric(output_bbg_serie(i, 2)) Then
                l_array_rel_perf(i) = output_bbg_serie(i, 2)
            End If
            
            
            'normalisation des series % -> *100, bsp 100*100
            l_array_daily_pnl(i)(1) = l_array_daily_pnl(i)(1) / 1000
            l_array_nav_pos(i)(1) = 100 * l_array_nav_pos(i)(1)
            l_array_nav_daily(i)(1) = 100 * 100 * l_array_nav_daily(i)(1)
            
        Next i
        
    End If
    
        
        ' construction du graph avec la methode precedente
        l_rows = l_array_index - 1
                         
        Set l_xls_sheet = Worksheets("Chart_Database")
         
         
        For i = 0 To UBound(chars_config(config_chart)(6), 1)
            For j = 0 To UBound(chars_config(config_chart)(6)(i), 1)
                If chars_config(config_chart)(6)(i)(j) <> "" Then
                    'clean de la column
                    For k = 13 To 32000
                        If Worksheets("Chart_Database").Cells(k, 240) = "" And Worksheets("Chart_Database").Cells(k + 1, 240) = "" And Worksheets("Chart_Database").Cells(k + 2, 240) = "" Then
                            Exit For
                        Else
                            Worksheets("Chart_Database").Cells(k, 240) = ""
                            Worksheets("Chart_Database").Cells(k, 241) = ""
                            Worksheets("Chart_Database").Cells(k, 242) = ""
                            Worksheets("Chart_Database").Cells(k, 243) = ""
                            Worksheets("Chart_Database").Cells(k, 244) = ""
                            Worksheets("Chart_Database").Cells(k, 245) = ""
                            Worksheets("Chart_Database").Cells(k, 246) = ""
                            Worksheets("Chart_Database").Cells(k, 247) = ""
                        End If
                    Next k
                End If
            Next j
        Next i
        
        
        
        
        'sort sur daily pnl
        Dim max_value As Double
        Dim max_pos As Integer
        Dim tmp_value As Variant
        
        For l_row = 0 To l_rows
            max_pos = l_row
            max_value = l_array_nav_daily(l_row)(1)
            
            For j = l_row + 1 To l_rows
                If l_array_nav_daily(j)(1) > max_value Then
                    max_value = l_array_nav_daily(j)(1)
                    max_pos = j
                End If
            Next j
            
            
            If max_pos <> l_row Then
                
                tmp_value = l_array_valeur_eur(l_row)
                l_array_valeur_eur(l_row) = l_array_valeur_eur(max_pos)
                l_array_valeur_eur(max_pos) = tmp_value
                
                tmp_value = l_array_daily_pnl(l_row)
                l_array_daily_pnl(l_row) = l_array_daily_pnl(max_pos)
                l_array_daily_pnl(max_pos) = tmp_value
                
                tmp_value = l_array_nav_pos(l_row)
                l_array_nav_pos(l_row) = l_array_nav_pos(max_pos)
                l_array_nav_pos(max_pos) = tmp_value
                
                tmp_value = l_array_nav_daily(l_row)
                l_array_nav_daily(l_row) = l_array_nav_daily(max_pos)
                l_array_nav_daily(max_pos) = tmp_value
                
                tmp_value = l_array_rel_perf(l_row)
                l_array_rel_perf(l_row) = l_array_rel_perf(max_pos)
                l_array_rel_perf(max_pos) = tmp_value
                
                tmp_value = l_array_rel_index(l_row)
                l_array_rel_index(l_row) = l_array_rel_index(max_pos)
                l_array_rel_index(max_pos) = tmp_value
                
                tmp_value = l_array_sectors(l_row)
                l_array_sectors(l_row) = l_array_sectors(max_pos)
                l_array_sectors(max_pos) = tmp_value
                
                
                tmp_value = l_array_line(l_row)
                l_array_line(l_row) = l_array_line(max_pos)
                l_array_line(max_pos) = tmp_value
                
                
            End If
            
        Next l_row
        
        
        
        'si plus de 20 resultats ne conserve que 20
        Dim limit_count_result As Integer
        limit_count_result = 20
        
        If UBound(l_array_valeur_eur, 1) + 1 > limit_count_result Then
            For i = (limit_count_result / 2) To UBound(l_array_valeur_eur, 1) - (limit_count_result / 2)
                
                l_array_valeur_eur(UBound(l_array_valeur_eur, 1)) = l_array_valeur_eur((limit_count_result / 2))
                ReDim Preserve l_array_valeur_eur(UBound(l_array_valeur_eur, 1) - 1)
                
                l_array_daily_pnl(UBound(l_array_daily_pnl, 1)) = l_array_daily_pnl((limit_count_result / 2))
                ReDim Preserve l_array_daily_pnl(UBound(l_array_daily_pnl, 1) - 1)
                
                
                
                l_array_nav_pos(UBound(l_array_nav_pos, 1)) = l_array_nav_pos((limit_count_result / 2))
                ReDim Preserve l_array_nav_pos(UBound(l_array_nav_pos, 1) - 1)
                
                l_array_nav_daily(UBound(l_array_nav_daily, 1)) = l_array_nav_daily((limit_count_result / 2))
                ReDim Preserve l_array_nav_daily(UBound(l_array_nav_daily, 1) - 1)
                
                
                
                l_array_rel_perf(UBound(l_array_rel_perf, 1)) = l_array_rel_perf((limit_count_result / 2))
                ReDim Preserve l_array_rel_perf(UBound(l_array_rel_perf, 1) - 1)
                
                
                
                l_array_rel_index(UBound(l_array_rel_index, 1)) = l_array_rel_index((limit_count_result / 2))
                ReDim Preserve l_array_rel_index(UBound(l_array_rel_index, 1) - 1)
                
                l_array_sectors(UBound(l_array_sectors, 1)) = l_array_sectors((limit_count_result / 2))
                ReDim Preserve l_array_sectors(UBound(l_array_sectors, 1) - 1)
                
                
                l_array_line(UBound(l_array_line, 1)) = l_array_line((limit_count_result / 2))
                ReDim Preserve l_array_line(UBound(l_array_line, 1) - 1)
                
                
            Next i
        End If
        
        l_rows = UBound(l_array_valeur_eur, 1)
        
        
        For l_row = 0 To l_rows
            
            
            'stock les resultats - donnee static
            Worksheets("Chart_Database").Cells(13 + l_row, 240) = l_array_valeur_eur(l_row)(0) 'name
            
            Worksheets("Chart_Database").Cells(13 + l_row, 241) = l_array_valeur_eur(l_row)(1) 'valeur eur
            Worksheets("Chart_Database").Cells(13 + l_row, 242) = l_array_daily_pnl(l_row)(1) 'daily pnl
            
'            Worksheets("Chart_Database").Cells(13 + l_row, 243) = l_array_nav_pos(l_row)(1) 'nav pos
'            Worksheets("Chart_Database").Cells(13 + l_row, 244) = l_array_nav_daily(l_row)(1) 'nav pnl
            
'            Worksheets("Chart_Database").Cells(13 + l_row, 245) = l_array_rel_perf(l_row) 'rel perf
            Worksheets("Chart_Database").Cells(13 + l_row, 246) = l_array_rel_index(l_row) 'rel index
            
            Worksheets("Chart_Database").Cells(13 + l_row, 247) = l_array_sectors(l_row)(1) 'sector code
            
            
            'stock les resultats - donnee live
            Worksheets("Chart_Database").Cells(13 + l_row, 243).FormulaLocal = "=100*" & l_array_line(l_row)(0) & "!AH" & l_array_line(l_row)(1)
            Worksheets("Chart_Database").Cells(13 + l_row, 244).FormulaLocal = "=100*100*" & l_array_line(l_row)(0) & "!AI" & l_array_line(l_row)(1)
            
            Worksheets("Chart_Database").Cells(13 + l_row, 245).FormulaLocal = "=BDP(IF" & 13 + l_row & " & "" Equity"";""rel_1d"")"
            
        Next l_row

    
    
    Dim count_serie_chart As Integer
    count_serie_chart = 0
    
    Application.ScreenUpdating = False
    
    Set l_xls_chart = Charts("Chart Perf")
    With l_xls_chart
        .ChartArea.ClearContents
        .Activate
        
        l_chart_rows = .SeriesCollection.count
        
        For l_chart_row = 1 To l_chart_rows Step 1
            .SeriesCollection(1).Delete
        Next l_chart_row
              
        
        
        'If chars_config(config_chart)(6)(0)(1) <> "" Then
            
            'NAV position
            Set l_xls_series = .SeriesCollection.NewSeries
            count_serie_chart = count_serie_chart + 1
            
            l_xls_series.name = l_xls_sheet.Range(xlColumnValue(243) & l_chart_database_header)
            
            l_xls_series.Values = l_xls_sheet.Range(xlColumnValue(243) & l_chart_database_header + 1 & ":" & xlColumnValue(243) & l_rows + l_chart_database_header + 1)
            l_xls_series.XValues = l_xls_sheet.Range(xlColumnValue(240) & l_chart_database_header + 1 & ":" & xlColumnValue(240) & l_rows + l_chart_database_header + 1)
            
            ActiveChart.SeriesCollection(count_serie_chart).DataLabels.Select
            Selection.NumberFormat = "#,##0.00"
            
            For l_row = 0 To l_rows
                l_xls_series.Points(l_row + 1).Interior.ColorIndex = format_Chart_ColorIndex(l_xls_sheet.Cells(l_row + l_chart_database_header + 1, 247).Value)
            Next l_row
            
            
            
            'NAV daily
            Set l_xls_series = .SeriesCollection.NewSeries
            count_serie_chart = count_serie_chart + 1
            
            l_xls_series.name = l_xls_sheet.Range(xlColumnValue(244) & l_chart_database_header)
            
            l_xls_series.Values = l_xls_sheet.Range(xlColumnValue(244) & l_chart_database_header + 1 & ":" & xlColumnValue(244) & l_rows + l_chart_database_header + 1)
            l_xls_series.XValues = l_xls_sheet.Range(xlColumnValue(240) & l_chart_database_header + 1 & ":" & xlColumnValue(240) & l_rows + l_chart_database_header + 1)
            
            ActiveChart.SeriesCollection(count_serie_chart).DataLabels.Select
            Selection.NumberFormat = "#,##0.00"
            
            'colorie vert/rouge
            For l_row = 0 To l_rows
                
                If l_array_nav_daily(l_row)(1) >= 0 Then 'utilise la donnee api static
                    l_xls_series.Points(l_row + 1).Interior.ColorIndex = 4
                    l_xls_series.DataLabels(l_row + 1).Font.ColorIndex = 10
                Else
                    l_xls_series.Points(l_row + 1).Interior.ColorIndex = 3
                    l_xls_series.DataLabels(l_row + 1).Font.ColorIndex = 9
                End If
                
            Next l_row
            
            
            'Rel Perf
            Set l_xls_series = .SeriesCollection.NewSeries
            count_serie_chart = count_serie_chart + 1
            
            l_xls_series.name = l_xls_sheet.Range(xlColumnValue(245) & l_chart_database_header)
            
            l_xls_series.Values = l_xls_sheet.Range(xlColumnValue(245) & l_chart_database_header + 1 & ":" & xlColumnValue(245) & l_rows + l_chart_database_header + 1)
            l_xls_series.XValues = l_xls_sheet.Range(xlColumnValue(240) & l_chart_database_header + 1 & ":" & xlColumnValue(240) & l_rows + l_chart_database_header + 1)
            
            ActiveChart.SeriesCollection(count_serie_chart).DataLabels.Select
            Selection.NumberFormat = "#,##0.00"
            
            'colorie vert/rouge
            For l_row = 0 To l_rows
                
                If l_array_rel_perf(l_row) >= 0 Then 'utilise la donnee api static
                    l_xls_series.Points(l_row + 1).Interior.ColorIndex = 4
                    l_xls_series.DataLabels(l_row + 1).Font.ColorIndex = 10
                Else
                    l_xls_series.Points(l_row + 1).Interior.ColorIndex = 3
                    l_xls_series.DataLabels(l_row + 1).Font.ColorIndex = 9
                End If
                
            Next l_row
            
            
            
        
        'End If
        
        Set l_xls_series = Nothing
        
        
    End With
    
    Set l_xls_series = Nothing
    Set l_xls_chart = Nothing
    Set l_xls_sheet = Nothing
    
    Application.ScreenUpdating = True
    
    Application.Calculation = xlCalculationAutomatic

End Sub


Public Sub agreg_stat_gics()

Worksheets("Simulation").Range("AA100:AZ5000").Clear

Dim tables As Variant

'max 2 dim
tables = Array(Array("Sector"), Array("Industry"), Array("Sector", "Industry"), Array("COUNTRY"), Array("GICS_SECTOR_NAME"), Array("GICS_SECTOR_NAME", "GICS_INDUSTRY_NAME"))
tables = Array(Array("COUNTRY"), Array("GICS_SECTOR_NAME", "GICS_INDUSTRY_NAME"))

Dim agreg_column As Variant
'agreg_column = Array("Valeur_Euro", "Daily Result", "Result Total", "Nav Position", "Nav Daily", "Vega_1%_ALL", "Theta_ALL")
agreg_column = Array("Nav Position", "Result Total")



Application.Calculation = xlCalculationManual

Dim l_eq_db_header As Integer
    l_eq_db_header = 25

Dim l_equity_db_last_line As Integer, c_equity_db_code As Integer, c_equity_db_nav_pos As Integer, c_equity_db_nav_daily As Integer, _
    c_equity_db_nav_vol As Integer

    l_equity_db_last_line = 5000
    c_equity_db_code = 4
    c_equity_db_nav_pos = 34
    c_equity_db_nav_daily = 35
    c_equity_db_nav_vol = 36

Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long

Dim l_sim_start_column As Integer, l_simu_line As Integer, c_simu_first As Integer

    l_sim_start_column = 100
    l_simu_line = l_sim_start_column
    
    c_simu_first = 27
    
    'converti les colonnes texts en colonne chiffre
    For i = 0 To UBound(agreg_column, 1)
        For j = 1 To 250
            If agreg_column(i) = Worksheets("Equity_Database").Cells(l_eq_db_header, j) Then
                agreg_column(i) = j
                Exit For
            End If
        Next j
    Next i
    
    'repere la colonne euro pour stat long / short
    For j = 1 To 250
        If Worksheets("Equity_Database").Cells(l_eq_db_header, j) = "Valeur_Euro" Then
            c_eq_db_valeur_eur = j
            Exit For
        End If
    Next j


Dim combi_columns() As Variant

Dim main_dim() As Variant
    
Dim sub_dim() As Variant

Dim c_concern_agreg As Integer



For i = 0 To UBound(tables, 1)
    
    ReDim main_dim(0)
        main_dim(0) = ""
    
    ReDim sub_dim(0)
        sub_dim(0) = ""
    
    For j = 0 To UBound(tables(i), 1)
        
        c_concern_agreg = -1
        
        'repere la colonne
        For m = 1 To 250
            If UCase(Worksheets("Equity_Database").Cells(l_eq_db_header, m)) = UCase(tables(i)(j)) Then
                c_concern_agreg = m
                Exit For
            End If
        Next m
        
        If c_concern_agreg = -1 Then
            MsgBox ("problem column")
            Exit Sub
        End If
        
        
        If j = 0 Then
            
            c_main_dim = c_concern_agreg
            
            For m = 27 To l_equity_db_last_line Step 2
                
                If Worksheets("Equity_Database").Cells(m, 1) = "" Then
                    Exit For
                Else
                    
                    If Worksheets("Equity_Database").Cells(m, c_equity_db_code) = 3 Then
                        Worksheets("Equity_Database").Cells(m, c_equity_db_nav_pos) = 0
                        Worksheets("Equity_Database").Cells(m, c_equity_db_nav_daily) = 0
                        Worksheets("Equity_Database").Cells(m, c_equity_db_nav_vol) = 0
                    End If
                    
                    For n = 0 To UBound(main_dim, 1)
                        
                        If Worksheets("Equity_Database").Cells(m, c_concern_agreg) <> "" And Left(Worksheets("Equity_Database").Cells(m, c_concern_agreg), 1) <> "#" Then
                        
                            If main_dim(n) = Worksheets("Equity_Database").Cells(m, c_concern_agreg) Then
                                Exit For
                            Else
                                If n = UBound(main_dim, 1) Then
                                    If main_dim(0) = "" Then
                                        main_dim(n) = Worksheets("Equity_Database").Cells(m, c_concern_agreg)
                                    Else
                                        ReDim Preserve main_dim(UBound(main_dim, 1) + 1)
                                        main_dim(UBound(main_dim, 1)) = Worksheets("Equity_Database").Cells(m, c_concern_agreg)
                                    End If
                                End If
                            End If
                        End If
                    Next n
                    
                    For p = 0 To UBound(main_dim, 1)
                            
                        min_value = main_dim(p)
                        min_pos = p
                        
                        For q = p + 1 To UBound(main_dim, 1)
                            
                            If main_dim(q) < min_value Then
                                min_value = main_dim(q)
                                min_pos = q
                            End If
                            
                        Next q
                        
                        If min_pos <> p Then
                            tmp_value = main_dim(p)
                            main_dim(p) = main_dim(min_pos)
                            main_dim(min_pos) = tmp_value
                        End If
                        
                    Next p
                    
                End If
                
            Next m
            
            'une seule données - country
            If j = UBound(tables(i), 1) Then
                
                For n = 0 To UBound(main_dim, 1)
                    
                    'header
                    If n = 0 Then
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first) = tables(i)(0)
                        
                        l_table_data_header = l_simu_line
                        
                        l_simu_line = l_simu_line + 1
                    End If
                    
                    
                    Worksheets("Simulation").Cells(l_simu_line, c_simu_first) = main_dim(n)

                    'calcul
                    k = 0
                    For q = 0 To UBound(agreg_column, 1)
                        
                        'long
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 1 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",IF(Equity_Database!R27C" & c_eq_db_valeur_eur & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_eq_db_valeur_eur & ">0,Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0),0))"
                            Worksheets("Simulation").Cells(l_table_data_header, c_simu_first + 1 + k) = Worksheets("Equity_Database").Cells(l_eq_db_header, agreg_column(q)).Value & " LONG"
                        k = k + 1
                        
                        'short
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 1 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",IF(Equity_Database!R27C" & c_eq_db_valeur_eur & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_eq_db_valeur_eur & "<0,Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0),0))"
                            Worksheets("Simulation").Cells(l_table_data_header, c_simu_first + 1 + k) = Worksheets("Equity_Database").Cells(l_eq_db_header, agreg_column(q)).Value & " SHORT"
                        k = k + 1
                        
                        'net
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 1 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0))"
                            Worksheets("Simulation").Cells(l_table_data_header, c_simu_first + 1 + k) = Worksheets("Equity_Database").Cells(l_eq_db_header, agreg_column(q)).Value & " NET"
                        k = k + 1

                    Next q
                    
                    l_simu_line = l_simu_line + 1
                    
                Next n
                
            End If
            
        Else
            
            c_sub_dim = c_concern_agreg
            
            For n = 0 To UBound(main_dim, 1)
                
                ReDim sub_dim(0)
                sub_dim(0) = ""
                
                For m = 27 To l_equity_db_last_line Step 2
                
                    If Worksheets("Equity_Database").Cells(m, 1) = "" Then
                        Exit For
                    Else
                        If Worksheets("Equity_Database").Cells(m, c_concern_agreg) <> "" And Left(Worksheets("Equity_Database").Cells(m, c_concern_agreg), 1) <> "#" Then
                            
                            If Worksheets("Equity_Database").Cells(m, c_main_dim) = main_dim(n) Then
                                
                                For p = 0 To UBound(sub_dim, 1)
                                    
                                    If sub_dim(p) = Worksheets("Equity_Database").Cells(m, c_concern_agreg) Then
                                        Exit For
                                    Else
                                        If p = UBound(sub_dim, 1) Then
                                            If sub_dim(0) = "" Then
                                                sub_dim(0) = Worksheets("Equity_Database").Cells(m, c_concern_agreg)
                                            Else
                                                ReDim Preserve sub_dim(UBound(sub_dim, 1) + 1)
                                                sub_dim(UBound(sub_dim, 1)) = Worksheets("Equity_Database").Cells(m, c_concern_agreg)
                                            End If
                                        End If
                                    End If
                                    
                                Next p
                                
                            End If
                            
                        End If
                    End If
                Next m
                
                
                'sort sub section
                Dim find_etf As Variant
                find_etf = False
                For p = 0 To UBound(sub_dim, 1)
                    
                    min_value = sub_dim(p)
                    min_pos = p
                    
                    For q = p + 1 To UBound(sub_dim, 1)
                        
                        If UCase(sub_dim(q)) = UCase("ETF") Then
                            find_etf = True
                        End If
                        
                        If UCase(sub_dim(q)) < UCase(min_value) Then
                            min_value = sub_dim(q)
                            min_pos = q
                        End If
                        
                    Next q
                    
                    If min_pos <> p Then
                        tmp_value = sub_dim(p)
                        sub_dim(p) = sub_dim(min_pos)
                        sub_dim(min_pos) = tmp_value
                    End If
                    
                Next p
                
                If find_etf = True Then
                    For p = 0 To UBound(sub_dim, 1)
                        
                        If UCase(sub_dim(p)) = UCase("ETF") Then
                            tmp_value_etf_sub_section = sub_dim(p)
                            
                            For q = p + 1 To UBound(sub_dim, 1)
                                sub_dim(q - 1) = sub_dim(q)
                            Next q
                            
                            sub_dim(UBound(sub_dim, 1)) = tmp_value_etf_sub_section
                            
                            Exit For
                        End If
                        
                    Next p
                End If
                
                
                
                'header
                If n = 0 Then
                    Worksheets("Simulation").Cells(l_simu_line, c_simu_first) = tables(i)(0)
                    Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 1) = tables(i)(1)
                    
                    l_table_data_header = l_simu_line
                    
                    l_simu_line = l_simu_line + 1
                End If
                
                For p = 0 To UBound(sub_dim, 1)
                    If p = 0 Then
                        'mise en place de l'entete main entry
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first) = main_dim(n)
                        l_table_data_main_dim = l_simu_line
                        
                    Else
                    End If
                    
                    
                    Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 1) = sub_dim(p)
                    
                    
                    'calcul
                    k = 0
                    For q = 0 To UBound(agreg_column, 1)
                        
                        'long
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 2 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",IF(Equity_Database!R27C" & c_sub_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_sub_dim & "=" & "R" & l_simu_line & "C28" & ",IF(Equity_Database!R27C" & c_eq_db_valeur_eur & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_eq_db_valeur_eur & ">0,Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0),0),0))"
                            Worksheets("Simulation").Cells(l_table_data_header, c_simu_first + 2 + k) = Worksheets("Equity_Database").Cells(l_eq_db_header, agreg_column(q)).Value & " LONG"
                        k = k + 1
                        
                        'short
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 2 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",IF(Equity_Database!R27C" & c_sub_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_sub_dim & "=" & "R" & l_simu_line & "C28" & ",IF(Equity_Database!R27C" & c_eq_db_valeur_eur & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_eq_db_valeur_eur & "<0,Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0),0),0))"
                            Worksheets("Simulation").Cells(l_table_data_header, c_simu_first + 2 + k) = Worksheets("Equity_Database").Cells(l_eq_db_header, agreg_column(q)).Value & " SHORT"
                        k = k + 1
                        
                        'net
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 2 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",IF(Equity_Database!R27C" & c_sub_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_sub_dim & "=" & "R" & l_simu_line & "C28" & ",Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0),0))"
                            Worksheets("Simulation").Cells(l_table_data_header, c_simu_first + 2 + k) = Worksheets("Equity_Database").Cells(l_eq_db_header, agreg_column(q)).Value & " NET"
                        k = k + 1

                    Next q
                    
                    l_simu_line = l_simu_line + 1
                    
                    
                Next p
                
                
                'total
                    'long
                    k = 0
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 1) = "TOTAL"
                        
                        For q = 0 To UBound(agreg_column, 1)
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 2 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",IF(Equity_Database!R27C" & c_eq_db_valeur_eur & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_eq_db_valeur_eur & ">0,Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0),0))"
                        k = k + 1
                        
                        'short
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 2 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",IF(Equity_Database!R27C" & c_eq_db_valeur_eur & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_eq_db_valeur_eur & "<0,Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0),0))"
                        k = k + 1
                        
                        'net
                        Worksheets("Simulation").Cells(l_simu_line, c_simu_first + 2 + k).FormulaArray = "=SUM(IF(Equity_Database!R27C" & c_main_dim & ":Equity_Database!R" & l_equity_db_last_line & "C" & c_main_dim & "=" & "R" & l_table_data_main_dim & "C27" & ",Equity_Database!R27C" & agreg_column(q) & ":Equity_Database!R" & l_equity_db_last_line & "C" & agreg_column(q) & ",0))"
                        k = k + 1
                    Next q
                    
                    l_simu_line = l_simu_line + 1
                
                
            Next n
        
        End If
        
    Next j
    
    l_simu_line = l_simu_line + 5
    
Next i

Call msci_insert_stats_in_simulation

Application.Calculation = xlCalculationAutomatic

End Sub


Public Sub load_Chart_GICS_sector()

Application.Calculation = xlCalculationManual

Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long

l_sim_start_column = 100
    l_simu_line = l_sim_start_column
    
    c_simu_first = 27


'repere la zone dans simulation sur GICS_SECTOR_NAME & GICS_INDUSTRY_NAME
Dim l_simu_data_gics_header As Integer
For i = l_sim_start_column To 5000
    If Worksheets("Simulation").Cells(i, c_simu_first) = "GICS_SECTOR_NAME" And Worksheets("Simulation").Cells(i, c_simu_first + 1) = "GICS_INDUSTRY_NAME" Then
        l_simu_data_gics_header = i
        
        'repere les colonnes de positions
        For j = c_simu_first To 52
            If Worksheets("Simulation").Cells(l_simu_data_gics_header, j) = "Nav Position LONG" Then
                c_simu_nav_pos_long = j
            ElseIf Worksheets("Simulation").Cells(l_simu_data_gics_header, j) = "Nav Position SHORT" Then
                c_simu_nav_pos_short = j
            ElseIf Worksheets("Simulation").Cells(l_simu_data_gics_header, j) = "Nav Position NET" Then
                c_simu_nav_pos_net = j
            ElseIf Worksheets("Simulation").Cells(l_simu_data_gics_header, j) = "MSCI Position" Then
                c_simu_msci_pos_net = j
            End If
        Next j
        
        Exit For
    End If
Next i

Dim vec_points() As Variant
k = 0
For i = l_simu_data_gics_header + 1 To 5000
    If Worksheets("Simulation").Cells(i, c_simu_first + 1) = "" Then
        l_simu_last_line = i - 1
        Exit For
    Else
        ReDim Preserve vec_points(k)
        For j = i To l_simu_data_gics_header + 1 Step -1
            If Worksheets("Simulation").Cells(j, c_simu_first) <> "" Then
                vec_points(k) = Array(Worksheets("Simulation").Cells(j, c_simu_first).Value, Worksheets("Simulation").Cells(i, c_simu_first + 1).Value)
                k = k + 1
                Exit For
            End If
        Next j
    End If
Next i



'bridge color bbg_sector <-> gics
Dim vec_bbg_sector_code As Variant
vec_bbg_sector_code = Array(Array("Basic Materials", 0), Array("Communications", 2), Array("Consumer, Cyclical", 3), Array("Consumer, Non-cyclical", 4), Array("Diversified", 5), Array("Energy", 6), Array("Financial", 7), Array("Industrial", 8), Array("Technology", 9), Array("Utilities", 10))
Dim vec_bbg_sector_gics As Variant
vec_bbg_sector_gics = Array(Array("Basic Materials", "Materials"), Array("Communications", "Telecommunication Services"), Array("Consumer, Cyclical", "Consumer Discretionary"), Array("Consumer, Non-cyclical", "Consumer Staples"), Array("Energy", "Energy"), Array("Financial", "Financials"), Array("Industrial", "Industrials"), Array("Technology", "Information Technology"), Array("Utilities", "Utilities"))




Dim chart_gics As Chart

Dim tmp_point_color As Integer

Dim tmp_serie As Series
Dim tmp_point As Point

k = 0
Dim count_serie_chart As Integer
    count_serie_chart = 0
    
    Application.ScreenUpdating = False
    
    Set chart_gics = Charts("Chart GICS")
    With chart_gics
        .ChartArea.ClearContents
        .Activate
        
        l_chart_rows = .SeriesCollection.count
        
        For l_chart_row = 1 To l_chart_rows Step 1
            .SeriesCollection(1).Delete
        Next l_chart_row
        
        .ChartType = xlColumnStacked
        
        'mise en place des 2 series long/short
        Set tmp_serie = .SeriesCollection.NewSeries
        k = k + 1
        
        .SeriesCollection(k).XValues = "=Simulation!R" & l_simu_data_gics_header + 1 & "C" & c_simu_first & ":R" & l_simu_last_line & "C" & c_simu_first + 1
        .SeriesCollection(k).Values = "=Simulation!R" & l_simu_data_gics_header + 1 & "C" & c_simu_nav_pos_long & ":R" & l_simu_last_line & "C" & c_simu_nav_pos_long
        .SeriesCollection(k).name = "=Simulation!R" & l_simu_data_gics_header & "C" & c_simu_nav_pos_long
        
        
        
        Set tmp_serie = .SeriesCollection.NewSeries
        k = k + 1
        
        .SeriesCollection(k).Values = "=Simulation!R" & l_simu_data_gics_header + 1 & "C" & c_simu_nav_pos_short & ":R" & l_simu_last_line & "C" & c_simu_nav_pos_short
        .SeriesCollection(k).name = "=Simulation!R" & l_simu_data_gics_header & "C" & c_simu_nav_pos_short
        
        
        
        'coloriage de tous les points
        For s = 1 To 2
            m = 0
            
            For Each tmp_point In .SeriesCollection(s).Points
                
                tmp_point_color = 0
                
                For p = 0 To UBound(vec_bbg_sector_gics, 1)
                    If vec_points(m)(0) = vec_bbg_sector_gics(p)(1) Then
                        
                        For q = 0 To UBound(vec_bbg_sector_code, 1)
                            If vec_bbg_sector_code(q)(0) = vec_bbg_sector_gics(p)(0) Then
                                tmp_point_color = format_Chart_ColorIndex(CInt(vec_bbg_sector_code(q)(1)))
                                Exit For
                            End If
                        Next q
                        
                        Exit For
                        
                    End If
                Next p
                
                
                If UCase(vec_points(m)(1)) <> "TOTAL" Then
                                    
                    With tmp_point.Border
                        .ColorIndex = tmp_point_color
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                    With tmp_point.Interior
                        .Pattern = xlSolid
                        .ColorIndex = tmp_point_color
                    End With
                Else
                    
'                    With tmp_point.Border
'                        .ColorIndex = tmp_point_color
'                        .Weight = xlThin
'                        .LineStyle = xlContinuous
'                    End With
'
'                    With tmp_point.Interior
'                        .ColorIndex = xlNone
'                    End With
                    
                    
                    With tmp_point.Border
                        .ColorIndex = 1
                        .Weight = xlThin
                        .LineStyle = xlContinuous
                    End With
                    
                    With tmp_point.Interior
                        .ColorIndex = tmp_point_color
                    End With
                End If
                
                
                
                m = m + 1
            Next
        Next s
        
        
        'mise en place serie net
        Set tmp_serie = .SeriesCollection.NewSeries
        k = k + 1
        
        .SeriesCollection(k).Values = "=Simulation!R" & l_simu_data_gics_header + 1 & "C" & c_simu_nav_pos_net & ":R" & l_simu_last_line & "C" & c_simu_nav_pos_net
        .SeriesCollection(k).name = "=Simulation!R" & l_simu_data_gics_header & "C" & c_simu_nav_pos_net
        .SeriesCollection(k).ChartType = xlXYScatter
        
        
        With .SeriesCollection(k).Border
            .Weight = xlHairline
            .LineStyle = xlNone
        End With
        
        With .SeriesCollection(k)
            .MarkerBackgroundColorIndex = 3
            .MarkerForegroundColorIndex = 3
            .MarkerStyle = xlCircle
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
        
        .SeriesCollection(k).ApplyDataLabels AutoText:=True, LegendKey:= _
        False, ShowSeriesName:=False, ShowCategoryName:=False, ShowValue:=True, _
        ShowPercentage:=False, ShowBubbleSize:=False
        
        
        
        With .SeriesCollection(k).DataLabels
            .AutoScaleFont = True
            .position = xlLabelPositionBelow
            
            With .Font
                .name = "Arial"
                .FontStyle = "Normal"
                .size = 7 '8
            End With
        End With
        
        
        With .Axes(xlCategory).TickLabels
            .AutoScaleFont = True
            
            With .Font
                .name = "Arial"
                .FontStyle = "Normal"
                .size = 7 '7
            End With
        End With
        
        
        
        'mise en place serie msci
        Set tmp_serie = .SeriesCollection.NewSeries
        k = k + 1
        
        .SeriesCollection(k).Values = "=Simulation!R" & l_simu_data_gics_header + 1 & "C" & c_simu_msci_pos_net & ":R" & l_simu_last_line & "C" & c_simu_msci_pos_net
        .SeriesCollection(k).name = "=Simulation!R" & l_simu_data_gics_header & "C" & c_simu_msci_pos_net
        .SeriesCollection(k).ChartType = xlXYScatter
        
        
        With .SeriesCollection(k).Border
            .Weight = xlHairline
            .LineStyle = xlNone
        End With
        
        With .SeriesCollection(k)
            .MarkerBackgroundColorIndex = 5
            .MarkerForegroundColorIndex = 5
            .MarkerStyle = xlDash
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
        
        .SeriesCollection(k).ApplyDataLabels AutoText:=True, LegendKey:= _
        False, ShowSeriesName:=False, ShowCategoryName:=False, ShowValue:=True, _
        ShowPercentage:=False, ShowBubbleSize:=False
        
        
        
        With .SeriesCollection(k).DataLabels
            .AutoScaleFont = True
            .position = xlLabelPositionAbove
            
            With .Font
                .name = "Arial"
                .FontStyle = "Normal"
                .size = 5
            End With
        End With
        
        
        With .Axes(xlCategory).TickLabels
            .AutoScaleFont = True
            
            With .Font
                .name = "Arial"
                .FontStyle = "Normal"
                .size = 7
            End With
        End With
        
        
        
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub


Public Sub update_gics_equity_db()

Application.Calculation = xlCalculationManual


Dim i As Long, j As Long, k As Long, m As Long, n As Long

Dim l_eq_db_header As Integer
Dim c_eq_db_ticker As Integer
Dim c_eq_db_industry_sector As Integer, c_eq_db_industry_group As Integer, c_eq_db_gics_sector_name As Integer, c_eq_db_gics_industry_group_name As Integer, c_eq_gics_industry_name As Integer, c_eq_gics_sub_industry_name As Integer, c_eq_country As Integer

l_eq_db_header = 25

c_eq_db_ticker = 47
c_eq_db_industry_sector = 53
c_eq_db_industry_group = 54
c_eq_db_gics_sector_name = 63
c_eq_db_gics_industry_group_name = 64
c_eq_gics_industry_name = 65
c_eq_gics_sub_industry_name = 66
c_eq_country = 67


Dim bbg_fields As Variant
    bbg_fields = Array("INDUSTRY_SECTOR", "INDUSTRY_GROUP", "GICS_SECTOR_NAME", "GICS_INDUSTRY_GROUP_NAME", "GICS_INDUSTRY_NAME", "GICS_SUB_INDUSTRY_NAME", "COUNTRY")


For i = 0 To UBound(bbg_fields, 1)
    If UCase(bbg_fields(i)) = UCase("INDUSTRY_SECTOR") Then
        dim_bbg_industry_sector = i
    ElseIf UCase(bbg_fields(i)) = UCase("INDUSTRY_GROUP") Then
        dim_bbg_industry_group = i
    ElseIf UCase(bbg_fields(i)) = UCase("GICS_SECTOR_NAME") Then
        dim_bbg_gics_sector_name = i
    ElseIf UCase(bbg_fields(i)) = UCase("GICS_INDUSTRY_GROUP_NAME") Then
        dim_bbg_gics_industry_group_name = i
    ElseIf UCase(bbg_fields(i)) = UCase("GICS_INDUSTRY_NAME") Then
        dim_bbg_gics_industry_name = i
    ElseIf UCase(bbg_fields(i)) = UCase("GICS_SUB_INDUSTRY_NAME") Then
        dim_bbg_gics_sub_industry_name = i
    ElseIf UCase(bbg_fields(i)) = UCase("COUNTRY") Then
        dim_bbg_country = i
    End If
Next i


Dim bridge_field_column As Variant
bridge_field_column = Array(Array(dim_bbg_industry_sector, c_eq_db_industry_sector), Array(dim_bbg_industry_group, c_eq_db_industry_group), Array(dim_bbg_gics_sector_name, c_eq_db_gics_sector_name), Array(dim_bbg_gics_industry_group_name, c_eq_db_gics_industry_group_name), Array(dim_bbg_gics_industry_name, c_eq_gics_industry_name), Array(dim_bbg_gics_sub_industry_name, c_eq_gics_sub_industry_name), Array(dim_bbg_country, c_eq_country))

'mise en place des entetes
For i = 2 To UBound(bridge_field_column, 1) 'eviter de renommer colonne de base
    Worksheets("Equity_Database").Cells(l_eq_db_header, bridge_field_column(i)(1)) = bbg_fields(bridge_field_column(i)(0))
Next i


Dim vec_ticker_and_line() As Variant
Dim vec_ticker() As Variant

k = 0
For i = l_eq_db_header + 2 To 32000 Step 2
    If Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        If Worksheets("Equity_Database").Cells(i, c_eq_db_industry_sector) = "" Or Left(Worksheets("Equity_Database").Cells(i, c_eq_db_industry_sector), 1) = "#" Or Worksheets("Equity_Database").Cells(i, c_eq_db_industry_group) = "" Or Left(Worksheets("Equity_Database").Cells(i, c_eq_db_industry_group), 1) = "#" Or Worksheets("Equity_Database").Cells(i, c_eq_db_gics_sector_name) = "" Or Left(Worksheets("Equity_Database").Cells(i, c_eq_db_gics_sector_name), 1) = "#" Or Worksheets("Equity_Database").Cells(i, c_eq_db_gics_industry_group_name) = "" Or Left(Worksheets("Equity_Database").Cells(i, c_eq_db_gics_industry_group_name), 1) = "#" Or Worksheets("Equity_Database").Cells(i, c_eq_gics_industry_name) = "" Or Left(Worksheets("Equity_Database").Cells(i, c_eq_gics_industry_name), 1) = "#" Or Worksheets("Equity_Database").Cells(i, c_eq_gics_sub_industry_name) = "" Or Left(Worksheets("Equity_Database").Cells(i, c_eq_gics_sub_industry_name), 1) = "#" Then
            If InStr(UCase(Worksheets("Equity_Database").Cells(i, c_eq_db_ticker)), " EQUITY") <> 0 Then
                
                ReDim Preserve vec_ticker_and_line(k)
                ReDim Preserve vec_ticker(k)
                
                vec_ticker_and_line(k) = Array(UCase(Worksheets("Equity_Database").Cells(i, c_eq_db_ticker)), i)
                vec_ticker(k) = UCase(Worksheets("Equity_Database").Cells(i, c_eq_db_ticker))
                
                k = k + 1
            End If
        End If
    End If
Next i


If k > 0 Then

    Dim output_bdp As Variant
    Dim oBBG As New cls_Bloomberg_Sync
    
    output_bdp = oBBG.bdp(vec_ticker, bbg_fields)
    
    'mise en place des nouvelles valeurs
    For i = 0 To UBound(vec_ticker_and_line, 1)
        
        For j = 0 To UBound(bridge_field_column, 1)
            
            If Worksheets("Equity_Database").Cells(vec_ticker_and_line(i)(1), bridge_field_column(j)(1)) = "" Or Left(Worksheets("Equity_Database").Cells(vec_ticker_and_line(i)(1), bridge_field_column(j)(1)), 1) = "#" Then
                
                'remplace avec la nouvelle valeur de bloomberg
                If Left(output_bdp(i + 1)(j + 1), 1) <> "#" Then
                    Worksheets("Equity_Database").Cells(vec_ticker_and_line(i)(1), bridge_field_column(j)(1)) = output_bdp(i + 1)(j + 1)
                End If
                
            End If
            
        Next j
        
    Next i

End If

End Sub
