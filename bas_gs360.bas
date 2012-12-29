Attribute VB_Name = "bas_gs360"
Public Const gs360_header_net_positions_account As String = "Account"
Public Const gs360_header_net_positions_current_net_qty As String = "Current Net Qty"
Public Const gs360_header_net_positions_ticker As String = "Bloomberg Code"
Public Const gs360_header_net_positions_product As String = "Product"
Public Const gs360_header_net_positions_contract_year As String = "Contract Year"
Public Const gs360_header_net_positions_contract_month As String = "Contract Month"
Public Const gs360_header_net_positions_contract_day As String = "Contract Day"
Public Const gs360_header_net_positions_put_call As String = "Put/Call"
Public Const gs360_header_net_positions_strike As String = "Strike Price"




Public Sub gs360_check_account_derivatives()

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer
Dim dico_gs360_pos As New Scripting.Dictionary

'repere si un document compatible avec la macro est deja ouvert
Dim gs360_net_pos_report As Workbook
Dim gs360_net_pos_report_name As String
    gs360_net_pos_report_name = ""
For Each gs360_net_pos_report In Workbooks
    
    If InStr(UCase(gs360_net_pos_report.name), UCase("Default View")) <> 0 Or InStr(UCase(gs360_net_pos_report.name), UCase("extract_gs360_net_position")) <> 0 Then
        gs360_net_pos_report_name = gs360_net_pos_report.name
        Exit For
    End If
    
Next

If gs360_net_pos_report_name = "" Then
    MsgBox ("no file")
    Exit Sub
End If


'repere la ligne header
Dim l_report_header As Integer
l_report_header = 0
For i = 1 To 50
    
    For j = 1 To 200
        
        
        If gs360_net_pos_report.Worksheets(1).Cells(i, j) = "Account" Then
            l_report_header = i
        End If
        
        If gs360_net_pos_report.Worksheets(1).Cells(i, j) = "" Then
            Exit For
        End If
        
    Next j
    
    If l_report_header <> 0 Then
        Exit For
    End If
    
Next i


If l_report_header = 0 Then
    MsgBox ("bug header")
    Exit Sub
End If


'repere strick min columns
Dim required_columns() As Variant
    required_columns = Array(Array(gs360_header_net_positions_account, 0), Array(gs360_header_net_positions_current_net_qty, 0), Array(gs360_header_net_positions_ticker, 0), Array(gs360_header_net_positions_product, 0), Array(gs360_header_net_positions_contract_year, 0), Array(gs360_header_net_positions_contract_month, 0), Array(gs360_header_net_positions_contract_day, 0), Array(gs360_header_net_positions_put_call, 0), Array(gs360_header_net_positions_strike, 0))

Dim dim_account As Integer, dim_curr_net_qty As Integer, dim_ticker As Integer, dim_product As Integer, dim_year As Integer, dim_month As Integer, dim_day As Integer, dim_putcall As Integer, dim_strike As Integer
    'update dim
    For i = 0 To UBound(required_columns, 1)
        If required_columns(i)(0) = gs360_header_net_positions_account Then
            dim_account = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_current_net_qty Then
            dim_curr_net_qty = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_ticker Then
            dim_ticker = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_product Then
            dim_product = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_contract_year Then
            dim_year = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_contract_month Then
            dim_month = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_contract_day Then
            dim_day = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_put_call Then
            dim_putcall = i
        ElseIf required_columns(i)(0) = gs360_header_net_positions_strike Then
            dim_strike = i
        End If
    Next i


Dim missing_column_check As Boolean, missing_column_count As Integer, vec_missing_column() As Variant
    missing_column_check = False
    missing_column_count = 0
    
For i = 0 To UBound(required_columns, 1)
    
    For j = 1 To 250
        If gs360_net_pos_report.Worksheets(1).Cells(l_report_header, j) = "" Then
            Exit For
        Else
            If gs360_net_pos_report.Worksheets(1).Cells(l_report_header, j) = required_columns(i)(0) Then
                
                required_columns(i)(1) = j
                Exit For
            End If
        End If
    Next j
    
    If required_columns(i)(1) = 0 Then
        ReDim Preserve vec_missing_column(missing_column_count)
        vec_missing_column(missing_column_count) = required_columns(i)(0)
        missing_column_count = missing_column_count + 1
        missing_column_check = True
    End If
    
Next i

's assure que toutes les colonnes ont ete trouvees
If missing_column_check = True Then

    For i = 0 To UBound(vec_missing_column, 1)
        MsgBox ("Missing column: " & vec_missing_column(i))
    Next i
    
    MsgBox ("->Exit.")
    Exit Sub
End If




'passe en revue les produits
Dim tmp_id As String
Dim vec_detailed_position() As Variant, vec_detailed_existing_position() As Variant

Dim vec_multiaccounts_product() As Variant
    Dim count_prob_multiaccounts_product As Integer
    count_prob_multiaccounts_product = 0

For i = l_report_header + 1 To 5000
    
    If gs360_net_pos_report.Worksheets(1).Cells(i, 1) = "NOTES:" Then
        Exit For
    Else
        
        'saute les fut
        If gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_putcall)(1)) = "" Or gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_strike)(1)) = "" Or CStr(gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_strike)(1))) = "0" Then
            'future
            Debug.Print gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_ticker)(1))
        Else
            'construction de l id
            tmp_id = ""
            If gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_ticker)(1)) = "" Then
                
                'ticker non dispo construction avec autres colonnes
                
                If gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_product)(1)) = "" Or gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_year)(1)) = "" Or gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_month)(1)) = "" Or gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_day)(1)) = "" Or gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_putcall)(1)) = "" Or gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_strike)(1)) = "" Then
                Else
                    
                        tmp_product = gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_product)(1))
                        tmp_year = gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_year)(1))
                        tmp_month = gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_month)(1))
                        tmp_day = gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_day)(1))
                        tmp_put_call = gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_putcall)(1))
                        tmp_strike = gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_strike)(1))
                    
                    tmp_id = tmp_product & tmp_year & tmp_month & tmp_day & tmp_put_call & tmp_strike
                End If
                
                
            Else
                tmp_id = gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_ticker)(1))
            End If
            
            If tmp_id = "" Then
                MsgBox ("bug id, line: " & i)
            Else
                
                vec_detailed_position = Array(gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_account)(1)).Value, CDbl(gs360_net_pos_report.Worksheets(1).Cells(i, required_columns(dim_curr_net_qty)(1))))
                
                If dico_gs360_pos.Exists(tmp_id) Then
                    
                    vec_detailed_existing_position = dico_gs360_pos.Item(tmp_id)
                    
                    'check if different account
                    If vec_detailed_existing_position(0) <> vec_detailed_position(0) Then
                        
                        'check if different sides
                        If (vec_detailed_existing_position(1) > 0 And vec_detailed_position(1) < 0) Or (vec_detailed_existing_position(1) < 0 And vec_detailed_position(1) > 0) Then
                        
                            'check if one leg is short
                            If vec_detailed_existing_position(1) < 0 Or vec_detailed_position(1) < 0 Then
                                
                                ReDim Preserve vec_multiaccounts_product(count_prob_multiaccounts_product)
                                vec_multiaccounts_product(count_prob_multiaccounts_product) = tmp_id
                                count_prob_multiaccounts_product = count_prob_multiaccounts_product + 1
                                
                            End If
                        
                        End If
                        
                    End If
                    
                    
                Else
                    dico_gs360_pos.Add tmp_id, vec_detailed_position
                End If
                
            End If
            
        End If
    End If
    
Next i


If count_prob_multiaccounts_product > 0 Then
    Dim msg_str As String
    
    msg_str = ""
    For i = 0 To UBound(vec_multiaccounts_product, 1)
        msg_str = msg_str & vec_multiaccounts_product(i) & vbCrLf
    Next i
    
    MsgBox (msg_str)
Else
    MsgBox ("everthing's fine")
End If


End Sub

