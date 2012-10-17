VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Central_mgmt_rank 
   Caption         =   "Custom rank"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   OleObjectBlob   =   "frm_Central_mgmt_rank.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_central_mgmt_rank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const form_small_view_height As Double = 306
Private Const form_big_view_height As Double = 385.5

Private Const column_weight_bbg_field As Integer = 100
Private Const column_weight_side As Integer = 55
Private Const column_weight_rank_if_not_available As Integer = 80
Private Const column_weight_weight As Integer = 50



Private Sub switch_view()

If Me.Height = form_small_view_height Then
    Me.Height = form_big_view_height
ElseIf Me.Height = form_big_view_height Then
    Me.Height = form_small_view_height
End If

End Sub


Private Sub expand_view()

If Me.Height = form_small_view_height Then
    Me.Height = form_big_view_height
End If

End Sub

Private Sub collapse_view()

If Me.Height = form_big_view_height Then
    Me.Height = form_small_view_height
End If

End Sub


Private Function check_fields() As Boolean

check_fields = True

Me.TB_field_bbg.Value = Replace(UCase(Me.TB_field_bbg.Value), " ", "_")

If Me.TB_field_bbg.Value = "" Then
    check_fields = False
End If

If Me.CB_field_side = "" Then
    check_fields = False
ElseIf Me.CB_field_side.Value <> "small is best" And Me.CB_field_side.Value <> "big is best" Then
    check_fields = False
End If


If Me.TB_field_rank_if_not_available.Value = "" Then
    check_fields = False
Else
    If IsNumeric(Me.TB_field_rank_if_not_available.Value) = False Then
        check_fields = False
    End If
End If


If Me.TB_field_weight.Value = "" Then
    check_fields = False
Else
    If IsNumeric(Me.TB_field_weight.Value) = False Then
        check_fields = False
    End If
End If


End Function




Private Sub btn_add_field_Click()

Call expand_view
    
    Me.TB_field_bbg.Value = ""
    Me.CB_field_side.Value = ""
    Me.TB_field_rank_if_not_available.Value = ""
    Me.TB_field_weight.Value = ""
    
    Me.btn_field.Caption = "Add"

End Sub

Private Sub btn_edit_field_Click()


If Me.LV_field.ListItems.count > 0 Then
    
    Call expand_view
    
    field_id = Me.LV_field.SelectedItem.key
    
    tmp_side = Me.LV_field.ListItems(field_id).ListSubItems("side_" & field_id).Text
    tmp_rank = CDbl(Me.LV_field.ListItems(field_id).ListSubItems("rank_" & field_id).Text)
    tmp_weight = CDbl(Me.LV_field.ListItems(field_id).ListSubItems("weight_" & field_id).Text)
    
    Me.TB_field_bbg.Value = field_id
    Me.CB_field_side.Value = tmp_side
    Me.TB_field_rank_if_not_available.Value = tmp_rank
    Me.TB_field_weight.Value = tmp_weight
    
    Me.btn_field.Caption = "Edit"
    
End If




End Sub

Private Sub btn_field_Click()

If UCase(Me.btn_field.Caption) = "EDIT" Then
    If check_fields = True Then
        
        'remplace les values
        Me.LV_field.ListItems(Me.TB_field_bbg.Value).ListSubItems("side_" & Me.TB_field_bbg.Value).Text = Me.CB_field_side.Value
        Me.LV_field.ListItems(Me.TB_field_bbg.Value).ListSubItems("rank_" & Me.TB_field_bbg.Value).Text = Me.TB_field_rank_if_not_available.Value
        Me.LV_field.ListItems(Me.TB_field_bbg.Value).ListSubItems("weight_" & Me.TB_field_bbg.Value).Text = Me.TB_field_weight.Value
        
        Call collapse_view
        
    End If
ElseIf UCase(Me.btn_field.Caption) = "ADD" Then
    
    
    
    If check_fields = True Then
        
        
            'mise en place header
            With Me.LV_field
                With .ColumnHeaders
                    If Me.LV_field.ListItems.count = 0 Then
                        .Clear
                        .Add , , "bbg_field", column_weight_bbg_field
                        .Add , , "side", column_weight_side
                        .Add , , "rank if not available", column_weight_rank_if_not_available
                        .Add , , "weight", column_weight_weight
                    End If
                End With
                
                .ListItems.Add , Me.TB_field_bbg.Value, Me.TB_field_bbg.Value
                
                .ListItems(Me.TB_field_bbg.Value).ListSubItems.Add , "side_" & Me.TB_field_bbg.Value, Me.CB_field_side.Value
                .ListItems(Me.TB_field_bbg.Value).ListSubItems.Add , "rank_" & Me.TB_field_bbg.Value, Me.TB_field_rank_if_not_available.Value
                .ListItems(Me.TB_field_bbg.Value).ListSubItems.Add , "weight_" & Me.TB_field_bbg.Value, Me.TB_field_weight.Value
                
            End With
        
        Call collapse_view
        
    End If
    
End If

End Sub

Private Sub btn_load_existing_Click()

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
Dim sql_query As String


If CB_existing_rank.Value <> "" Then
    
    sql_query = "SELECT * FROM " & t_central_rank & " WHERE " & f_central_rank_id & "=""" & CB_existing_rank.Value & """"
    Dim extract_view_details As Variant
    extract_view_details = sqlite3_query(central_get_db_fullpath, sql_query)
    
    'detect des dim
    For i = 0 To UBound(extract_view_details(0), 1)
        If extract_view_details(0)(i) = f_central_rank_bbg_field Then
            dim_extract_bbg_field = i
        ElseIf extract_view_details(0)(i) = f_central_rank_order Then
            dim_extract_side = i
        ElseIf extract_view_details(0)(i) = f_central_rank_rank_if_not_available Then
            dim_extract_rank_if_not_available = i
        ElseIf extract_view_details(0)(i) = f_central_rank_weight Then
            dim_extract_weight = i
        End If
    Next i
    
    With Me.LV_field
        
        .ListItems.Clear
        
        With .ColumnHeaders
            .Clear
            
            .Add , , "bbg_field", column_weight_bbg_field
            .Add , , "side", column_weight_side
            .Add , , "rank if not available", column_weight_rank_if_not_available
            .Add , , "weight", column_weight_weight
            
        End With
        
            For i = 1 To UBound(extract_view_details, 1)
                .ListItems.Add , extract_view_details(i)(dim_extract_bbg_field), extract_view_details(i)(dim_extract_bbg_field)
                
                If extract_view_details(i)(dim_extract_side) = central_order_rank.big_is_best Then
                    .ListItems(i).ListSubItems.Add , "side_" & extract_view_details(i)(dim_extract_bbg_field), "big is best"
                ElseIf extract_view_details(i)(dim_extract_side) = central_order_rank.small_is_best Then
                    .ListItems(i).ListSubItems.Add , "side_" & extract_view_details(i)(dim_extract_bbg_field), "small is best"
                End If
                
                .ListItems(i).ListSubItems.Add , "rank_" & extract_view_details(i)(dim_extract_bbg_field), extract_view_details(i)(dim_extract_rank_if_not_available)
                .ListItems(i).ListSubItems.Add , "weight_" & extract_view_details(i)(dim_extract_bbg_field), extract_view_details(i)(dim_extract_weight)
                
            Next i
            
    End With
    
End If

End Sub

Private Sub btn_remove_field_Click()

If Me.LV_field.ListItems.count > 0 Then
    Me.LV_field.ListItems.Remove (Me.LV_field.SelectedItem.key)
End If

End Sub

Private Sub btn_run_report_Click()

If CB_existing_rank.Value <> "" Then
    debug_test = central_load_rank(CB_existing_rank.Value, central_get_ticker_rank)
    Me.Hide
End If

End Sub

Private Sub btn_save_Click()

Dim rank_name As String

If Me.LV_field.ListItems.count > 0 Then

    If Me.CB_existing_rank.Value = "" Then
get_rank_name:
        rank_name = InputBox("Name?")
    Else
        rank_name = Me.CB_existing_rank.Value
    End If
    
    If rank_name = "" Then
        GoTo get_rank_name
    End If
    
    
    'remonte les champs
    Dim tmp_bbg_field As Variant
    Dim vec_fields() As Variant
    k = 0
    For Each tmp_bbg_field In Me.LV_field.ListItems
        
        tmp_side = Me.LV_field.ListItems(CStr(tmp_bbg_field)).ListSubItems("side_" & tmp_bbg_field).Text
        
        If InStr(tmp_side, "small") <> 0 Then
            tmp_side = central_order_rank.small_is_best
        ElseIf InStr(tmp_side, "big") <> 0 Then
            tmp_side = central_order_rank.big_is_best
        End If
        
        tmp_rank = CDbl(Me.LV_field.ListItems(CStr(tmp_bbg_field)).ListSubItems("rank_" & tmp_bbg_field).Text)
        tmp_weight = CDbl(Me.LV_field.ListItems(CStr(tmp_bbg_field)).ListSubItems("weight_" & tmp_bbg_field).Text)
        
        ReDim Preserve vec_fields(k)
        vec_fields(k) = Array(tmp_bbg_field, tmp_bbg_field, tmp_side, tmp_rank, tmp_weight)
        k = k + 1
        
    Next
    
    If k > 0 Then
        debug_test = central_create_custom_rank(rank_name, vec_fields)
    End If
    
End If


End Sub

Private Sub CB_existing_rank_Change()

End Sub

Private Sub CB_existing_rank_DropButtonClick()

Dim i As Integer, j As Integer, k As Integer

Dim tmp_value As String

tmp_value = CB_existing_rank.Value
CB_existing_rank.Clear

Dim sql_query As String
sql_query = "SELECT DISTINCT " & f_central_rank_id & " FROM " & t_central_rank & " ORDER BY " & f_central_rank_id & " ASC"
Dim extract_rank As Variant
extract_rank = sqlite3_query(central_get_db_fullpath, sql_query)

For i = 1 To UBound(extract_rank, 1)
    CB_existing_rank.AddItem extract_rank(i)(0)
Next i

CB_existing_rank.Value = tmp_value

End Sub

Private Sub CB_field_side_Change()

End Sub

Private Sub CB_field_side_DropButtonClick()

If CB_field_side.ListCount = 0 Then
    CB_field_side.AddItem "big is best"
    CB_field_side.AddItem "small is best"
End If

End Sub


Private Sub LV_field_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub UserForm_Click()

End Sub
