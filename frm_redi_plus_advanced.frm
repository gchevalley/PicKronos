VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_redi_plus_advanced 
   Caption         =   "RediPlus - Advanced Trading"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   OleObjectBlob   =   "frm_redi_plus_advanced.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_redi_plus_advanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Const book_name As String = "Kronos.xls"
Private Const sparkline_tmp_file As String = "sparkline_tmp.jpg"
Private Const tmp_sheet_sparkline As String = "sparkline_tmp_sheet"
Private Const tmp_sheet_chart As String = "chart_tmp_sheet"

Private Const form_caption_base As String = "RediPlus - Advanced Trading"
Private Const frame_market_data_caption As String = "Market Datas"

Private Const quick_btn_tooltip_stairs_size As String = "size"
Private Const quick_btn_tooltip_stairs_increment As String = "incr"
Private Const quick_btn_tooltip_stairs_step As String = "step"


Private Const weight_serie As Double = 0.8
Private Const weight_moving_average As Double = 0.01

Private Const minute_timebase_intraday As Integer = 15

    Private Const color_r_serie As Integer = 0
    Private Const color_g_serie As Integer = 0
    Private Const color_b_serie As Integer = 0
    
    Private Const color_r_ma_1 As Integer = 0
    Private Const color_g_ma_1 As Integer = 191
    Private Const color_b_ma_1 As Integer = 255
    
    Private Const color_r_ma_2 As Integer = 255
    Private Const color_g_ma_2 As Integer = 0
    Private Const color_b_ma_2 As Integer = 285
    
    Private Const color_r_ma_3 As Integer = 255
    Private Const color_g_ma_3 As Integer = 153
    Private Const color_b_ma_3 As Integer = 0
    
Private Const dim_r_plus_trades_symbol As Integer = 0
Private Const dim_r_plus_trades_qty As Integer = 1
Private Const dim_r_plus_trades_price As Integer = 2

Public mode_emsx As Integer




Private Sub LineChart_VBA_to_picture(array_serie1 As Variant, ByVal picture_file_path As String, Optional MA1_period As Integer, Optional MA2_period As Integer, Optional MA3_period As Integer)

Dim i As Long, j As Long, k As Long

Const cMargin = 2

Dim rng As Range

Dim dblMin As Double, dblMax As Double

Dim shp As Shape
Dim SegmntShp_serie1 As Shape
Dim SegmntShp_MA1 As Shape
Dim SegmntShp_MA2 As Shape
Dim SegmntShp_MA3 As Shape


Dim timebase As Integer 'pour ma

Dim arr_element_serie1() As Variant

Dim arr_element_ma1() As Variant
Dim arr_element_ma2() As Variant
Dim arr_element_ma3() As Variant

Dim arr_all_elements() As Variant
Dim nbre_all_elements As Integer
nbre_all_elements = 0

Dim sum_for_ma As Double


's'assure que la sheet tmp_sparline existe bien sinon la cree
Dim tmp_wrksht As Worksheet, tmp_wrksht_chart As Chart, find_wrksht_tmp_sparkline As Boolean, find_wrksht_tmp_chart As Boolean
find_wrksht_tmp_sparkline = False
For Each tmp_wrksht In Workbooks(book_name).Worksheets
    If tmp_wrksht.name = tmp_sheet_sparkline Then
        find_wrksht_tmp_sparkline = True
    End If
Next


For Each tmp_wrksht_chart In Workbooks(book_name).Charts
    If tmp_wrksht_chart.name = tmp_sheet_chart Then
        find_wrksht_tmp_chart = True
    End If
Next


If find_wrksht_tmp_sparkline = False Then
    Set tmp_wrksht = Workbooks(book_name).Worksheets.Add
    tmp_wrksht.name = tmp_sheet_sparkline
    tmp_wrksht.Visible = xlSheetHidden
End If


If find_wrksht_tmp_chart = False Then
    Set tmp_wrksht_chart = Workbooks(book_name).Charts.Add
    tmp_wrksht_chart.name = tmp_sheet_chart
    tmp_wrksht_chart.Visible = xlSheetHidden
End If



Set rng = Workbooks(book_name).Worksheets(tmp_sheet_sparkline).Cells(1, 1)  'impose la cellule 1,1


'repère le min/max des 2 séries

    dblMin = array_serie1(0)
    dblMax = array_serie1(0)

For i = 1 To UBound(array_serie1, 1)
    
    'repère Min/Max de la série
    If array_serie1(i) > dblMax Then
        dblMax = array_serie1(i)
    End If
    
    If array_serie1(i) < dblMin Then
        dblMin = array_serie1(i)
    End If
 
Next
    
        


proc_1_serie:
    ' graph serie1
    With rng.Worksheet.Shapes
        For i = 0 To UBound(array_serie1, 1) - 2
            Set SegmntShp_serie1 = .AddLine( _
                cMargin + rng.Left + (i * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                cMargin + rng.Top + (dblMax - array_serie1(i + 1)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin), _
                cMargin + rng.Left + ((i + 1) * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                cMargin + rng.Top + (dblMax - array_serie1(i + 2)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin))
            
            
            
            'SegmntShp_serie1.Line.Weight = weight_serie
            SegmntShp_serie1.line.Weight = 8
            
            On Error Resume Next
            j = 0: j = UBound(arr_element_serie1) + 1
            On Error GoTo 0
            ReDim Preserve arr_element_serie1(j)
            arr_element_serie1(j) = SegmntShp_serie1.name
            
            ReDim Preserve arr_all_elements(nbre_all_elements)
            arr_all_elements(nbre_all_elements) = SegmntShp_serie1.name
            nbre_all_elements = nbre_all_elements + 1
            
        Next
 
        With rng.Worksheet.Shapes.Range(arr_element_serie1)
            .line.ForeColor.RGB = RGB(255, 255, 255)    'blanc
        End With
 
    End With
    
MA_1:
    'construction des moving average
    If MA1_period <> 0 Then
        timebase = MA1_period
        
        If timebase > UBound(array_serie1, 1) Then
            GoTo MA_2
        End If

        Dim arr_ma1() As Variant
        ReDim Preserve arr_ma1(1 To UBound(array_serie1, 1))

        For i = timebase To UBound(array_serie1, 1)
            sum_for_ma = 0

            For j = i - (timebase - 1) To i
                sum_for_ma = sum_for_ma + array_serie1(j)
            Next j

            arr_ma1(i) = sum_for_ma / timebase

        Next i



        'graph de la ma1
        With rng.Worksheet.Shapes

           For i = (timebase) - 1 To UBound(arr_ma1, 1) - 2

            Set SegmntShp_MA1 = .AddLine( _
                cMargin + rng.Left + (i * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                cMargin + rng.Top + (dblMax - arr_ma1(i + 1)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin), _
                cMargin + rng.Left + ((i + 1) * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                cMargin + rng.Top + (dblMax - arr_ma1(i + 2)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin))

            SegmntShp_MA1.line.Weight = 6

                On Error Resume Next
                j = 0: j = UBound(arr_element_ma1) + 1
                On Error GoTo 0
                ReDim Preserve arr_element_ma1(j)
                arr_element_ma1(j) = SegmntShp_MA1.name

                ReDim Preserve arr_all_elements(nbre_all_elements)
                arr_all_elements(nbre_all_elements) = SegmntShp_MA1.name
                nbre_all_elements = nbre_all_elements + 1

            Next i



            With rng.Worksheet.Shapes.Range(arr_element_ma1)
                .line.ForeColor.RGB = RGB(color_r_ma_1, color_g_ma_1, color_b_ma_1)
            End With


        End With

    End If




MA_2:
    If MA2_period <> 0 Then
        timebase = MA2_period
        
        If timebase > UBound(array_serie1, 1) Then
            GoTo MA_3
        End If
        
        Dim arr_ma2() As Variant
        ReDim Preserve arr_ma2(1 To UBound(array_serie1, 1))

        For i = timebase To UBound(array_serie1, 1)
            sum_for_ma = 0

            For j = i - (timebase - 1) To i
                sum_for_ma = sum_for_ma + array_serie1(j)
            Next j

            arr_ma2(i) = sum_for_ma / timebase

        Next i



        'graph de la ma2
        With rng.Worksheet.Shapes

            For i = (timebase) - 1 To UBound(arr_ma2, 1) - 2

                Set SegmntShp_MA2 = .AddLine( _
                    cMargin + rng.Left + (i * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                    cMargin + rng.Top + (dblMax - arr_ma2(i + 1)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin), _
                    cMargin + rng.Left + ((i + 1) * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                    cMargin + rng.Top + (dblMax - arr_ma2(i + 2)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin))

                SegmntShp_MA2.line.Weight = 6

                On Error Resume Next
                j = 0: j = UBound(arr_element_ma2) + 1
                On Error GoTo 0
                ReDim Preserve arr_element_ma2(j)
                arr_element_ma2(j) = SegmntShp_MA2.name

                ReDim Preserve arr_all_elements(nbre_all_elements)
                arr_all_elements(nbre_all_elements) = SegmntShp_MA2.name
                nbre_all_elements = nbre_all_elements + 1

            Next i



            With rng.Worksheet.Shapes.Range(arr_element_ma2)
                .line.ForeColor.RGB = RGB(color_r_ma_2, color_g_ma_2, color_b_ma_2)
            End With


        End With

    End If
    
    
MA_3:
        If MA3_period <> 0 Then
        timebase = MA3_period
        
        If timebase > UBound(array_serie1, 1) Then
            GoTo group_all_elements
        End If
        
        Dim arr_ma3() As Variant
        ReDim Preserve arr_ma3(1 To UBound(array_serie1, 1))

        For i = timebase To UBound(array_serie1, 1)
            sum_for_ma = 0

            For j = i - (timebase - 1) To i
                sum_for_ma = sum_for_ma + array_serie1(j)
            Next j

            arr_ma3(i) = sum_for_ma / timebase

        Next i



        'graph de la ma3
        With rng.Worksheet.Shapes

            For i = (timebase) - 1 To UBound(arr_ma3, 1) - 2

                Set SegmntShp_MA3 = .AddLine( _
                    cMargin + rng.Left + (i * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                    cMargin + rng.Top + (dblMax - arr_ma3(i + 1)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin), _
                    cMargin + rng.Left + ((i + 1) * (rng.Width - (cMargin * 2)) / (UBound(array_serie1, 1) - 1)), _
                    cMargin + rng.Top + (dblMax - arr_ma3(i + 2)) * (rng.Height - (cMargin * 2)) / (dblMax - dblMin))

                SegmntShp_MA3.line.Weight = 6

                On Error Resume Next
                j = 0: j = UBound(arr_element_ma3) + 1
                On Error GoTo 0
                ReDim Preserve arr_element_ma3(j)
                arr_element_ma3(j) = SegmntShp_MA3.name

                ReDim Preserve arr_all_elements(nbre_all_elements)
                arr_all_elements(nbre_all_elements) = SegmntShp_MA3.name
                nbre_all_elements = nbre_all_elements + 1

            Next i



            With rng.Worksheet.Shapes.Range(arr_element_ma3)
                .line.ForeColor.RGB = RGB(color_r_ma_3, color_g_ma_3, color_b_ma_3)
            End With


        End With

    End If



group_all_elements:
'regroup tous les éléments
With rng.Worksheet.Shapes.Range(arr_all_elements)
    .Group
    .name = "SparklineGS" & rng.Address(, , xlR1C1) & "Shape"
End With

Dim PicHeight As Double
Dim PicWidth As Double

With rng.Worksheet.Shapes("SparklineGS" & rng.Address(, , xlR1C1) & "Shape")
    PicHeight = .Height
    PicWidth = .Width
    
    .Copy
    .Delete
End With

For Each tmp_shape In Workbooks(book_name).Charts(tmp_sheet_chart).Shapes
    If tmp_shape.name = "SparklineGSR1C1Shape" Then
        Workbooks(book_name).Charts(tmp_sheet_chart).Shapes("SparklineGSR1C1Shape").Delete
    End If
Next


Workbooks(book_name).Charts(tmp_sheet_chart).Paste
Workbooks(book_name).Charts(tmp_sheet_chart).Export FileName:=picture_file_path, FilterName:="jpg"


''avec objectchart in sheet
'ThisWorkbook.Sheets("sparklines_tmp_sheet").ChartObjects("sprkl_tmp").Activate
'ActiveChart.Shapes("SparklineGSR1C1Shape").Delete
'ActiveChart.Paste
'ActiveChart.Export filename:="Q:\Sag\Financial Engineering\front\greg\Addins\sparkline_tmp.png", FilterName:="PNG"


End Sub


Private Sub draw_sparkline_to_picture(ByVal ticker As Variant, ByVal field_bbg As Variant, ByVal timebase As String, ByVal working_days_histo As Integer, ByVal picture_file_path As String, Optional ByVal moving_average_1 As Integer, Optional ByVal moving_average_2 As Integer, Optional ByVal moving_average_3 As Integer)

Dim oBBG As New cls_Bloomberg_Sync

Dim bdh_data As Variant

Dim date_tmp As Date


Dim vect_data_sprkl() As Variant

If timebase = "DAILY" Then
    'bdh_data = bbh_multi_tickers_and_multi_fields(ticker, field_bbg, workday_custom(FormatDateTime(Date, vbShortDate), -working_days_histo), Date)
    bdh_data = oBBG.bdh(ticker, field_bbg, Date - 365, Date)
    
        If UBound(bdh_data(0), 1) > 2 Then
            
            Erase vect_data_sprkl
            
            k = 0
            For i = 0 To UBound(bdh_data(0), 1)
                If IsNumeric(bdh_data(0)(i)(2)) = True Then
                    ReDim Preserve vect_data_sprkl(k)
                    vect_data_sprkl(k) = bdh_data(0)(i)(2)
                    k = k + 1
                End If
                
            Next i
            
            'graph
            If k > 3 Then
                Call LineChart_VBA_to_picture(vect_data_sprkl, picture_file_path, moving_average_1, moving_average_2, moving_average_3)
            End If
            
        End If
        
ElseIf timebase = "INTRADAY" Then
    Dim date_from As Date
    date_from = workday_custom(FormatDateTime(Date, vbShortDate), -working_days_histo)
    
    'bdh_data = bbh_intraday_multi_tickers_and_multi_fields(ticker, field_bbg, CDate(date_from & " 09:00:00"), 15, Now)
    bdh_data = oBBG.intraday(ticker, CDate(date_from & " 08:00:00"), Now, 15)
    
    Erase vect_data_sprkl
    k = 0
    If UBound(bdh_data(0), 1) > 2 Then
        For i = 0 To UBound(bdh_data(0), 1)
            If IsNumeric(bdh_data(0)(i)(5)) = True Then
                ReDim Preserve vect_data_sprkl(k)
                vect_data_sprkl(k) = bdh_data(0)(i)(5)
                k = k + 1
            End If
        Next i
        
        'graph
        If k > 3 Then
            Call LineChart_VBA_to_picture(vect_data_sprkl, picture_file_path, moving_average_1, moving_average_2, moving_average_3)
        End If
        
    End If
    
End If

End Sub


Private Sub btn_add_suffix_f10_index_Click()

Dim ticker As String
ticker = UCase(TB_order_ticker.Value)

If ticker <> "" Then
    ticker = Replace(ticker, " INDEX", "")
    ticker = Replace(ticker, " EQUITY", "")
    TB_order_ticker.Value = ticker & " INDEX"
    Call load_ticker_datas
End If

End Sub


Private Sub btn_add_suffix_f8_equity_Click()

Dim ticker As String
ticker = UCase(TB_order_ticker.Value)

If ticker <> "" Then
    ticker = Replace(ticker, " INDEX", "")
    ticker = Replace(ticker, " EQUITY", "")
    TB_order_ticker.Value = ticker & " EQUITY"
    Call load_ticker_datas
End If

End Sub


Private Sub btn_append_to_csv_Click()

Dim debug_test As Variant

Dim emsx_csv_filename As String
    emsx_csv_filename = "01emsx_oi.csv"

Dim default_aim_account As String
    default_aim_account = "C6414GSJ"

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim FNum As Integer
Dim wholeLine  As String

'check requirements
Dim vec_trade() As Variant
Dim vec_trades_emsx() As Variant

k = 0
If CB_Order_Side.Value <> "" And TB_order_qty.Value <> "" And IsNumeric(TB_order_qty.Value) And TB_order_ticker.Value <> "" And Right(UCase(TB_order_ticker.Value), 6) = "EQUITY" And TB_custom_price.Value <> "" And IsNumeric(TB_custom_price.Value) And CB_exec_broker.Value <> "" And CB_aim_strategy.Value <> "" Then
    
    
    If CB_Order_Side.Value = "BUY" Then
        tmp_side = "B"
    ElseIf CB_Order_Side.Value = "SELL" Then
        tmp_side = "H" 'check si deja long
    End If
    
    
    vec_trade = Array(UCase(TB_order_ticker.Value), default_aim_account, CB_aim_strategy.Value, "LMT", tmp_side, CDbl(TB_order_qty.Value), CDbl(TB_custom_price.Value), "DAY", CB_exec_broker.Value)
    
    
    If mode_emsx = 0 Then
        vec_trades_emsx = Array(vec_trade)
        'new file
        debug_test = array_to_csv(vec_trades_emsx, base_path & emsx_csv_filename)
    Else
        'append
        'mount le contenu actuel
        Dim already_in_csv() As Variant
        already_in_csv = csv_to_array(base_path & emsx_csv_filename)
        
        ReDim Preserve already_in_csv(UBound(already_in_csv, 1) + 1)
        already_in_csv(UBound(already_in_csv, 1)) = vec_trade
        
        debug_test = array_to_csv(already_in_csv, base_path & emsx_csv_filename)
        
    End If
    
    mode_emsx = 1
    
Else
    MsgBox ("missing field(s)")
    Exit Sub
End If


End Sub


Private Sub btn_clear_Click()

Call clear_form_idea

End Sub


Sub clear_form_idea()

Call clear_market_datas

CB_Order_Side.Value = ""
TB_order_qty.Value = ""
TB_order_ticker.Value = ""
TB_custom_price.Value = ""

CB_exec_broker.Value = "NEGOCE"
CB_aim_strategy.Value = "GROWTH STOCK"


End Sub


Sub clear_market_datas()

TgglBtn_Price_Last.Value = False
TgglBtn_Price_Bid.Value = False
TgglBtn_Price_Ask.Value = False
TgglBtn_Price_High.Value = False
TgglBtn_Price_Low.Value = False

TgglBtn_Price_R3.Value = False
TgglBtn_Price_R2.Value = False
TgglBtn_Price_R1.Value = False
TgglBtn_Price_PP.Value = False
TgglBtn_Price_S1.Value = False
TgglBtn_Price_S2.Value = False
TgglBtn_Price_S3.Value = False

frame_market_data.Caption = "Market Datas"
TgglBtn_Price_Last.Caption = ""
L_change_value.Caption = ""
TgglBtn_Price_Bid.Caption = ""
TgglBtn_Price_Ask.Caption = ""
TgglBtn_Price_High.Caption = ""
TgglBtn_Price_Low.Caption = ""
L_volume_value.Caption = ""

TgglBtn_Price_R3.Caption = ""
TgglBtn_Price_R2.Caption = ""
TgglBtn_Price_R1.Caption = ""
TgglBtn_Price_PP.Caption = ""
TgglBtn_Price_S1.Caption = ""
TgglBtn_Price_S2.Caption = ""
TgglBtn_Price_S3.Caption = ""

TgglBtn_Price_MAVG20D.Caption = ""
TgglBtn_Price_MAVG100D.Caption = ""
TgglBtn_Price_MAVG200D.Caption = ""

L_yesterday_close_value.Caption = ""
L_open_price_value.Caption = ""
L_pct_since_open_value.Caption = ""
L_pct_since_last_close_value.Caption = ""

TgglBtn_position_current.Caption = ""
TgglBtn_position_yst_close.Caption = ""
TgglBtn_position_delta.Caption = ""

L_postion_pnl_daily_value.Caption = ""
L_postion_pnl_ytd_value.Caption = ""


End Sub


Private Sub btn_sparkline_gpo_Click()

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

'prend plutot le fullpath du book car le form ne sera pas forcement lancer depuis le wrbk du book (possible depuis partout, prt etc.)
Dim tmp_wrbk As Workbook
For Each tmp_wrbk In Workbooks
    If UCase(tmp_wrbk.name) = UCase(book_name) Then
        base_path = StrReverse(Mid(StrReverse(tmp_wrbk.FullName), InStr(StrReverse(tmp_wrbk.FullName), "\")))
        Exit For
    End If
Next

Dim path_picture As String
path_picture = base_path & sparkline_tmp_file

If TB_order_ticker.Value <> "" Then

    Call draw_sparkline_to_picture(Array(TB_order_ticker.Value), Array("px_last"), "DAILY", 252, path_picture, 20, 50, 200)
    
    Me.img_sparklines.Picture = LoadPicture(path_picture)
    
End If

End Sub


Private Sub btn_sparkline_intraday_Click()

Dim base_path As String
base_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

'prend plutot le fullpath du book car le form ne sera pas forcement lancer depuis le wrbk du book (possible depuis partout, prt etc.)
Dim tmp_wrbk As Workbook
For Each tmp_wrbk In Workbooks
    If UCase(tmp_wrbk.name) = UCase(book_name) Then
        base_path = StrReverse(Mid(StrReverse(tmp_wrbk.FullName), InStr(StrReverse(tmp_wrbk.FullName), "\")))
        Exit For
    End If
Next

Dim path_picture As String
path_picture = base_path & sparkline_tmp_file

If TB_order_ticker.Value <> "" Then
    Call draw_sparkline_to_picture(Array(TB_order_ticker.Value), Array("last price"), "INTRADAY", 3, path_picture, 20, 60)
    Me.img_sparklines.Picture = LoadPicture(path_picture)
End If

End Sub



Private Function check_ticker() As Boolean

check_ticker = False

If TB_order_ticker.Value <> "" Then
    If Right(UCase(TB_order_ticker.Value), 6) = "EQUITY" Or Right(UCase(TB_order_ticker.Value), 5) = "INDEX" Then
        check_ticker = True
        Exit Function
    Else
        check_ticker = False
        Exit Function
    End If
Else
    check_ticker = False
    Exit Function
End If

End Function


Private Function check_price() As Boolean

check_price = False

If TB_custom_price.Value <> "" Then
    If IsNumeric(TB_custom_price.Value) Then
        If CDbl(TB_custom_price.Value) > 0 Then
            check_price = True
            Exit Function
        Else
            check_price = False
            Exit Function
        End If
    Else
        check_price = False
        Exit Function
    End If
Else
    check_price = False
    Exit Function
End If


End Function


Private Function check_side() As Boolean

check_side = False

If CB_Order_Side.Value = "" Then
    If TB_order_qty.Value <> "" And IsNumeric(TB_order_qty.Value) = True Then
        If CDbl(TB_order_qty.Value) < 0 Then
            CB_Order_Side.Value = "SELL"
            check_side = True
            Exit Function
        ElseIf CDbl(TB_order_qty.Value) > 0 Then
            CB_Order_Side.Value = "BUY"
            check_side = True
            Exit Function
        Else
            check_side = False
            Exit Function
        End If
    Else
        check_side = False
        Exit Function
    End If
Else
    's'assure que dans le bon sens
    If TB_order_qty.Value <> "" And IsNumeric(TB_order_qty.Value) = True Then
        If CDbl(TB_order_qty.Value) < 0 Then
            If UCase(CB_Order_Side.Value) = "SELL" Then
                check_side = True
                Exit Function
            Else
                check_side = False
                Exit Function
            End If
        ElseIf CDbl(TB_order_qty.Value) > 0 Then
            If UCase(CB_Order_Side.Value) = "BUY" Then
                check_side = True
                Exit Function
            Else
                check_side = False
                Exit Function
            End If
        Else
            check_side = False
            Exit Function
        End If
    Else
        check_side = False
        Exit Function
    End If
End If

End Function


Private Function check_qty() As Boolean

check_qty = False

If TB_order_qty.Value <> "" And IsNumeric(TB_order_qty.Value) = True Then
    If CDbl(TB_order_qty.Value) > 0 Or CDbl(TB_order_qty.Value) < 0 Then
        check_qty = True
        Exit Function
    Else
        check_qty = False
        Exit Function
    End If
Else
    check_qty = False
    Exit Function
End If


End Function


Private Sub btn_trade_it_with_redi_plus_Click()

Dim vec_trades() As Variant

If check_side And check_qty And check_ticker And check_price Then
    
    ReDim Preserve vec_trades(0)
    vec_trades(0) = Array("", 0, 0)
    
    vec_trades(0)(dim_r_plus_trades_symbol) = UCase(TB_order_ticker.Value)
    vec_trades(0)(dim_r_plus_trades_qty) = CDbl(TB_order_qty.Value)
    vec_trades(0)(dim_r_plus_trades_price) = CDbl(TB_custom_price.Value)
    
    Dim exec_orders As Variant
    exec_orders = universal_trades_r_plus(vec_trades)
End If

End Sub


Private Sub CB_aim_strategy_Change()

End Sub

Private Sub CB_aim_strategy_DropButtonClick()

Dim i As Integer

Dim list_strategy() As Variant
list_strategy = Array("HEDGING", "TRADING", "LONG GAMMA", "SHORT GAMMA", "BROKER IDEA", "GROWTH STOCK", "IPO PLACING", "SHORT", "DIV FUTURE")

If CB_aim_strategy.ListCount = 0 Then
    For i = 0 To UBound(list_strategy, 1)
        CB_aim_strategy.AddItem list_strategy(i)
    Next i
End If



End Sub

Private Sub CB_exec_broker_Change()

End Sub

Private Sub CB_exec_broker_DropButtonClick()

Dim i As Integer

list_brokers = Array("BARCLAYS", "CS", "DB", "GOLDMAN", "JPM", "ML", "MS", "NEGOCE", "WONEIL")

If CB_exec_broker.ListCount = 0 Then
    For i = 0 To UBound(list_brokers, 1)
        CB_exec_broker.AddItem list_brokers(i)
    Next i
End If

End Sub

Private Sub CB_Order_Side_Change()

If CB_Order_Side.Value <> "" And (UCase(CB_Order_Side.Value) = "SELL" Or UCase(CB_Order_Side.Value) = "BUY") Then
    If TB_order_qty.Value <> "" And IsNumeric(TB_order_qty.Value) Then
        If UCase(CB_Order_Side.Value) = "SELL" Then
            TB_order_qty.Value = -Abs(CDbl(TB_order_qty.Value))
        ElseIf UCase(CB_Order_Side.Value) = "BUY" Then
            TB_order_qty.Value = Abs(CDbl(TB_order_qty.Value))
        End If
    End If
Else
    CB_Order_Side.Value = ""
End If

End Sub


Private Sub CB_Order_Side_DropButtonClick()


If CB_Order_Side.ListCount = 0 Then

    CB_Order_Side.Clear
    
    CB_Order_Side.AddItem "BUY"
    CB_Order_Side.AddItem "SELL"

End If

End Sub


Private Sub frame_order_type_Click()

End Sub

Private Sub L_order_qty_Click()

Dim qty As Double

If IsNumeric(TB_order_qty.Value) And TB_order_qty.Value <> "" Then
    qty = -TB_order_qty.Value
    TB_order_qty.Value = qty
End If

End Sub


Private Sub L_Side_Click()

Dim i As Integer, j As Integer, k As Integer

Dim lst_possibilities As Variant
lst_possibilities = Array(Array("BUY", 1), Array("SELL", -1))

If CB_Order_Side.Value <> "" Then
    For i = 0 To UBound(lst_possibilities, 1)
        If CB_Order_Side.Value = lst_possibilities(i)(0) Then
            k = i
            
            For j = 0 To UBound(lst_possibilities, 1)
                If j <> k Then
                    CB_Order_Side.Value = lst_possibilities(j)(0)
                    
                    If TB_order_qty.Value <> "" And IsNumeric(TB_order_qty.Value) Then
                        TB_order_qty.Value = lst_possibilities(j)(1) * Abs(CDbl(TB_order_qty.Value))
                    End If
                    
                End If
            Next j
            
            Exit For
        Else
            If i = UBound(lst_possibilities, 1) Then
                CB_Order_Side.Value = ""
            End If
        End If
    Next i
Else
End If

End Sub


Private Sub quick_btn_pivot_Click()

Dim size As Double, ticker As String

Dim vec_trades() As Variant

k = 0
If check_ticker And quick_btn_pivot_qty_size.Value <> "" And IsNumeric(quick_btn_pivot_qty_size.Value) Then
    If CDbl(quick_btn_pivot_qty_size.Value) > 0 Then
        
        size = CDbl(quick_btn_pivot_qty_size.Value)
        ticker = UCase(TB_order_ticker.Value)
        
        
        Dim s3 As Double, s2 As Double, s1 As Double, pp As Double
        Dim r1 As Double, r2 As Double, r3 As Double
        
        
        'BUY
        If TgglBtn_Price_S3.Caption <> "" And IsNumeric(TgglBtn_Price_S3.Value) Then
            s3 = CDbl(TgglBtn_Price_S3.Caption)
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(ticker, Abs(size), s3)
            k = k + 1
        End If
        
        If TgglBtn_Price_S2.Caption <> "" And IsNumeric(TgglBtn_Price_S2.Value) Then
            s2 = CDbl(TgglBtn_Price_S2.Caption)
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(ticker, Abs(size), s2)
            k = k + 1
        End If
        
        If TgglBtn_Price_S1.Caption <> "" And IsNumeric(TgglBtn_Price_S1.Value) Then
            s1 = CDbl(TgglBtn_Price_S1.Caption)
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(ticker, Abs(size), s1)
            k = k + 1
        End If
        
        If TgglBtn_Price_PP.Caption <> "" And IsNumeric(TgglBtn_Price_PP.Value) Then
            pp = CDbl(TgglBtn_Price_PP.Caption)
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(ticker, Abs(size), pp)
            k = k + 1
        End If
        
        
        'SELL
        If TgglBtn_Price_R1.Caption <> "" And IsNumeric(TgglBtn_Price_R1.Value) Then
            r1 = CDbl(TgglBtn_Price_R1.Caption)
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(ticker, -Abs(size), r1)
            k = k + 1
        End If
        
        If TgglBtn_Price_R2.Caption <> "" And IsNumeric(TgglBtn_Price_R2.Value) Then
            r2 = CDbl(TgglBtn_Price_R2.Caption)
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(ticker, -Abs(size), r2)
            k = k + 1
        End If
        
        If TgglBtn_Price_R1.Caption <> "" And IsNumeric(TgglBtn_Price_R1.Value) Then
            r1 = CDbl(TgglBtn_Price_R1.Caption)
            ReDim Preserve vec_trades(k)
            vec_trades(k) = Array(ticker, -Abs(size), r1)
            k = k + 1
        End If
        
        If k > 0 Then
            Dim exec_orders As Variant
            exec_orders = universal_trades_r_plus(vec_trades)
        End If
        
    End If
End If

End Sub


Private Sub quick_btn_pivot_qty_size_Enter()

If IsNumeric(quick_btn_pivot_qty_size.Value) Then
Else
    quick_btn_pivot_qty_size.Value = ""
End If

End Sub


Private Sub quick_btn_pivot_qty_size_Exit(ByVal Cancel As msforms.ReturnBoolean)

If quick_btn_pivot_qty_size.Value = "" Or IsNumeric(quick_btn_pivot_qty_size.Value) = False Then
    quick_btn_pivot_qty_size = quick_btn_tooltip_stairs_size
End If

End Sub


Private Sub quick_btn_pivot_qty_size_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With quick_btn_pivot_qty_size
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub quick_btn_stairs_buy_Click()

Dim i As Integer, j As Integer, k As Integer

Dim start_price As Double, size As Double, incr As Double, step As Integer
Dim ticker As String

If check_ticker Then
    
    ticker = UCase(TB_order_ticker.Value)
    
    If TB_custom_price.Value <> "" Then
        If IsNumeric(TB_custom_price.Value) Then
            If CDbl(TB_custom_price.Value) > 0 Then
                start_price = CDbl(TB_custom_price.Value)
            Else
                MsgBox ("limit price error")
                Exit Sub
            End If
        Else
            MsgBox ("limit price error")
            Exit Sub
        End If
    Else
        'prendre caption last price si dispo
        If TgglBtn_Price_Last.Caption <> "" Then
            If IsNumeric(TgglBtn_Price_Last.Caption) Then
                If CDbl(TgglBtn_Price_Last.Caption) > 0 Then
                    start_price = CDbl(TgglBtn_Price_Last.Caption)
                Else
                    MsgBox ("last price error")
                    Exit Sub
                End If
            Else
                MsgBox ("no price available")
                Exit Sub
            End If
        Else
            MsgBox ("no price available")
            Exit Sub
        End If
    End If
    
    
    
    
    If IsNumeric(quick_btn_stairs_buy_qty_size.Value) Then
        If CDbl(quick_btn_stairs_buy_qty_size.Value) > 0 Then
            size = CDbl(quick_btn_stairs_buy_qty_size.Value)
        Else
            MsgBox ("size problem")
            Exit Sub
        End If
    Else
        MsgBox ("size problem")
        Exit Sub
    End If
    
    
    If IsNumeric(quick_btn_stairs_buy_incr.Value) Then
        If CDbl(quick_btn_stairs_buy_incr.Value) > 0 Then
            incr = quick_btn_stairs_buy_incr.Value
        Else
            MsgBox ("increment problem")
            Exit Sub
        End If
    Else
        MsgBox ("increment problem")
        Exit Sub
    End If
    
    
    If IsNumeric(quick_btn_stairs_buy_step.Value) Then
        If CDbl(quick_btn_stairs_buy_step.Value) > 0 Then
            step = CInt(quick_btn_stairs_buy_step.Value)
        Else
            MsgBox ("step problem")
            Exit Sub
        End If
    Else
        MsgBox ("step problem")
        Exit Sub
    End If
    
    
    Dim vec_trades() As Variant
    k = 0
    For i = 1 To step
        ReDim Preserve vec_trades(k)
        vec_trades(k) = Array(ticker, Abs(size), start_price - (i - 1) * incr)
        k = k + 1
    Next i
    
    If k > 0 Then
        Dim exec_orders As Variant
        exec_orders = universal_trades_r_plus(vec_trades)
    End If
    
Else
    MsgBox ("ticker error")
    Exit Sub
End If

End Sub

Private Sub quick_btn_stairs_buy_incr_Enter()

If IsNumeric(quick_btn_stairs_buy_incr.Value) Then
Else
    quick_btn_stairs_buy_incr.Value = ""
End If

End Sub


Private Sub quick_btn_stairs_buy_incr_Exit(ByVal Cancel As msforms.ReturnBoolean)

If quick_btn_stairs_buy_incr.Value = "" Or IsNumeric(quick_btn_stairs_buy_incr.Value) = False Then
    quick_btn_stairs_buy_incr = quick_btn_tooltip_stairs_increment
End If

End Sub


Private Sub quick_btn_stairs_buy_incr_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With quick_btn_stairs_buy_incr
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub quick_btn_stairs_buy_qty_size_Enter()

If IsNumeric(quick_btn_stairs_buy_qty_size.Value) Then
Else
    quick_btn_stairs_buy_qty_size.Value = ""
End If

End Sub


Private Sub quick_btn_stairs_buy_qty_size_Exit(ByVal Cancel As msforms.ReturnBoolean)

If quick_btn_stairs_buy_qty_size.Value = "" Or IsNumeric(quick_btn_stairs_buy_qty_size.Value) = False Then
    quick_btn_stairs_buy_qty_size = quick_btn_tooltip_stairs_size
End If

End Sub


Private Sub quick_btn_stairs_buy_qty_size_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With quick_btn_stairs_buy_qty_size
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub quick_btn_stairs_buy_step_Enter()

If IsNumeric(quick_btn_stairs_buy_step.Value) Then
Else
    quick_btn_stairs_buy_step.Value = ""
End If

End Sub


Private Sub quick_btn_stairs_buy_step_Exit(ByVal Cancel As msforms.ReturnBoolean)

If quick_btn_stairs_buy_step.Value = "" Or IsNumeric(quick_btn_stairs_buy_step.Value) = False Then
    quick_btn_stairs_buy_step = quick_btn_tooltip_stairs_step
End If

End Sub


Private Sub quick_btn_stairs_buy_step_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With quick_btn_stairs_buy_step
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub quick_btn_stairs_sell_Click()

Dim i As Integer, j As Integer, k As Integer

Dim start_price As Double, size As Double, incr As Double, step As Integer
Dim ticker As String

If check_ticker Then
    
    ticker = UCase(TB_order_ticker.Value)
    
    If TB_custom_price.Value <> "" Then
        If IsNumeric(TB_custom_price.Value) Then
            If CDbl(TB_custom_price.Value) > 0 Then
                start_price = CDbl(TB_custom_price.Value)
            Else
                MsgBox ("limit price error")
                Exit Sub
            End If
        Else
            MsgBox ("limit price error")
            Exit Sub
        End If
    Else
        'prendre caption last price si dispo
        If TgglBtn_Price_Last.Caption <> "" Then
            If IsNumeric(TgglBtn_Price_Last.Caption) Then
                If CDbl(TgglBtn_Price_Last.Caption) > 0 Then
                    start_price = CDbl(TgglBtn_Price_Last.Caption)
                Else
                    MsgBox ("last price error")
                    Exit Sub
                End If
            Else
                MsgBox ("no price available")
                Exit Sub
            End If
        Else
            MsgBox ("no price available")
            Exit Sub
        End If
    End If
    
    
    
    
    If IsNumeric(quick_btn_stairs_sell_qty_size.Value) Then
        If CDbl(quick_btn_stairs_sell_qty_size.Value) > 0 Then
            size = CDbl(quick_btn_stairs_sell_qty_size.Value)
        Else
            MsgBox ("size problem")
            Exit Sub
        End If
    Else
        MsgBox ("size problem")
        Exit Sub
    End If
    
    
    If IsNumeric(quick_btn_stairs_sell_incr.Value) Then
        If CDbl(quick_btn_stairs_sell_incr.Value) > 0 Then
            incr = quick_btn_stairs_sell_incr.Value
        Else
            MsgBox ("increment problem")
            Exit Sub
        End If
    Else
        MsgBox ("increment problem")
        Exit Sub
    End If
    
    
    If IsNumeric(quick_btn_stairs_sell_step.Value) Then
        If CDbl(quick_btn_stairs_sell_step.Value) > 0 Then
            step = CInt(quick_btn_stairs_sell_step.Value)
        Else
            MsgBox ("step problem")
            Exit Sub
        End If
    Else
        MsgBox ("step problem")
        Exit Sub
    End If
    
    
    Dim vec_trades() As Variant
    k = 0
    For i = 1 To step
        ReDim Preserve vec_trades(k)
        vec_trades(k) = Array(ticker, -Abs(size), start_price + (i - 1) * incr)
        k = k + 1
    Next i
    
    If k > 0 Then
        Dim exec_orders As Variant
        exec_orders = universal_trades_r_plus(vec_trades)
    End If
    
Else
    MsgBox ("ticker error")
    Exit Sub
End If

End Sub

Private Sub quick_btn_stairs_sell_incr_Enter()

If IsNumeric(quick_btn_stairs_sell_incr.Value) Then
Else
    quick_btn_stairs_sell_incr.Value = ""
End If

End Sub


Private Sub quick_btn_stairs_sell_incr_Exit(ByVal Cancel As msforms.ReturnBoolean)

If quick_btn_stairs_sell_incr.Value = "" Or IsNumeric(quick_btn_stairs_sell_incr.Value) = False Then
    quick_btn_stairs_sell_incr = quick_btn_tooltip_stairs_increment
End If

End Sub


Private Sub quick_btn_stairs_sell_incr_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With quick_btn_stairs_sell_incr
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub quick_btn_stairs_sell_qty_size_Enter()

If IsNumeric(quick_btn_stairs_sell_qty_size.Value) Then
Else
    quick_btn_stairs_sell_qty_size.Value = ""
End If

End Sub


Private Sub quick_btn_stairs_sell_qty_size_Exit(ByVal Cancel As msforms.ReturnBoolean)

If quick_btn_stairs_sell_qty_size.Value = "" Or IsNumeric(quick_btn_stairs_sell_qty_size.Value) = False Then
    quick_btn_stairs_sell_qty_size = quick_btn_tooltip_stairs_size
End If

End Sub


Private Sub quick_btn_stairs_sell_qty_size_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With quick_btn_stairs_sell_qty_size
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub quick_btn_stairs_sell_step_Enter()

If IsNumeric(quick_btn_stairs_sell_step.Value) Then
Else
    quick_btn_stairs_sell_step.Value = ""
End If

End Sub


Private Sub quick_btn_stairs_sell_step_Exit(ByVal Cancel As msforms.ReturnBoolean)

If quick_btn_stairs_sell_step.Value = "" Or IsNumeric(quick_btn_stairs_sell_step.Value) = False Then
    quick_btn_stairs_sell_step = quick_btn_tooltip_stairs_step
End If

End Sub


Private Sub quick_btn_stairs_sell_step_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With quick_btn_stairs_sell_step
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub TB_custom_price_Exit(ByVal Cancel As msforms.ReturnBoolean)

'passe en revue les différents prix des tggle btn
Dim tmp_control As Control

Dim label As String
label = ""

For Each tmp_control In Me.Controls
    If TypeOf tmp_control Is msforms.ToggleButton And Left(tmp_control.name, 14) = "TgglBtn_Price_" Then
        
        If tmp_control.Caption = TB_custom_price Then
            label = UCase(Replace(tmp_control.name, "TgglBtn_Price_", ""))
            Exit For
        End If
        
    End If
Next

If label <> "" Then
    L_limit_price.Caption = "LIMIT (" & label & ")"
Else
    L_limit_price.Caption = "LIMIT"
End If


End Sub


Private Sub TB_custom_price_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With TB_custom_price
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub TB_order_qty_Change()

If IsNumeric(TB_order_qty.Value) And TB_order_qty.Value <> "" Then
    If Left(TB_order_qty.Value, 1) = "-" Then
        Call update_side
        update_qty_everywhere
    Else
        If UCase(CB_Order_Side.Value) = "SELL" Then
            TB_order_qty.Value = "-" & TB_order_qty.Value
            Call update_side
            update_qty_everywhere
        End If
    End If
End If

End Sub


Private Sub update_side()

Dim qty As Double

If IsNumeric(TB_order_qty.Value) And TB_order_qty.Value <> "" Then
    qty = CDbl(TB_order_qty.Value)
    
    If qty < 0 Then
        CB_Order_Side.Value = "SELL"
    Else
        CB_Order_Side.Value = "BUY"
    End If
End If

End Sub


Private Sub update_qty_everywhere()

Dim tmp_control As Control

With TB_order_qty
    If IsNumeric(TB_order_qty.Value) And TB_order_qty.Value <> "" Then
        
        With TB_order_qty
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.TextBox And InStr(UCase(tmp_control.name), "QTY") And tmp_control.name <> .name Then
                    tmp_control.Value = Abs(CDbl(TB_order_qty.Value))
                End If
            Next
        End With
    End If
End With

End Sub


Private Sub TB_order_qty_Exit(ByVal Cancel As msforms.ReturnBoolean)

If IsNumeric(TB_order_qty.Value) = False Then
    TB_order_qty.Value = ""
End If

End Sub


Private Sub load_ticker_datas()

Dim oBBG As New cls_Bloomberg_Sync

Dim oReg As New VBScript_RegExp_55.RegExp
    oReg.Global = True
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

If UCase(Right(TB_order_ticker.Value, 6)) = "EQUITY" Or UCase(Right(TB_order_ticker.Value, 5)) = "INDEX" Then
    
    Dim ticker As String
    ticker = UCase(TB_order_ticker.Value)
    
    'patch si sous forme de symbol pour le marche us
    If UCase(Right(TB_order_ticker.Value, 6)) = "EQUITY" Then
        oReg.Pattern = "^[A-Za-z0-9]+(\s)EQUITY$"
        Set matches = oReg.Execute(ticker)
        
        For Each match In matches
            ticker = Left(ticker, InStr(ticker, " ") - 1) & " US EQUITY"
        Next
    End If
    
    Dim bdp_data As Variant
    
    Me.Caption = form_caption_base & " - Please wait..."
    
    
    Dim bbg_field() As Variant
    bbg_field = Array("name", "px_last", "CHG_PCT_1D", "PX_BID", "PX_ASK", "PX_OPEN", "PX_HIGH", "PX_LOW", "PX_VOLUME", _
        "PX_YEST_CLOSE", "PX_YEST_HIGH", "PX_YEST_LOW", "MOV_AVG_20D", "MOV_AVG_100D", "MOV_AVG_200D")
    
        
    
    'bdp_data = bbg_multi_tickers_and_multi_fields(Array(ticker), bbg_field)
    bdp_data = oBBG.bdp(Array(ticker), bbg_field, output_format.of_vec_without_header)
        
        
        For i = 0 To UBound(bbg_field, 1)
            If UCase(bbg_field(i)) = UCase("name") Then
                dim_compagny_name = i
            ElseIf UCase(bbg_field(i)) = UCase("px_last") Then
                dim_px_last = i
            ElseIf UCase(bbg_field(i)) = UCase("CHG_PCT_1D") Then
                dim_chg_pct = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_BID") Then
                dim_bid = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_ASK") Then
                dim_ask = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_OPEN") Then
                dim_open = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_HIGH") Then
                dim_high = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_LOW") Then
                dim_low = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_VOLUME") Then
                dim_volume = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_CLOSE") Then
                dim_yest_close = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_HIGH") Then
                dim_yest_high = i
            ElseIf UCase(bbg_field(i)) = UCase("PX_YEST_LOW") Then
                dim_yest_low = i
            ElseIf UCase(bbg_field(i)) = UCase("MOV_AVG_20D") Then
                dim_mov_avg_20d = i
            ElseIf UCase(bbg_field(i)) = UCase("MOV_AVG_100D") Then
                dim_mov_avg_100d = i
            ElseIf UCase(bbg_field(i)) = UCase("MOV_AVG_200D") Then
                dim_mov_avg_200d = i
            End If
        Next i
        
        
    'ajuste les valeurs des boutons
    frame_market_data.Caption = bdp_data(0)(dim_compagny_name)
    
    Dim px_last As Double
    px_last = 0
    
    If IsNumeric(bdp_data(0)(dim_px_last)) = True And Left(bdp_data(0)(dim_px_last), 1) <> "#" Then
        TgglBtn_Price_Last.Caption = bdp_data(0)(dim_px_last)
        px_last = bdp_data(0)(dim_px_last)
    End If
    
    If IsNumeric(bdp_data(0)(dim_chg_pct)) = True And Left(bdp_data(0)(dim_chg_pct), 1) <> "#" Then
        L_change_value.Caption = Round(bdp_data(0)(dim_chg_pct), 2) & "%"
    End If
    
    If IsNumeric(bdp_data(0)(dim_bid)) = True And Left(bdp_data(0)(dim_bid), 1) <> "#" Then
        TgglBtn_Price_Bid.Caption = bdp_data(0)(dim_bid)
    End If
    
    If IsNumeric(bdp_data(0)(dim_ask)) = True And Left(bdp_data(0)(dim_ask), 1) <> "#" Then
        TgglBtn_Price_Ask.Caption = bdp_data(0)(dim_ask)
    End If
    
    If IsNumeric(bdp_data(0)(dim_high)) = True And Left(bdp_data(0)(dim_high), 1) <> "#" Then
        TgglBtn_Price_High.Caption = bdp_data(0)(dim_high)
    End If
    
    If IsNumeric(bdp_data(0)(dim_low)) = True And Left(bdp_data(0)(dim_low), 1) <> "#" Then
        TgglBtn_Price_Low.Caption = bdp_data(0)(dim_low)
    End If
    
    If IsNumeric(bdp_data(0)(dim_volume)) = True And Left(bdp_data(0)(dim_volume), 1) <> "#" Then
        L_volume_value.Caption = bdp_data(0)(dim_volume)
    End If
    
    If IsNumeric(bdp_data(0)(dim_mov_avg_20d)) = True And Left(bdp_data(0)(dim_mov_avg_20d), 1) <> "#" Then
        TgglBtn_Price_MAVG20D.Caption = bdp_data(0)(dim_mov_avg_20d)
        
        'signal
        With TgglBtn_Price_MAVG20D
            If px_last <> 0 And px_last <= 1.001 * bdp_data(0)(dim_mov_avg_20d) And px_last >= 0.999 * bdp_data(0)(dim_mov_avg_20d) Then
                .BackStyle = fmBackStyleOpaque
                .BackColor = ActiveWorkbook.Colors(23)
                .ForeColor = ActiveWorkbook.Colors(2)
            End If
        End With
        
    End If
    
    If IsNumeric(bdp_data(0)(dim_mov_avg_100d)) = True And Left(bdp_data(0)(dim_mov_avg_100d), 1) <> "#" Then
        TgglBtn_Price_MAVG100D.Caption = bdp_data(0)(dim_mov_avg_100d)
        
        'signal
        With TgglBtn_Price_MAVG100D
            If px_last <> 0 And px_last <= 1.001 * bdp_data(0)(dim_mov_avg_100d) And px_last >= 0.999 * bdp_data(0)(dim_mov_avg_100d) Then
                .BackStyle = fmBackStyleOpaque
                .BackColor = ActiveWorkbook.Colors(5)
                .ForeColor = ActiveWorkbook.Colors(2)
            End If
        End With
    End If
    
    If IsNumeric(bdp_data(0)(dim_mov_avg_200d)) = True And Left(bdp_data(0)(dim_mov_avg_200d), 1) <> "#" Then
        TgglBtn_Price_MAVG200D.Caption = bdp_data(0)(dim_mov_avg_200d)
        
        'signal
        With TgglBtn_Price_MAVG200D
            If px_last <> 0 And px_last <= 1.001 * bdp_data(0)(dim_mov_avg_200d) And px_last >= 0.999 * bdp_data(0)(dim_mov_avg_200d) Then
                .BackStyle = fmBackStyleOpaque
                .BackColor = ActiveWorkbook.Colors(25)
                .ForeColor = ActiveWorkbook.Colors(2)
            End If
        End With
    End If
    
    If IsNumeric(bdp_data(0)(dim_yest_close)) = True And Left(bdp_data(0)(dim_yest_close), 1) <> "#" Then
        L_yesterday_close_value.Caption = bdp_data(0)(dim_yest_close)
    End If
    
    
    If IsNumeric(bdp_data(0)(dim_open)) = True And Left(bdp_data(0)(dim_open), 1) <> "#" Then
        L_open_price_value.Caption = bdp_data(0)(dim_open)
    End If
    
    
    'calcul %
    If IsNumeric(bdp_data(0)(dim_open)) = True And IsNumeric(bdp_data(0)(dim_px_last)) = True Then
        L_pct_since_open_value.Caption = Round(100 * ((bdp_data(0)(dim_px_last) / bdp_data(0)(dim_open)) - 1), 2) & "%"
    End If
    
    If IsNumeric(bdp_data(0)(dim_yest_close)) = True And IsNumeric(bdp_data(0)(dim_px_last)) = True Then
        L_pct_since_last_close_value.Caption = Round(100 * ((bdp_data(0)(dim_px_last) / bdp_data(0)(dim_yest_close)) - 1), 2) & "%"
    End If
    
    
    'calcul des pivots
    If IsNumeric(bdp_data(0)(dim_yest_close)) = True And IsNumeric(bdp_data(0)(dim_yest_high)) = True And IsNumeric(bdp_data(0)(dim_yest_low)) = True Then
        Dim pp As Double
        pp = (bdp_data(0)(dim_yest_close) + bdp_data(0)(dim_yest_high) + bdp_data(0)(dim_yest_low)) / 3
        
        Dim s1 As Double
        s1 = 2 * pp - bdp_data(0)(dim_yest_high)
        
        Dim r1 As Double
        r1 = 2 * pp - bdp_data(0)(dim_yest_low)
        
        Dim s2 As Double
        s2 = pp + (s1 - r1)
        
        Dim r2 As Double
        r2 = pp - (s1 - r1)
        
        Dim s3 As Double
        's3 = pp + (s1 - r2)
        s3 = pp + (s2 - r2)
        
        Dim r3 As Double
        'r3 = pp - (s2 - r1)
        r3 = pp - (s2 - r2)
        
        
        TgglBtn_Price_R3.Caption = Round(r3, 2)
            'signal
            With TgglBtn_Price_R3
                If px_last <> 0 And px_last <= 1.001 * r3 And px_last >= 0.999 * r3 Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = ActiveWorkbook.Colors(51)
                    .ForeColor = ActiveWorkbook.Colors(2)
                End If
            End With
            
            
        TgglBtn_Price_R2.Caption = Round(r2, 2)
            'signal
            With TgglBtn_Price_R2
                If px_last <> 0 And px_last <= 1.001 * r2 And px_last >= 0.999 * r2 Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = ActiveWorkbook.Colors(43)
                    .ForeColor = ActiveWorkbook.Colors(2)
                End If
            End With
            
            
        TgglBtn_Price_R1.Caption = Round(r1, 2)
            'signal
            With TgglBtn_Price_R1
                If px_last <> 0 And px_last <= 1.001 * r1 And px_last >= 0.999 * r1 Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = ActiveWorkbook.Colors(4)
                    .ForeColor = ActiveWorkbook.Colors(1)
                End If
            End With
            
        TgglBtn_Price_PP.Caption = Round(pp, 2)
            'signal
            With TgglBtn_Price_PP
                If px_last <> 0 And px_last <= 1.001 * pp And px_last >= 0.999 * pp Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = ActiveWorkbook.Colors(6)
                    .ForeColor = ActiveWorkbook.Colors(1)
                End If
            End With
        
        TgglBtn_Price_S1.Caption = Round(s1, 2)
            'signal
            With TgglBtn_Price_S1
                If px_last <> 0 And px_last <= 1.001 * s1 And px_last >= 0.999 * s1 Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = ActiveWorkbook.Colors(7)
                    .ForeColor = ActiveWorkbook.Colors(1)
                End If
            End With
            
        TgglBtn_Price_S2.Caption = Round(s2, 2)
            'signal
            With TgglBtn_Price_S2
                If px_last <> 0 And px_last <= 1.001 * s2 And px_last >= 0.999 * s2 Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = ActiveWorkbook.Colors(29)
                    .ForeColor = ActiveWorkbook.Colors(2)
                End If
            End With
            
        TgglBtn_Price_S3.Caption = Round(s3, 2)
            'signal
            With TgglBtn_Price_S3
                If px_last <> 0 And px_last <= 1.001 * s3 And px_last >= 0.999 * s3 Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = ActiveWorkbook.Colors(22)
                    .ForeColor = ActiveWorkbook.Colors(1)
                End If
            End With
    
    End If
    
    
    'recupere si une pos existe deja dans le book
    Dim position_current As Double, position_yst_close As Double, pnl_daily As Double, pnl_ytd As Double, delta As Double
        
        position_current = 0
        position_yst_close = 0
        pnl_daily = 0
        pnl_ytd = 0
        delta = 0
    
    If UCase(Right(TB_order_ticker.Value, 6)) = "EQUITY" Then
        For i = 27 To 32000
            If Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 47) = UCase(TB_order_ticker.Value) Then
                
                If IsError(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 24)) = False Then
                    If IsNumeric(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 24)) = True Then
                        position_current = Round(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 24), 0)
                    End If
                End If
                
                If IsError(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 25)) = False Then
                    If IsNumeric(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 25)) = True Then
                        position_yst_close = Round(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 25), 0)
                    End If
                End If
                
                
                If IsError(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 6)) = False Then
                    If IsNumeric(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 6)) = True Then
                        delta = Round(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 6), 0)
                    End If
                End If
                
                
                If IsError(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 13)) = False Then
                    If IsNumeric(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 13)) = True Then
                        pnl_daily = Round(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 13), 0)
                    End If
                End If
                
                
                If IsError(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 14)) = False Then
                    If IsNumeric(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 14)) = True Then
                        pnl_ytd = Round(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 14), 0)
                    End If
                End If
                
                
                
                
                Exit For
            End If
        Next i
    End If
    
    
    TgglBtn_position_current.Caption = position_current
    TgglBtn_position_yst_close.Caption = position_yst_close
    L_postion_pnl_daily_value.Caption = pnl_daily
    L_postion_pnl_ytd_value.Caption = pnl_ytd
    TgglBtn_position_delta.Caption = delta
    
    
    Me.Caption = form_caption_base
Else
    Call clear_market_datas
    Exit Sub
End If

End Sub


Private Sub TB_order_qty_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With TB_order_qty
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub TB_order_ticker_AfterUpdate()

Call load_ticker_datas

End Sub

Private Sub TB_order_ticker_Change()

'TB_order_ticker.value = UCase(TB_order_ticker.value) 'evite de recharger 10x les données API
'Call load_ticker_datas

End Sub


Private Sub load_name_product_in_ticker_list()

'Dim tmp_value_cbox As String
'tmp_value_cbox = TB_order_ticker.value

'charge la liste des equities_name
Dim i As Integer, j As Integer, k As Integer

Dim vec_name() As Variant

k = 0
For i = 27 To 32000 Step 2
    If Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_name(k)
        vec_name(k) = Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 2)
        k = k + 1
    End If
Next i

'sort ABC
Dim tmp_value As String
Dim tmp_pos As Integer

For i = 0 To UBound(vec_name, 1)
    tmp_pos = i
    tmp_value = vec_name(i)
    
    For j = i + 1 To UBound(vec_name, 1)
        If vec_name(j) < tmp_value Then
            tmp_value = vec_name(j)
            tmp_pos = j
        End If
    Next j
    
    If tmp_pos <> i Then
        tmp_value = vec_name(i)
        vec_name(i) = vec_name(tmp_pos)
        vec_name(tmp_pos) = tmp_value
    End If
Next i

TB_order_ticker.Clear
For i = 0 To UBound(vec_name, 1)
    TB_order_ticker.AddItem vec_name(i)
Next i

'TB_order_ticker.value = tmp_value_cbox

End Sub


Private Sub TB_order_ticker_DropButtonClick()

If TB_order_ticker.ListCount = 0 Then
    Call load_name_product_in_ticker_list
End If


End Sub


Private Sub TB_order_ticker_Enter()

'Dim i As Integer
'
'Dim list_suffix As Variant
'list_suffix = Array("EQUITY", "INDEX")
'
'For i = 0 To UBound(list_suffix, 1)
'    If TB_order_ticker.value <> "" Then
'        If UCase(Right(TB_order_ticker.value, 6)) = list_suffix(i) Then
'            TB_order_ticker.value = Mid(TB_order_ticker.value, 1, InStr(UCase(TB_order_ticker.value), list_suffix(i)) - 1)
'        End If
'    End If
'Next i

End Sub


Private Sub TB_order_ticker_Exit(ByVal Cancel As msforms.ReturnBoolean)

Dim i As Integer, j As Integer, k As Integer

Dim list_column_matching As Variant
list_column_matching = Array("Identifier", "Equities_Name", "BLOOMBERG", "ISIN")

Dim l_equity_db_header As Integer
l_equity_db_header = 25

For i = 0 To UBound(list_column_matching, 1)
    For j = 1 To 260
        If IsNumeric(list_column_matching(i)) Then
            Exit For
        Else
            If Workbooks(book_name).Worksheets("Equity_Database").Cells(l_equity_db_header, j) = list_column_matching(i) Then
                list_column_matching(i) = j
                Exit For
            End If
        End If
    Next j
Next i



If TB_order_ticker.Value <> "" And Right(UCase(TB_order_ticker.Value), 6) <> "EQUITY" And Right(TB_order_ticker.Value, 5) <> "INDEX" Then
    'id ou name
    For i = 27 To 32000 Step 2
        If Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 1) = "" Then
            Exit For
        Else
'            If Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 1) = TB_order_ticker.value Or Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 2) = TB_order_ticker.value Then
'                TB_order_ticker.value = CStr(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 47))
'                Exit For
'            End If
            
            For j = 0 To UBound(list_column_matching, 1)
                If UCase(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, list_column_matching(j))) = UCase(TB_order_ticker.Value) Then
                    If UCase(TB_order_ticker.Value) <> UCase(CStr(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 47))) Then
                        TB_order_ticker.Value = CStr(Workbooks(book_name).Worksheets("Equity_Database").Cells(i, 47))
                    End If
                    Exit Sub
                End If
            Next j
            
        End If
    Next i
Else
    If Right(UCase(TB_order_ticker.Value), 6) = "EQUITY" Or Right(TB_order_ticker.Value, 5) = "INDEX" Then
        'Call load_ticker_datas
    End If
End If

End Sub


Private Sub TB_order_ticker_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

With TB_order_ticker
    If .Value <> "" Then
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub


Private Sub TgglBtn_Price_Ask_Click()

Dim tmp_control As Control

With TgglBtn_Price_Ask
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With


End Sub


Private Sub TgglBtn_Price_Bid_Click()

Dim tmp_control As Control

With TgglBtn_Price_Bid
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_position_delta_Click()

Dim tmp_control As Control

With TgglBtn_position_delta
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "POSITION") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_order_qty.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_High_Click()

Dim tmp_control As Control

With TgglBtn_Price_High
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_Last_Click()


Dim tmp_control As Control

With TgglBtn_Price_Last
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_Low_Click()

Dim tmp_control As Control

With TgglBtn_Price_Low
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_MAVG200D_Click()

Dim tmp_control As Control

With TgglBtn_Price_MAVG200D
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_MAVG20D_Click()

Dim tmp_control As Control

With TgglBtn_Price_MAVG20D
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_MAVG100D_Click()

Dim tmp_control As Control

With TgglBtn_Price_MAVG100D
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_position_current_Click()

Dim tmp_control As Control

With TgglBtn_position_current
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "POSITION") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_order_qty.Value = .Caption
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_position_yst_close_Click()

Dim tmp_control As Control

With TgglBtn_position_yst_close
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "POSITION") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_order_qty.Value = .Caption
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_PP_Click()

Dim tmp_control As Control

With TgglBtn_Price_PP
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_R1_Click()

Dim tmp_control As Control

With TgglBtn_Price_R1
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_R2_Click()

Dim tmp_control As Control

With TgglBtn_Price_R2
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_R3_Click()

Dim tmp_control As Control

With TgglBtn_Price_R3
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_S1_Click()

Dim tmp_control As Control

With TgglBtn_Price_S1
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_S2_Click()

Dim tmp_control As Control

With TgglBtn_Price_S2
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub TgglBtn_Price_S3_Click()

Dim tmp_control As Control

With TgglBtn_Price_S3
    If .Caption <> "" Then
        If .Value = True Then
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.ToggleButton And InStr(UCase(tmp_control.name), "PRICE") <> 0 And tmp_control.name <> .name Then
                    tmp_control.Value = False
                End If
            Next
            
            TB_custom_price.Value = .Caption
            
            'repere le label joint pour rennomer le caption LIMIT
            For Each tmp_control In Me.Controls
                If TypeOf tmp_control Is msforms.label And tmp_control.name = "L_market_data_" & Replace(.name, "TgglBtn_Price_", "") Then
                    L_limit_price.Caption = "LIMIT (" & tmp_control.Caption & ")"
                    Exit For
                End If
            Next
        End If
    Else
        .Value = False
    End If
End With

End Sub


Private Sub UserForm_Activate()

Application.Calculation = xlCalculationManual

Dim tmp_worbook As Workbook, find_book As Boolean
find_book = False

For Each tmp_worbook In Workbooks
    If UCase(tmp_worbook.name) = UCase(book_name) Then
        find_book = True
        Exit For
    End If
Next

If find_book = False Then
    MsgBox ("le book " & book_name & " n'est pas ouvert, il est nécessaire pour déterminer les short sells etc.")
    Me.Hide
    Exit Sub
End If

'charge les info du trader (name / r+ id / accounts)
Dim id_rplus As String
id_rplus = Workbooks(book_name).Worksheets("FORMAT2").Cells(7, 21)

Dim account_equities As String, account_derivatives As String
account_equities = Workbooks(book_name).Worksheets("FORMAT2").Cells(6, 23)
account_derivatives = Workbooks(book_name).Worksheets("FORMAT2").Cells(6, 24)

Dim trader_name As String
trader_name = Workbooks(book_name).Worksheets("Parametres").Cells(17, 18)

L_trader_info.Caption = UCase(trader_name) & " - R+ id : " & id_rplus & " - equities account : " & account_equities & " - derivatives account : " & account_derivatives

If TB_order_ticker.Value <> "" Then
    load_ticker_datas
End If

mode_emsx = 0

End Sub


Private Function workday_between_2_dates(ByVal start_date As Date, ByVal end_date As Date) As Long


Dim start_date_long As Long, end_date_long As Long

start_date_long = start_date
end_date_long = end_date

Dim d As Long, dCount As Long
dCount = 0

    For d = start_date_long To end_date_long
        If Weekday(d, vbMonday) < 6 Then
            dCount = dCount + 1
        End If
    Next d

workday = dCount

End Function


Private Function workday_custom(ByVal start_date As Date, ByVal nbre_working_days As Long) As Date

workday_custom = start_date + CInt((nbre_working_days / 5) * 7)

End Function


Private Sub UserForm_Terminate()

Application.Calculation = xlCalculationAutomatic

End Sub
