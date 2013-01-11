Attribute VB_Name = "bas_pivot_external_file"
Sub Macro1()
'
' Macro1 Macro
' Macro recorded 24.06.2010 by stouff
'


 Sheets("spx").Activate
    Range("P3:P59").Select
    Selection.Copy
    Range("T3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R3:R59").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("U3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("T3:U59").Select
    Application.CutCopyMode = False
    Selection.Sort Key1:=Range("U3"), Order1:=xlDescending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
End Sub


Sub Macro2()
'
' Macro2 Macro
' Macro recorded 29.11.2005 by X01221106
'

'
Application.ScreenUpdating = False

 ' Initialize the Bloomberg data control
    Set objBloomberg = New BlpData

Sheets("table").Select
Range(Cells(1, 1), Cells(13, 18)).Select
Selection.Copy
Range(Cells(34, 1), Cells(46, 18)).PasteSpecial xlPasteValues


    Dim stocks
    Dim taille
    Dim Fields
    Dim periode
    periode = 30
    Sheets("Sheet1").Select
    Fields = Array("OPEN", "LOW", "HIGH", "PX_Last")
    Range("A20").Select
    Selection.End(xlDown).Select
    taille = ActiveCell.row - 20
    ReDim stocks(taille)
    Range("A20").Select
    For i = 0 To taille
    stocks(i) = ActiveCell.Offset(i, 0).Text
    Next
    

    ' Make the synchronous call for the historical price/volume
    ' of Microsoft and Dell over the last 30 days.
    objBloomberg.Periodicity = bbDaily
    vtResults = objBloomberg.BLPGetHistoricalData(stocks, Fields, Date - 30, Date - 1)

    x = LBound(vtResults, 2) - UBound(vtResults, 2)
    Dim Dates
    ReDim Dates(periode, taille)
    Dim Openp
    ReDim Openp(periode, taille)
    Dim Lowp
    ReDim Lowp(periode, taille)
    Dim Highp
    ReDim Highp(periode, taille)
    Dim Lastp
    ReDim Lastp(periode, taille)
    
    For nDate = LBound(vtResults, 1) To UBound(vtResults, 1)
        For i = 0 To taille
        Dates(nDate, i) = vtResults(UBound(vtResults, 1) - nDate, i, 0)
        Next i
    Next nDate
    
    For nDate = LBound(vtResults, 1) To UBound(vtResults, 1)
        For i = 0 To taille
            Openp(nDate, i) = vtResults(UBound(vtResults, 1) - nDate, i, 1)
        Next i
    Next nDate
    
    For nDate = LBound(vtResults, 1) To UBound(vtResults, 1)
        For i = 0 To taille
            Lowp(nDate, i) = vtResults(UBound(vtResults, 1) - nDate, i, 2)
        Next i
    Next nDate
    
    For nDate = LBound(vtResults, 1) To UBound(vtResults, 1)
        For i = 0 To taille
            Highp(nDate, i) = vtResults(UBound(vtResults, 1) - nDate, i, 3)
        Next i
    Next nDate
    
    For nDate = LBound(vtResults, 1) To UBound(vtResults, 1)
        For i = 0 To taille
            Lastp(nDate, i) = vtResults(UBound(vtResults, 1) - nDate, i, 4)
        Next i
    Next nDate
   
   
    'Ecriture Last
    'Date
    Sheets("Sheet1").Select
    Range("A2") = Dates(0, 0)
    
    'Prix
    Sheets("Sheet1").Select
    Range("C3").Select
    For x = 0 To (taille)
        ActiveCell.Offset(x, 0) = Openp(0, x)
    Next
    
    Range("D3").Select
    For x = 0 To (taille)
        ActiveCell.Offset(x, 0) = Lowp(0, x)
    Next
    
    Range("E3").Select
    For x = 0 To (taille)
        ActiveCell.Offset(x, 0) = Highp(0, x)
    Next
    
    Range("F3").Select
    For x = 0 To (taille)
        ActiveCell.Offset(x, 0) = Lastp(0, x)
    Next
    
    
    'Ecriture Max
    'Date
    Sheets("Sheet1").Select
    'Range("A18") = Dates(periode, 0)
    'Range("A19") = Dates(0, 0)
    
    'Prix
    Dim max
    Sheets("Sheet1").Select
    Range("C20").Select
    For x = 0 To (taille) 'boucle produit
        
        'find last friday
        For i = 0 To UBound(Dates, 1)
            If Weekday(Dates(i, x)) = 6 Then
                idx_last_friday_for_current_produt = i
                'Debug.Print "last friday: " & Dates(i, x)
                Range("A19") = Dates(i, x)
                Exit For
            End If
        Next i
        
        'find last monday before the last friday
        For i = idx_last_friday_for_current_produt + 1 To UBound(Dates, 1)
            If Weekday(Dates(i, x)) = 2 Then
                idx_last_monday_for_current_produt = i
                'Debug.Print "last monday before last friday: " & Dates(i, x)
                Range("A18") = Dates(i, x)
                Exit For
            End If
        Next i
        
        'low
        min_value = 1000000
        For i = idx_last_friday_for_current_produt To idx_last_monday_for_current_produt
            
            If IsNumeric(Lowp(i, x)) Then
                If Lowp(i, x) < min_value Then
                    min_value = Lowp(i, x)
                End If
            End If
            
        Next i
        
        'high
        max_value = -1
        For i = idx_last_friday_for_current_produt To idx_last_monday_for_current_produt
            
            If IsNumeric(Highp(i, x)) Then
                If Highp(i, x) > max_value Then
                    max_value = Highp(i, x)
                End If
            End If
            
        Next i
        
        
        
        'close
        close_value = 0
        For i = idx_last_friday_for_current_produt To idx_last_monday_for_current_produt
            
            If IsNumeric(Lastp(i, x)) Then
                close_value = Lastp(i, x)
                Exit For
            End If
            
        Next i
        
        
        
        'impression rapport
        Worksheets("Sheet1").Cells(20 + x, 1) = stocks(x)
        Worksheets("Sheet1").Cells(20 + x, 4) = min_value
        Worksheets("Sheet1").Cells(20 + x, 5) = max_value
        Worksheets("Sheet1").Cells(20 + x, 6) = close_value
        
    Next
    
'    Range("D20").Select   'low
'    For x = 0 To (taille)
'        i = 5
'        Do Until Weekday(Dates(i, x)) = 2
'        i = i + 1
'        Loop
'        Range("A18") = Dates(i, x)
'        j = 5
'        Do Until Weekday(Dates(j, x)) = 6
'        j = j - 1
'        Loop
'        Range("A19") = Dates(j, x)
'        max = Lowp(j, x)
'        For z = j To i - 1
'        If max < Lowp(z + 1, x) Then
'        ElseIf Lowp(z + 1, x) = "#N/A N.A." Then
'        Else
'        max = Lowp(z + 1, x)
'        End If
'        Next
'        ActiveCell.Offset(x, 0) = max
'    Next
'
'    Range("E20").Select
'    For x = 0 To (taille)
'        i = 5
'        Do Until Weekday(Dates(i, x)) = 2
'        i = i + 1
'        Loop
'        Range("A18") = Dates(i, x)
'        j = 5
'        Do Until Weekday(Dates(j, x)) = 6
'        j = j - 1
'        Loop
'        Range("A19") = Dates(j, x)
'        max = Highp(j, x)
'        For z = j To i - 1
'        If max > Highp(z + 1, x) Then
'        ElseIf Highp(z + 1, x) = "#N/A N.A." Then
'        Else
'        max = Highp(z + 1, x)
'        End If
'        Next
'        ActiveCell.Offset(x, 0) = max
'    Next
'
'    Range("F20").Select
'    For x = 0 To (taille)
'        i = 5
'        Do Until Weekday(Dates(i, x)) = 2
'        i = i + 1
'        Loop
'        Range("A18") = Dates(i, x)
'        j = 5
'        Do Until Weekday(Dates(j, x)) = 6
'        j = j - 1
'        Loop
'        Range("A19") = Dates(j, x)
'        Do Until Lastp(i, x) <> "#N/A N.A."
'        i = i + 1
'        Loop
'        ActiveCell.Offset(x, 0) = Lastp(j, x)
'    Next
    
Application.ScreenUpdating = True
'MsgBox "Bloomberg Data Downloded"
Call Macro1
Call histo

End Sub


Sub histo()   ' base données historiques
Dim i, last

Worksheets("histo").Activate
Range(Cells(1, 2), Cells(1, 9)).Copy
For i = 2 To 25

If Cells(i, 2) = "" Then
 last = i
 Exit For
End If
Next i
last = i
Range(Cells(1, 2), Cells(1, 9)).Copy
Range(Cells(last, 2), Cells(last, 9)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False





End Sub
Sub Macro3()
'
' Macro3 Macro
' Macro recorded 29.11.2005 by X01221106
'

'
Application.ScreenUpdating = False

 ' Initialize the Bloomberg data control
    Set objBloomberg = New BlpData

    Dim stocks
    Dim taille
    Dim Fields
    Dim periode
    Dim obs
    periode = 80
    Sheets("Sheet1").Select
    Range("K3").Select
    Selection.End(xlToRight).Select
    taille = ActiveCell.column - 11
    ReDim stocks(taille)
    ReDim obs(taille)
    ReDim Fields(taille)
    Range("K1").Select
    For i = 0 To taille
    Fields(i) = ActiveCell.Offset(0, i).Text
    Next
    Range("K2").Select
    For i = 0 To taille
    obs(i) = ActiveCell.Offset(0, i).Text
    Next
    Range("K3").Select
    For i = 0 To taille
    stocks(i) = ActiveCell.Offset(0, i).Text
    Next
    

    ' Make the synchronous call for the historical price/volume
    ' of Microsoft and Dell over the last 30 days.
    objBloomberg.Periodicity = bbDaily
    vtResults = objBloomberg.BLPGetHistoricalData(stocks, Fields, Date - 80, Date - 1)

    x = LBound(vtResults, 2) - UBound(vtResults, 2)
    Dim Dates
    ReDim Dates(periode, taille)
    Dim Prix
    ReDim Prix(periode, taille)
    
    For nDate = LBound(vtResults, 1) To UBound(vtResults, 1)
        For i = 0 To taille
        Dates(nDate, i) = vtResults(UBound(vtResults, 1) - nDate, i, 0)
        Next i
    Next nDate

    For nDate = LBound(vtResults, 1) To UBound(vtResults, 1)
        For i = 0 To taille
            Prix(nDate, i) = vtResults(UBound(vtResults, 1) - nDate, i, i + 1)
        Next i
    Next nDate
   
   
    'Ecriture Moyenne mobile 14j
    
    Dim moyenne
    Dim j
    Sheets("Sheet1").Select
    Range("K4").Select
    For x = 0 To (taille)
        j = CInt(obs(x))
        i = 0
        moyenne = 0
        Do Until i = j
        If Prix(i, x) = "#N/A N.A." Then
        j = j + 1
        i = i + 1
        Else
        moyenne = moyenne + Prix(i, x)
        i = i + 1
        End If
        Loop
        moyenne = moyenne / CInt(obs(x))
        ActiveCell.Offset(0, x) = moyenne
    Next

    
Application.ScreenUpdating = True
'MsgBox "Bloomberg Data Downloded"
 Call Macro1
End Sub

