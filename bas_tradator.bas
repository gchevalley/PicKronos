Attribute VB_Name = "bas_Tradator"
Public Const tradator_wrbk As String = "Tradator.xls"

Public Const tradator_sheet_parameters As String = "PARAMETERS"
    Public Const c_tradator_parameters_mention As Integer = 1

Public Const tradator_sheet As String = "portfolio live"
Public Const tradator_archive_sheet As String = "archives"

Public Const l_tradator_header As Integer = 9


Public Const c_tradator_idea_id As Integer = 1
Public Const c_tradator_source As Integer = 2
Public Const c_tradator_ticker As Integer = 3
Public Const c_tradator_asset As Integer = 4 'deprecie
Public Const c_tradator_idea_date As Integer = 5
Public Const c_tradator_idea_time As Integer = 6
Public Const c_tradator_side As Integer = 7
Public Const c_tradator_qty_exec As Integer = 8
Public Const c_tradator_nominal_base As Integer = 9
Public Const c_tradator_pct_nav As Integer = 10
Public Const c_tradator_trigger As Integer = 11
Public Const c_tradator_last_price As Integer = 12
Public Const c_tradator_theo_stop As Integer = 13
Public Const c_tradator_theo_target As Integer = 14
Public Const c_tradator_pnl_base As Integer = 15
Public Const c_tradator_pct_potential_profit As Integer = 16
Public Const c_tradator_pct_potential_loss As Integer = 17
Public Const c_tradator_stars As Integer = 18 ' ????
Public Const c_tradator_rrr As Integer = 19
Public Const c_tradator_risk As Integer = 20 ' ???
Public Const c_tradator_room As Integer = 21
Public Const c_tradator_nav_target As Integer = 22
Public Const c_tradator_currency As Integer = 23 'deprecier
Public Const c_tradator_change_rate As Integer = 24 'deprecie
Public Const c_tradator_contract_size As Integer = 25 'deprecie
Public Const c_tradator_option_strike As Integer = 26 'deprecie
Public Const c_tradator_central_rank_eps As Integer = 27
'Public Const c_tradator_pct_return As Integer = 28
'Public Const c_tradator_avg_pnl As Integer = 29
Public Const c_tradator_nav_pnl As Integer = 28
Public Const c_tradator_avg_pnl As Integer = 29



Public Function tradator_get_vec_mention_to_track() As Variant

Application.Calculation = xlCalculationManual

tradator_get_vec_mention_to_track = Empty

Dim i As Integer, j As Integer

Dim tmp_vec() As Variant
k = 0
For i = 2 To 1500
    If Workbooks(tradator_wrbk).Worksheets(tradator_sheet_parameters).Cells(i, c_tradator_parameters_mention) = "" Then
        Exit For
    Else
        ReDim Preserve tmp_vec(k)
        tmp_vec(k) = "@" & Replace(UCase(Workbooks(tradator_wrbk).Worksheets(tradator_sheet_parameters).Cells(i, c_tradator_parameters_mention).Value), "@", "")
        k = k + 1
    End If
Next i

If k > 0 Then
    tradator_get_vec_mention_to_track = tmp_vec
Else
    'deprecie
    tradator_get_vec_mention_to_track = Array("@DT", "@EYE", "@CORTO", "@MS", "@WOW", "@GS", "@HUET", "@ATLAS", "@JEF", "@AM", "@ONEIL", "@BRAZIL", "@JPM")
End If

End Function


Public Function tradator_get_last_line() As Integer

Dim i As Integer

For i = l_tradator_header + 1 To 12000
    If Worksheets(tradator_sheet).Cells(i, c_tradator_idea_id) = "" Then
        tradator_get_last_line = i - 1
        
        Exit For
    End If
Next i

End Function


Public Function tradator_get_tweet_id_already_in_sheet() As Variant

Application.Calculation = xlCalculationManual

Dim i As Integer, j As Integer, k As Integer


tradator_get_tweet_id_already_in_sheet = Empty

Dim tmp_vec_tweet_id_already_in_sheet() As Variant
k = 0
For i = l_tradator_header + 1 To 12000
    If Worksheets(tradator_sheet).Cells(i, c_tradator_idea_id) = "" Then
        Exit For
    Else
        ReDim Preserve tmp_vec_tweet_id_already_in_sheet(k)
        tmp_vec_tweet_id_already_in_sheet(k) = Worksheets(tradator_sheet).Cells(i, c_tradator_idea_id)
        k = k + 1
    End If
Next i


For i = l_tradator_header + 1 To 12000
    If Worksheets(tradator_archive_sheet).Cells(i, c_tradator_idea_id) = "" Then
        Exit For
    Else
        ReDim Preserve tmp_vec_tweet_id_already_in_sheet(k)
        tmp_vec_tweet_id_already_in_sheet(k) = Worksheets(tradator_archive_sheet).Cells(i, c_tradator_idea_id)
        k = k + 1
    End If
Next i



If k > 0 Then
    tradator_get_tweet_id_already_in_sheet = tmp_vec_tweet_id_already_in_sheet
End If


End Function


Private Function tradator_get_pivots_from_vec_tickers(ByVal vec_tickers As Variant) As Scripting.Dictionary

Dim oBBG As New cls_Bloomberg_Sync



End Function



Private Sub test_ticker_option()

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
oReg.Global = True



End Sub


Private Sub test_find_beginning_auto()

Dim sql_query As String
sql_query = "SELECT " & f_tweet_id & ", " & f_tweet_text & " FROM " & t_tweet & " WHERE " & f_tweet_text & " LIKE ""%$APA.US%"""
Dim extract_tweet As Variant
extract_tweet = sqlite3_query(twitter_get_db_path, sql_query)

End Sub


Private Function tradator_get_tweet_datas_if_fullfill_requirements(ByVal id As Long, ByVal tweet As String, ByVal datetime As Date, Optional ByVal vec_ticker As Variant, Optional ByVal vec_hashtag As Variant, Optional ByVal vec_mention As Variant) As Object

Set tradator_get_tweet_datas_if_fullfill_requirements = Nothing

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
oReg.Global = True

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
        array_sell_hashtags = Array("#S", "#SELL", "#SHORT", "#SS", "#SHORTSELL")
find_ticker = False
find_side = False
find_stop = False
    Dim array_stop_hashtags() As Variant
        array_stop_hashtags = Array("#STP", "#STOP")
find_tgt = False
    Dim array_target_hashtags() As Variant
        array_target_hashtags = Array("#TGT", "#TARGET")
find_room = False
    Dim array_room_hashtags() As Variant
        array_room_hashtags = Array("#ROOM")
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
            
            
            'room
            For j = 0 To UBound(array_room_hashtags, 1)
                If vec_hashtag(i) = array_room_hashtags(j) Then
                    
                    'regexp pour checker si bien suivi d un prix
                    oReg.Pattern = array_room_hashtags(j) & "\s+[^\s]+"
                    Set matches = oReg.Execute(tweet)
                    
                    For Each match In matches
                        find_room = Replace(match.Value, array_room_hashtags(j) & " ", "")
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

Dim dico_ticker As New Scripting.Dictionary

Dim find_datas As Boolean
find_datas = False


Dim tmp_pattern_stock As String, tmp_pattern_option As String
    tmp_pattern_stock = "[A-Za-z0-9]+\s[A-Za-z]{2}\sEQUITY"
    tmp_pattern_option = "[A-Za-z0-9]+\s[A-Za-z]{2}\s([\d]{1,2}/|)[\d]{1,2}(/[\d]{1,2}|)\s(c|C|p|P)[\d]+(\.[\d]+|)\sEQUITY"



Set matches = oReg.Execute(UCase("nesn sw 9/12 c52.3 equity"))

For Each match In matches
    Debug.Print match.Value
Next



If find_ticker <> False And find_side <> False And find_mention <> False Then
    
    dico_ticker.Add "id", id
    dico_ticker.Add "tweet", tweet
    dico_ticker.Add "datetime", datetime
    dico_ticker.Add "src", vec_mention(0)
    
    dico_ticker.Add "ticker", find_ticker
    dico_ticker.Add "side", find_side
    
    'asset equity / option
    oReg.Pattern = tmp_pattern_stock
    Set matches = oReg.Execute(find_ticker)
    
    For Each match In matches
        dico_ticker.Add "asset", "stock"
        Exit For
    Next
    
    
    oReg.Pattern = tmp_pattern_option
    Set matches = oReg.Execute(find_ticker)
    
    For Each match In matches
        dico_ticker.Add "asset", "option"
        Exit For
    Next
    
    
    dico_ticker.Add "vec_hashtag", vec_hashtag
    dico_ticker.Add "vec_ticker", vec_ticker
    dico_ticker.Add "vec_mention", vec_mention
    
    'vec_simple_trade ? /stop / target etc.
    
    
    find_datas = True
    
    If find_stop <> False Then
        dico_ticker.Add "stop", find_stop
    End If
    
    If find_tgt <> False Then
        dico_ticker.Add "target", find_tgt
    End If
    
    If find_room <> False Then
        dico_ticker.Add "room", find_room
    End If
    
End If

If find_datas = True Then
    Set tradator_get_tweet_datas_if_fullfill_requirements = dico_ticker
End If


End Function


Private Sub test_tradator_get_all_tweet()


Dim debug_test As Variant

'debug_test = Worksheets("portfolio live").Cells(20, 14).Font.ColorIndex

debug_test = tradator_get_all_tweet(tradator_get_vec_mention_to_track())

End Sub


'structure id / tweet / datetime / vec_tickers / vec_hashtags / vec_mentions
Public Function tradator_get_all_tweet(ByVal vec_mention As Variant) As Variant

Application.Calculation = xlCalculationManual

Workbooks(tradator_wrbk).Worksheets(tradator_sheet).Activate

Dim sql_query As String
Dim exec_query As Variant

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim oBBG As New cls_Bloomberg_Sync

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
oReg.Global = True


sql_query = "DELETE FROM " & t_twitter_helper
exec_query = sqlite3_query(twitter_get_db_path, sql_query)

Dim id_already_in_sheet As Variant
id_already_in_sheet = tradator_get_tweet_id_already_in_sheet

If IsEmpty(id_already_in_sheet) = False Then
    
    Dim db_id_already_in_sheet() As Variant
    For i = 0 To UBound(id_already_in_sheet, 1)
        ReDim Preserve db_id_already_in_sheet(i)
        db_id_already_in_sheet(i) = Array(id_already_in_sheet(i))
    Next i
    
    insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_twitter_helper, db_id_already_in_sheet, Array(f_twitter_helper_numeric1))
    'debug_test = sqlite3_query(twitter_get_db_path, "SELECT " & f_twitter_helper_numeric1 & " FROM " & t_twitter_helper)
    
End If

Dim extract_new_id As Variant
'pas uniquement ceux qui remplisse les conditions mais en tout cas pas ceux qui ne les remplissent pas
sql_query = "SELECT " & f_tweet_id & " FROM " & t_tweet & " WHERE " & f_tweet_id & " NOT IN (SELECT " & f_twitter_helper_numeric1 & " FROM " & t_twitter_helper & ")"
extract_new_id = sqlite3_query(twitter_get_db_path, sql_query)



'remonte vec currency
Dim vec_currency() As Variant
dim_currency_txt = 0
dim_currency_code = 1
dim_currency_line = 2
dim_currency_rate = 3

k = 0
For i = 14 To 32
    If Workbooks("Kronos.xls").Worksheets("Parametres").Cells(i, 1) = "" Then
        Exit For
    Else
        ReDim Preserve vec_currency(k)
        vec_currency(k) = Array(Workbooks("Kronos.xls").Worksheets("Parametres").Cells(i, 1).Value, Workbooks("Kronos.xls").Worksheets("Parametres").Cells(i, 5).Value, i, Workbooks("Kronos.xls").Worksheets("Parametres").Cells(i, 6).Value)
        k = k + 1
    End If
Next i
        



Dim data_bbg As Variant




Dim extract_tweets As Variant

k = 0
Dim final_tweet_to_add As New Collection

Dim last_line_tradator As Integer
last_line_tradator = tradator_get_last_line()

Dim tmp_check_fullfill_requirements As Scripting.Dictionary

For i = 0 To UBound(vec_mention, 1)
    
    dim_tweet_id = 0
    dim_tweet_tweet = 1
    dim_tweet_datetime = 2
    dim_tweet_tickers = 3
    dim_tweet_hashtags = 4
    dim_tweet_mentions = 5
    
    
    extract_tweets = get_specific_tweet_content(Array(f_tweet_id, f_tweet_text, f_tweet_datetime, f_tweet_json_tickers, f_tweet_json_hashtags, f_tweet_json_mentions), Array(vec_mention(i)))
    
    'passe en revue les tweets trouves pour s assurer
    If IsEmpty(extract_tweets) Then
        
    Else
        
        Dim tmp_mani_tweet_vec() As Variant
        
        Dim tmp_tweet_id As Long
        Dim tmp_tweet As String
        Dim tmp_tweet_date As Date
        Dim tmp_vec_hashtags() As Variant
        Dim tmp_vec_mentions() As Variant
        Dim tmp_vec_tickers() As Variant
        
        
        For j = 0 To UBound(extract_tweets(dim_tweet_id), 1)
            
            For m = 1 To UBound(extract_new_id, 1)
                'check si nouvel id
                If extract_tweets(dim_tweet_id)(j)(0) = extract_new_id(m)(0) Then
                    
                
                    'check si filled requirements (mention / ticker / side
                    If IsEmpty(extract_tweets(dim_tweet_hashtags)(j)) = False And IsEmpty(extract_tweets(dim_tweet_tickers)(j)) = False And IsEmpty(extract_tweets(dim_tweet_mentions)(j)) = False Then
                        
                        tmp_tweet_id = extract_tweets(dim_tweet_id)(j)(0)
                        tmp_tweet = extract_tweets(dim_tweet_tweet)(j)(0)
                        tmp_tweet_date = FromJulianDay(CDbl(extract_tweets(dim_tweet_datetime)(j)(0)))
                        
                        For n = 0 To UBound(extract_tweets(dim_tweet_hashtags)(j), 1)
                            ReDim Preserve tmp_vec_hashtags(n)
                            tmp_vec_hashtags(n) = extract_tweets(dim_tweet_hashtags)(j)(n)
                        Next n
                        
                        
                        For n = 0 To 0 'UBound(extract_tweets(dim_tweet_mentions)(j), 1)
                            ReDim Preserve tmp_vec_mentions(n)
                            tmp_vec_mentions(n) = extract_tweets(dim_tweet_mentions)(j)(n)
                        Next n
                        
                        
                        For n = 0 To 0 'UBound(extract_tweets(dim_tweet_mentions)(j), 1)
                            ReDim Preserve tmp_vec_tickers(n)
                            tmp_vec_tickers(n) = extract_tweets(dim_tweet_tickers)(j)(n)
                        Next n
                        
                        
                        Set tmp_check_fullfill_requirements = tradator_get_tweet_datas_if_fullfill_requirements(tmp_tweet_id, tmp_tweet, tmp_tweet_date, tmp_vec_tickers, tmp_vec_hashtags, tmp_vec_mentions)
                        
                        If tmp_check_fullfill_requirements Is Nothing Or tmp_tweet_id <= 168 Then
                        Else
                            
                            Dim bbg_fields() As Variant
                            bbg_fields = Array("CRNCY", "OPT_CONT_SIZE_REAL", "OPT_STRIKE_PX")
                                
                                For n = 0 To UBound(bbg_fields, 1)
                                    If UCase(bbg_fields(n)) = UCase("CRNCY") Then
                                        dim_bbg_CRNCY = n
                                    ElseIf UCase(bbg_fields(n)) = UCase("OPT_CONT_SIZE_REAL") Then
                                        dim_bbg_OPT_CONT_SIZE_REAL = n
                                    ElseIf UCase(bbg_fields(n)) = UCase("OPT_STRIKE_PX") Then
                                        dim_bbg_OPT_STRIKE_PX = n
                                    End If
                                Next n
                                
                            
                            data_bbg = oBBG.bdp(Array(UCase(tmp_check_fullfill_requirements.Item("ticker"))), bbg_fields, output_format.of_vec_without_header)
                            
                            
                            'peut etre insere
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_idea_id) = tmp_check_fullfill_requirements.Item("id")
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_source) = Replace(UCase(tmp_check_fullfill_requirements.Item("src")), "@", "")
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_ticker) = UCase(tmp_check_fullfill_requirements.Item("ticker"))
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_asset) = tmp_check_fullfill_requirements.Item("asset")
                            
                            year_int = year(tmp_check_fullfill_requirements.Item("datetime"))
                            month_int = Month(tmp_check_fullfill_requirements.Item("datetime"))
                            day_int = day(tmp_check_fullfill_requirements.Item("datetime"))
                            hour_int = Hour(tmp_check_fullfill_requirements.Item("datetime"))
                            minute_int = Minute(tmp_check_fullfill_requirements.Item("datetime"))
                            second_int = Second(tmp_check_fullfill_requirements.Item("datetime"))
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_idea_date) = DateSerial(year_int, month_int, day_int)
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_idea_date).NumberFormat = "d-mmm"
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_idea_time) = TimeSerial(hour_int, minute_int, second_int)
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_idea_time).NumberFormat = "h:mm"
                            
                            tmp_formula_pct_profit = "IF(AND(" & xlColumnValue(c_tradator_theo_target) & last_line_tradator + 1 & "<>"""";" & xlColumnValue(c_tradator_trigger) & last_line_tradator + 1 & "<>"""");(" & xlColumnValue(c_tradator_theo_target) & last_line_tradator + 1 & "/" & xlColumnValue(c_tradator_trigger) & last_line_tradator + 1 & "-1);"""")"
                            tmp_formula_pct_loss = "IF(AND(" & xlColumnValue(c_tradator_theo_stop) & last_line_tradator + 1 & "<>"""";" & xlColumnValue(c_tradator_trigger) & last_line_tradator + 1 & "<>"""");(" & xlColumnValue(c_tradator_trigger) & last_line_tradator + 1 & "/" & xlColumnValue(c_tradator_theo_stop) & last_line_tradator + 1 & "-1);"""")"
                            
                            If tmp_check_fullfill_requirements.Item("side") = "B" Then
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_side) = "long"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_potential_profit).FormulaLocal = "=" & tmp_formula_pct_profit
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_potential_loss).FormulaLocal = "=" & tmp_formula_pct_loss
                                
                                
                            ElseIf tmp_check_fullfill_requirements.Item("side") = "S" Then
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_side) = "SHORT"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_potential_profit).FormulaLocal = "=-" & tmp_formula_pct_profit
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_potential_loss).FormulaLocal = "=" & tmp_formula_pct_loss
                                
                                
                            End If
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_potential_profit).NumberFormat = "0.00%"
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_potential_loss).NumberFormat = "0.00%"
                            
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_qty_exec) = 0
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_qty_exec).NumberFormat = "#,##0"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nominal_base).NumberFormat = "#,##0"
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_currency) = UCase(data_bbg(0)(dim_bbg_CRNCY))
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_currency).Font.ColorIndex = 11
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_currency).Font.Bold = True
                            
                            'rate de kronos
                            For n = 0 To UBound(vec_currency, 1)
                                If UCase(vec_currency(n)(dim_currency_txt)) = UCase(data_bbg(0)(dim_bbg_CRNCY)) Then
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_change_rate).FormulaLocal = "=[Kronos.xls]Parametres!$F$" & vec_currency(n)(dim_currency_line)
                                    Exit For
                                End If
                            Next n
                            
                            
                            If UCase(tmp_check_fullfill_requirements.Item("asset")) = "STOCK" Then
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_contract_size) = 1
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_last_price).FormulaLocal = "=BDP(" & xlColumnValue(c_tradator_ticker) & last_line_tradator + 1 & ";""LAST_PRICE"")"
                                'Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nominal_base).FormulaLocal = "=-" & xlColumnValue(c_tradator_qty_exec) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_last_price) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_change_rate) & last_line_tradator + 1
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nominal_base).FormulaLocal = "=" & xlColumnValue(c_tradator_qty_exec) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_last_price) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_change_rate) & last_line_tradator + 1
                                
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_central_rank_eps).FormulaLocal = "=CENTRAL(" & xlColumnValue(c_tradator_ticker) & last_line_tradator + 1 & ";""rank_eps"")"
                                
                            ElseIf UCase(tmp_check_fullfill_requirements.Item("asset")) = "OPTION" Then
                                
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_last_price).FormulaLocal = "=BDP(" & xlColumnValue(c_tradator_ticker) & last_line_tradator + 1 & ";""PX_MID"")"
                                
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_contract_size) = data_bbg(0)(dim_bbg_OPT_CONT_SIZE_REAL)
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_option_strike) = data_bbg(0)(dim_bbg_OPT_STRIKE_PX)
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_option_strike).Font.ColorIndex = 11
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_option_strike).Font.Bold = True
                                    
                                
                                
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nominal_base).FormulaLocal = "=-" & xlColumnValue(c_tradator_qty_exec) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_option_strike) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_contract_size) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_change_rate) & last_line_tradator + 1
                                
                                
                                'extraction underlying ticker pour note central
                                oReg.Pattern = "[A-Za-z0-9]+\s[A-Za-z]{2}\s"
                                Set matches = oReg.Execute(UCase(tmp_check_fullfill_requirements.Item("ticker")))
                                For Each match In matches
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_central_rank_eps).FormulaLocal = "=CENTRAL(""" & match.Value & "EQUITY"";""rank_eps"")"
                                Next
                                
                                
                            End If
                            
                                'format
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_last_price).NumberFormat = "#,##0.00"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_last_price).Font.ColorIndex = 11
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_last_price).Font.Bold = True
                            
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_nav).FormulaLocal = "=" & xlColumnValue(c_tradator_nominal_base) & last_line_tradator + 1 & "/$I$2"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pct_nav).NumberFormat = "0.00%"
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_trigger) = ""
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_trigger).NumberFormat = "#,##0.00"
                            
                            
                            If tmp_check_fullfill_requirements.Exists("stop") Then
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop) = tmp_check_fullfill_requirements.Item("stop")
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop).NumberFormat = "#,##0.00"
                                
                                    'format
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop).Font.ColorIndex = 3
                                
                                    'cond format
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop).FormatConditions.Delete
                                    If tmp_check_fullfill_requirements.Item("side") = "B" Then
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & xlColumnValue(c_tradator_last_price) & "$" & last_line_tradator + 1
                                    ElseIf tmp_check_fullfill_requirements.Item("side") = "S" Then
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & xlColumnValue(c_tradator_last_price) & "$" & last_line_tradator + 1
                                    End If
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop).FormatConditions(1).Interior.ColorIndex = 3
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_stop).FormatConditions(1).Font.ColorIndex = 2
                            End If
                            
                            If tmp_check_fullfill_requirements.Exists("target") Then
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_target) = tmp_check_fullfill_requirements.Item("target")
                                
                                'format
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_target).Font.ColorIndex = 12
                                
                                'cond format
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_target).FormatConditions.Delete
                                    If tmp_check_fullfill_requirements.Item("side") = "B" Then
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_target).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:="=$" & xlColumnValue(c_tradator_last_price) & "$" & last_line_tradator + 1
                                    ElseIf tmp_check_fullfill_requirements.Item("side") = "S" Then
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_target).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:="=$" & xlColumnValue(c_tradator_last_price) & "$" & last_line_tradator + 1
                                    End If
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_target).FormatConditions(1).Interior.ColorIndex = 12
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_theo_target).FormatConditions(1).Font.ColorIndex = 2
                                
                            End If
                            
                            If tmp_check_fullfill_requirements.Exists("room") Then
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_room) = tmp_check_fullfill_requirements.Item("room")
                            End If
                            
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).FormulaLocal = "=" & xlColumnValue(c_tradator_qty_exec) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_contract_size) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_change_rate) & last_line_tradator + 1 & "*(" & xlColumnValue(c_tradator_last_price) & last_line_tradator + 1 & "-" & xlColumnValue(c_tradator_trigger) & last_line_tradator + 1 & ")"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).NumberFormat = "#,##0"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).Font.Bold = True
                                
                                
                                'cond format
                                Dim limit_pnl_color As Double
                                    limit_pnl_color = 500
                                
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).FormatConditions.Delete
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=limit_pnl_color
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).FormatConditions(1).Interior.ColorIndex = 6
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=-limit_pnl_color
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_pnl_base).FormatConditions(2).Interior.ColorIndex = 3
                            
                            
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nav_target).FormulaLocal = "=" & xlColumnValue(c_tradator_pct_nav) & last_line_tradator + 1 & "*" & xlColumnValue(c_tradator_pct_potential_profit) & last_line_tradator + 1
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nav_target).NumberFormat = "0.0000%"
                            
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nav_pnl).FormulaLocal = "=" & xlColumnValue(c_tradator_pnl_base) & last_line_tradator + 1 & "/$I$2*100"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_nav_pnl).NumberFormat = "#,##0.0%"
                            
                            
                            Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_avg_pnl).FormulaLocal = "=IF(" & xlColumnValue(c_tradator_nominal_base) & last_line_tradator + 1 & "<>0;" & xlColumnValue(c_tradator_pnl_base) & last_line_tradator + 1 & "/ABS(" & xlColumnValue(c_tradator_nominal_base) & last_line_tradator + 1 & ");"""")"
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_avg_pnl).NumberFormat = "#,##0.00%"
                                
                                'cond format
                                Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_avg_pnl).FormatConditions.Delete
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_avg_pnl).FormatConditions.Add type:=xlCellValue, Operator:=xlGreater, Formula1:=0
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_avg_pnl).FormatConditions(1).Interior.ColorIndex = 6
                                    Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_avg_pnl).FormatConditions.Add type:=xlCellValue, Operator:=xlLess, Formula1:=0
                                        Worksheets(tradator_sheet).Cells(last_line_tradator + 1, c_tradator_avg_pnl).FormatConditions(2).Interior.ColorIndex = 3
                            
                            
                            
                            
                            
                            
                            
                            
                            last_line_tradator = last_line_tradator + 1
                            
                        End If
                        
                    End If
                End If
            Next m
            
        Next j
    End If
    
    
Next i

Application.Calculation = xlCalculationAutomatic

End Function



Public Sub tradator_insert_qty_price_from_form()


Dim tmp_order_line As String

Dim tmp_side As String, tmp_qty As Double, tmp_symbol As String, tmp_price As Double
'If frm_Tradator_choose_qty_price.CB_side_symbol_qty_price.Value <> "" Then
    
    tmp_order_line = frm_Tradator_choose_qty_price.CB_side_symbol_qty_price.Value
    tmp_order_line = "B 10 GOOG 742.23"
    
    tmp_side = Left(Left(tmp_order_line, InStr(tmp_order_line, " ") - 1), 1)
    space_side = InStr(tmp_order_line, " ")
    
    space_qty = InStr(space_side + 1, tmp_order_line, " ")
    tmp_qty = CDbl(Mid(tmp_order_line, space_side + 1, space_qty - 1 - space_side))
    
    space_ticker = InStr(space_qty + 1, tmp_order_line, " ")
    tmp_symbol = Mid(tmp_order_line, space_qty + 1, space_ticker - 1 - space_qty)
    
    
    space_price = Len(tmp_order_line)
    tmp_price = CDbl(Mid(tmp_order_line, space_ticker + 1, space_price - space_ticker))
    
'End If

'frm_Tradator_choose_qty_price.Hide

End Sub


Public Sub tradator_mount_form_chose_order_to_complete_qty_price()

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim tmp_ticker As String

If ActiveWorkbook.name = "Tradator.xls" Then
    If ActiveSheet.name = "portfolio live" Then
        If ActiveCell.row > 9 Then
            
            'remonte le symbol
            tmp_ticker = ActiveWorkbook.ActiveSheet.Cells(ActiveCell.row, c_tradator_ticker).Value
            
            tmp_ticker = "GOOG Us equity"
            
            Dim redi_order_from_ticker As Variant
            redi_order_from_ticker = tradator_get_combi_order_qty_price(tmp_ticker)
            
            If IsEmpty(redi_order_from_ticker) Then
            Else
                
                'construit les combi et affiche form
                Dim vec_combi_side_symbol_qty_price() As Variant
                k = 0
                For i = 0 To UBound(redi_order_from_ticker, 1)
                    If i = 0 Then
                        ReDim Preserve vec_combi_side_symbol_qty_price(k)
                        vec_combi_side_symbol_qty_price(k) = Array(redi_order_from_ticker(i)(5), redi_order_from_ticker(i)(6), redi_order_from_ticker(i)(7), redi_order_from_ticker(i)(8))
                        k = k + 1
                    Else
                        
                        'check si new
                        For j = 0 To UBound(vec_combi_side_symbol_qty_price, 1)
                            
                            If redi_order_from_ticker(i)(5) = vec_combi_side_symbol_qty_price(j)(0) And redi_order_from_ticker(i)(6) = vec_combi_side_symbol_qty_price(j)(1) And redi_order_from_ticker(i)(7) = vec_combi_side_symbol_qty_price(j)(2) And redi_order_from_ticker(i)(8) = vec_combi_side_symbol_qty_price(j)(3) Then
                                Exit For
                            Else
                                If j = UBound(vec_combi_side_symbol_qty_price, 1) Then
                                    ReDim Preserve vec_combi_side_symbol_qty_price(k)
                                    vec_combi_side_symbol_qty_price(k) = Array(redi_order_from_ticker(i)(5), redi_order_from_ticker(i)(6), redi_order_from_ticker(i)(7), redi_order_from_ticker(i)(8))
                                    k = k + 1
                                End If
                            End If
                            
                        Next j
                        
                    End If
                Next i
                
                If k > 0 Then
                    frm_Tradator_choose_qty_price.CB_side_symbol_qty_price.Clear
                    For i = 0 To UBound(vec_combi_side_symbol_qty_price, 1)
                        frm_Tradator_choose_qty_price.CB_side_symbol_qty_price.AddItem vec_combi_side_symbol_qty_price(i)(0) & " " & vec_combi_side_symbol_qty_price(i)(2) & " " & vec_combi_side_symbol_qty_price(i)(1) & " " & vec_combi_side_symbol_qty_price(i)(3)
                    Next i
                    
                    frm_Tradator_choose_qty_price.Show
                    
                End If
                
            End If
            
        End If
    End If
End If

End Sub


Public Function tradator_get_combi_order_qty_price(ByVal ticker As String) As Variant

Dim redi_symbol As String
redi_symbol = Replace(UCase(ticker), " EQUITY", "")
redi_symbol = Replace(redi_symbol, " ", ".")
redi_symbol = Replace(redi_symbol, ".US", "")



tradator_get_combi_order_qty_price = Empty

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
        Exit Function
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
            
            
            '00 - RefNum
            '01 - OrderRefKey
            '02 - Desc
            '03 - BranchSequence
            '04 - datetime
            '05 - SideAbrev
            '06 - symbol
            '07 - OrderQty
            '08 - OrderPrice
            '09 - ExecQty
            '10 - ExecPrice
            '11 - PriceType
            '12 - Status
            '13 - UserID
            
            vec_redi_exec(k) = Array(CStr(tmp_RefNum), tmp_OrderRefKey, tmp_Desc, tmp_BranchSequence, tmp_datetime, tmp_SideAbrev, _
                tmp_symbol, tmp_OrderQty, tmp_OrderPrice, tmp_ExecQty, tmp_ExecPrice, tmp_PriceType, tmp_Status, tmp_UserID)
            k = k + 1
            
        End If
        
    Next i
    
    
    If k > 0 Then
        
        k = 0
        Dim vec_combi() As Variant
        For i = 0 To UBound(vec_redi_exec, 1)
            If vec_redi_exec(i)(6) = redi_symbol Then
                ReDim Preserve vec_combi(k)
                vec_combi(k) = vec_redi_exec(i)
                k = k + 1
            End If
        Next i
        
        If k > 0 Then
            tradator_get_combi_order_qty_price = vec_combi
        End If
        
    End If

End If

End Function


