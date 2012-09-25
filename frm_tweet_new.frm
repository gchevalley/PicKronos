VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Tweet_new 
   Caption         =   "Tweet"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14550
   OleObjectBlob   =   "frm_Tweet_new.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Tweet_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












Private Const advanced_form_height As Integer = 615
Private Const small_form_height As Integer = 190


Private Sub clean_form_content()

TB_tweet.Value = ""
LB_attach.Clear
LB_Helpers.Clear

End Sub


Private Sub load_last_tweets_in_form()

Dim date_tmp As Date

Dim list_last_tweets As Variant
list_last_tweets = get_last_tweets(10)
    

If IsEmpty(list_last_tweets) = False Then
    
    'detection des dim
    For i = 0 To UBound(list_last_tweets(0), 1)
        If list_last_tweets(0)(i) = f_tweet_id Then
            dim_tweet_id = i
        ElseIf list_last_tweets(0)(i) = f_tweet_datetime Then
            dim_tweet_datetime = i
        ElseIf list_last_tweets(0)(i) = f_tweet_from Then
            dim_tweet_user = i
        ElseIf list_last_tweets(0)(i) = f_tweet_text Then
            dim_tweet_text = i
        ElseIf InStr(list_last_tweets(0)(i), "attach") <> 0 Then
            dim_tweet_attachements = i
        End If
    Next i
    
        With frm_Tweet_new.LV_last_tweet
            
            .ListItems.Clear
            
            With .ColumnHeaders
                .Clear

                .Add , , "user", 65
                .Add , , "date and time", 70
                .Add , , "tweet", 415
                .Add , , "attach", 35
            End With

            For i = 1 To UBound(list_last_tweets, 1)

                With .ListItems
                    .Add , "user_" & CStr(list_last_tweets(i)(dim_tweet_id)), list_last_tweets(i)(dim_tweet_user)  'user
                End With

                date_tmp = FromJulianDay(CDbl(list_last_tweets(i)(dim_tweet_datetime)))
                
                'formattage pour la date
                day_txt = day(date_tmp)
                    If Len(day_txt) = 1 Then
                        day_txt = "0" & day_txt
                    End If
                
                month_txt = Month(date_tmp)
                    If Len(month_txt) = 1 Then
                        month_txt = "0" & month_txt
                    End If
                    
                hour_txt = Hour(date_tmp)
                    If Len(hour_txt) = 1 Then
                        hour_txt = "0" & hour_txt
                    End If
                    
                minute_txt = Minute(date_tmp)
                    If Len(minute_txt) = 1 Then
                        minute_txt = "0" & minute_txt
                    End If
                
                
                .ListItems(i).ListSubItems.Add , "date_" & CStr(list_last_tweets(i)(dim_tweet_id)), day_txt & "." & month_txt & "." & Right(year(date_tmp), 2) & " " & hour_txt & ":" & minute_txt 'date
                '.ListItems(i).ListSubItems.Add , "tweet_" & CStr(list_last_tweets(i)(dim_tweet_id)), list_last_tweets(i)(dim_tweet_text) 'tweet
                .ListItems(i).ListSubItems.Add , "tweet_" & CStr(list_last_tweets(i)(dim_tweet_id)), show_tweet_with_infos(list_last_tweets(i)(dim_tweet_id))  'tweet
                
                If IsEmpty(list_last_tweets(i)(dim_tweet_attachements)) Then
                    .ListItems(i).ListSubItems.Add , , "0"
                Else
                    'count
                    .ListItems(i).ListSubItems.Add , "attachement_" & CStr(list_last_tweets(i)(dim_tweet_id)), CStr(UBound(list_last_tweets(i)(dim_tweet_attachements), 1) + 1)
                End If
            Next i

            '.view = lvwReport
            .FullRowSelect = True
        End With

End If

End Sub


Private Sub btn_attach_file_Click()
already_attach = False
Dim db_path As Variant

db_path = Application.GetOpenFilename(FileFilter:="All Files,*.*", Title:="Chose File(s)", MultiSelect:=False)

If db_path <> False Then
    's'assure que n'existe pas deja
    If LB_attach.ListCount = 0 Then
        LB_attach.AddItem db_path
    Else
        For Each tmp_file In LB_attach.list
            If tmp_file = db_path Then
                already_attach = True
                Exit For
            Else
            End If
        Next
        
        If already_attach = False Then
            LB_attach.AddItem db_path
        End If
    End If
End If

End Sub

Private Sub btn_filter_real_twitter_follow_links_Click()

Dim real_twitter_tweet As String

If LV_real_twitter.ListItems.count > 0 Then
    real_twitter_tweet = LV_real_twitter.SelectedItem.Text
    
    Dim list_links As Variant
    list_links = get_links_from_tweet(real_twitter_tweet)
    
    If IsEmpty(list_links) Then
        MsgBox ("no links in the selected tweet")
    Else
        For i = 0 To UBound(list_links, 1)
            ActiveWorkbook.FollowHyperlink list_links(i)(1), , True
        Next i
    End If
    
End If

End Sub

Private Sub btn_filter_real_twitter_hashtags_Click()

Dim real_twitter_tweet As String

If LV_real_twitter.ListItems.count > 0 Then
    real_twitter_tweet = LV_real_twitter.SelectedItem.Text
    
    'extraction des hashtags
    Dim list_hashtags As Variant
    list_hashtags = get_hashtags_from_tweet(real_twitter_tweet)
    
    If IsEmpty(list_hashtags) Then
        MsgBox ("No hashtag # in the selected tweet")
    Else
        If UBound(list_hashtags, 1) = 0 Then
            TB_search.Value = list_hashtags(0)(1)
            Call search_interal_db_and_twitter
        Else
            'inputbox list
            txt_inputbox = ""
            For i = 0 To UBound(list_hashtags, 1)
                txt_inputbox = txt_inputbox & "[" & i + 1 & "] " & list_hashtags(i)(1) & vbCrLf
            Next i
            
            answer = InputBox(txt_inputbox, "Which one ?")
            
            If IsNumeric(answer) And answer <= UBound(list_hashtags, 1) + 1 Then
                TB_search.Value = list_hashtags(CDbl(answer) - 1)(1)
                Call search_interal_db_and_twitter
            End If
        End If
    End If
Else
    MsgBox ("Nothing !")
End If

End Sub

Private Sub btn_filters_on_hashtag_Click()

tweet_id_txt = LV_last_tweet.SelectedItem.key
tweet_id_txt = Replace(tweet_id_txt, "user_", "")
tweet_id_long = CLng(tweet_id_txt)


Dim extract_hashtags_in_tweet As Variant
extract_hashtags_in_tweet = get_specific_tweet_content(Array(f_tweet_json_hashtags), , , , , , tweet_id_long)

If IsEmpty(extract_hashtags_in_tweet(0)(0)(0)) Then
    MsgBox ("No # in the selected tweet")
Else
    If UBound(extract_hashtags_in_tweet(0)(0), 1) > 0 Then
        'msgbox pour la section du hashtag desire
        txt_inputbox = ""
        For i = 0 To UBound(extract_hashtags_in_tweet(0)(0), 1)
            txt_inputbox = txt_inputbox & "[" & i + 1 & "] " & extract_hashtags_in_tweet(0)(0)(i) & vbCrLf
        Next i
        
        answer = InputBox(txt_inputbox, "Which one ?")
        
        If IsNumeric(answer) And answer <= UBound(extract_hashtags_in_tweet(0)(0), 1) + 1 Then
            TB_search.Value = extract_hashtags_in_tweet(0)(0)(CDbl(answer) - 1)
            Call search_interal_db_and_twitter
        End If
        
    Else
        TB_search.Value = extract_hashtags_in_tweet(0)(0)(0)
        Call search_interal_db_and_twitter
    End If
End If


End Sub

Private Sub btn_filters_on_mention_Click()

tweet_id_txt = LV_last_tweet.SelectedItem.key
tweet_id_txt = Replace(tweet_id_txt, "user_", "")
tweet_id_long = CLng(tweet_id_txt)


Dim extract_mentions_in_tweet As Variant
extract_mentions_in_tweet = get_specific_tweet_content(Array(f_tweet_json_mentions), , , , , , tweet_id_long)

If IsEmpty(extract_mentions_in_tweet(0)(0)(0)) Then
    MsgBox ("No @ in the selected tweet")
Else
    If UBound(extract_mentions_in_tweet(0)(0), 1) > 0 Then
        'msgbox pour la section du hashtag desire
        txt_inputbox = ""
        For i = 0 To UBound(extract_mentions_in_tweet(0)(0), 1)
            txt_inputbox = txt_inputbox & "[" & i + 1 & "] " & extract_mentions_in_tweet(0)(0)(i) & vbCrLf
        Next i
        
        answer = InputBox(txt_inputbox, "Which one ?")
        
        If IsNumeric(answer) And answer <= UBound(extract_mentions_in_tweet(0)(0), 1) + 1 Then
            TB_search.Value = extract_mentions_in_tweet(0)(0)(CDbl(answer) - 1)
            Call search_interal_db_and_twitter
        End If
        
    Else
        TB_search.Value = extract_mentions_in_tweet(0)(0)(0)
        Call search_interal_db_and_twitter
    End If
End If

End Sub

Private Sub btn_filters_on_ticker_Click()

tweet_id_txt = LV_last_tweet.SelectedItem.key
tweet_id_txt = Replace(tweet_id_txt, "user_", "")
tweet_id_long = CLng(tweet_id_txt)


Dim extract_tickers_in_tweet As Variant
extract_tickers_in_tweet = get_specific_tweet_content(Array(f_tweet_json_tickers), , , , , , tweet_id_long)

If IsEmpty(extract_tickers_in_tweet(0)(0)(0)) Then
    MsgBox ("No $ in the selected tweet")
Else
    If UBound(extract_tickers_in_tweet(0)(0), 1) > 0 Then
        'msgbox pour la section du hashtag desire
        txt_inputbox = ""
        For i = 0 To UBound(extract_tickers_in_tweet(0)(0), 1)
            txt_inputbox = txt_inputbox & "[" & i + 1 & "] " & extract_tickers_in_tweet(0)(0)(i) & vbCrLf
        Next i
        
        answer = InputBox(txt_inputbox, "Which one ?")
        
        If IsNumeric(answer) And answer <= UBound(extract_tickers_in_tweet(0)(0), 1) + 1 Then
            TB_search.Value = extract_tickers_in_tweet(0)(0)(CDbl(answer) - 1)
            Call search_interal_db_and_twitter
        End If
        
    Else
        TB_search.Value = extract_tickers_in_tweet(0)(0)(0)
        Call search_interal_db_and_twitter
    End If
End If

End Sub

Private Sub btn_follow_links_Click()

tweet_id_txt = LV_last_tweet.SelectedItem.key
tweet_id_txt = Replace(tweet_id_txt, "user_", "")
tweet_id_long = CLng(tweet_id_txt)


Dim extract_links_in_tweet As Variant
extract_links_in_tweet = get_specific_tweet_content(Array(f_tweet_json_links), , , , , , tweet_id_long)

If IsEmpty(extract_links_in_tweet(0)(0)(0)) Then
    MsgBox ("No hyperlinks in the selected tweet")
Else

    For i = 0 To UBound(extract_links_in_tweet(0)(0), 1)
        ActiveWorkbook.FollowHyperlink extract_links_in_tweet(0)(0)(i), , True
    Next i
End If

End Sub

Private Sub btn_insert_at_Click()

TB_tweet.Value = TB_tweet.Value & "@"
TB_tweet.SetFocus

End Sub

Private Sub btn_insert_diese_Click()

TB_tweet.Value = TB_tweet.Value & "#"
TB_tweet.SetFocus

End Sub

Private Sub btn_insert_dollar_Click()

TB_tweet.Value = TB_tweet.Value & "$"
TB_tweet.SetFocus

End Sub

Private Sub btn_new_tweet_Click()


If TB_tweet.Value <> "" Then
    test_insert = create_tweet(TB_tweet.Value, get_username_from_tweet, get_attach_file)
    
    Call clean_form_content
    
    Call load_last_tweets_in_form
End If

End Sub

Private Sub LB_Hashtags_Click()

End Sub



Private Sub search_interal_db_and_twitter()

Dim date_tmp  As Date

frm_Tweet_new.LV_real_twitter.ListItems.Clear

If frm_Tweet_new.TB_search.Value = "" Then
    
    Dim list_last_tweets As Variant
    list_last_tweets = get_last_tweets(10)
    
    
    'detection des dim
    For i = 0 To UBound(list_last_tweets(0), 1)
        If list_last_tweets(0)(i) = f_tweet_id Then
            dim_tweet_id = i
        ElseIf list_last_tweets(0)(i) = f_tweet_datetime Then
            dim_tweet_datetime = i
        ElseIf list_last_tweets(0)(i) = f_tweet_from Then
            dim_tweet_user = i
        ElseIf list_last_tweets(0)(i) = f_tweet_text Then
            dim_tweet_text = i
        ElseIf InStr(list_last_tweets(0)(i), "attach") <> 0 Then
            dim_tweet_attachements = i
        End If
    Next i
    
    
    If IsEmpty(list_last_tweets) = False Then
        With frm_Tweet_new.LV_last_tweet
            
            .ListItems.Clear
            
            With .ColumnHeaders
                .Clear

                .Add , , "user", 65
                .Add , , "date and time", 70
                .Add , , "tweet", 415
                .Add , , "attach", 35
            End With

            For i = 1 To UBound(list_last_tweets, 1)

                With .ListItems
                    .Add , "user_" & CStr(list_last_tweets(i)(dim_tweet_id)), list_last_tweets(i)(dim_tweet_user)  'user
                End With

                date_tmp = FromJulianDay(CDbl(list_last_tweets(i)(dim_tweet_datetime)))
                
                'formattage pour la date
                day_txt = day(date_tmp)
                    If Len(day_txt) = 1 Then
                        day_txt = "0" & day_txt
                    End If
                
                month_txt = Month(date_tmp)
                    If Len(month_txt) = 1 Then
                        month_txt = "0" & month_txt
                    End If
                    
                hour_txt = Hour(date_tmp)
                    If Len(hour_txt) = 1 Then
                        hour_txt = "0" & hour_txt
                    End If
                    
                minute_txt = Minute(date_tmp)
                    If Len(minute_txt) = 1 Then
                        minute_txt = "0" & minute_txt
                    End If
                
                
                .ListItems(i).ListSubItems.Add , "date_" & CStr(list_last_tweets(i)(dim_tweet_id)), day_txt & "." & month_txt & "." & Right(year(date_tmp), 2) & " " & hour_txt & ":" & minute_txt 'date
                '.ListItems(i).ListSubItems.Add , "tweet_" & CStr(list_last_tweets(i)(dim_tweet_id)), list_last_tweets(i)(dim_tweet_text) 'tweet
                .ListItems(i).ListSubItems.Add , "tweet_" & CStr(list_last_tweets(i)(dim_tweet_id)), show_tweet_with_infos(list_last_tweets(i)(dim_tweet_id)) 'tweet
                
                If IsEmpty(list_last_tweets(i)(dim_tweet_attachements)) Then
                    .ListItems(i).ListSubItems.Add , , "0"
                Else
                    'count
                    .ListItems(i).ListSubItems.Add , "attachement_" & CStr(list_last_tweets(i)(dim_tweet_id)), CStr(UBound(list_last_tweets(i)(dim_tweet_attachements), 1) + 1)
                End If
            Next i

            '.view = lvwReport
            .FullRowSelect = True
        End With
    
    End If
    
    Exit Sub
End If

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True


'eclate les mots grace a une regexp
oReg.Pattern = "[\S]+"

Set matches = oReg.Execute(frm_Tweet_new.TB_search.Value)
k = 0
Dim vec_words() As Variant
For Each match In matches
    ReDim Preserve vec_words(k)
    vec_words(k) = match.Value
    k = k + 1
Next


'envoi dans les differents array
Dim count_hashtag As Integer, count_mention As Integer, count_ticker As Integer, count_regexp As Integer
    count_hashtag = 0
    count_mention = 0
    count_ticker = 0
    count_regexp = 0

Dim tmp_vec_hashtag() As Variant, tmp_vec_mention() As Variant, tmp_vec_ticker() As Variant, tmp_vec_regexp() As Variant



Dim call_vec_hashtag As Variant, call_vec_mention As Variant, call_vec_ticker As Variant, call_vec_regexp As Variant
    call_vec_hashtag = Empty
    call_vec_mention = Empty
    call_vec_ticker = Empty
    call_vec_regexp = Empty




If k > 0 Then
    For i = 0 To UBound(vec_words, 1)
        If Left(vec_words(i), 1) = "#" Then
            ReDim Preserve tmp_vec_hashtag(count_hashtag)
            tmp_vec_hashtag(count_hashtag) = UCase(vec_words(i))
            count_hashtag = count_hashtag + 1
        ElseIf Left(vec_words(i), 1) = "@" Then
            ReDim Preserve tmp_vec_mention(count_mention)
            tmp_vec_mention(count_mention) = LCase(vec_words(i))
            count_mention = count_mention + 1
        ElseIf Left(vec_words(i), 1) = "$" Then
            ReDim Preserve tmp_vec_ticker(count_ticker)
            tmp_vec_ticker(count_ticker) = UCase(vec_words(i))
'            If InStr(vec_words(i), ".") <> 0 Then
'
'            Else
'
'            End If
            
            count_ticker = count_ticker + 1
        Else
            ReDim Preserve tmp_vec_regexp(count_regexp)
            tmp_vec_regexp(count_regexp) = Array(vec_words(i), 1) 'pattern + qty
            count_regexp = count_regexp + 1
        End If
    Next i

End If

'retourne les tweets
k = 0
Dim tmp_ticker As String

Dim query_twitter As String
    query_twitter = ""
    
If count_hashtag = 0 Then
    call_vec_hashtag = Empty
Else
    call_vec_hashtag = tmp_vec_hashtag
    For i = 0 To UBound(tmp_vec_hashtag, 1)
        If k = 0 Then
            query_twitter = tmp_vec_hashtag(i)
        Else
            query_twitter = query_twitter & " " & tmp_vec_hashtag(i)
        End If
        
        k = k + 1
    Next i
End If

If count_mention = 0 Then
    call_vec_mention = Empty
Else
    call_vec_mention = tmp_vec_mention
End If

If count_ticker = 0 Then
    call_vec_ticker = Empty
Else
    call_vec_ticker = tmp_vec_ticker
    For i = 0 To UBound(tmp_vec_ticker, 1)
        If k = 0 Then
            query_twitter = tmp_vec_ticker(i)
        Else
            query_twitter = query_twitter & " " & tmp_vec_ticker(i)
        End If
        
        k = k + 1
    Next i
End If

If count_regexp = 0 Then
    call_vec_regexp = Empty
Else
    call_vec_regexp = tmp_vec_regexp
    For i = 0 To UBound(tmp_vec_regexp, 1)
        If k = 0 Then
            query_twitter = tmp_vec_regexp(i)(0)
        Else
            query_twitter = query_twitter & " " & tmp_vec_regexp(i)(0)
        End If
        
        k = k + 1
    Next i
End If

Dim extract_search_tweets As Variant
extract_search_tweets = get_specific_tweet_content(Array(f_tweet_id, f_tweet_datetime, f_tweet_from, f_tweet_text, "attachement"), call_vec_mention, call_vec_ticker, call_vec_hashtag, call_vec_regexp)


'reorder content
Dim tmp_bridge_vec As Variant

For i = 0 To UBound(extract_search_tweets, 1)
    For j = UBound(extract_search_tweets(i), 1) To CLng(UBound(extract_search_tweets(i), 1) / 2) Step -1
        tmp_bridge_vec = extract_search_tweets(i)(UBound(extract_search_tweets(i), 1) - j)
        extract_search_tweets(i)(UBound(extract_search_tweets(i), 1) - j) = extract_search_tweets(i)(j)
        extract_search_tweets(i)(j) = tmp_bridge_vec
    Next j
Next i


dim_tweet_search_id = 0
dim_tweet_search_datetime = 1
dim_tweet_search_from = 2
dim_tweet_search_text = 3
dim_tweet_search_attachement = 4


With frm_Tweet_new.LV_last_tweet
    .ListItems.Clear
    
    With .ColumnHeaders
        .Clear

        .Add , , "user", 65
        .Add , , "date and time", 70
        .Add , , "tweet", 230
        .Add , , "attach", 35
    End With
End With


If IsEmpty(extract_search_tweets) Then
    frm_Tweet_new.LV_last_tweet.ListItems.Clear
    'Exit Sub
Else
    
    With frm_Tweet_new.LV_last_tweet

        For i = 0 To UBound(extract_search_tweets(dim_tweet_search_id), 1) 'for each results - boucle sur une des dim
            
            With .ListItems
                .Add , "user_" & CStr(extract_search_tweets(dim_tweet_search_id)(i)(0)), extract_search_tweets(dim_tweet_search_from)(i)(0) 'user
            End With
            
            date_tmp = FromJulianDay(CDbl(extract_search_tweets(dim_tweet_search_datetime)(i)(0)))
            
            
            'formattage pour la date
            day_txt = day(date_tmp)
                If Len(day_txt) = 1 Then
                    day_txt = "0" & day_txt
                End If
            
            month_txt = Month(date_tmp)
                If Len(month_txt) = 1 Then
                    month_txt = "0" & month_txt
                End If
                
            hour_txt = Hour(date_tmp)
                If Len(hour_txt) = 1 Then
                    hour_txt = "0" & hour_txt
                End If
                
            minute_txt = Minute(date_tmp)
                If Len(minute_txt) = 1 Then
                    minute_txt = "0" & minute_txt
                End If
            
            
            .ListItems(i + 1).ListSubItems.Add , "date_" & CStr(extract_search_tweets(dim_tweet_search_id)(i)(0)), day_txt & "." & month_txt & "." & Right(year(date_tmp), 2) & " " & hour_txt & ":" & minute_txt 'date
            '.ListItems(i + 1).ListSubItems.Add , "tweet_" & CStr(extract_search_tweets(dim_tweet_search_id)(i)(0)), extract_search_tweets(dim_tweet_search_text)(i)(0) 'tweet
            .ListItems(i + 1).ListSubItems.Add , "tweet_" & CStr(extract_search_tweets(dim_tweet_search_id)(i)(0)), show_tweet_with_infos(extract_search_tweets(dim_tweet_search_id)(i)(0)) 'tweet
            
            If IsEmpty(extract_search_tweets(dim_tweet_search_attachement)(i)) Then
                .ListItems(i + 1).ListSubItems.Add , , "0"
            Else
                'count
                .ListItems(i + 1).ListSubItems.Add , "attachement_" & CStr(extract_search_tweets(dim_tweet_search_id)(i)(0)), CStr(UBound(extract_search_tweets(dim_tweet_search_attachement)(i), 1) + 1)
            End If
    
            '.view = lvwReport
            .FullRowSelect = True
            
        Next i
    
    End With
End If


'appel egalement sur le vrai twitter
If query_twitter <> "" Then
    Dim extract_tweets_from_twitter As Variant
    extract_tweets_from_twitter = get_twitter_content(Array("created_at", "from_user", "text"), query_twitter)
    
    dim_live_twitter_date = 0
    dim_live_twitter_user = 1
    dim_live_twitter_text = 2
    
    If IsEmpty(extract_tweets_from_twitter) = False Then
    
        With frm_Tweet_new.LV_real_twitter
            .ListItems.Clear
        
            With .ColumnHeaders
                .Clear
        
                '.Add , , "user", 65
                '.Add , , "date and time", 70
                .Add , , "tweet", 500
            End With
            
            
            For i = 0 To UBound(extract_tweets_from_twitter, 1)
                
                's'assure qu'il n'existe pas deja un tweet similaire
                If i > 0 Then
                    
                    Dim limist_distance_levenshtein As Long
                    limit_distance_levenshtein = 10
                    
                    For j = 0 To i - 1
                        If DistanceDeLevenshtein(extract_tweets_from_twitter(i)(dim_live_twitter_text), extract_tweets_from_twitter(j)(dim_live_twitter_text), limit_distance_levenshtein) < limit_distance_levenshtein Then
                            GoTo check_next_tweet_real_twitter
                        End If
                    Next j
                    
                End If
                
                
                With .ListItems
                    '.Add , "user_" & CStr(i), "@" & extract_tweets_from_twitter(i)(dim_live_twitter_user) 'user
                    .Add , "tweet_" & CStr(i), extract_tweets_from_twitter(i)(dim_live_twitter_text) 'tweet
                End With
                
                
                
                '.ListItems(i + 1).ListSubItems.Add , "date_" & CStr(i), extract_tweets_from_twitter(i)(dim_live_twitter_date) 'date
                '.ListItems(i + 1).ListSubItems.Add , "tweet_" & CStr(i), extract_tweets_from_twitter(i)(dim_live_twitter_text) 'tweet
check_next_tweet_real_twitter:
            Next i
            
            .FullRowSelect = True
            
        End With
    
    Else
        frm_Tweet_new.LV_real_twitter.ListItems.Clear
    End If
    
    
End If

End Sub

Private Sub btn_open_attachements_Click()

Dim i As Long, j As Long, k As Long, m As Long, n As Long

tweet_id_txt = LV_last_tweet.SelectedItem.key
tweet_id_txt = Replace(tweet_id_txt, "user_", "")
tweet_id_long = CLng(tweet_id_txt)

'remonte les attachements
dim_attach_tweet_id = 0
dim_attach_source = 1
dim_attach_tinyurl = 2
dim_attach_local_copy = 3

Dim sql_query As String
sql_query = "SELECT " & f_hyperlink_and_file_tweet_id & ", " & f_hyperlink_and_file_source & ", " & f_hyerplink_and_file_tinyurl & ", " & f_hyperlink_and_file_local_copy
    sql_query = sql_query & " FROM " & t_hyperlink_and_file
    sql_query = sql_query & " WHERE " & f_hyperlink_and_file_tweet_id & "=" & tweet_id_long
Dim extract_tweet_attachements As Variant
extract_tweet_attachements = sqlite3_query(db_path_base & db_twitter, sql_query)

k = 0
If UBound(extract_tweet_attachements, 1) > 0 Then
    For i = 1 To UBound(extract_tweet_attachements, 1)
        
        If IsNull(extract_tweet_attachements(i)(dim_attach_tinyurl)) Then
            
            k = k + 1
            
            'il s'agit bien d'un fichier et non d'un hyperlinks
            If exist_file(extract_tweet_attachements(i)(dim_attach_source)) Then
                ActiveWorkbook.FollowHyperlink extract_tweet_attachements(i)(dim_attach_source), NewWindow:=True
                
            Else
                ActiveWorkbook.FollowHyperlink extract_tweet_attachements(i)(dim_attach_local_copy), NewWindow:=True
            End If
        End If
        
    Next i
    
    If k = 0 Then
        MsgBox ("No valid attachement in this tweet")
    End If
    
Else
    MsgBox ("No attachements in this tweet")
End If



End Sub

Private Sub btn_search_Click()

Call search_interal_db_and_twitter

End Sub

Private Sub btn_sync_with_tradator_Click()

Dim tmp_wrbk As Workbook

For Each tmp_wrbk In Workbooks
    If tmp_wrbk.name = "Tradator.xls" Then
        Me.Hide
        tradator_get_all_tweet (tradator_get_vec_mention_to_track())
        Exit Sub
    End If
Next

MsgBox ("open Tradator !")

End Sub

Private Sub btn_trade_securities_Click()

Dim oReg As New VBScript_RegExp_55.RegExp
Dim matches As VBScript_RegExp_55.MatchCollection
Dim match As VBScript_RegExp_55.match
oReg.Global = True
oReg.IgnoreCase = True


Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer, p As Integer, q As Integer

Dim oBBG As New cls_Bloomberg_Sync

Dim tweet_id_txt As String, tweet_id_long As Long

tweet_id_txt = LV_last_tweet.SelectedItem.key
tweet_id_txt = Replace(tweet_id_txt, "user_", "")
tweet_id_long = CLng(tweet_id_txt)

Dim prt As String

Dim tweet_txt As String
tweet_txt = LV_last_tweet.ListItems("user_" & tweet_id_txt).ListSubItems("tweet_" & tweet_id_txt).Text

Dim tickers_in_tweets() As Variant

Dim vec_ticker_api() As Variant
Dim output_api_bbg As Variant

k = 0
If InStr(UCase(tweet_txt), UCase("@portfolio")) <> 0 Then
    Dim extract_hashtags_in_tweet As Variant
    extract_hashtags_in_tweet = get_specific_tweet_content(Array(f_tweet_json_hashtags), , , , , , tweet_id_long)
    
    'la premiere entree correspond au prt
    If IsEmpty(extract_hashtags_in_tweet(0)) Then
        
        Exit Sub
    Else
        For i = 0 To UBound(extract_hashtags_in_tweet(0), 1)
            prt = extract_hashtags_in_tweet(0)(i)(0)
            Exit For
        Next i
    End If
    
    'Dim extract_tickers_from_prt As Variant
    'extract_tickers_from_prt = get_list_tickers_from_tweeted_portfolio(prt)
    
    'goto format
    Application.Calculation = xlCalculationManual
    
    Worksheets("FORMAT2").CB_format2_twitter_basket.Value = prt
    
    Worksheets("FORMAT2").CB_format2_portfolio.Value = ""
    Worksheets("FORMAT2").CB_format2_region.Value = ""
    Worksheets("FORMAT2").TB_format2_valeur_eur_value.Value = ""
    Worksheets("FORMAT2").TB_format2_delta_value.Value = ""
    Worksheets("FORMAT2").TB_format2_vega_value.Value = ""
    Worksheets("FORMAT2").TB_format2_theta_value.Value = ""
    Worksheets("FORMAT2").CB_format2_sector.Value = ""
    Worksheets("FORMAT2").CB_format2_industry.Value = ""
    Worksheets("FORMAT2").TB_format2_beta_value.Value = ""
    Worksheets("FORMAT2").CB_format2_tag.Value = ""
    Worksheets("FORMAT2").TB_format2_rel_1d_value.Value = ""
    Worksheets("FORMAT2").TB_format2_rel_5d_value.Value = ""
    Worksheets("FORMAT2").TB_format2_rank_eps_value.Value = ""
    Worksheets("FORMAT2").TB_format2_rank_overall_value.Value = ""
    
    Worksheets("FORMAT2").CB_buy_s3.Value = True
    Worksheets("FORMAT2").CB_buy_s2.Value = True
    Worksheets("FORMAT2").CB_buy_s1.Value = True
    Worksheets("FORMAT2").CB_smart_p.Value = True
    Worksheets("FORMAT2").CB_sell_R3.Value = True
    Worksheets("FORMAT2").CB_sell_R2.Value = True
    Worksheets("FORMAT2").CB_sell_R1.Value = True
    
    
    Call preparation_trades_with_filters
    
    Me.Hide
    Worksheets("FORMAT2").Cells(100, 1).Select
    
    Application.Calculation = xlCalculationAutomatic
    
Else
    
    Dim array_buy_hashtags() As Variant, array_sell_hashtags() As Variant
        array_buy_hashtags = Array("#BUY", "#B", "#LONG")
        array_sell_hashtags = Array("#S", "#SELL", "#SHORT", "#SS", "#SHORTSELL")
    
    Dim array_stop_hashtags() As Variant
        array_stop_hashtags = Array("#STP", "#STOP")

    Dim array_target_hashtags() As Variant
            array_target_hashtags = Array("#TGT", "#TARGET")
    
    
    Dim extract_tickers_in_tweet As Variant, extract_mentions_in_tweet As Variant
    extract_tickers_in_tweet = get_specific_tweet_content(Array(f_tweet_json_tickers), , , , , , tweet_id_long)
    extract_mentions_in_tweet = get_specific_tweet_content(Array(f_tweet_json_mentions), , , , , , tweet_id_long)
    extract_hashtags_in_tweet = get_specific_tweet_content(Array(f_tweet_json_hashtags), , , , , , tweet_id_long)
    
    
    
    If IsEmpty(extract_tickers_in_tweet(0)) Then
        Exit Sub
    Else
        
        'appel api pour connaitre les last price
        k = 0
        For i = 0 To UBound(extract_tickers_in_tweet(0)(0), 1)
            ReDim Preserve vec_ticker_api(k)
            vec_ticker_api(k) = get_clean_ticker_bloomberg(extract_tickers_in_tweet(0)(0)(i))
            k = k + 1
        Next i
        
        'appel api
        'output_api_bbg = bbg_multi_tickers_and_multi_fields(vec_ticker_api, Array("last_price"))
        output_api_bbg = oBBG.bdp(vec_ticker_api, Array("last_price"), output_format.of_vec_without_header)
        
        For i = 0 To UBound(vec_ticker_api, 1)
            
            'If IsNumeric(output_api_bbg(i)(0)) Then
                
                
'                'quick trade module
'                frm_redi_plus.TB_ticker.Value = vec_ticker_api(i)
'
'                frm_redi_plus.TB_price.Value = "Price"
'
'                If IsNumeric(output_api_bbg(i)(0)) Then
'                    frm_redi_plus.TB_price.Value = output_api_bbg(i)(0)
'                End If
'
'                frm_redi_plus.Show

                'advanced trade module
                Application.Calculation = xlCalculationManual
                frm_redi_plus_advanced.TB_order_ticker.Value = vec_ticker_api(i)
                If IsNumeric(output_api_bbg(i)(0)) Then
                    frm_redi_plus_advanced.TB_custom_price.Value = output_api_bbg(i)(0)
                    frm_redi_plus_advanced.L_limit_price.Caption = "LIMIT (LAST)"
                Else
                    frm_redi_plus_advanced.L_limit_price.Caption = "LIMIT"
                End If
                
                
                
                'check si possible de trouver side
                Dim tmp_side As String
                If IsEmpty(extract_hashtags_in_tweet(0)(0)) Then
                Else
                    For m = 0 To UBound(array_buy_hashtags, 1)
                        For n = 0 To UBound(extract_hashtags_in_tweet(0)(0), 1)
                            If array_buy_hashtags(m) = extract_hashtags_in_tweet(0)(0)(n) Then
                                frm_redi_plus_advanced.CB_Order_Side.Value = "BUY"
                                Exit For
                            End If
                            
                        Next n
                    Next m
                    
                    
                    For m = 0 To UBound(array_sell_hashtags, 1)
                        For n = 0 To UBound(extract_hashtags_in_tweet(0)(0), 1)
                            If array_sell_hashtags(m) = extract_hashtags_in_tweet(0)(0)(n) Then
                                frm_redi_plus_advanced.CB_Order_Side.Value = "SELL"
                                Exit For
                            End If
                        Next n
                    Next m
                    
                End If
                
                
                'mise en place stop si dispo
                Dim tmp_stop_price As Double
                If IsEmpty(extract_hashtags_in_tweet(0)(0)) Then
                Else
                    For m = 0 To UBound(array_stop_hashtags, 1)
                        For n = 0 To UBound(extract_hashtags_in_tweet(0)(0), 1)
                            
                            If array_stop_hashtags(m) = extract_hashtags_in_tweet(0)(0)(n) Then
                                
                                'suivi d un prix ?
                                oReg.Pattern = array_stop_hashtags(m) & "\s+[\d]+(\.[\d]+|)"
                                Set matches = oReg.Execute(tweet_txt)
                                
                                For Each match In matches
                                    'extraction du prix
                                    tmp_stop_price = CDbl(Replace(match.Value, array_stop_hashtags(m), ""))
                                    frm_redi_plus_advanced.TB_custom_stop.Value = tmp_stop_price
                                    
                                    Exit For
                                Next
                                
                                Exit For
                            End If
                            
                        Next n
                    Next m
                End If
                
                
                
                'mise en place target si dispo
                Dim tmp_target_price As Double
                If IsEmpty(extract_hashtags_in_tweet(0)(0)) Then
                Else
                    For m = 0 To UBound(array_target_hashtags, 1)
                        For n = 0 To UBound(extract_hashtags_in_tweet(0)(0), 1)
                            
                            If array_target_hashtags(m) = extract_hashtags_in_tweet(0)(0)(n) Then
                                
                                'suivi d un prix ?
                                oReg.Pattern = array_target_hashtags(m) & "\s+[\d]+(\.[\d]+|)"
                                Set matches = oReg.Execute(tweet_txt)
                                
                                For Each match In matches
                                    'extraction du prix
                                    tmp_target_price = CDbl(Replace(match.Value, array_target_hashtags(m), ""))
                                    frm_redi_plus_advanced.TB_custom_target.Value = tmp_target_price
                                    
                                    Exit For
                                Next
                                
                                Exit For
                            End If
                            
                        Next n
                    Next m
                End If
                
                
                
                'check si le broker peut en etre deduit du tweet grace au mention
                Dim list_brokers As Variant
                list_brokers = Array(Array("BARCLAYS", Array("BARCLAYS", "BARCLAY", "BARC", "BARCAP")), Array("CS", Array("CS", "CREDITSUISSE")), Array("DB", Array("DB", "DBK")), Array("JPM", Array("JPM", "JP")), Array("ML", Array("ML", "BOA", "BOAML", "MERRILL")), Array("MS", Array("MS", "MORGANSTANLEY")), Array("NEGOCE", Array("NEGOCE", "PICTET", "PMA")), Array("WONEIL", Array("WONEIL", "ONEIL")))
                
                Dim tmp_mention As String, find_broker As Boolean
                find_broker = False
                If IsEmpty(extract_mentions_in_tweet(0)(0)) Then
                Else
                    
                    
                    For m = 0 To UBound(list_brokers, 1)
                        For n = 0 To UBound(list_brokers(m)(1), 1)
                            
                            If "@" & UCase(list_brokers(m)(1)(n)) = UCase(extract_mentions_in_tweet(0)(0)(0)) Then
                                find_broker = True
                                Exit For
                            End If
                            
                        Next n
                        
                        If find_broker = True Then
                            'ajustement du champs broker
                            frm_redi_plus_advanced.CB_exec_broker.Value = list_brokers(m)(0)
                            Exit For
                        End If
                        
                    Next m
                    
                    
                End If
                
                
                frm_redi_plus_advanced.Show
            
            'End If
            
        Next i
        
    End If
    
    
End If

End Sub



Private Sub L_attach_Click()

End Sub

Private Sub L_show_advanced_features_Click()

If Me.Height > 1.1 * small_form_height And Me.Height <= 1.1 * advanced_form_height Then
    Me.Height = small_form_height
Else
    Me.Height = advanced_form_height
End If

End Sub

Private Sub LB_Helpers_DblClick(ByVal Cancel As msforms.ReturnBoolean)

If LB_Helpers.ListIndex <> -1 Then
    replace_last_word_in_new_tweet (LB_Helpers.list(LB_Helpers.ListIndex) & " ")
    TB_tweet.SetFocus
End If

End Sub


Private Sub TB_search_Enter()

Call search_interal_db_and_twitter

End Sub

Private Sub TB_tweet_Change()

LB_Helpers.Clear

If TB_tweet.Value <> "" Then
    
    Dim last_word As String
    last_word = get_last_word_new_tweet
    
    If IsEmpty(last_word) = False Then
        
        Call twitter_new_autocomplete_structure(last_word)
        
        Dim test_char As String
        test_char = Left(last_word, 1)
        
        If InStr(test_char, "@") <> 0 Or InStr(test_char, "#") <> 0 Or InStr(test_char, "$") <> 0 Then
            
            Dim list_helpers As Variant
            list_helpers = tweet_autocomplete(last_word)
            
            If IsEmpty(list_helpers) = False Then
                For i = 0 To UBound(list_helpers, 1)
                    LB_Helpers.AddItem list_helpers(i)
                Next i
            End If
            
        End If
    End If
End If


End Sub


Private Sub twitter_new_autocomplete_structure(ByVal last_word As String)

Dim i As Integer

'trigger / precommand
Dim autocomplete() As Variant

autocomplete = Array(Array("@DT", "@DT #BUYSELL $ticker.us #STOP s #TGT t #ROOM sb"))


Dim vbanswer As Variant
For i = 0 To UBound(autocomplete, 1)
    If UCase(last_word) = UCase(autocomplete(i)(0)) Then
        vbanswer = MsgBox("autocomplete available. Use it ?", vbYesNo)
        
        If vbanswer = vbYes Then
            TB_tweet.Value = autocomplete(i)(1)
        End If
        
        Exit For
    End If
Next i


End Sub


Private Function replace_last_word_in_new_tweet(ByVal text_to_add As String)

If IsEmpty(get_last_word_new_tweet) = False And IsEmpty(text_to_add) = False Then
    TB_tweet.Value = Replace(TB_tweet.Value, get_last_word_new_tweet, text_to_add)
End If

End Function


Private Function get_last_word_new_tweet() As String

Dim tmp_str As String

If TB_tweet.Value <> "" Then
    
    If Len(TB_tweet.Value) = 1 Then
        get_last_word_new_tweet = Left(TB_tweet.Value, 1)
    Else
        
        'remonte jusqu'au dernier espace ou postion 1
        tmp_str = StrReverse(TB_tweet.Value)
        
        If InStr(tmp_str, " ") = 0 Then '1 seul mot dans la box
            get_last_word_new_tweet = TB_tweet.Value
            Exit Function
        Else
            tmp_str = Left(tmp_str, InStr(tmp_str, " ") - 1)
            get_last_word_new_tweet = StrReverse(tmp_str)
            Exit Function
        End If
        
    End If
    
Else
    sub_get_last_word = Empty
End If

End Function


Private Function get_attach_file() As Variant

get_attach_file = Empty

k = 0
Dim list_files()
If LB_attach.ListCount > 0 Then
    
    For Each tmp_file In LB_attach.list
        If IsEmpty(tmp_file) = False And IsNull(tmp_file) = False Then
            ReDim Preserve list_files(k)
            list_files(k) = tmp_file
            k = k + 1
        End If
    Next
    
    If k > 0 Then
        get_attach_file = list_files
    Else
        get_attach_file = Empty
    End If
    
    Exit Function
End If

End Function

Private Sub UserForm_Click()

End Sub
