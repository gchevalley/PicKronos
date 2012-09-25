Attribute VB_Name = "Bas_Twitter"
'FOLDERS & FILES
Public Const db_path_base As String = "Q:\front\greg\Twitter\"
Public Const db_twitter As String = "db_twitter.sqlt3"

Public Const g_col_bw As String = "directory_tmp_download"
Public Const directory_local_copy As String = "local_datas"



'TABLES
Public Const t_tweet_cache As String = "t_tweet_cache"
    Public Const f_tweet_cache_cache_id As String = "tweet_cache_cache_id"
    Public Const f_tweet_cache_raw_tweet As String = "tweet_cache_raw_tweet"
    Public Const f_tweet_cache_parsed_status As String = "tweet_cache_parse_status"

Public Const t_tweet As String = "t_tweet"
    Public Const f_tweet_id As String = "tweet_id"
    Public Const f_tweet_datetime As String = "tweet_datetime"
    Public Const f_tweet_from As String = "tweet_from"
    Public Const f_tweet_text = "tweet_text"
    Public Const f_tweet_json_tickers = "tweet_json_tickers"
    Public Const f_tweet_json_hashtags = "tweet_json_hashtags"
    Public Const f_tweet_json_mentions = "tweet_json_mentions"
    Public Const f_tweet_json_links = "tweet_json_links"
    
Public Const t_user As String = "t_user" '@
    Public Const f_user_id As String = "user_id"
    Public Const f_user_first_name As String = "user_first_name"
    Public Const f_user_name As String = "user_name"
    Public Const f_user_pco_bloomberg = "user_pco_bloomberg"
    
Public Const t_mention As String = "t_mention"
    Public Const f_mention_tweet_id As String = "mention_tweet_id"
    'Public Const f_mention_source As String = "mention_source" 'username QUI appelle / cite
    Public Const f_mention_target As String = "mention_target" 'username cite / appele

Public Const t_hyperlink_and_file As String = "t_hyperlink_and_file"
    Public Const f_hyperlink_and_file_tweet_id As String = "hyperlink_and_file_tweet_id"
    Public Const f_hyperlink_and_file_source As String = "hyperlink_and_file_source"
    Public Const f_hyerplink_and_file_tinyurl As String = "hyperlink_and_file_tinyurl"
    Public Const f_hyperlink_and_file_local_copy As String = "hyperlink_and_file_local_copy"

Public Const t_hashtag As String = "t_hashtag" '#
    Public Const f_hashtag_id As String = "hashtag_id"

Public Const t_category As String = "t_category" '|
    Public Const f_category_id As String = "category_id"
    
Public Const t_hashtag_category As String = "t_hashtag_category"
    Public Const f_hashtag_category_hashtag As String = "f_hashtag_category_hashtag"
    Public Const f_hashtag_category_category As String = "f_hashtag_category_category"
    
Public Const t_ticker As String = "t_ticker" '$
    Public Const f_ticker_twitter As String = "ticker_twitter"
    Public Const f_ticker_bloomberg As String = "ticker_bloomberg"

Public Const t_market_data As String = "t_market_data"
    Public Const f_market_data_ticker_twitter As String = "market_data_twitter"
    Public Const f_market_data_datetime As String = "market_data_datetime"
    Public Const f_market_data_px_last As String = "market_data_px_last"
    Public Const f_market_data_impl_vol As String = "market_data_impl_vol"
    Public Const f_market_data_histo_vol_30d As String = "market_data_histo_vol_30d"
    Public Const f_market_data_central_rank_eps As String = "market_data_central_rank_eps"

Public Const t_twitter_helper As String = "t_twitter_helper"
    Public Const f_twitter_helper_text1 As String = "f_twitter_helper_text1"
    Public Const f_twitter_helper_text2 As String = "f_twitter_helper_text2"
    Public Const f_twitter_helper_text3 As String = "f_twitter_helper_text3"
    Public Const f_twitter_helper_numeric1 As String = "f_twitter_helper_numeric1"
    Public Const f_twitter_helper_numeric2 As String = "f_twitter_helper_numeric2"
    Public Const f_twitter_helper_numeric3 As String = "f_twitter_helper_numeric3"
    


'VIEWS
Public Const v_last_tweet_ticker As String = "last_tweet_ticker"


'COMMON VAR
Public t_current_username As Variant
Public t_current_ticker As Variant
Public t_current_category As Variant
Public t_current_hashtag As Variant
Public system_mode As Variant


'OFFLINE MODE
Public Enum internal_kronos_local_mode
    online_with_db = 1
    offline_with_xml = 0
End Enum

Public Const offline_xml_file As String = "internal_kronos_local_twitter.xml"
Public Const offline_xml_tag_root As String = "internal_kronos_local_twitter"
Public Const offline_xml_tag_tweet As String = "tweet"
    Public Const offline_xml_tag_tweet_id As String = "id"
    Public Const offline_xml_tag_tweet_from As String = "from"
    Public Const offline_xml_tag_tweet_datetime As String = "datetime"
    Public Const offline_xml_tag_tweet_text As String = "text"




Public Function twitter_get_db_path() As String

If exist_file(db_path_base & db_twitter) Then
    twitter_get_db_path = db_path_base & db_twitter
ElseIf exist_file(ThisWorkbook.path & "\" & db_twitter) Then
    twitter_get_db_path = ThisWorkbook.path & "\" & db_twitter
Else
    twitter_get_db_path = ThisWorkbook.path & "\" & db_twitter
End If

End Function


Public Function check_mode() As Integer

'si office -> sqlite sinon XML offline
If exist_file(db_path_base & db_twitter) Then
    check_mode = internal_kronos_local_mode.online_with_db
Else
    check_mode = internal_kronos_local_mode.offline_with_xml
End If

End Function


Sub show_form_Twitter_new_tweet()

Dim date_tmp As Date

frm_Tweet_new.TB_tweet.Value = ""
frm_Tweet_new.LB_Helpers.Clear
frm_Tweet_new.LB_attach.Clear

Dim list_last_tweets As Variant
list_last_tweets = get_last_tweets(15)
    

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



Dim twitter_trends_daily As Variant, twitter_trends_weekly As Variant
    twitter_trends_daily = get_twitter_trends("daily", 0)
    twitter_trends_weekly = get_twitter_trends("weekly", 0)
    
    'fusion pour avoir sous la forme d'un seul vecteur
    '-> 10 trends
    limit_nbre_trends = 10
    
    Dim vec_trends_daily() As Variant
    Dim vec_trends_weekly() As Variant
    
    k = 0
    If IsEmpty(twitter_trends_daily) = False Then
        For i = 0 To UBound(twitter_trends_daily, 1)
            
            For j = 0 To UBound(twitter_trends_daily(1), 1)
                If k < limit_nbre_trends Then
                    ReDim Preserve vec_trends_daily(k)
                    vec_trends_daily(k) = twitter_trends_daily(i)(1)(j)
                    k = k + 1
                End If
            Next j
                
        Next i
    End If
    
    k = 0
    If IsEmpty(twitter_trends_weekly) = False Then
        For i = 0 To UBound(twitter_trends_weekly, 1)
            
            For j = 0 To UBound(twitter_trends_weekly(1), 1)
                If k < limit_nbre_trends Then
                    ReDim Preserve vec_trends_weekly(k)
                    vec_trends_weekly(k) = twitter_trends_weekly(i)(1)(j)
                    k = k + 1
                End If
            Next j
                
        Next i
    End If
    
    Dim max_trends As Integer
    
    If IsEmpty(twitter_trends_daily) = False And IsEmpty(twitter_trends_weekly) = False Then
        If UBound(vec_trends_daily, 1) = UBound(vec_trends_weekly, 1) Then
            max_trends = UBound(vec_trends_daily, 1)
        Else
            If UBound(vec_trends_daily, 1) < UBound(twitter_trends_weekly, 1) Then
                max_trends = UBound(vec_trends_daily, 1)
            Else
                max_trends = UBound(vec_trends_weekly, 1)
            End If
        End If
    End If
    
    
    'affichage de l'output
    If k > 0 Then
        
        With frm_Tweet_new.LV_real_twitter
            
            With .ColumnHeaders
                .Clear

                .Add , , "daily trend", 150
                .Add , , "weekly trends", 150
            End With
            
            For i = 0 To max_trends
                With .ListItems
                    .Add , , vec_trends_daily(i)  'user
                End With
                
                .ListItems(i + 1).ListSubItems.Add , , vec_trends_weekly(i)
                
            Next i
        End With
        
    End If


frm_Tweet_new.Show

End Sub



Public Function init_offline_xml() As String

Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

If exist_file(src_path & offline_xml_file) = False Then
    
    'etablissement du squelette du fichier
    Dim oDOM As New DOMDocument
    Dim oRootElement As IXMLDOMElement
    
    Set oRootElement = oDOM.createElement(offline_xml_tag_root)
    oDOM.appendChild oRootElement
    
    oDOM.Save (src_path & offline_xml_file)
    
    init_offline_xml = src_path & offline_xml_file
    
Else
    init_offline_xml = src_path & offline_xml_file
End If

End Function


Sub init_db_twitter()

Dim create_db_status As Variant
create_db_status = sqlite3_create_db(twitter_get_db_path)


'creation des tables
Dim create_table As String
Dim exec_query_return As Variant

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_user) = False Then
    create_table = sqlite3_get_query_create_table(t_user, Array(Array(f_user_id, "TEXT", ""), Array(f_user_first_name, "TEXT", ""), Array(f_user_name, "TEXT", ""), Array(f_user_pco_bloomberg, "TEXT", "")), Array(Array(f_user_id, "ASC")))
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
    
    Call init_db_twitter_with_datas
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_tweet) = False Then
    create_table = "CREATE TABLE " & t_tweet & " (" & f_tweet_id & " REAL, " & f_tweet_datetime & " NUMERIC, " & f_tweet_from & " TEXT, " & f_tweet_text & " TEXT, " & f_tweet_json_tickers & " TEXT, " & f_tweet_json_hashtags & " TEXT, " & f_tweet_json_mentions & " TEXT, " & f_tweet_json_links & " TEXT, FOREIGN KEY(" & f_tweet_from & ") REFERENCES " & t_user & " (" & f_user_id & "), PRIMARY KEY(" & f_tweet_id & "))"
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_hyperlink_and_file) = False Then
    create_table = "CREATE TABLE " & t_hyperlink_and_file & " (" & f_hyperlink_and_file_tweet_id & " REAL, " & f_hyperlink_and_file_source & " TEXT, " & f_hyerplink_and_file_tinyurl & " TEXT, " & f_hyperlink_and_file_local_copy & " TEXT, FOREIGN KEY(" & f_hyperlink_and_file_tweet_id & ") REFERENCES " & t_tweet & " (" & f_tweet_id & "), PRIMARY KEY(" & f_hyperlink_and_file_tweet_id & ", " & f_hyperlink_and_file_source & "))"
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_hashtag) = False Then
    create_table = sqlite3_get_query_create_table(t_hashtag, Array(Array(f_hashtag_id, "TEXT", "")), Array(Array(f_hashtag_id, "ASC")))
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_category) = False Then
    create_table = sqlite3_get_query_create_table(t_category, Array(Array(f_category_id, "TEXT", "")), Array(Array(f_category_id, "ASC")))
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_hashtag_category) = False Then
    create_table = "CREATE TABLE " & t_hashtag_category & " (" & f_hashtag_category_hashtag & " TEXT, " & f_hashtag_category_category & " TEXT, FOREIGN KEY (" & f_hashtag_category_hashtag & ") REFERENCES " & t_hashtag & "(" & f_hashtag_id & "), FOREIGN KEY(" & f_hashtag_category_category & ") REFERENCES " & t_category & " (" & f_category_id & "), PRIMARY KEY(" & f_hashtag_category_hashtag & ", " & f_hashtag_category_category & "))"
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_ticker) = False Then
    create_table = sqlite3_get_query_create_table(t_ticker, Array(Array(f_ticker_bloomberg, "TEXT", ""), Array(f_ticker_twitter, "TEXT", "")), Array(Array(f_ticker_twitter, "ASC")))
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_mention) = False Then
    create_table = "CREATE TABLE " & t_mention & " (" & f_mention_tweet_id & " REAL, " & f_mention_target & " TEXT, FOREIGN KEY (" & f_mention_tweet_id & ") REFERENCES " & t_tweet & "(" & f_tweet_id & "), FOREIGN KEY (" & f_mention_target & ") REFERENCES " & t_user & " (" & f_user_id & ") , PRIMARY KEY(" & f_mention_tweet_id & ", " & f_mention_target & "))"
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If


If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_market_data) = False Then
    create_table = "CREATE TABLE " & t_market_data & " (" & f_market_data_ticker_twitter & " TEXT, " & f_market_data_datetime & " NUMERIC, " & f_market_data_px_last & " REAL, " & f_market_data_impl_vol & " REAL, " & f_market_data_histo_vol_30d & " REAL, " & f_market_data_central_rank_eps & " REAL)"
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If

If sqlite3_check_if_table_already_exist(twitter_get_db_path, t_twitter_helper) = False Then
    create_table = "CREATE TABLE " & t_twitter_helper & " (" & f_twitter_helper_text1 & " TEXT, " & f_twitter_helper_text2 & " TEXT, " & f_twitter_helper_text3 & " TEXT, " & f_twitter_helper_numeric1 & " NUMERIC, " & f_twitter_helper_numeric2 & " NUMERIC, " & f_twitter_helper_numeric3 & " NUMERIC)"
    exec_query_return = sqlite3_create_tables(twitter_get_db_path, Array(create_table))
End If



End Sub


Sub init_db_twitter_with_datas()

Dim i As Integer, j As Integer, k As Integer

Dim users() As Variant
k = 0
ReDim Preserve users(k)
users(k) = Array("@amorange", "Alexis", "Morange", "x01221248")
k = k + 1

ReDim Preserve users(k)
users(k) = Array("@gchevalley", "Gregory", "Chevalley", "x01231024")
k = k + 1

ReDim Preserve users(k)
users(k) = Array("@landreasson", "Lennart", "Andreasson", "x01231003")
k = k + 1

ReDim Preserve users(k)
users(k) = Array("@jstouff", "Julien", "Stouff", "x01221179")
k = k + 1

Dim status_insert_with_transac As Variant
status_insert_with_transac = sqlite3_insert_with_transaction(twitter_get_db_path, t_user, users, Array(f_user_id, f_user_first_name, f_user_name, f_user_pco_bloomberg))

End Sub


Sub check_db_status()

Dim oJSON As New JSONLib
Dim col_from_json As Collection

Dim extract_tickers As Variant
extract_tickers = sqlite3_query(twitter_get_db_path, "SELECT * FROM " & t_ticker)

Dim extract_users As Variant
extract_users = sqlite3_query(twitter_get_db_path, "SELECT * FROM " & t_user)

Dim extract_mentions As Variant
extract_mentions = sqlite3_query(twitter_get_db_path, "SELECT * FROM " & t_mention)

Dim extract_hyperlinks_and_files As Variant
extract_hyperlinks_and_files = sqlite3_query(twitter_get_db_path, "SELECT * FROM " & t_hyperlink_and_file)

Dim extract_hashtags As Variant
extract_hashtags = sqlite3_query(twitter_get_db_path, "SELECT * FROM " & t_hashtag)



    Dim extract_tweet_t_structure As Variant
    extract_tweet_t_structure = sqlite3_get_table_structure(twitter_get_db_path, t_tweet)
Dim extract_tweets As Variant
extract_tweets = sqlite3_query(twitter_get_db_path, "SELECT * FROM " & t_tweet)

    If UBound(extract_tweets, 1) > 0 Then
        
        For i = 0 To UBound(extract_tweets(0), 1)
            If InStr(UCase(extract_tweets(0)(i)), UCase("JSON")) <> 0 Then
                Set col_from_json = oJSON.parse(decode_json_from_DB(extract_tweets(1)(i)))
                debug_test = "test"
            End If
        Next i
        
    End If


End Sub


Public Function get_last_tweet_id() As Long

Dim extract_last_tweet As Variant
extract_last_tweet = sqlite3_query(twitter_get_db_path, "SELECT MAX(" & f_tweet_id & ") FROM " & t_tweet)

If UBound(extract_last_tweet, 1) = 0 Then
    get_last_tweet_id = 0
Else
    If IsNull(extract_last_tweet(1)(0)) Then
        get_last_tweet_id = 0
    Else
        get_last_tweet_id = extract_last_tweet(1)(0)
    End If
End If

End Function


Public Function get_mentions_from_tweet(ByRef tweet As String)

Dim k As Integer


Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

oReg.Pattern = "@[\S]+"

Set matches = oReg.Execute(tweet)

k = 0
Dim tmp_mention As String
Dim list_mentions() As Variant

Dim db_usernamne_already_mount
    db_usernamne_already_mount = False


For Each match In matches
    
    If db_usernamne_already_mount = False Then
        Dim return_mount_usernames As Variant
        return_mount_usernames = mount_usernames()
        db_usernamne_already_mount = True
    End If
    
    's'assure qu'il existe deja tous
    tmp_mention = create_username(match.Value)
    
    ReDim Preserve list_mentions(k)
    list_mentions(k) = Array(tmp_mention, match.Value)
    
    If InStr(tweet, list_mentions(k)(1)) > 1 Then
        If Mid(tweet, InStr(tweet, list_mentions(k)(1)) - 1, 1) = " " Then
            tweet = Replace(tweet, list_mentions(k)(1), list_mentions(k)(0))
        Else
            tweet = Replace(tweet, list_mentions(k)(1), " " & list_mentions(k)(0))
        End If
    Else
        tweet = Replace(tweet, list_mentions(k)(1), list_mentions(k)(0))
    End If
    
    k = k + 1
Next

If k = 0 Then
    get_mentions_from_tweet = Empty
Else
    get_mentions_from_tweet = list_mentions
End If

End Function


Public Function mount_usernames(Optional ByVal mount_query As Variant)

Dim sql_query As String

If IsMissing(mount_query) Then
    sql_query = "SELECT " & f_user_id & " FROM " & t_user & " ORDER BY " & f_user_id & " COLLATE NOCASE ASC"
Else
    sql_query = mount_query
End If

t_current_username = sqlite3_query(twitter_get_db_path, sql_query)
mount_usernames = t_current_username

End Function


Public Function create_username(ByVal username As String, Optional ByVal first_name As String = "auto", Optional ByVal name As String = "auto", Optional ByVal pco_bloomberg As Variant = Empty) As String

Dim i As Long, j As Long, k As Long

username = LCase(username)
If Left(username, 1) <> "@" Then
    username = "@" & username
End If

create_username = username

Dim insert_status As Variant

If IsEmpty(t_current_username) = True Then
    Dim mount_usernames_status As Variant
    mount_usernames_status = mount_usernames
End If

If UBound(t_current_username, 1) = 0 Then
    '1ere entree
    insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_user, Array(Array(username, first_name, name, pco_bloomberg)), Array(f_user_id, f_user_first_name, f_user_name, f_user_pco_bloomberg))
    mount_usernames_status = mount_usernames
Else
    's'assure que n'existe pas deja
    For i = 1 To UBound(t_current_username, 1)
        If UCase(username) = UCase(t_current_username(i)(0)) Then
            Exit For
        Else
            If i = UBound(t_current_username, 1) Then
                insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_user, Array(Array(username, first_name, name, pco_bloomberg)), Array(f_user_id, f_user_first_name, f_user_name, f_user_pco_bloomberg))
                mount_usernames_status = mount_usernames
            End If
        End If
    Next i
End If

End Function


Public Function get_username_from_tweet()

Dim sql_query As String

Dim extract_users As Variant

If check_mode = internal_kronos_local_mode.online_with_db Then

    sql_query = "SELECT " & f_user_id & ", " & f_user_pco_bloomberg & " FROM " & t_user
    extract_users = sqlite3_query(twitter_get_db_path, sql_query)
    
    
    If UCase(Trim(Environ("userdomain"))) = "PCO" Then 'au bureau
        
        'maintenant distinguer machine bbg / pco lotus
        system_mode = "PICTET"
        
        For i = 1 To UBound(extract_users, 1)
        
            If UCase(Mid(extract_users(i)(0), 2)) = UCase(Trim(Environ("username"))) Or UCase(extract_users(i)(1)) = UCase(Trim(Environ("username"))) Then
                get_username_from_tweet = extract_users(i)(0)
                Exit Function
            End If
        Next i
    
    Else
        GoTo inputbox_get_user
    End If
    
ElseIf check_mode = internal_kronos_local_mode.offline_with_xml Then
xml_offline_mode:
    
    'home without db
    
    extract_users = Array(f_user_id, Array("@amorange"), Array("@gchevalley"), Array("@jstouff"), Array("@landreasson"))
        
inputbox_get_user:

    system_mode = "HOME"
    
    Dim msg_inputbox As String
    msg_inputbox = "Chose your number [" & 1 & "-" & UBound(extract_users, 1) & "] " & vbCrLf
    
    For i = 1 To UBound(extract_users, 1)
        msg_inputbox = msg_inputbox & i & " : " & extract_users(i)(0) & vbCrLf
    Next i
    
    Dim answer_user As Variant
    answer_user = InputBox(msg_inputbox, "Chose your number")
    
    If IsNumeric(answer_user) Then
        answer_user = CDbl(answer_user)
check_answer_chose_user_name:
        If answer_user <= UBound(extract_users, 1) And answer_user > 0 Then
            get_username_from_tweet = extract_users(answer_user)(0)
            Exit Function
        Else
            answer_user = InputBox(msg_inputbox, "Error")
            GoTo check_answer_chose_user_name
        End If
    End If
    
    
    
End If

End Function


Public Function get_hashtags_from_tweet(ByRef tweet As String) As Variant

Dim k As Integer


Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

'oReg.Pattern = "#[A-Za-z0-9]+"
oReg.Pattern = "#[\S]+"

Set matches = oReg.Execute(tweet)

k = 0
Dim tmp_hashtag As String
Dim list_hashtags() As Variant

Dim db_hashtag_already_mount
    db_hashtag_already_mount = False


For Each match In matches
    
    If db_hashtag_already_mount = False Then
        Dim return_mount_hashtags As Variant
        return_mount_hashtags = mount_hashtags()
        db_hashtag_already_mount = True
    End If
    
    's'assure qu'il existe deja tous
    tmp_hashtag = create_hashtag(UCase(match.Value))
    
    ReDim Preserve list_hashtags(k)
    list_hashtags(k) = Array(tmp_hashtag, match.Value)
    
    If InStr(tweet, list_hashtags(k)(1)) > 1 Then
        If Mid(tweet, InStr(tweet, list_hashtags(k)(1)) - 1, 1) = " " Then
            tweet = Replace(tweet, list_hashtags(k)(1), list_hashtags(k)(0))
        Else
            tweet = Replace(tweet, list_hashtags(k)(1), " " & list_hashtags(k)(0))
        End If
    Else
        tweet = Replace(tweet, list_hashtags(k)(1), list_hashtags(k)(0))
    End If
    
    k = k + 1
Next

If k = 0 Then
    get_hashtags_from_tweet = Empty
Else
    get_hashtags_from_tweet = list_hashtags
End If

End Function


Public Function mount_tickers(Optional ByVal mount_query As Variant) As Variant

Dim sql_query As String

If IsMissing(mount_query) Then
    sql_query = "SELECT " & f_ticker_twitter & ", " & f_ticker_bloomberg & " FROM " & t_ticker & " ORDER BY " & f_ticker_twitter & " COLLATE NOCASE ASC"
Else
    sql_query = mount_query
End If

t_current_ticker = sqlite3_query(twitter_get_db_path, sql_query)
mount_tickers = t_current_ticker

End Function


Private Sub test_get_tickers_from_tweet()

Dim debug_test As Variant
debug_test = get_tickers_from_tweet("#buy test $abx")

End Sub


Public Function get_tickers_from_tweet(ByRef tweet As String) As Variant

Dim k As Integer


Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

'oReg.Pattern = "\$[A-Za-z0-9]+(\.[A-Za-z]{2}|)" 'based only equity
'with option support
oReg.Pattern = "\$[A-Za-z0-9]+((\.[A-Za-z]{2}|)\.([\d]{1,2}/|)[\d]{1,2}(/[\d]{1,2}|)\.[c|C|p|P][\d]+(\.[\d]+|)|\.[A-Za-z]{2}|)"

Set matches = oReg.Execute(tweet)

Dim ticker_bbg As String


k = 0
Dim list_tickers() As Variant
For Each match In matches
    's'assure qu'il existe deja tous
    ticker_bbg = get_clean_ticker_bloomberg(UCase(match.Value))
    create_ticker get_ticker_twitter(UCase(match.Value)), ticker_bbg
    
    ReDim Preserve list_tickers(k)
    list_tickers(k) = Array(get_ticker_twitter(UCase(match.Value)), match.Value, ticker_bbg)
    
    If InStr(tweet, list_tickers(k)(1)) > 1 Then
        If Mid(tweet, InStr(tweet, list_tickers(k)(1)) - 1, 1) = " " Then
            tweet = Replace(tweet, list_tickers(k)(1), list_tickers(k)(0))
        Else
            tweet = Replace(tweet, list_tickers(k)(1), " " & list_tickers(k)(0))
        End If
    Else
        tweet = Replace(tweet, list_tickers(k)(1), list_tickers(k)(0))
    End If
    
    k = k + 1
Next

If k = 0 Then
    get_tickers_from_tweet = Empty
Else
    get_tickers_from_tweet = list_tickers
End If

End Function


Private Function check_valid_url(ByVal url As String) As Boolean

If Left(UCase(url), 7) <> UCase("http://") Then
    url = "http://" & url
End If


Dim oHTTP As New MSXML2.XMLHTTP30

check_valid_url = False

On Error GoTo ErrorHandler

oHTTP.Open "HEAD", url, False
oHTTP.send

If oHTTP.Status <> 404 Then
    check_valid_url = True
    Exit Function
End If


Exit Function

ErrorHandler:
check_valid_url = False


End Function

Public Function get_links_from_tweet(ByRef tweet As String) As Variant



Dim k As Integer

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

oReg.Pattern = "([\w]+://|)[\S]+\.[A-Za-z]{2,3}(\S|)+"

Set matches = oReg.Execute(tweet)

k = 0
Dim list_links() As Variant
For Each match In matches
    If Left(match.Value, 1) <> "$" And Len(match.Value) > 8 And InStr(match.Value, " ") = 0 Then
        
        
        'check si valide url
        If check_valid_url(match.Value) Then
        
            ReDim Preserve list_links(k)
            list_links(k) = Array(get_tinyurl_from_weblink(get_clean_weblink(match.Value)), match.Value)
            
            If InStr(tweet, list_links(k)(1)) > 1 Then
                If Mid(tweet, InStr(tweet, list_links(k)(1)) - 1, 1) = " " Then
                    tweet = Replace(tweet, list_links(k)(1), list_links(k)(0))
                Else
                    tweet = Replace(tweet, list_links(k)(1), " " & list_links(k)(0))
                End If
            Else
                tweet = Replace(tweet, list_links(k)(1), list_links(k)(0))
            End If
            
            k = k + 1
        End If
    End If
Next

If k = 0 Then
    get_links_from_tweet = Empty
Else
    get_links_from_tweet = list_links
End If

End Function


Sub test_load_xml_offline()

Dim debug_test As Variant
debug_test = load_xml_offline_twitter_into_db("Q:\front\Kronos\internal_kronos_local_twitter.xml")

End Sub

Public Function load_xml_offline_twitter_into_db(ByVal path_xml_offline_twitter As String) As Variant

Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim sql_query As String

Dim oXML As New DOMDocument
Dim oXMLRoot As IXMLDOMElement
Dim oXMLTweet As IXMLDOMElement
Dim oXMLTweetText As IXMLDOMElement
Dim oXMLTweetFrom As IXMLDOMElement
Dim oXMLTweetDatetime As IXMLDOMElement

Dim oXMLtmp As IXMLDOMElement

Dim tweet_text As String, tweet_from As String, tweet_datetime As Double

oXML.load (path_xml_offline_twitter)
Set oXMLRoot = oXML.documentElement

Dim max_date As Double, min_date As Double
min_date = ToJulianDay(Date + 100)
max_date = ToJulianDay(Date - 100)



If check_mode = internal_kronos_local_mode.online_with_db Then

    Dim list_tweets_to_import() As Variant
        
        dim_tweets_import_from = 0
        dim_tweets_import_datetime = 1
        dim_tweets_import_text = 2
    
    k = 0
    For Each oXMLTweet In oXMLRoot.getElementsByTagName(offline_xml_tag_tweet)
        
        For Each oXMLtmp In oXMLTweet.childNodes
            If oXMLtmp.baseName = offline_xml_tag_tweet_from Then
                tweet_from = oXMLtmp.Text
            ElseIf oXMLtmp.baseName = offline_xml_tag_tweet_datetime Then
                tweet_datetime = CDbl(oXMLtmp.Text)
                
                If tweet_datetime < min_date Then
                    min_date = tweet_datetime
                End If
                
                If tweet_datetime > max_date Then
                    max_date = tweet_datetime
                End If
                
            ElseIf oXMLtmp.baseName = offline_xml_tag_tweet_text Then
                tweet_text = oXMLtmp.Text
            End If
            
        Next
        
        ReDim Preserve list_tweets_to_import(k)
        list_tweets_to_import(k) = Array(tweet_from, tweet_datetime, tweet_text)
        k = k + 1
        
    Next
    
    If k > 0 Then
        'extraction pour voir si le tweet n'est pas deja dans la base
        Dim extract_tweets As Variant
        sql_query = "SELECT " & f_tweet_datetime & ", " & f_tweet_from & ", " & f_tweet_text
            sql_query = sql_query & " FROM " & t_tweet
            sql_query = sql_query & " WHERE " & f_tweet_datetime & ">=" & min_date - 0.5
            sql_query = sql_query & " AND " & f_tweet_datetime & "<=" & max_date + 0.5
        
        extract_tweets = sqlite3_query(db_path_base & db_twitter, sql_query)
        
        If UBound(extract_tweets, 1) > 0 Then
            
            For i = 0 To UBound(extract_tweets(0), 1)
                If extract_tweets(0)(i) = f_tweet_datetime Then
                    dim_extract_tweet_datetime = i
                ElseIf extract_tweets(0)(i) = f_tweet_from Then
                    dim_extract_tweet_from = i
                ElseIf extract_tweets(0)(i) = f_tweet_text Then
                    dim_extract_tweet_text = i
                End If
            Next i
        
            For i = 0 To UBound(list_tweets_to_import, 1)
                    
                For j = 1 To UBound(extract_tweets, 1)
                    
                    If list_tweets_to_import(i)(dim_tweets_import_from) = extract_tweets(j)(dim_extract_tweet_from) And FormatDateTime(FromJulianDay(CDbl(list_tweets_to_import(i)(dim_tweets_import_datetime))), vbShortDate) = FormatDateTime(FromJulianDay(CDbl(extract_tweets(j)(dim_extract_tweet_datetime))), vbShortDate) Then
                        
                        If DistanceDeLevenshtein(list_tweets_to_import(i)(dim_tweets_import_text), extract_tweets(j)(dim_extract_tweet_text), 3) < 3 Then
                            
                            'ne pas prendre le tweet deja present dans la DB
                            list_tweets_to_import(i) = Array(Empty, Empty, Empty)
                            Exit For
                        End If
                        
                    End If
                Next j
                
            Next i
        
        End If
        
        
        'insertion
        m = 0
        Dim final_import() As Variant
        For i = 0 To UBound(list_tweets_to_import, 1)
            If IsEmpty(list_tweets_to_import(i)(dim_tweets_import_text)) = False Then
                
                ReDim Preserve final_import(m)
                final_import(m) = create_tweet(list_tweets_to_import(i)(dim_tweets_import_text), list_tweets_to_import(i)(dim_tweets_import_from), , list_tweets_to_import(i)(dim_tweets_import_datetime))
                m = m + 1
            End If
        Next i
    End If

End If

If m > 0 Then
    load_xml_offline_twitter_into_db = final_import
Else
    load_xml_offline_twitter_into_db = Empty
End If

End Function


Public Function create_tweet(ByVal tweet As String, Optional ByVal user As String = "", Optional ByVal attach_files As Variant, Optional ByVal override_datetime As Variant) As Variant ' retourne *

Dim wrbk_book As String
    wrbk_book = "Kronos.xls"

Dim date_tmp As Date

If tweet = "" Then
    create_tweet = Empty
    Exit Function
End If

If Right(tweet, 1) = " " Then
    tweet = Left(tweet, Len(tweet) - 1)
End If

If user = "" Then
    user = get_username_from_tweet
End If

Dim oBBG As New cls_Bloomberg_Sync
Dim output_bbg As Variant

Dim data_central As Variant

Dim vec_ticker_bbg() As Variant

Dim oJSON As New JSONLib

Dim i As Long, j As Long, k As Long, m As Long, n As Long


If check_mode = internal_kronos_local_mode.online_with_db Then

    Call init_db_twitter
    
    'comme <tweet> est passe en ref, chaque fonction profite d'optimiser la structure
    
    Dim tmp_vec() As Variant
    
    Dim list_tickers As Variant
        list_tickers = get_tickers_from_tweet(tweet)
        
            If IsEmpty(list_tickers) = True Then
                json_list_tickers = Empty
            Else
                Dim vec_list_tickers() As Variant
                ReDim vec_list_tickers(UBound(list_tickers, 1))
                
                For i = 0 To UBound(list_tickers, 1)
                    vec_list_tickers(i) = list_tickers(i)(0)
                Next i
                json_list_tickers = encode_json_for_DB(oJSON.toString(vec_list_tickers))
            End If
            
    Dim list_hashtags As Variant
        list_hashtags = get_hashtags_from_tweet(tweet)
        
            If IsEmpty(list_hashtags) = True Then
                json_list_hashtags = Empty
            Else
                Dim vec_list_hashtags() As Variant
                ReDim vec_list_hashtags(UBound(list_hashtags, 1))
                
                For i = 0 To UBound(list_hashtags, 1)
                    vec_list_hashtags(i) = list_hashtags(i)(0)
                Next i
                json_list_hashtags = encode_json_for_DB(oJSON.toString(vec_list_hashtags))
                
            End If
            
        
    Dim list_mentions As Variant
        list_mentions = get_mentions_from_tweet(tweet)
        
            If IsEmpty(list_mentions) = True Then
                json_list_mentions = Empty
            Else
                Dim vec_list_mentions() As Variant
                ReDim vec_list_mentions(UBound(list_mentions, 1))
                
                For i = 0 To UBound(list_mentions, 1)
                    vec_list_mentions(i) = list_mentions(i)(0)
                Next i
                json_list_mentions = encode_json_for_DB(oJSON.toString(vec_list_mentions))
            End If
        
        
    Dim list_links As Variant
        list_links = get_links_from_tweet(tweet)
            
            If IsEmpty(list_links) = True Then
                json_list_links = Empty
            Else
                Dim vec_list_links() As Variant
                ReDim vec_list_links(UBound(list_links, 1))
                
                For i = 0 To UBound(list_links, 1)
                    vec_list_links(i) = list_links(i)(0)
                Next i
                json_list_links = encode_json_for_DB(oJSON.toString(vec_list_links))
            End If
        
        
    
    Dim tweet_last_entry As Double
        tweet_last_entry = get_last_tweet_id
    
    
    'insertion du tweet dans la DB
    Dim insert_status As Variant
    
    Dim tweet_datetime As Double
    tweet_datetime = ToJulianDay(Now)
    
    If IsMissing(override_datetime) = False Then
        If IsNumeric(override_datetime) Then
            tweet_datetime = CDbl(override_datetime)
        End If
    End If
    
    insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_tweet, Array(Array(tweet_last_entry + 1, tweet_datetime, user, tweet, json_list_tickers, json_list_hashtags, json_list_mentions, json_list_links)), Array(f_tweet_id, f_tweet_datetime, f_tweet_from, f_tweet_text, f_tweet_json_tickers, f_tweet_json_hashtags, f_tweet_json_mentions, f_tweet_json_links))
    
    'si tout c'est bien deroule - attach des files + hyperlinks
    If insert_status = 101 Then
        
        Dim tmp_local_copy As Variant
        
        If IsEmpty(list_links) = False Then
            
            'creation des attachements
            For i = 0 To UBound(list_links, 1)
                tmp_local_copy = get_local_copy(list_links(i)(1))
                
                If IsEmpty(tmp_local_copy) = False Then
                    create_hyperlink_and_file tweet_last_entry + 1, list_links(i)(1), list_links(i)(0), tmp_local_copy
                End If
            Next i
        End If
        
        If IsEmpty(list_tickers) = False Then
            
            For i = 0 To UBound(list_tickers, 1)
                ReDim Preserve vec_ticker_bbg(i)
                vec_ticker_bbg(i) = list_tickers(i)(2)
            Next i
            
            
            'complete la table market_data
            output_bbg = oBBG.bdp(vec_ticker_bbg, Array("PX_LAST", "HIST_PUT_IMP_VOL", "VOLATILITY_30D"), output_format.of_vec_without_header)
            data_central = mount_sqlite_central()
            
            For i = 1 To UBound(data_central, 1)
                data_central(i)(0) = UCase(patch_ticker_marketplace(data_central(i)(0)))
            Next i
            
            
            'insertion with transaction
            Dim tmp_last_price As Double, tmp_impl_vol As Double, tmp_hist_vol As Double, tmp_central_eps As Double, tmp_ticker As String
            k = 0
            Dim vec_data_market_data() As Variant
            For i = 0 To UBound(output_bbg, 1)
                If IsNumeric(output_bbg(i)(0)) Then
                    
                    tmp_ticker = UCase(patch_ticker_marketplace(vec_ticker_bbg(i)))
                    
                    tmp_central_eps = -1
                    
                    For j = 1 To UBound(data_central, 1)
                        If UCase(data_central(j)(0)) = tmp_ticker Then
                            
                            For m = 0 To UBound(data_central(0), 1)
                                If data_central(0)(m) = "Rank_EPS" Then
                                    tmp_central_eps = data_central(j)(m)
                                    Exit For
                                End If
                            Next m
                            
                            Exit For
                        End If
                    Next j
                    
                    tmp_last_price = output_bbg(i)(0)
                    
                    If IsNumeric(output_bbg(i)(1)) Then
                        tmp_impl_vol = Round(output_bbg(i)(1), 2)
                    Else
                        tmp_impl_vol = -1
                    End If
                    
                    If IsNumeric(output_bbg(i)(2)) Then
                        tmp_hist_vol = Round(output_bbg(i)(2), 2)
                    Else
                        tmp_hist_vol = -1
                    End If
                    
                    
                    ReDim Preserve vec_data_market_data(k)
                    vec_data_market_data(k) = Array(list_tickers(i)(0), tweet_datetime, tmp_last_price, tmp_impl_vol, tmp_hist_vol, tmp_central_eps)
                    k = k + 1
                    
                End If
            Next i
            
            Dim check_insert_status_market_data As Variant
            If k > 0 Then
                check_insert_status_market_data = sqlite3_insert_with_transaction(twitter_get_db_path, t_market_data, vec_data_market_data, Array(f_market_data_ticker_twitter, f_market_data_datetime, f_market_data_px_last, f_market_data_impl_vol, f_market_data_histo_vol_30d, f_market_data_central_rank_eps))
            End If
            
            
            c_open_underlying_id = 1
            c_open_product_id = 2
            c_open_description = 7
            c_open_underlying_ticker = 104
            c_open_product_ticker = 105
            
            Dim rng_description As Range
            
            'rajoute un comment dans open si une ligne avec le ticker est trouvee
            For i = 0 To UBound(list_tickers, 1)
                For j = 26 To 3200
                    If Workbooks(wrbk_book).Worksheets("Open").Cells(j, 1) = "" And Workbooks(wrbk_book).Worksheets("Open").Cells(j + 2, 1) = "" And Workbooks(wrbk_book).Worksheets("Open").Cells(j + 3, 1) = "" Then
                        Exit For
                    Else
                        
'                        tmp_ticker = Replace(UCase(list_tickers(i)(2)), " EQUITY", "")
'                        tmp_ticker = Replace(UCase(list_tickers(i)(2)), " INDEX", "")
                        
'                        If InStr(Workbooks(wrbk_book).Worksheets("Open").Cells(j, c_open_underlying_ticker), test) <> 0 Then
'
'                            Exit For
'                        End If
                        
                        If UCase(Workbooks(wrbk_book).Worksheets("Open").Cells(j, c_open_underlying_ticker)) = UCase(list_tickers(i)(2)) Then
                            
                            Set rng_description = Worksheets("Open").Cells(j, c_open_description)
                            
                            If rng_description.comment Is Nothing Then rng_description.AddComment
                            rng_description.comment.Visible = False
                            rng_description.comment.Text user & Chr(10) & tweet & Chr(10) & "at " & FormatDateTime(FromJulianDay(CDbl(tweet_datetime)), vbShortDate) & " " & FormatDateTime(FromJulianDay(CDbl(tweet_datetime)), vbShortTime)
                            rng_description.comment.Shape.Width = 250
                            rng_description.comment.Shape.Height = 75
                            
                            
                            Exit For
                        End If
                        
                    End If
                Next j
            Next i
        End If
        
        If IsMissing(attach_files) = False Then
            If IsEmpty(attach_files) = False Then
                If IsArray(attach_files) = True Then
                    
                    For i = 0 To UBound(attach_files, 1)
                        tmp_local_copy = get_local_copy(attach_files(i))
                
                        If IsEmpty(tmp_local_copy) = False Then
                            create_hyperlink_and_file tweet_last_entry + 1, attach_files(i), Empty, tmp_local_copy
                        End If
                    Next i
                    
                End If
            End If
        End If
        
        
        If IsEmpty(list_mentions) = False Then
            
            k = 0
            Dim list_mention_for_export_table_mention() As Variant
            For i = 0 To UBound(list_mentions, 1)
                ReDim Preserve list_mention_for_export_table_mention(k)
                list_mention_for_export_table_mention(k) = Array(tweet_last_entry + 1, list_mentions(i)(0))
                k = k + 1
            Next i
            
            If k > 0 Then
                insert_mentions_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_mention, list_mention_for_export_table_mention, Array(f_mention_tweet_id, f_mention_target))
            End If
            
        End If
        
        debug_test = tweet_trigger(tweet_last_entry + 1)
        
    End If
    
    
    create_tweet = Array(tweet, list_tickers, list_hashtags, list_mentions, list_links, attach_files)

ElseIf check_mode = internal_kronos_local_mode.offline_with_xml Then
    
    Dim xml_offline_path As String
        xml_offline_path = init_offline_xml
        
        Dim oXML As New DOMDocument
        Dim oXMLRoot As IXMLDOMElement
        Dim oXMLTweet As IXMLDOMElement
        Dim oXMLTweetId As IXMLDOMElement
        Dim oXMLTweetFrom As IXMLDOMElement
        Dim oXMLTweetDatetime As IXMLDOMElement
        Dim oXMLTweetText As IXMLDOMElement
        
        oXML.async = False
        oXML.load (xml_offline_path)
        
        Set oXMLRoot = oXML.documentElement
        
        date_tmp = Now
        
        Dim date_sqlite As Double
        date_sqlite = ToJulianDay(Now)
        
        'creation de la nouvelle entree
        Set oXMLTweet = oXML.createElement(offline_xml_tag_tweet)
        oXMLRoot.appendChild oXMLTweet
        
            'different composants
            Set oXMLTweetId = oXML.createElement(offline_xml_tag_tweet_id)
            oXMLTweet.appendChild oXMLTweetId
            oXMLTweetId.Text = user & "_" & date_sqlite
            
            Set oXMLTweetFrom = oXML.createElement(offline_xml_tag_tweet_from)
            oXMLTweet.appendChild oXMLTweetFrom
            oXMLTweetFrom.Text = user
            
            Set oXMLTweetDatetime = oXML.createElement(offline_xml_tag_tweet_datetime)
            oXMLTweet.appendChild oXMLTweetDatetime
            oXMLTweetDatetime.Text = date_sqlite
            
            Set oXMLTweetText = oXML.createElement(offline_xml_tag_tweet_text)
            oXMLTweet.appendChild oXMLTweetText
            oXMLTweetText.Text = tweet
        
        
        oXML.Save (xml_offline_path)
End If

End Function


Public Function mount_hashtags(Optional ByVal mount_query As Variant) As Variant

Dim sql_query As String
If IsMissing(mount_query) Then
    sql_query = "SELECT " & f_hashtag_id & " FROM " & t_hashtag & " ORDER BY " & f_hashtag_id & " COLLATE NOCASE ASC"
Else
    sql_query = mount_query
End If

t_current_hashtag = sqlite3_query(twitter_get_db_path, sql_query)
mount_hashtags = t_current_hashtag

End Function


Public Function create_hashtag(ByVal hashtag As String) As String

Dim i As Long, j As Long, k As Long

create_hashtag = UCase(hashtag)

Dim insert_status As Variant

If IsEmpty(t_current_hashtag) = True Then
    Dim mount_hashtags_status As Variant
    mount_hashtags_status = mount_hashtags
End If

If UBound(t_current_hashtag, 1) = 0 Then
    '1ere entree
    insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_hashtag, Array(Array(UCase(hashtag))), Array(f_hashtag_id))
    mount_hashtags_status = mount_hashtags
Else
    's'assure que n'existe pas deja
    For i = 1 To UBound(t_current_hashtag, 1)
        If UCase(hashtag) = UCase(t_current_hashtag(i)(0)) Then
            Exit For
        Else
            If i = UBound(t_current_hashtag, 1) Then
                insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_hashtag, Array(Array(UCase(hashtag))), Array(f_hashtag_id))
                mount_hashtags_status = mount_hashtags
            End If
        End If
    Next i
End If

End Function


'si ticker_twitter existe deja mais que ticker bloomberg est different -> remplacer
Public Function create_ticker(ByVal ticker_twitter As String, ByVal ticker_bloomberg As String, Optional ByVal replace_if_different As Boolean = True) As Long

Dim i As Long, j As Long, k As Long

If IsEmpty(t_current_ticker) Then
    Dim mount_ticker_status As Variant
    mount_ticker_status = mount_tickers
End If


If UBound(t_current_ticker, 1) = 0 Then
    '1ere entree
    insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_ticker, Array(Array(UCase(ticker_twitter), UCase(ticker_bloomberg))), Array(f_ticker_twitter, f_ticker_bloomberg))
    mount_ticker_status = mount_tickers
Else
    's'assure que n'existe pas deja
    For i = 1 To UBound(t_current_ticker, 1)
        If UCase(ticker_twitter) = UCase(t_current_ticker(i)(0)) Then
            
'            Dim update_ticker_status As Variant
'
'            'upadte if new
'            If replace_if_different = True Then
'                If UCase(t_current_ticker(i)(1)) <> ticker_bloomberg Then
'                    sql_query = "UPDATE " & t_ticker & " SET " & f_ticker_bloomberg & "=""" & ticker_bloomberg & """ WHERE " & f_ticker_twitter & "=""" & ticker_twitter & """"
'                    update_ticker_status = sqlite3_query(db_path_base & db_twitter, sql_query)
'                    mount_ticker_status = mount_tickers
'                End If
'            End If
            
            Exit For
        Else
            If i = UBound(t_current_ticker, 1) Then
                insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_ticker, Array(Array(UCase(ticker_twitter), UCase(ticker_bloomberg))), Array(f_ticker_twitter, f_ticker_bloomberg))
                mount_ticker_status = mount_tickers
            End If
        End If
    Next i
End If



End Function


Public Function create_hyperlink_and_file(ByVal tweet_id As Long, ByVal source_link As String, Optional ByVal tinyurl As Variant = Empty, Optional ByVal local_copy_path As Variant = Empty) As Long

Dim insert_status As Variant
insert_status = sqlite3_insert_with_transaction(twitter_get_db_path, t_hyperlink_and_file, Array(Array(tweet_id, source_link, tinyurl, local_copy_path)), Array(f_hyperlink_and_file_tweet_id, f_hyperlink_and_file_source, f_hyerplink_and_file_tinyurl, f_hyperlink_and_file_local_copy))

End Function


Public Function get_clean_weblink(ByVal weblink As String) As String

If InStr(weblink, "://") = 0 Then
    weblink = "http://" & weblink
End If

get_clean_weblink = weblink

End Function


Public Function get_tinyurl_from_weblink(ByVal weblink As String) As String

Dim link_base_tinyurl As String
    link_base_tinyurl = "http://tinyurl.com"
Dim link_api_tinyurl As String
    link_api_tinyurl = link_base_tinyurl & "/api-create.php?url="

Dim XMLHttpRequest As New MSXML2.XMLHTTP


If InStr(weblink, link_base_tinyurl) <> 0 Then
    'deja en tinyurl
    get_tinyurl_from_weblink = weblink
Else
    
    XMLHttpRequest.Open "GET", link_api_tinyurl & weblink, False
    XMLHttpRequest.send
    
    If UCase(XMLHttpRequest.statusText) = UCase("OK") Then
        get_tinyurl_from_weblink = XMLHttpRequest.responseText
    Else
        get_tinyurl_from_weblink = weblink
    End If
End If


End Function




Public Function get_clean_ticker_bloomberg(ByVal ticker_raw As String) As String

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

Dim ticker_symbol As String, ticker_marketplace As String, ticker_product As String


ticker_raw = UCase(ticker_raw)
ticker_raw = Replace(ticker_raw, "$", "")
ticker_raw = Replace(ticker_raw, ".", " ")

If Right(ticker_raw, 6) = "EQUITY" Then
    ticker_product = "EQUITY"
ElseIf Right(ticker_raw, 5) = "INDEX" Then
    ticker_product = "INDEX"
Else
    ticker_product = "EQUITY"
End If

Dim symbol As String, marketplace As String


If InStr(ticker_raw, " ") <> 0 Then
    ticker_symbol = Left(ticker_raw, InStr(ticker_raw, " ") - 1)
    
    ticker_marketplace = ticker_raw
    ticker_marketplace = Replace(ticker_marketplace, ticker_symbol & " ", "")
    ticker_marketplace = Left(ticker_marketplace, 2)
Else
    ticker_symbol = ticker_raw
    ticker_marketplace = "US"
End If


If UCase(ticker_product) = "EQUITY" Then
    
    get_clean_ticker_bloomberg = patch_ticker_marketplace(ticker_symbol & " " & ticker_marketplace & " " & ticker_product)
    
    'check si option
    oReg.Pattern = "[A-Za-z0-9]+((\s[A-Za-z]{2}|)\s([\d]{1,2}/|)[\d]{1,2}(/[\d]{1,2}|)\s[c|C|p|P][\d]+(\s[\d]+|))"
    
    Set matches = oReg.Execute(ticker_raw)
    For Each match In matches
        get_clean_ticker_bloomberg = ticker_raw & " " & ticker_product
        Exit Function
    Next
    
    
    
ElseIf UCase(ticker_product) = "INDEX" Then
    get_clean_ticker_bloomberg = ticker_symbol & " " & ticker_product
End If

End Function


Public Function get_ticker_twitter(ByVal ticker_bloomberg As String)

Dim clean_ticker As String
clean_ticker = get_clean_ticker_bloomberg(ticker_bloomberg)

    clean_ticker = Replace(UCase(clean_ticker), " EQUITY", "")
    clean_ticker = Replace(UCase(clean_ticker), " INDEX", "")


get_ticker_twitter = "$" & Replace(clean_ticker, " ", ".")

End Function


Private Function update_tweet(ByVal tweet_id As Variant, ByVal new_text_content As String) As Variant

Dim sql_query As String

sql_query = "UPDATE " & t_tweet & " SET " & f_tweet_text & "=""" & new_text_content & """ WHERE " & f_tweet_id & "=" & tweet_id
update_tweet = sqlite3_query(twitter_get_db_path, sql_query)



End Function


Public Function tweet_trigger(ByVal tweet_id_or_tweet As Variant) As Variant

Dim sql_query As String

Dim oJSON As New JSONLib
Dim colTickers As Collection, colHashtags As Collection, colMentions As Collection, colLinks As Collection, _
    col_tmp_mention As Variant, col_tmp_ticker As Variant, col_tmp_hashtag As Variant, col_tmp_link As Variant


Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

Dim custom_msg As String
Dim vec_ticker() As Variant



If IsNumeric(tweet_id_or_tweet) Then
    'remonte le tweet_txt de la base de donnees
    sql_query = "SELECT * FROM " & t_tweet & " WHERE " & f_tweet_id & "=" & CDbl(tweet_id_or_tweet)
    Dim extract_tweet As Variant
    extract_tweet = sqlite3_query(twitter_get_db_path, sql_query)
    
    If UBound(extract_tweet, 1) = 0 Then
        Exit Function
    Else
        For i = 0 To UBound(extract_tweet(0), 1)
            If extract_tweet(0)(i) = f_tweet_id Then
                dim_extract_id = i
            ElseIf extract_tweet(0)(i) = f_tweet_datetime Then
                dim_extract_datetime = i
            ElseIf extract_tweet(0)(i) = f_tweet_from Then
                dim_extract_from = i
            ElseIf extract_tweet(0)(i) = f_tweet_text Then
                dim_extract_tweet = i
                'tweet_id_or_tweet = extract_tweet(1)(i)
            ElseIf extract_tweet(0)(i) = f_tweet_json_tickers Then
                dim_extract_json_tickers = i
            ElseIf extract_tweet(0)(i) = f_tweet_json_hashtags Then
                dim_extract_json_hashtags = i
            ElseIf extract_tweet(0)(i) = f_tweet_json_mentions Then
                dim_extract_json_mentions = i
            ElseIf extract_tweet(0)(i) = f_tweet_json_links Then
                dim_extract_json_links = i
            End If
        Next i
    End If
Else
    Exit Function 'a dev
End If


'redi plus
test_content = get_specific_tweet_content(Array(f_tweet_json_tickers, f_tweet_json_hashtags, "regexp"), Array("@redi"), , , Array(Array("[\d]+(\.[\d]+|)", 2)), , tweet_id_or_tweet)
If IsEmpty(test_content) = False Then
    dim_ticker = 0
    dim_hashtag = 1
    dim_regexp = 2
    
    'check si hashtags contient #buy / #sell
    For Each tmp_hashtag In test_content(dim_hashtag)(0)
        If InStr(UCase(tmp_hashtag), UCase("buy")) <> 0 Or InStr(UCase(tmp_hashtag), UCase("sell")) <> 0 Then
            
            If IsNull(test_content(dim_ticker)(0)(0)) = False Then
                tmp_ticker = get_clean_ticker_bloomberg(test_content(dim_ticker)(0)(0))
                
                If InStr(UCase(tmp_hashtag), UCase("buy")) <> 0 Then
                    tmp_qty = Abs(CDbl(test_content(dim_regexp)(0)(1).Item(0)))
                ElseIf InStr(UCase(tmp_hashtag), UCase("sell")) <> 0 Then
                    tmp_qty = -Abs(CDbl(test_content(dim_regexp)(0)(1).Item(0)))
                End If
                
                tmp_price = CDbl(test_content(dim_regexp)(0)(1).Item(1))
                
                universal_trades_r_plus Array(Array(tmp_ticker, tmp_qty, tmp_price))
            End If
            
            Exit For
        End If
    Next
    
End If


'tag
test_content = get_specific_tweet_content(Array(f_tweet_json_tickers, f_tweet_json_hashtags), Array("@tag"), , , , , tweet_id_or_tweet)
If IsEmpty(test_content) = False Then
    
    c_equity_db_delta = 6
    c_equity_db_result_ytd = 14
    c_equity_db_json_tag = 137
    
    Dim colTag As Collection
    Dim EBJsonTag As Collection
    Dim EBJsonTagElement As Variant
    
    Dim l_ticker_concern_open_line As Long
    Dim l_ticker_concern_equity_db As Long
    
    Dim found_equity_db_line As Boolean
    Dim found_open_line As Boolean
    
    dim_ticker = 0
    dim_hashtag = 1
    
    If IsEmpty(test_content(dim_ticker)) = False And IsEmpty(test_content(dim_hashtag)) = False Then
        
        Dim vec_tickers() As Variant
        Dim vec_tags() As Variant
        
        Dim json_tag_for_equity_db As String
        Dim str_open_tag As String
        
        'remonte les tickers
        k = 0
        For i = 0 To UBound(test_content(dim_ticker)(0), 1)
            ReDim Preserve vec_tickers(k)
            vec_tickers(k) = get_clean_ticker_bloomberg(test_content(dim_ticker)(0)(i))
            k = k + 1
        Next i
        
        
        'remonte les tags
        k = 0
        For i = 0 To UBound(test_content(dim_hashtag)(0), 1)
            ReDim Preserve vec_tags(k)
            vec_tags(k) = test_content(dim_hashtag)(0)(i)
            k = k + 1
        Next i
        
        
        'mise a jour d'equity database
        found_equity_db_line = False
        found_open_line = False
        
        For i = 0 To UBound(vec_tickers, 1)
            If ActiveSheet.name = "Open" And ActiveCell.row > 25 Then
                
                'tente de prendre directement  la ligne dans CY
                If UCase(patch_ticker_marketplace(Worksheets("Open").Cells(ActiveCell.row, 104))) = UCase(vec_tickers(i)) Then
                    
                    If IsNumeric(Worksheets("Open").Cells(ActiveCell.row, 103)) Then
                        
                        l_ticker_concern_equity_db = Worksheets("Open").Cells(ActiveCell.row, 103)
                        
                        If UCase(patch_ticker_marketplace(Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, 47))) = UCase(vec_tickers(i)) Then
                            l_ticker_concern_open_line = ActiveCell.row
                            
                            found_equity_db_line = True
                            found_open_line = True
                        End If
                        
                    End If
                    
                End If
                
            End If
            
            
            If found_equity_db_line = False Then
                'passage en revue d'equity db
                For j = 27 To 10000 Step 2
                    If Worksheets("Equity_Database").Cells(j, 1) = "" And Worksheets("Equity_Database").Cells(j + 2, 1) = "" Then
                        l_equity_db_last_line = j - 2
                        Exit For
                    Else
                        If UCase(patch_ticker_marketplace(Worksheets("Equity_Database").Cells(j, 47))) = UCase(vec_tickers(i)) Then
                            l_ticker_concern_equity_db = j
                            found_equity_db_line = True
                            Exit For
                        End If
                    End If
                Next j
            End If
            
            If found_equity_db_line = True Then
                
                str_open_tag = ""
                k = 0
                
                'regarde si contient deja des tags
                If Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_json_tag) <> "" Then
                    
                    If IsError(Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_delta)) = False Then
                        If IsNumeric(Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_delta)) Then
                            tmp_delta = Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_delta)
                        Else
                            tmp_delta = 0
                        End If
                    Else
                        tmp_delta = 0
                    End If
                    
                    
                    If IsError(Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_result_ytd)) = False Then
                        If IsNumeric(Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_result_ytd)) Then
                            tmp_result_ytd = Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_result_ytd)
                        Else
                            tmp_result_ytd = 0
                        End If
                    Else
                        tmp_result_ytd = 0
                    End If
                    
                    
                    
                    Set colTag = oJSON.parse(Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_json_tag))
                    
                    If colTag Is Nothing Then
                        
                    Else
                        'tranforme la collection en array
                        Dim tmp_vec_tag() As Variant
                        Dim vec_tags_for_edb() As Variant
                        
                        k = 0
                        
                        For Each EBJsonTag In colTag
                            
                            ReDim Preserve vec_tags_for_edb(k)
                            
                            m = 0
                            For Each EBJsonTagElement In EBJsonTag
                                ReDim Preserve tmp_vec_tag(m)
                                tmp_vec_tag(m) = EBJsonTagElement
                                
                                m = m + 1
                            Next
                            
                            vec_tags_for_edb(k) = tmp_vec_tag
                            k = k + 1
                            
                        Next
                    End If
                End If
                
                
                'ajoute le nouveau tag
                For j = 0 To UBound(vec_tags, 1)
                    ReDim Preserve vec_tags_for_edb(k)
                    vec_tags_for_edb(k) = Array(Date, "OPEN", vec_tags(j), tmp_delta, "OPEN", tmp_result_ytd, "OPEN")
                    k = k + 1
                Next j
                
                
                'transforme en json
                json_tag_for_equity_db = oJSON.toString(vec_tags_for_edb)
                
                'inscrit dans la cellule
                Worksheets("Equity_Database").Cells(l_ticker_concern_equity_db, c_equity_db_json_tag) = json_tag_for_equity_db
                
                
                'mise a jour dans open
                str_open_tag = ""
                k = 0
                For j = 0 To UBound(vec_tags_for_edb, 1)
                    If vec_tags_for_edb(j)(1) = "OPEN" Then
                        k = k + 1
                        
                        If k = 1 Then
                            str_open_tag = vec_tags_for_edb(j)(2)
                        Else
                            str_open_tag = str_open_tag & ", " & vec_tags_for_edb(j)(2)
                        End If
                    End If
                Next j
                
                
                If found_open_line = False Then
                    For j = 26 To 5000
                        If Worksheets("Open").Cells(j, 1) = "" And Worksheets("Open").Cells(j + 1, 1) = "" And Worksheets("Open").Cells(j + 2, 1) = "" Then
                            Exit For
                        Else
                            If UCase(patch_ticker_marketplace(Worksheets("Open").Cells(j, 104))) = UCase(vec_tickers(i)) Then
                                found_open_line = True
                                l_ticker_concern_open_line = j
                                Exit For
                            End If
                        End If
                    Next j
                End If
                
                
                If found_open_line = True Then
                    Worksheets("Open").Cells(l_ticker_concern_open_line, 15) = str_open_tag
                End If
                
            End If
            
            
        Next i
    End If
    
End If



'performance
test_content = get_specific_tweet_content(Array(f_tweet_json_hashtags), Array("@perf", "@portfolio"), , , , , tweet_id_or_tweet)
If IsEmpty(test_content) = False Then
    
    dim_hashtag = 0
    
    If IsEmpty(test_content(dim_hashtag)) = False Then
        
        'second appel pour remonter * les tickers du portfolio mentionne dans le tweet
        For i = 0 To UBound(test_content(dim_hashtag)(0), 1)
            
            extract_tickers_twitter = get_specific_tweet_content(Array(f_tweet_json_tickers), , , Array(test_content(dim_hashtag)(0)(i)))
            
            ReDim vec_ticker(0)
                vec_ticker(0) = ""
            k = 0
            If IsNull(extract_tickers_twitter) = False Then
                
                For j = 0 To UBound(extract_tickers_twitter(0)(0), 1)
                    If IsEmpty(extract_tickers_twitter(0)(0)(j)) = False Then
                        For m = 0 To UBound(vec_ticker, 1)
                            If vec_ticker(m) = get_clean_ticker_bloomberg(extract_tickers_twitter(0)(0)(j)) Then
                                Exit For
                            Else
                                If m = UBound(vec_ticker, 1) Then
                                    ReDim Preserve vec_ticker(k)
                                    vec_ticker(k) = get_clean_ticker_bloomberg(extract_tickers_twitter(0)(0)(j))
                                    k = k + 1
                                End If
                            End If
                        Next m
                    End If
                Next j
                
                
                'appel bdp pour calcul de la perf
                If k > 0 Then
                    Dim output_bdp As Variant
                    output_bdp = bbg_multi_tickers_and_multi_fields(vec_ticker, Array("LAST_PRICE", "CHG_PCT_1D", "REL_1D"))
                    
                    dim_bdp_last_price = 0
                    dim_bdp_chg_pct_1d = 1
                    dim_bdp_chg_rel_1d = 2
                    
                    custom_msg = "@performance @portfolio " & test_content(dim_hashtag)(0)(i) & vbCrLf & vbCrLf
                    
                    Dim sum_perf As Double, sum_rel_perf As Double, count_perf As Integer, count_rel_perf As Integer
                        sum_perf = 0
                        count_perf = 0
                        
                        sum_rel_perf = 0
                        count_rel_perf = 0
                    
                    For j = 0 To UBound(output_bdp, 1)
                        custom_msg = custom_msg & vec_ticker(j) & " " & output_bdp(j, dim_bdp_chg_pct_1d) & vbCrLf
                        
                        If IsNumeric(output_bdp(j, dim_bdp_chg_pct_1d)) Then
                            sum_perf = sum_perf + output_bdp(j, dim_bdp_chg_pct_1d)
                            count_perf = count_perf + 1
                        End If
                        
                        If IsNumeric(output_bdp(j, dim_bdp_chg_rel_1d)) Then
                            sum_rel_perf = sum_rel_perf + output_bdp(j, dim_bdp_chg_rel_1d)
                            count_rel_perf = count_rel_perf + 1
                        End If
                        
                    Next j
                    
                    'portfolio stat
                    custom_msg = custom_msg & vbCrLf
                    If count_perf > 0 Then
                        custom_msg = custom_msg & test_content(dim_hashtag)(0)(i) & "=" & Round(sum_perf / count_perf, 2)
                        
                        'edition du tweet avec la performance
                        debug_test = update_tweet(extract_tweet(1)(dim_extract_id), extract_tweet(1)(dim_extract_tweet) & " perf=" & Round(sum_perf / count_perf, 2))
                    End If
                    
                    MsgBox (custom_msg)
                    
                End If
                
                
            End If
            
        Next i
        
    End If
    
    
End If

'If IsNull(extract_tweet(1)(dim_extract_json_mentions)) = False And IsNull(extract_tweet(1)(dim_extract_json_hashtags)) = False And IsNull(extract_tweet(1)(dim_extract_json_tickers)) = False Then
'
'    Set colTickers = oJSON.parse(decode_json_from_DB(extract_tweet(1)(dim_extract_json_tickers)))
'    Set colHashtags = oJSON.parse(decode_json_from_DB(extract_tweet(1)(dim_extract_json_hashtags)))
'    Set colMentions = oJSON.parse(decode_json_from_DB(extract_tweet(1)(dim_extract_json_mentions)))
'
'    For Each col_tmp_mention In colMentions
'        If col_tmp_mention = "@redi" Then
'
'            For Each col_tmp_hashtag In colHashtags
'                If col_tmp_hashtag = "#BUY" Or col_tmp_hashtag = "#SELL" Then
'                    For Each col_tmp_ticker In colTickers
'
'                        ticker_to_trade = get_clean_ticker_bloomberg(col_tmp_ticker)
'
'                        'extraction de la qty & prix
'                        oReg.Pattern = "[\d]+(\.[\d]+|)"
'                        Set Matches = oReg.Execute(extract_tweet(1)(dim_extract_tweet))
'
'                        k = 0
'                        For Each Match In Matches
'                            If k = 0 Then
'                                qty_to_trade = CDbl(Match.value)
'                                k = k + 1
'                            ElseIf k = 1 Then
'                                price_to_trade = CDbl(Match.value)
'                                k = k + 1
'                            End If
'                        Next
'
'                        If k >= 2 Then
'                            MsgBox ("trade with r+ " & qty_to_trade & " " & ticker_to_trade & "@" & price_to_trade)
'                        End If
'
'                        Exit For
'                    Next
'
'                    Exit For
'                End If
'            Next
'
'            Exit For
'        End If
'    Next
'
'End If

'basket




End Function



Public Function get_local_copy(ByVal filepath_or_hyperlink) As String

CreateFolder (db_path_base & directory_local_copy)

Dim new_file_path As String, new_file_name As String

Dim XMLHttpRequest As New MSXML2.XMLHTTP


If InStr(filepath_or_hyperlink, "\") <> 0 Then
    'fichier local
    If exist_file(filepath_or_hyperlink) Then
    
        Dim file_extension As String
        For i = 1 To Len(filepath_or_hyperlink)
            If Mid(StrReverse(filepath_or_hyperlink), i, 1) = "." Then
                file_extension = Right(filepath_or_hyperlink, i)
                Exit For
            End If
        Next i
        
        new_file_name = "attachements_" & year(Now) & Month(Now) & day(Now) & Hour(Now) & Minute(Now) & Second(Now) & file_extension
        
        FileCopy filepath_or_hyperlink, db_path_base & directory_local_copy & "\" & new_file_name
        
        get_local_copy = db_path_base & directory_local_copy & "\" & new_file_name
        Exit Function
    End If
    
    
ElseIf InStr(filepath_or_hyperlink, "://") <> 0 Then
    'lien web
    XMLHttpRequest.Open "GET", filepath_or_hyperlink, False
    XMLHttpRequest.send
    
    new_file_name = "attachements_" & year(Now) & Month(Now) & day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".html"
    
    FNum = FreeFile()
    
    Open db_path_base & directory_local_copy & "\" & new_file_name For Output As FNum
    Write #FNum, XMLHttpRequest.responseText
    Close #FNum
    
    get_local_copy = db_path_base & directory_local_copy & "\" & new_file_name
    Exit Function
    
End If

End Function


Public Function get_specific_tweet_content(content_to_get As Variant, Optional cond_mentions As Variant, Optional cond_tickers As Variant, Optional cond_hashtags As Variant, Optional cond_regexp As Variant, Optional tweets_from As Variant, Optional tweet_id As Variant, Optional range_date As Variant, Optional has_links As Boolean = False, Optional has_attachements As Boolean = False) As Variant

Dim oJSON As New JSONLib
Dim colTickers As Collection, colHashtags As Collection, colMentions As Collection, colLinks As Collection, colTmp As Collection, _
    col_tmp_mention As Variant, col_tmp_ticker As Variant, col_tmp_hashtag As Variant, col_tmp_link As Variant, col_tmp As Variant

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True


Dim i As Long, j As Long, k As Long, m As Long, n As Long

Dim sql_query As String

sql_query = "SELECT * FROM " & t_tweet
k = 0

If IsMissing(cond_mentions) = False Then
    
    If IsArray(cond_mentions) Then
        
        For i = 0 To UBound(cond_mentions, 1)
        
            If k = 0 Then
                sql_query = sql_query & " WHERE "
                k = k + 1
            Else
                sql_query = sql_query & " AND "
            End If
            
            sql_query = sql_query & f_tweet_json_mentions & " LIKE ""%" & LCase(cond_mentions(i)) & "%"""
        
        Next i
    End If
    
End If



If IsMissing(range_date) = False Then
    Dim date_from As Date, date_to As Date
    
    If UBound(range_date) = 0 Then
        date_from = Date - 365 * 10
        date_to = range_date(0)
    ElseIf UBound(range_date) = 1 Then
        date_from = range_date(0)
        date_to = range_date(1)
    End If
    
    
    If IsArray(range_date) Then
        
        If k = 0 Then
            sql_query = sql_query & " WHERE "
            k = k + 1
        Else
            sql_query = sql_query & " AND "
        End If
        
        sql_query = sql_query & f_tweet_datetime & ">=" & ToJulianDay(CDbl(date_from))
        sql_query = sql_query & " AND " & f_tweet_datetime & "<=" & ToJulianDay(CDbl(date_to))
        
    End If
    
End If


If IsMissing(cond_hashtags) = False Then
    
    If IsArray(cond_hashtags) Then
        
        For i = 0 To UBound(cond_hashtags, 1)
        
            If k = 0 Then
                sql_query = sql_query & " WHERE "
                k = k + 1
            Else
                sql_query = sql_query & " AND "
            End If
            
            sql_query = sql_query & f_tweet_json_hashtags & " LIKE ""%" & UCase(cond_hashtags(i)) & "%"""
        
        Next i
    End If
    
End If


If IsMissing(cond_tickers) = False Then
    
    If IsArray(cond_tickers) Then
        
        For i = 0 To UBound(cond_tickers, 1)
        
            If k = 0 Then
                sql_query = sql_query & " WHERE "
                k = k + 1
            Else
                sql_query = sql_query & " AND "
            End If
            
            sql_query = sql_query & f_tweet_json_tickers & " LIKE ""%" & UCase(cond_tickers(i)) & "%"""
        
        Next i
    End If
    
End If




If IsMissing(tweets_from) = False Then
    
    tweets_from = LCase(tweets_from)
    
    If Left(tweets_from, 1) <> "@" Then
        tweets_from = "@" & tweets_from
    End If
    
    If k = 0 Then
        sql_query = sql_query & " WHERE "
        k = k + 1
    Else
        sql_query = sql_query & " AND "
    End If
            
    sql_query = sql_query & f_tweet_from & " =""" & tweets_from & """"
    
End If



If has_links = True Then
    
    If k = 0 Then
        sql_query = sql_query & " WHERE "
        k = k + 1
    Else
        sql_query = sql_query & " AND "
    End If
    
    sql_query = sql_query & f_tweet_json_links & " IS NOT NULL"
    
End If


If IsMissing(tweet_id) = False Then
    
    If k = 0 Then
        sql_query = sql_query & " WHERE "
        k = k + 1
    Else
        sql_query = sql_query & " AND "
    End If
    
    sql_query = sql_query & f_tweet_id & " =" & tweet_id & ""
    
End If



Dim extract_tweet As Variant
extract_tweet = sqlite3_query(twitter_get_db_path, sql_query)

If UBound(extract_tweet, 1) = 0 Then
    get_specific_tweet_content = Empty
    Exit Function
End If

Dim filter_tweet() As Variant
ReDim filter_tweet(UBound(extract_tweet, 1) - 1)

For i = 1 To UBound(extract_tweet, 1)
    filter_tweet(i - 1) = extract_tweet(i)
Next i


'dimension du tweet
For i = 0 To UBound(extract_tweet(0), 1)
    If extract_tweet(0)(i) = f_tweet_text Then
        dim_tweet_text = i
        'Exit For
    ElseIf extract_tweet(0)(i) = f_tweet_json_links Then
        dim_tweet_json_links = i
    ElseIf extract_tweet(0)(i) = f_tweet_id Then
        dim_tweet_id = i
    End If
Next i



If IsMissing(cond_regexp) = False Then
    
    If IsArray(cond_regexp) Then
    
        For i = 0 To UBound(cond_regexp, 1)
            If UBound(cond_regexp(i), 1) = 0 Then
                cond_regexp(i) = Array(cond_regexp(i)(0), 1)
                count_cond_regexp = 1
            Else
                count_cond_regexp = cond_regexp(i)(1)
            End If
            
            oReg.Pattern = cond_regexp(i)(0)
            oReg.IgnoreCase = True
            
            Dim vec_tweet_row_cond_regexp() As Variant
            Dim tmp_tweet As String
            v = 0
            For j = 1 To UBound(extract_tweet, 1)
                
                'si des liens tinyurl sont presents, les retirer pour eviter des vrai/faux
                If IsNull(extract_tweet(j)(dim_tweet_json_links)) Then
                    tmp_tweet = extract_tweet(j)(dim_tweet_text)
                Else
                    Set colLinks = oJSON.parse(decode_json_from_DB(extract_tweet(j)(dim_tweet_json_links)))
                    
                    For Each col_tmp_link In colLinks
                        tmp_tweet = Replace(extract_tweet(j)(dim_tweet_text), col_tmp_link, "")
                    Next
                End If
                
                u = 0
                'Set Matches = oReg.Execute(extract_tweet(j)(dim_tweet_text))
                Set matches = oReg.Execute(tmp_tweet)
                For Each match In matches
                    u = u + 1
                Next
                
                If u >= count_cond_regexp Then
                    ReDim Preserve vec_tweet_row_cond_regexp(v)
                    vec_tweet_row_cond_regexp(v) = Array(j, matches)
                    v = v + 1
                End If
                
            Next j
            
            If v = 0 Then
                cond_regexp(i)(1) = Empty
            Else
                cond_regexp(i)(1) = vec_tweet_row_cond_regexp
            End If
            
        Next i
        
        
        'fusion
        Dim vec_line_cond_regexp() As Variant, tmp_vec_var() As Variant
        u = 0
        For i = 0 To UBound(cond_regexp, 1)
            If IsEmpty(cond_regexp(i)(1)) Then
                ReDim Preserve vec_line_cond_regexp(0)
                vec_line_cond_regexp(0) = Empty
                Exit For
            Else
                If i = 0 Then
                    ReDim Preserve vec_line_cond_regexp(UBound(cond_regexp(i)(1), 1))
                    For j = 0 To UBound(cond_regexp(i)(1), 1)
                        vec_line_cond_regexp(j) = cond_regexp(i)(1)(j)
                    Next j
                Else
                    v = 0
                    For j = 0 To UBound(cond_regexp(i)(1), 1)
                        For m = 0 To UBound(vec_line_cond_regexp, 1)
                            If cond_regexp(i)(1)(j)(0) = vec_line_cond_regexp(m)(0) Then
                                ReDim Preserve tmp_vec_var(v)
                                tmp_vec_var(v) = cond_regexp(i)(1)(j)
                                v = v + 1
                            End If
                        Next m
                    Next j
                    
                    If v = 0 Then
                        ReDim Preserve vec_line_cond_regexp(0)
                        vec_line_cond_regexp(0) = Empty
                        Exit For
                    Else
                        vec_line_cond_regexp = tmp_vec_var
                    End If
                    
                End If
            End If
        Next i
        
        If IsEmpty(vec_line_cond_regexp(0)) Then
            '0 Match
            get_specific_tweet_content = Empty
            Exit Function
        Else
            'ajuste filter tweet
            ReDim filter_tweet(UBound(vec_line_cond_regexp, 1))
            For i = 0 To UBound(vec_line_cond_regexp, 1)
                filter_tweet(i) = extract_tweet(vec_line_cond_regexp(i)(0))
            Next i
        End If
    
    End If
    
End If





'check l'array d'ouput est retourne les donnees
Dim final_output() As Variant
Dim tmp_vec_var_sub_element() As Variant

Dim extract_attachements As Variant
extract_attachements = Empty

For i = 0 To UBound(content_to_get, 1)
    
    ReDim Preserve final_output(i)
    
    'field de la table tweet
    For j = 0 To UBound(extract_tweet(0), 1)
        If content_to_get(i) = extract_tweet(0)(j) Then
            
            'filters
            ReDim tmp_vec_var(UBound(filter_tweet, 1))
            For m = 0 To UBound(filter_tweet, 1)
                
                If InStr(UCase(extract_tweet(0)(j)), UCase("json")) <> 0 Then
                    If IsNull(filter_tweet(m)(j)) Then
                        ReDim tmp_vec_var_sub_element(0)
                        tmp_vec_var_sub_element(0) = Empty
                    Else
                        Set colTmp = oJSON.parse(decode_json_from_DB(filter_tweet(m)(j)))
                        v = 0
                        For Each col_tmp In colTmp
                            ReDim Preserve tmp_vec_var_sub_element(v)
                            tmp_vec_var_sub_element(v) = col_tmp
                            v = v + 1
                        Next
                    End If
                Else
                    ReDim tmp_vec_var_sub_element(0)
                    tmp_vec_var_sub_element(0) = filter_tweet(m)(j)
                End If
                
                ReDim Preserve tmp_vec_var(m)
                tmp_vec_var(m) = tmp_vec_var_sub_element
                
            Next m
            
            final_output(i) = tmp_vec_var
        End If
    Next j
    
    
    'GESTION DES CUSTOMS
    If content_to_get(i) = "regexp" Then
        
        ReDim tmp_vec_var(UBound(filter_tweet, 1))
        For m = 0 To UBound(filter_tweet, 1)
            
            ReDim Preserve tmp_vec_var(m)
            tmp_vec_var(m) = vec_line_cond_regexp(m)
            
        Next m
        
        final_output(i) = tmp_vec_var
    
    ElseIf InStr(UCase(content_to_get(i)), UCase("attach")) <> 0 Then
        
        sql_query = "SELECT " & f_hyperlink_and_file_tweet_id & ", " & f_hyperlink_and_file_source & ", " & f_hyerplink_and_file_tinyurl & ", " & f_hyperlink_and_file_local_copy
            sql_query = sql_query & " FROM " & t_hyperlink_and_file
            sql_query = sql_query & " ORDER BY " & f_hyperlink_and_file_tweet_id & " ASC"
        
        
        For n = 0 To UBound(filter_tweet, 1)
            
            If IsEmpty(extract_attachements) Then
                extract_attachements = sqlite3_query(twitter_get_db_path, sql_query)
            End If
            
            Dim vec_attatchements() As Variant
            Dim count_attachements As Integer
            count_attachements = 0
            If UBound(extract_attachements, 1) > 0 Then
                For m = 1 To UBound(extract_attachements, 1)
                    If filter_tweet(n)(dim_tweet_id) > extract_attachements(m)(0) Then
                        Exit For
                    Else
                        
                        If filter_tweet(n)(dim_tweet_id) = extract_attachements(m)(0) Then
                            ReDim Preserve vec_attatchements(count_attachements)
                            vec_attatchements(count_attachements) = extract_attachements(m)
                            count_attachements = count_attachements + 1
                        End If
                    End If
                Next m
            End If
            
            ReDim Preserve tmp_vec_var(n)
            If count_attachements > 0 Then
                tmp_vec_var(n) = vec_attatchements
            Else
                tmp_vec_var(n) = Empty
            End If
        
        Next n
        
        final_output(i) = tmp_vec_var
        
    End If
    
    
Next i

get_specific_tweet_content = final_output

End Function


Public Function tweet_autocomplete(seach_str As String) As Variant

tweet_autocomplete = Empty

Dim return_data() As Variant

If check_mode = internal_kronos_local_mode.online_with_db Then

    If Left(seach_str, 1) = "#" Then
        
        If Len(seach_str) = 1 Then
            mount_hashtags
        Else
            mount_hashtags ("SELECT " & f_hashtag_id & " FROM " & t_hashtag & " WHERE " & f_hashtag_id & " LIKE ""%" & seach_str & "%""")
        End If
        
        If UBound(t_current_hashtag, 1) = 0 Then
            tweet_autocomplete = Empty
            Exit Function
        Else
            ReDim Preserve return_data(UBound(t_current_hashtag, 1) - 1)
            For i = 1 To UBound(t_current_hashtag, 1)
                return_data(i - 1) = t_current_hashtag(i)(0)
            Next i
            
            tweet_autocomplete = return_data
            Exit Function
        End If
        
    
    ElseIf Left(seach_str, 1) = "@" Then
        
        If Len(seach_str) = 1 Then
            mount_usernames
        Else
             mount_usernames ("SELECT " & f_user_id & " FROM " & t_user & " WHERE " & f_user_id & " LIKE ""%" & seach_str & "%""")
        End If
        
        If UBound(t_current_username, 1) = 0 Then
            tweet_autocomplete = Empty
            Exit Function
        Else
            ReDim Preserve return_data(UBound(t_current_username, 1) - 1)
            For i = 1 To UBound(t_current_username, 1)
                return_data(i - 1) = t_current_username(i)(0)
            Next i
            
            tweet_autocomplete = return_data
            Exit Function
        End If
        
    ElseIf Left(seach_str, 1) = "$" Then
        
        If Len(seach_str) = 1 Then
            mount_tickers
        Else
             mount_tickers ("SELECT " & f_ticker_twitter & " FROM " & t_ticker & " WHERE " & f_ticker_twitter & " LIKE ""%" & seach_str & "%""")
        End If
        
        If UBound(t_current_ticker, 1) = 0 Then
            tweet_autocomplete = Empty
            Exit Function
        Else
            ReDim Preserve return_data(UBound(t_current_ticker, 1) - 1)
            For i = 1 To UBound(t_current_ticker, 1)
                return_data(i - 1) = t_current_ticker(i)(0)
            Next i
            
            tweet_autocomplete = return_data
            Exit Function
        End If
        
    End If

End If

End Function



Public Function get_last_tweets(ByVal nbre As Long) As Variant

get_last_tweets = Empty


If check_mode = internal_kronos_local_mode.online_with_db Then

    Dim sql_query As String
    sql_query = "SELECT * FROM " & t_tweet & " ORDER BY " & f_tweet_datetime & " DESC LIMIT " & nbre
    Dim extract_tweets As Variant
    extract_tweets = sqlite3_query(twitter_get_db_path, sql_query)
    
    
    'extract les attachements
    sql_query = "SELECT " & f_hyperlink_and_file_tweet_id & ", " & f_hyperlink_and_file_source & ", " & f_hyerplink_and_file_tinyurl & ", " & f_hyperlink_and_file_local_copy
        sql_query = sql_query & " FROM " & t_hyperlink_and_file
        sql_query = sql_query & " WHERE " & f_hyperlink_and_file_tweet_id & " IN (SELECT " & f_tweet_id & " FROM " & t_tweet & " ORDER BY " & f_tweet_datetime & " DESC LIMIT " & nbre & ")"
        sql_query = sql_query & " ORDER BY " & f_hyperlink_and_file_tweet_id & " ASC"
    
    Dim extract_tweets_hyperlinks_and_file As Variant
    extract_tweets_hyperlinks_and_file = sqlite3_query(twitter_get_db_path, sql_query)
    
    
    
    
    
    
    Dim vec_tweets() As Variant
    Dim tmp_vec() As Variant
    
    Dim vec_attachement() As Variant
    
    k = 0
    If UBound(extract_tweets, 1) = 0 Then
        get_last_tweets = Empty
    Else
        
        'detection des dimensions
        For i = 0 To UBound(extract_tweets(0), 1)
            If extract_tweets(0)(i) = f_tweet_id Then
                dim_tweet_id = i
            End If
        Next i
        
        For i = 0 To nbre 'y compris headers
            
            n = 0
            For m = 0 To UBound(extract_tweets(i), 1)
                ReDim Preserve tmp_vec(n)
                tmp_vec(n) = extract_tweets(i)(m)
                n = n + 1
            Next m
            
            
            'attachements
            u = 0
            For m = 1 To UBound(extract_tweets_hyperlinks_and_file, 1)
                If extract_tweets_hyperlinks_and_file(m)(0) > extract_tweets(i)(dim_tweet_id) Then
                    Exit For
                Else
                    If extract_tweets_hyperlinks_and_file(m)(0) = extract_tweets(i)(dim_tweet_id) Then
                        ReDim Preserve vec_attachement(u)
                        vec_attachement(u) = extract_tweets_hyperlinks_and_file(m)
                        u = u + 1
                    End If
                End If
            Next m
            
            ReDim Preserve tmp_vec(n)
            
            If u > 0 Then
                tmp_vec(n) = vec_attachement
            Else
                If i = 0 Then
                    'header
                    tmp_vec(n) = "attachements"
                Else
                    tmp_vec(n) = Empty
                End If
            End If
            
            n = n + 1
            
            ReDim Preserve vec_tweets(k)
            vec_tweets(k) = tmp_vec
            k = k + 1
        Next i
    End If
    
    
    If k > 0 Then
        get_last_tweets = vec_tweets
    End If

End If

End Function


Public Function URLEncode(ByVal data_to_encode As String, Optional SpaceAsPlus As Boolean = False) As String


'URLEncode("!*'();:@&=+$,/?%#[]"))
'URLEncode("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"))

URLEncode = ""

Dim result() As Variant

If Len(data_to_encode) > 0 Then
    ReDim result(Len(data_to_encode))
    
    Dim i As Long
    Dim space As String
    
    If SpaceAsPlus Then
        space = "+"
    Else
        space = "%20"
    End If
    
    For i = 1 To Len(data_to_encode)
        
        Select Case Asc(Mid$(data_to_encode, i, 1))
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 61, 95, 129
                result(i) = Mid$(data_to_encode, i, 1)
            Case 32
                result(i) = space
            Case 0 To 15
                result(i) = "%0" & Hex(Asc(Mid$(data_to_encode, i, 1)))
            Case Else
                result(i) = "%" & Hex(Asc(Mid$(data_to_encode, i, 1)))
        End Select
    Next i
    
    URLEncode = Join(result, "")
    
End If

End Function


'contenu possible'
    ' created_at
    ' from_user
    ' from_user_id
    ' from_user_id_str
    ' from_user_id_name
    ' geo
    ' id
    ' id_str
    ' iso_language_code
    ' metadata
    ' profile_image_url
    ' source
    ' text
    ' to_user
    ' to_user_id
    ' to_user_id_str
    ' to_user_name
Public Function get_twitter_content(ByVal which_content As Variant, ByVal query As String) As Variant

Dim oJSON As New JSONLib

If UCase(Left(query, 2)) <> UCase("q=") Then
    query = "q=" & query
End If

Dim url As String
    url = "http://search.twitter.com/search.json?" & URLEncode(query)

Dim XMLHttpRequest As MSXML2.XMLHTTP
Set XMLHttpRequest = New MSXML2.XMLHTTP

XMLHttpRequest.Open "GET", url, False
XMLHttpRequest.send

Set api_output = CreateObject("Scripting.Dictionary")
Set api_output = oJSON.parse(XMLHttpRequest.responseText)

Set api_output_tweets = api_output.Item("results")


Dim output() As Variant

Dim tmp_vec_return_data() As Variant
ReDim Preserve tmp_vec_return_data(UBound(which_content))

k = 0
For Each tweet In api_output_tweets
    For j = 0 To UBound(which_content, 1)
        tmp_vec_return_data(j) = tweet.Item(which_content(j))
    Next j
    
    ReDim Preserve output(k)
    output(k) = tmp_vec_return_data
    k = k + 1
Next

If k = 0 Then
    get_twitter_content = Empty
Else
    get_twitter_content = output
End If

End Function


'weekly or daily
Public Function get_twitter_trends(timebase As String, Optional ByVal nbre_block As Integer = 0) As Variant

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection

oReg.Global = True
oReg.IgnoreCase = True

Dim oJSON As New JSONLib

Dim url As String
    url = "http://api.twitter.com/1/trends/" & LCase(timebase) & ".json"

Dim XMLHttpRequest As MSXML2.XMLHTTP
Set XMLHttpRequest = New MSXML2.XMLHTTP

XMLHttpRequest.Open "GET", url, False
XMLHttpRequest.send

Dim output_web As String
    output_web = XMLHttpRequest.responseText

'patch des key sous forme de date pour eviter bug parsing de la JSONLib
oReg.Pattern = """[\d]{4}-[\d]{2}-[\d]{2}\s[\d]{2}:[\d]{2}"""
Set matches = oReg.Execute(output_web)

For Each match In matches
    output_web = Replace(output_web, match.Value, Left(match.Value, 11) & "_" & Mid(match.Value, 13, 2) & "h" & Right(match.Value, 3))
Next

Set api_output = CreateObject("Scripting.Dictionary")


If InStr(output_web, "Rate limit exceeded. Clients may not make more than 150 requests per hour.") = 0 Then
    
    Set api_output = oJSON.parse(output_web)
    Set api_output_trends = api_output.Item("trends")
    
    
    m = 0
    k = 0
    Dim vec_block() As Variant
    Dim vec_trends() As Variant
    For Each twitter_trend_hour In api_output_trends
        
        If nbre_block = 0 Or m < nbre_block Then
            
            k = 0
            For Each twitter_trend_hour_trend In api_output_trends.Item(twitter_trend_hour)
                ReDim Preserve vec_trends(k)
                vec_trends(k) = twitter_trend_hour_trend.Item("name")
                k = k + 1
            Next
        Else
            Exit For
        End If
        
        ReDim Preserve vec_block(m)
        vec_block(m) = Array(twitter_trend_hour, vec_trends)
        
        m = m + 1
    Next
End If


If k > 0 Then
    get_twitter_trends = vec_block
Else
    get_twitter_trends = Empty
End If

End Function


Public Sub refresh_last_tweet_in_open()

Dim oJSON As New JSONLib
Dim colTicker As Collection, tmp_ticker As Variant

Dim i As Long, j As Long, k As Long, m As Long, n As Long

Dim sql_query As String
sql_query = "SELECT " & f_tweet_id & ", " & f_tweet_datetime & ", " & f_tweet_from & ", " & f_tweet_text & ", " & f_tweet_json_tickers
    sql_query = sql_query & " FROM " & t_tweet
    sql_query = sql_query & " WHERE " & f_tweet_json_tickers & " IS NOT NULL"
    sql_query = sql_query & " ORDER BY " & f_tweet_datetime & " DESC"

Dim extract_tweets As Variant
extract_tweets = sqlite3_query(twitter_get_db_path, sql_query)


    'detection des dim
    For i = 0 To UBound(extract_tweets(0), 1)
        If extract_tweets(0)(i) = f_tweet_id Then
            dim_tweet_id = i
        ElseIf extract_tweets(0)(i) = f_tweet_datetime Then
            dim_tweet_datetime = i
        ElseIf extract_tweets(0)(i) = f_tweet_from Then
            dim_tweet_from = i
        ElseIf extract_tweets(0)(i) = f_tweet_text Then
            dim_tweet_text = i
        ElseIf extract_tweets(0)(i) = f_tweet_json_tickers Then
            dim_tweet_json_tickers = i
        End If
    Next i

Dim vec_distinct_tickers() As Variant
ReDim Preserve vec_distinct_tickers(0)
vec_distinct_tickers(0) = Array("")
    dim_last_tweet_ticker = 0
    dim_last_tweet_text = 1
    dim_last_tweet_date = 2
    dim_last_tweet_from = 3
    
k = 0
If UBound(extract_tweets, 1) > 0 Then
    For i = 1 To UBound(extract_tweets, 1)
        Set colTicker = oJSON.parse(decode_json_from_DB(extract_tweets(i)(dim_tweet_json_tickers)))
        
        For Each tmp_ticker In colTicker
            
            tmp_ticker = get_clean_ticker_bloomberg(tmp_ticker)
            
            For m = 0 To UBound(vec_distinct_tickers, 1)
                If vec_distinct_tickers(m)(0) = tmp_ticker Then
                    Exit For
                Else
                    If m = UBound(vec_distinct_tickers, 1) Then
                        ReDim Preserve vec_distinct_tickers(k)
                        vec_distinct_tickers(k) = Array(tmp_ticker, extract_tweets(i)(dim_tweet_text), extract_tweets(i)(dim_tweet_datetime), extract_tweets(i)(dim_tweet_from))
                        k = k + 1
                    End If
                End If
            Next m
            
        Next
        
    Next i
End If



'passe en revue les ligne d'open
'remonte les ticker + ligne d'open
k = 0
Dim vec_open_ticker_line() As Variant
tmp_ticker = ""
For i = 26 To 3000
    If Worksheets("Open").Cells(i, 1) = "" And Worksheets("Open").Cells(i + 1, 1) = "" And Worksheets("Open").Cells(i + 2, 1) = "" Then
        Exit For
    Else
        If UCase(Worksheets("Open").Cells(i, 104)) <> tmp_ticker Then
            tmp_ticker = UCase(Worksheets("Open").Cells(i, 104))
            
            ReDim Preserve vec_open_ticker_line(k)
            vec_open_ticker_line(k) = Array(UCase(Worksheets("Open").Cells(i, 104)), i)
            k = k + 1
            
        End If
    End If
Next i


'matching + affichage
Dim rng_description As Range
c_open_description = 7

For i = 0 To UBound(vec_distinct_tickers, 1)
    For j = 0 To UBound(vec_open_ticker_line, 1)
        If vec_distinct_tickers(i)(dim_last_tweet_ticker) = vec_open_ticker_line(j)(0) Then
            
            
            Set rng_description = Worksheets("Open").Cells(vec_open_ticker_line(j)(1), c_open_description)
            
            If rng_description.comment Is Nothing Then rng_description.AddComment
            rng_description.comment.Visible = False
            rng_description.comment.Text vec_distinct_tickers(i)(dim_last_tweet_from) & Chr(10) & vec_distinct_tickers(i)(dim_last_tweet_text) & Chr(10) & "at " & FormatDateTime(FromJulianDay(CDbl(vec_distinct_tickers(i)(dim_last_tweet_date))), vbShortDate) & " " & FormatDateTime(FromJulianDay(CDbl(vec_distinct_tickers(i)(dim_last_tweet_date))), vbShortTime)
            rng_description.comment.Shape.Width = 250
            rng_description.comment.Shape.Height = 75
            
            Exit For
        End If
    Next j
Next i


End Sub


Public Sub load_related_tweets_from_activecell_into_form()

Dim search_terms As String

If ActiveSheet.name = "Open" And ActiveCell.row > 25 Then
    search_terms = get_ticker_twitter(Worksheets("Open").Cells(ActiveCell.row, 104))
Else
    search_terms = ActiveCell.Value
    
    If InStr(UCase(search_terms), "EQUITY") <> 0 Then
        search_terms = get_ticker_twitter(search_terms)
    End If
End If


If search_terms = "" Then
    Exit Sub
End If


frm_Tweet_new.Height = 609

frm_Tweet_new.TB_search = search_terms

Call search_interal_db_and_twitter_in_form

frm_Tweet_new.Show


End Sub



Private Sub search_interal_db_and_twitter_in_form()

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
                .Add , , "tweet", 230
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
                .ListItems(i).ListSubItems.Add , "tweet_" & CStr(list_last_tweets(i)(dim_tweet_id)), list_last_tweets(i)(dim_tweet_text) 'tweet
                
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
            .ListItems(i + 1).ListSubItems.Add , "tweet_" & CStr(extract_search_tweets(dim_tweet_search_id)(i)(0)), extract_search_tweets(dim_tweet_search_text)(i)(0) 'tweet
            
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




Public Function show_tweet_with_infos(ByVal tweet_id As Integer) As String

show_tweet_with_infos = ""

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

Dim sql_query As String

Dim date_dbl As Double

Dim extract_tweet As Variant
sql_query = "SELECT * FROM " & t_tweet & " WHERE " & f_tweet_id & "=" & tweet_id
extract_tweet = sqlite3_query(twitter_get_db_path, sql_query)

Dim tmp_tweet As String

If UBound(extract_tweet) > 0 Then

    'detection des dim
    For i = 0 To UBound(extract_tweet(0), 1)
        If extract_tweet(0)(i) = f_tweet_id Then
            dim_tweet_id = i
        ElseIf extract_tweet(0)(i) = f_tweet_datetime Then
            dim_tweet_datetime = i
        ElseIf extract_tweet(0)(i) = f_tweet_text Then
            dim_tweet_text = i
        ElseIf extract_tweet(0)(i) = f_tweet_json_tickers Then
            dim_tweet_json_tickers = i
        End If
    Next i
    
    show_tweet_with_infos = extract_tweet(1)(dim_tweet_text)
    
    If IsNull(extract_tweet(1)(dim_tweet_json_tickers)) = False Then
        sql_query = "SELECT * FROM " & t_market_data & " WHERE " & f_market_data_datetime & ">=" & extract_tweet(1)(dim_tweet_datetime) - 0.0000001 & " AND " & f_market_data_datetime & "<=" & extract_tweet(1)(dim_tweet_datetime) + 0.0000001
        Dim extract_market_data As Variant
        extract_market_data = sqlite3_query(twitter_get_db_path, sql_query)
        
        If UBound(extract_market_data, 1) > 0 Then
        'detection des dim
            For i = 0 To UBound(extract_market_data(0), 1)
                If extract_market_data(0)(i) = f_market_data_ticker_twitter Then
                    dim_market_data_ticker = i
                ElseIf extract_market_data(0)(i) = f_market_data_px_last Then
                    dim_market_data_px_last = i
                ElseIf extract_market_data(0)(i) = f_market_data_impl_vol Then
                    dim_market_data_impl_vol = i
                ElseIf extract_market_data(0)(i) = f_market_data_histo_vol_30d Then
                    dim_market_data_histo_vol_30d = i
                ElseIf extract_market_data(0)(i) = f_market_data_central_rank_eps Then
                    dim_market_data_central_rank_eps = i
                End If
            Next i
            
            Dim tmp_ticker_with_info As String
            
            
            For i = 1 To UBound(extract_market_data, 1)
                count_info = 0
                tmp_ticker_with_info = extract_market_data(i)(dim_market_data_ticker) & " "
                
                If extract_market_data(i)(dim_market_data_px_last) <> -1 Then
                    If count_info = 0 Then
                        tmp_ticker_with_info = tmp_ticker_with_info & "["
                    Else
                        tmp_ticker_with_info = tmp_ticker_with_info & ", "
                    End If
                    
                    tmp_ticker_with_info = tmp_ticker_with_info & "P=" & extract_market_data(i)(dim_market_data_px_last)
                    count_info = count_info + 1
                End If
                
                If extract_market_data(i)(dim_market_data_impl_vol) <> -1 Then
                    If count_info = 0 Then
                        tmp_ticker_with_info = tmp_ticker_with_info & "["
                    Else
                        tmp_ticker_with_info = tmp_ticker_with_info & ", "
                    End If
                    
                    tmp_ticker_with_info = tmp_ticker_with_info & "IV=" & Round(extract_market_data(i)(dim_market_data_impl_vol), 2)
                    count_info = count_info + 1
                End If
                
'                If extract_market_data(i)(dim_market_data_histo_vol_30d) <> -1 Then
'                    If count_info = 0 Then
'                        tmp_ticker_with_info = tmp_ticker_with_info & "["
'                    Else
'                        tmp_ticker_with_info = tmp_ticker_with_info & ", "
'                    End If
'
'                    tmp_ticker_with_info = tmp_ticker_with_info & "HV=" & Round(extract_market_data(i)(dim_market_data_histo_vol_30d), 2)
'                    count_info = count_info + 1
'                End If
                
                If extract_market_data(i)(dim_market_data_central_rank_eps) <> -1 Then
                    If count_info = 0 Then
                        tmp_ticker_with_info = tmp_ticker_with_info & "["
                    Else
                        tmp_ticker_with_info = tmp_ticker_with_info & ", "
                    End If
                    
                    tmp_ticker_with_info = tmp_ticker_with_info & "EPS=" & extract_market_data(i)(dim_market_data_central_rank_eps)
                    count_info = count_info + 1
                End If
                
                
                If count_info > 0 Then
                    tmp_ticker_with_info = tmp_ticker_with_info & "]"
                    
                    show_tweet_with_infos = Replace(show_tweet_with_infos, extract_market_data(i)(dim_market_data_ticker), tmp_ticker_with_info)
                End If
                
            Next i
        End If
    End If
Else
    show_tweet_with_infos = ""
End If


End Function



Private Sub manipulate_db()

Dim sql_query As String
Dim exec_query As Variant
Dim date_tmp As Date

sql_query = "SELECT * FROM " & t_market_data
extract_market_data = sqlite3_query(twitter_get_db_path, sql_query)


sql_query = "SELECT * FROM " & t_tweet & " ORDER BY " & f_tweet_id & " DESC"
extract_tweeets = sqlite3_query(twitter_get_db_path, sql_query)

'sql_query = "DELETE FROM " & t_hyperlink_and_file & " WHERE " & f_hyperlink_and_file_tweet_id & "=89"
'exec_query = sqlite3_query(twitter_get_db_path, sql_query)
'
'sql_query = "DELETE FROM " & t_tweet & " WHERE " & f_tweet_id & "=89"
'exec_query = sqlite3_query(twitter_get_db_path, sql_query)

'sql_query = "DELETE FROM " & t_hyperlink_and_file & " WHERE " & f_hyperlink_and_file_source & "=""0700.hk"""
'exec_query = sqlite3_query(twitter_get_db_path, sql_query)

'date_tmp = Now() - 0.08097
'sql_query = "UPDATE " & t_tweet & " SET " & f_tweet_datetime & "=" & ToJulianDay(date_tmp) & " WHERE " & f_tweet_id & "=39"
'exec_query = sqlite3_query(twitter_get_db_path, sql_query)
'
'sql_query = "UPDATE " & t_tweet & " SET " & f_tweet_from & "=""@jstouff""" & " WHERE " & f_tweet_id & "=39"
'exec_query = sqlite3_query(twitter_get_db_path, sql_query)

'sql_query = "UPDATE " & t_hyperlink_and_file & " SET " & f_hyperlink_and_file_tweet_id & "=39 WHERE " & f_hyperlink_and_file_tweet_id & "=42"
'exec_query = sqlite3_query(twitter_get_db_path, sql_query)

End Sub
