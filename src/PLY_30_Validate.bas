Attribute VB_Name = "PLY_30_Validate"
Option Explicit

'==================================================
' SCRIPT構造チェック
'
' <SCRIPT>  と  </SCRIPT> が
' それぞれ1個ずつ存在することを確認する。
'
' <SCRIPT> は制御タグなので、
' Playerの表示対象にはしない。
'
' ScriptStartLine / ScriptEndLine に
' SCRIPTの境界行を保存する。
'==================================================

'==================================================
' SCRIPT構造チェック
'
' <SCRIPT> / </SCRIPT> は
' 行頭の空白・タブを除いた位置から
' 開始している必要がある。
'
' 例
'
'   <SCRIPT>          OK
'    <SCRIPT>         OK
'   [TAB]<SCRIPT>     OK
'
'   <SCRIPT>hello     NG
'   hello <SCRIPT>    NG
'
' SCRIPTタグの後ろには
' 何も置かない。
'==================================================
Function ValidateScriptStructure() As Boolean

    Dim LineNo As Long
    Dim SourceText As String

    Dim StartPos As Long
    Dim EndPos As Long


    ValidateScriptStructure = False


    '==================================
    ' 初期化
    '==================================
    ScriptStartLine = 0
    ScriptEndLine = 0


    '==================================
    ' スクリプトを走査
    '==================================
    For LineNo = 1 To TextLastLine

        '----------------------------------
        ' 行頭の空白・タブを許可
        '----------------------------------
        SourceText = _
            LTrim$(ScriptText(LineNo))


        '==================================
        ' <SCRIPT> をチェック
        '==================================
        If Left$(SourceText, Len("<SCRIPT>")) = _
            "<SCRIPT>" Then

            '----------------------------------
            ' <SCRIPT> の後ろに
            ' 何かある場合はNG
            '----------------------------------
            If Len(SourceText) <> Len("<SCRIPT>") Then

                Debug.Print _
                    "ValidateScriptStructure : NG" & _
                    " / <SCRIPT> の後ろに文字があります。" & _
                    " / Line=" & LineNo

                Exit Function

            End If


            '----------------------------------
            ' 2個目の<SCRIPT>はNG
            '----------------------------------
            If StartPos > 0 Then

                Debug.Print _
                    "ValidateScriptStructure : " & _
                    "<SCRIPT> が複数あります。"

                Exit Function

            End If

            StartPos = LineNo

        End If


        '==================================
        ' </SCRIPT> をチェック
        '==================================
        If Left$(SourceText, Len("</SCRIPT>")) = _
            "</SCRIPT>" Then

            '----------------------------------
            ' </SCRIPT> の後ろに
            ' 何かある場合はNG
            '----------------------------------
            If Len(SourceText) <> Len("</SCRIPT>") Then

                Debug.Print _
                    "ValidateScriptStructure : NG" & _
                    " / </SCRIPT> の後ろに文字があります。" & _
                    " / Line=" & LineNo

                Exit Function

            End If


            '----------------------------------
            ' 2個目の</SCRIPT>はNG
            '----------------------------------
            If EndPos > 0 Then

                Debug.Print _
                    "ValidateScriptStructure : " & _
                    "</SCRIPT> が複数あります。"

                Exit Function

            End If

            EndPos = LineNo

        End If

    Next LineNo


    '==================================
    ' <SCRIPT> の存在確認
    '==================================
    If StartPos = 0 Then

        Debug.Print _
            "ValidateScriptStructure : " & _
            "<SCRIPT> がありません。"

        Exit Function

    End If


    '==================================
    ' </SCRIPT> の存在確認
    '==================================
    If EndPos = 0 Then

        Debug.Print _
            "ValidateScriptStructure : " & _
            "</SCRIPT> がありません。"

        Exit Function

    End If


    '==================================
    ' 順序確認
    '==================================
    If StartPos >= EndPos Then

        Debug.Print _
            "ValidateScriptStructure : " & _
            "<SCRIPT> / </SCRIPT> の順序が不正です。"

        Exit Function

    End If


    '==================================
    ' SCRIPT境界を保存
    '==================================
    ScriptStartLine = StartPos
    ScriptEndLine = EndPos


    '==================================
    ' OK
    '==================================
    Debug.Print _
        "ValidateScriptStructure : OK" & _
        " / Start=" & ScriptStartLine & _
        " / End=" & ScriptEndLine

    Debug.Print _
        "  Display Start : " & _
        ScriptStartLine + 1

    ValidateScriptStructure = True

End Function

'==================================================
' タイムカウンターチェック
'
' スクリプト範囲内に存在する
' タイムカウンターの値を確認する。
'
' IsTimeCounter
'   ↓
' 形式としてタイムカウンターか確認
'
' ValidateTimeCounter
'   ↓
' HH / MM / SS の値が正しいか確認
'
' 許容範囲
'   HH = 00～99
'   MM = 00～59
'   SS = 00～59
'
' タイムカウンターは「時刻」ではなく
' 経過時間として扱う。
'
' 例
'   24:00:00 → OK
'   99:59:59 → OK
'   00:60:00 → NG
'   00:00:60 → NG
'==================================================
Function ValidateTimeCounter() As Boolean

    Dim LineNo As Long

    Dim SourceText As String
    Dim TimeText As String

    Dim HH As Long
    Dim MM As Long
    Dim SS As Long

    ValidateTimeCounter = False


    '==================================
    ' スクリプト範囲をチェック
    '
    ' <SCRIPT>       = ScriptStartLine
    ' 最初のデータ  = ScriptStartLine + 1
    '
    ' </SCRIPT>      = ScriptEndLine
    '==================================
    For LineNo = _
        ScriptStartLine + 1 To _
        ScriptEndLine - 1

        SourceText = _
            ScriptText(LineNo)


        '==================================
        ' タイムカウンター判定
        '==================================
        If IsTimeCounter(SourceText) Then


            '----------------------------------
            ' 時刻部分8文字を取得
            '----------------------------------
            TimeText = _
                Left$(SourceText, 8)


            '----------------------------------
            ' HH / MM / SSを取得
            '----------------------------------
            HH = CLng( _
                Left$(TimeText, 2))

            MM = CLng( _
                Mid$(TimeText, 4, 2))

            SS = CLng( _
                Mid$(TimeText, 7, 2))


            Debug.Print _
                "ValidateTimeCounter : Line=" & _
                LineNo & _
                " / " & _
                TimeText


            '==================================
            ' 時の範囲
            '
            ' 00～99
            '==================================
            If HH < 0 Or HH > 99 Then

                Debug.Print _
                    "ValidateTimeCounter : NG" & _
                    " / Invalid HH : " & HH & _
                    " / Line=" & LineNo

                Exit Function

            End If


            '==================================
            ' 分の範囲
            '
            ' 00～59
            '==================================
            If MM < 0 Or MM > 59 Then

                Debug.Print _
                    "ValidateTimeCounter : NG" & _
                    " / Invalid MM : " & MM & _
                    " / Line=" & LineNo

                Exit Function

            End If


            '==================================
            ' 秒の範囲
            '
            ' 00～59
            '==================================
            If SS < 0 Or SS > 59 Then

                Debug.Print _
                    "ValidateTimeCounter : NG" & _
                    " / Invalid SS : " & SS & _
                    " / Line=" & LineNo

                Exit Function

            End If


        End If

    Next LineNo


    '==================================
    ' 全チェックOK
    '==================================
    Debug.Print _
        "ValidateTimeCounter : OK"

    ValidateTimeCounter = True

End Function

Function ValidateTIME() As Boolean

    Dim LineNo As Long

    Dim SourceText As String

    Dim StartPos As Long
    Dim EndPos As Long

    Dim TimeText As String

    Dim HH As Long
    Dim MM As Long
    Dim SS As Long

    ValidateTIME = False


    '==================================
    ' SCRIPT 内だけをチェック
    '==================================
    For LineNo = _
        ScriptStartLine + 1 To _
        ScriptEndLine - 1

        SourceText = ScriptText(LineNo)

        StartPos = InStr( _
            1, _
            SourceText, _
            "<TIME>", _
            vbTextCompare)

        '----------------------------------
        ' <TIME> がない行
        '----------------------------------
        If StartPos = 0 Then

            ' </TIME>だけある場合はエラー
            If InStr( _
                1, _
                SourceText, _
                "</TIME>", _
                vbTextCompare) > 0 Then

                Debug.Print _
                    "ValidateTIME : NG" & _
                    " / </TIME> のみ存在" & _
                    " / Line=" & LineNo

                Exit Function

            End If

        Else

            '----------------------------------
            ' </TIME> を探す
            '----------------------------------
            EndPos = InStr( _
                StartPos + Len("<TIME>"), _
                SourceText, _
                "</TIME>", _
                vbTextCompare)

            If EndPos = 0 Then

                Debug.Print _
                    "ValidateTIME : NG" & _
                    " / </TIME> がありません" & _
                    " / Line=" & LineNo

                Exit Function

            End If

            '----------------------------------
            ' TIME の中身を取得
            '----------------------------------
            TimeText = Mid$( _
                SourceText, _
                StartPos + Len("<TIME>"), _
                EndPos - _
                (StartPos + Len("<TIME>")))

            Debug.Print _
                "ValidateTIME : Line=" & _
                LineNo & _
                " / " & TimeText

            '----------------------------------
            ' 時刻フォーマットチェック
            '----------------------------------
            If Len(TimeText) <> TimeCounterFormat Then

                Debug.Print _
                    "ValidateTIME : NG" & _
                    " / Invalid Format" & _
                    " / Line=" & LineNo & _
                    " / Value=" & TimeText

                Exit Function

            End If

            If TimeCounterFormat = _
                TIME_COUNTER_FORMAT_8 Then

                If Not TimeText Like "##:##:##" Then

                    Debug.Print _
                        "ValidateTIME : NG" & _
                        " / Invalid TIME Format" & _
                        " / Line=" & LineNo

                    Exit Function

                End If

            ElseIf TimeCounterFormat = _
                TIME_COUNTER_FORMAT_9 Then

                If Not TimeText Like "##:##:##:" Then

                    Debug.Print _
                        "ValidateTIME : NG" & _
                        " / Invalid TIME Format" & _
                        " / Line=" & LineNo

                    Exit Function

                End If

            Else

                Debug.Print _
                    "ValidateTIME : NG" & _
                    " / Invalid TimeCounterFormat"

                Exit Function

            End If

            '----------------------------------
            ' 時刻値を取得
            '----------------------------------
            TimeText = Left$(TimeText, 8)

            HH = CLng( _
                Left$(TimeText, 2))

            MM = CLng( _
                Mid$(TimeText, 4, 2))

            SS = CLng( _
                Mid$(TimeText, 7, 2))

            '----------------------------------
            ' 時刻範囲チェック
            '
            ' HH : 00～99
            ' MM : 00～59
            ' SS : 00～59
            '----------------------------------
            If HH < 0 Or HH > 99 Then

                Debug.Print _
                    "ValidateTIME : NG" & _
                    " / Invalid HH : " & HH & _
                    " / Line=" & LineNo

                Exit Function

            End If

            If MM < 0 Or MM > 59 Then

                Debug.Print _
                    "ValidateTIME : NG" & _
                    " / Invalid MM : " & MM & _
                    " / Line=" & LineNo

                Exit Function

            End If

            If SS < 0 Or SS > 59 Then

                Debug.Print _
                    "ValidateTIME : NG" & _
                    " / Invalid SS : " & SS & _
                    " / Line=" & LineNo

                Exit Function

            End If

            Debug.Print _
                "ValidateTIME : Valid Time = " & _
                TimeText

        End If

    Next LineNo

    Debug.Print _
        "ValidateTIME : OK"

    ValidateTIME = True

End Function

Function ValidateTimeline() As Boolean

    Dim LineNo As Long

    Dim CurrentSeconds As Long
    Dim PreviousSeconds As Long

    Dim HasPrevious As Boolean

    ValidateTimeline = False


    '==================================
    ' Timelineモード以外では
    ' このチェックは不要
    '==================================
    If TimeCounterMode <> TIME_TIMELINE Then

        Debug.Print _
            "ValidateTimeline : Skip" & _
            " / TimeCounterMode=" & _
            TimeCounterMode

        ValidateTimeline = True

        Exit Function

    End If

    '==================================
    ' Dictionaryが存在しない場合
    '==================================
    If TimeCounterDictionary Is Nothing Then

        Debug.Print _
            "ValidateTimeline : NG" & _
            " / Dictionaryがありません。"

        Exit Function

    End If

    '==================================
    ' TimeCounterが存在しない場合
    '
    ' BuildTimeCounterDictionary側で
    ' TIME_IGNOREへフォールバックするため、
    ' 通常ここには来ない。
    '==================================
    If TimeCounterDictionary.Count = 0 Then

        Debug.Print _
            "ValidateTimeline : OK" & _
            " / TimeCounterなし"

        ValidateTimeline = True

        Exit Function

    End If

    '==================================
    ' Scriptを上から走査
    '
    ' DictionaryのキーはScriptの行番号
    '==================================
    PreviousSeconds = 0
    HasPrevious = False

    For LineNo = _
        ScriptStartLine + 1 To _
        ScriptEndLine - 1

        If TimeCounterDictionary.Exists(LineNo) Then

            CurrentSeconds = _
                CLng(TimeCounterDictionary(LineNo))

            Debug.Print _
                "ValidateTimeline : Line=" & _
                LineNo & _
                " / Time=" & _
                CurrentSeconds

            '----------------------------------
            ' 時刻の順序をチェック
            '
            ' 同時刻はOK
            '
            ' Previous <= Current
            '----------------------------------
            If HasPrevious Then

                If CurrentSeconds < PreviousSeconds Then

                    Debug.Print _
                        "ValidateTimeline : NG" & _
                        " / Timeが逆行しています。" & _
                        " / Line=" & LineNo & _
                        " / Previous=" & PreviousSeconds & _
                        " / Current=" & CurrentSeconds

                    Exit Function

                End If

            End If

            PreviousSeconds = CurrentSeconds
            HasPrevious = True

        End If

    Next LineNo

    Debug.Print _
        "ValidateTimeline : OK"

    ValidateTimeline = True

End Function

'==================================================
' タグチェック
'
' 行の先頭が "<" で始まる行を
' 制御タグ候補としてチェックする。
'
' 今回のバージョンで認識する予約タグ
'
'   <SCRIPT>
'   </SCRIPT>
'   <TIME>
'   </TIME>
'
' 行の途中に存在する "<" は
' 今回はタグとして扱わない。
'
' 例
'
'   <TIME>00:04:19</TIME>Hello
'       → OK
'
'   Hello <TIME>00:04:19</TIME>
'       → 通常の文字列として扱う
'
'   <SCRIP>
'       → NG
'
'   <TIME
'       → NG
'==================================================
Function ValidateTag() As Boolean

    Dim LineNo As Long
    Dim SourceText As String

    Dim TagEndPos As Long
    Dim TagText As String


    ValidateTag = False


    '==================================
    ' SCRIPT範囲をチェック
    '
    ' <SCRIPT>       = ScriptStartLine
    ' 最初のデータ  = ScriptStartLine + 1
    '
    ' </SCRIPT>      = ScriptEndLine
    '==================================
    For LineNo = _
        ScriptStartLine + 1 To _
        ScriptEndLine - 1


        SourceText = ScriptText(LineNo)


        '==================================
        ' 行頭が "<" でない場合
        '
        ' 今回はタグとして扱わない。
        '==================================
        If Left$(SourceText, 1) <> "<" Then

            GoTo ContinueLine

        End If


        '==================================
        ' ">" を探す
        '
        ' タグの終端がない場合は破損タグ
        '==================================
        TagEndPos = InStr( _
            1, _
            SourceText, _
            ">", _
            vbBinaryCompare)


        If TagEndPos = 0 Then

            Debug.Print _
                "ValidateTag : NG" & _
                " / タグの終端 '>' がありません。" & _
                " / Line=" & LineNo & _
                " / Text=" & SourceText

            Exit Function

        End If


        '==================================
        ' タグ部分だけ取得
        '==================================
        TagText = Left$( _
            SourceText, _
            TagEndPos)


        Debug.Print _
            "ValidateTag : Line=" & _
            LineNo & _
            " / Tag=" & TagText


        '==================================
        ' 予約タグか確認
        '==================================
        If Not IsReservedTag(TagText) Then

            Debug.Print _
                "ValidateTag : NG" & _
                " / 未定義のタグです。" & _
                " / Line=" & LineNo & _
                " / Tag=" & TagText

            Exit Function

        End If


ContinueLine:

    Next LineNo


    '==================================
    ' 全チェックOK
    '==================================
    Debug.Print _
        "ValidateTag : OK"

    ValidateTag = True

End Function

Private Function IsReservedTag(ByVal TagText As String) As Boolean

    Dim ReservedTags As Variant
    Dim i As Long


    IsReservedTag = False


    '==================================
    ' 現在のバージョンで使用する
    ' 予約タグ
    '==================================
    ReservedTags = Array( _
        "<SCRIPT>", _
        "</SCRIPT>", _
        "<TIME>", _
        "</TIME>" _
    )


    '==================================
    ' 予約タグと比較
    '==================================
    For i = LBound(ReservedTags) To UBound(ReservedTags)

        If StrComp( _
            TagText, _
            ReservedTags(i), _
            vbTextCompare) = 0 Then

            IsReservedTag = True

            Exit Function

        End If

    Next i

End Function
