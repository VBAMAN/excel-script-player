Attribute VB_Name = "PLY_20_Dictionary"
Option Explicit

'==================================================
' タイムカウンターDictionaryを作成する。
' <TIME>...</TIME> と通常の行頭カウンターの両方に対応。
'==================================================
Public Sub BuildTimeCounterDictionary()

    Dim ws As Worksheet
    Dim LineNo As Long
    Dim SourceText As String
    Dim TimeText As String
    Dim TotalSeconds As Long

    Dim StartPos As Long
    Dim EndPos As Long

    '==================================
    ' Dictionaryを作成
    '==================================
    Set TimeCounterDictionary = _
        CreateObject("Scripting.Dictionary")

    Set ws = _
        ThisWorkbook.Worksheets("Sheet1")

    Debug.Print "TextLastLine : " & TextLastLine
    Debug.Print _
        "TimeCounterFormat : " & _
        TimeCounterFormat


    '==================================
    ' スクリプトを走査
    '==================================
    For LineNo = _
        ScriptStartLine + 1 To _
        ScriptEndLine - 1


SourceText = ScriptText(LineNo)


        Debug.Print _
            "Line : " & LineNo & _
            " / " & SourceText


        '==================================
        ' <TIME> タグ
        '
        ' 行頭にある <TIME> だけ有効
        '==================================
        If Left$( _
            SourceText, _
            Len("<TIME>")) = "<TIME>" Then

            StartPos = Len("<TIME>") + 1

            EndPos = InStr( _
                StartPos, _
                SourceText, _
                "</TIME>", _
                vbTextCompare)

            If EndPos > 0 Then

                '----------------------------------
                ' TIMEタグ内部の時刻を取得
                '----------------------------------
                TimeText = Mid$( _
                    SourceText, _
                    StartPos, _
                    EndPos - StartPos)

                '----------------------------------
                ' 実際の時刻部分は8文字
                '
                ' 8文字形式
                ' 00:00:07
                '
                ' 9文字形式
                ' 00:00:07:
                '----------------------------------
                TimeText = _
                    Left$(TimeText, 8)

                '----------------------------------
                ' 秒へ変換
                '----------------------------------
                TotalSeconds = _
                    TimeTextToSeconds(TimeText)

                '----------------------------------
                ' Dictionaryへ登録
                '----------------------------------
                TimeCounterDictionary.Add _
                    LineNo, _
                    TotalSeconds

                Debug.Print _
                    "  TIME : " & LineNo & _
                    " = " & TotalSeconds

            End If


        '==================================
        ' 通常のタイムカウンター
        '==================================
        ElseIf IsTimeCounter(SourceText) Then

            '----------------------------------
            ' 時刻部分8文字だけを取得
            '
            ' 8文字形式
            ' 00:00:07
            '
            ' 9文字形式
            ' 00:00:07:
            '
            ' どちらも時刻そのものは8文字
            '----------------------------------
            TimeText = _
                Left$(SourceText, 8)

            '----------------------------------
            ' 秒へ変換
            '----------------------------------
            TotalSeconds = _
                TimeTextToSeconds(TimeText)

            '----------------------------------
            ' Dictionaryへ登録
            '----------------------------------
            TimeCounterDictionary.Add _
                LineNo, _
                TotalSeconds

            Debug.Print _
                "  Counter : " & LineNo & _
                " = " & TotalSeconds

        End If

    Next LineNo


    '==================================
    ' 結果
    '==================================
    Debug.Print _
        "Dictionary Count : " & _
        TimeCounterDictionary.Count


    '==================================
    ' タイムカウンターが存在しない場合
    '
    ' TIME_TIMELINEの場合は
    ' TIME_IGNOREへフォールバック
    '==================================
    If TimeCounterDictionary.Count = 0 Then

        If TimeCounterMode = TIME_TIMELINE Then

            Debug.Print _
                "TIME_TIMELINE : No TimeCounter"

            TimeCounterMode = TIME_IGNORE

        End If

    End If

End Sub

Public Function IsTimeCounter( _
    ByVal SourceText As String) As Boolean

    Dim TimeText As String

    IsTimeCounter = False

    '----------------------------------
    ' 8文字 / 9文字の長さを確認
    '----------------------------------
    If TimeCounterFormat <> TIME_COUNTER_FORMAT_8 _
       And TimeCounterFormat <> TIME_COUNTER_FORMAT_9 Then

        Debug.Print _
            "IsTimeCounter : Invalid Format = " & _
            TimeCounterFormat

        Exit Function

    End If

    '----------------------------------
    ' 必要な文字数があるか確認
    '----------------------------------
    If Len(SourceText) < TimeCounterFormat Then
        Exit Function
    End If

    '----------------------------------
    ' 指定された文字数を取得
    '----------------------------------
    TimeText = _
        Left$(SourceText, TimeCounterFormat)

    '----------------------------------
    ' 8文字フォーマット
    '
    ' 00:00:00
    '----------------------------------
    If TimeCounterFormat = TIME_COUNTER_FORMAT_8 Then

        If TimeText Like "##:##:##" Then

            IsTimeCounter = True

        End If

    '----------------------------------
    ' 9文字フォーマット
    '
    ' 00:00:00:
    '----------------------------------
    ElseIf TimeCounterFormat = TIME_COUNTER_FORMAT_9 Then

        If TimeText Like "##:##:##:" Then

            IsTimeCounter = True

        End If

    End If

End Function

Public Function TimeTextToSeconds( _
    ByVal TimeText As String) As Long

    Dim HH As Long
    Dim MM As Long
    Dim SS As Long

    '----------------------------------
    ' 時刻部分は8文字
    '
    ' 00:00:00
    '----------------------------------
    If Len(TimeText) < 8 Then

        TimeTextToSeconds = 0

        Exit Function

    End If

    '----------------------------------
    ' 9文字フォーマットの場合
    '
    ' 00:00:00:
    '          ↑
    '          最後の : は時刻ではない
    '
    ' 時刻部分8文字だけを使用する
    '----------------------------------
    TimeText = Left$(TimeText, 8)

    HH = CLng(Left$(TimeText, 2))
    MM = CLng(Mid$(TimeText, 4, 2))
    SS = CLng(Mid$(TimeText, 7, 2))

    TimeTextToSeconds = _
        HH * 3600 + _
        MM * 60 + _
        SS

End Function
