Attribute VB_Name = "PLY_10_Tag"
Option Explicit

'======================================
' タグ表示設定
'======================================

Public ShowTime As Boolean


'======================================
' タグを処理する
'
' 戻り値：
'   表示する文字列
'
' タグによって非表示の場合：
'   ""
'======================================

Public Function ParseEndRollLine( _
    ByVal SourceText As String) As String

    ' 新コード

    Dim StartPos As Long
    Dim EndPos As Long
    Dim TagValue As String

    ' <TIME>を処理した行で、行頭カウンターを二重処理しないためのフラグ。
    Dim HasTimeTag As Boolean


    '----------------------------------
    ' 初期値
    '----------------------------------
    HasTimeTag = False


    '----------------------------------
    ' TIMEタグを探す
    '----------------------------------
    StartPos = InStr( _
        1, _
        SourceText, _
        "<TIME>", _
        vbTextCompare)

    If StartPos > 0 Then

        EndPos = InStr( _
            StartPos, _
            SourceText, _
            "</TIME>", _
            vbTextCompare)

        If EndPos > 0 Then

            '----------------------------------
            ' TIMEタグあり
            '----------------------------------
            HasTimeTag = True


            '----------------------------------
            ' タグの中身
            '----------------------------------
            TagValue = Mid$( _
                SourceText, _
                StartPos + Len("<TIME>"), _
                EndPos - _
                (StartPos + Len("<TIME>")))


            '----------------------------------
            ' TIMEを表示する場合
            '----------------------------------
            If ShowTime Then

                SourceText = _
                    Left$(SourceText, StartPos - 1) _
                    & TagValue _
                    & Mid$( _
                        SourceText, _
                        EndPos + Len("</TIME>"))

            Else

                '----------------------------------
                ' TIMEを表示しない
                '----------------------------------
                SourceText = _
                    Left$(SourceText, StartPos - 1) _
                    & Mid$( _
                        SourceText, _
                        EndPos + Len("</TIME>"))

            End If

        End If

    End If


    '----------------------------------
    ' 行頭タイムカウンターの処理
    '
    ' <TIME>タグを使用した行は
    ' ここでは処理しない。
    '
    ' 8文字
    ' 00:00:07Hello
    '
    ' 9文字
    ' 00:00:07:Hello
    '----------------------------------
    If Not HasTimeTag Then

        If Not ShowTimeCounter Then

            If Len(SourceText) >= TimeCounterFormat Then

                If IsTimeCounter(SourceText) Then

                    SourceText = _
                        Mid$( _
                            SourceText, _
                            TimeCounterFormat + 1)

                End If

            End If

        End If

    End If


    '----------------------------------
    ' SCRIPT ENDタグを表示しない
    '----------------------------------
    If InStr( _
        1, _
        SourceText, _
        "</SCRIPT>", _
        vbTextCompare) > 0 Then

        SourceText = ""

    End If


    '----------------------------------
    ' 最終結果
    '----------------------------------
    ParseEndRollLine = SourceText

End Function
