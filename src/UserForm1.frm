VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------
' Label管理
'--------------------------------------
Dim labelControls() As MSForms.Label
Dim rowTop() As Long

'--------------------------------------
' スクロール管理
'--------------------------------------
Dim scrollTimerRunning As Boolean


'======================================
' 定数
'======================================

Const ROWS As Long = 11
Const COLS As Long = 40

Const LABEL_WIDTH As Long = 13
Const LABEL_HEIGHT As Long = 13

Const MARGIN As Long = 1

Const SCROLL_SPEED As Long = 20


'==================================================
' UserForm 初期化
'
' preview側ですでに
'   Load
'   Validate
'   BuildTimeCounterDictionary
' が完了している前提。
'==================================================
Private Sub UserForm_Initialize()

    '----------------------------------
    ' Timeline初期化
    '----------------------------------
    TimeLineStarted = False
    TimeLineReady = False
    TimeLineFirstSeconds = 0

    '----------------------------------
    ' TIME_WAIT初期化
    '----------------------------------
    TimeWaitActive = False
    TimeWaitLine = 0
    TimeWaitStart = 0

    '----------------------------------
    ' Label配列
    '----------------------------------
    ReDim labelControls(1 To ROWS * COLS)
    ReDim rowTop(1 To ROWS)

    '----------------------------------
    ' Label作成
    '----------------------------------
    CreateLabels

    '----------------------------------
    ' 再生位置初期化
    '----------------------------------
    TextLine = ScriptStartLine + 1

    scrollTimerRunning = False

End Sub


'==================================================
' UserForm 表示開始
'==================================================
Private Sub UserForm_Activate()

    Debug.Print _
        "UserForm_Activate : scrollTimerRunning = " & _
        scrollTimerRunning

    If Not scrollTimerRunning Then

        scrollTimerRunning = True

        StartScrolling

    End If

End Sub


'======================================
' フォーム終了時
'======================================

Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)

    scrollTimerRunning = False

End Sub

Private Sub UserForm_Terminate()

    scrollTimerRunning = False

End Sub


'======================================
' Labelコントロール生成
'======================================


Private Sub CreateLabels()

    Dim r As Long
    Dim c As Long
    Dim lbl As MSForms.Label

    ReDim labelControls(1 To ROWS, 1 To COLS)
    ReDim rowTop(1 To ROWS)

    For r = 1 To ROWS

        '----------------------------------
        ' Labelは常に固定位置
        '
        ' FirstLineRowsは
        ' 初期表示を開始する行を決めるだけ
        ' Labelの位置には影響させない
        '
        ' LabelOffsetY / X は
        ' Label全体の左上位置を指定する
        '----------------------------------
        rowTop(r) = _
            LabelOffsetY + _
            (r - 1) * _
            (LABEL_HEIGHT + MARGIN)

        For c = 1 To COLS

            Set lbl = Me.Controls.Add( _
                "Forms.Label.1", _
                "Label_" & r & "_" & c, _
                True)

            With lbl

                .Width = LABEL_WIDTH
                .Height = LABEL_HEIGHT

                .Top = rowTop(r)

                .Left = _
                    LabelOffsetX + _
                    (c - 1) * _
                    (LABEL_WIDTH + MARGIN)

                .Caption = ""

                .TextAlign = fmTextAlignCenter

            End With

            Set labelControls(r, c) = lbl

        Next c

    Next r

End Sub


'======================================
' 初期表示
'======================================

Private Sub SetInitialText()

    Dim r As Long

    Debug.Print "SetInitialText"
    Debug.Print "DisplayMode : " & DisplayMode
    Debug.Print "TimeCounterMode : " & TimeCounterMode


    '==================================
    ' スクリプト読み出し位置を先頭へ
    '==================================
    TextLine = ScriptStartLine + 1


    Select Case DisplayMode


        '==================================
        ' DISPLAY_MESSAGE
        '
        ' FirstLineRowsから下方向へ
        ' 1行ずつ表示する。
        '
        ' TIME_WAIT
        '   表示 → 待つ → 次の行
        '
        ' TIME_TIMELINE
        '   時刻を待つ → 表示
        '
        ' Repeatはここでは行わない。
        '==================================
        Case DISPLAY_MESSAGE


            For r = FirstLineRows To ROWS


                '==================================
                ' SCRIPT END
                '
                ' 初期表示中に
                ' </SCRIPT>へ到達した場合、
                ' ここで初期表示を終了する。
                '
                ' Repeatの判断は
                ' StartScrollingで行う。
                '==================================
                If TextLine >= ScriptEndLine Then

                    Debug.Print _
                        "SetInitialText : SCRIPT END" & _
                        " / TextLine=" & TextLine

                    Exit For

                End If


                '==================================
                ' TIME_TIMELINE
                '==================================
                If TimeCounterMode = TIME_TIMELINE Then

                    Do While Not IsTimeLineReady

                        If Not scrollTimerRunning Then

                            Debug.Print _
                                "SetInitialText : " & _
                                "TIME_TIMELINE 中断"

                            Exit Sub

                        End If

                        DoEvents
                        wait 10

                    Loop

                End If


                '==================================
                ' 現在のスクリプト行を表示
                '==================================
                Debug.Print _
                    "Initial Message Row : " & _
                    r & _
                    " / TextLine : " & _
                    TextLine


                '==================================
                ' SCRIPT END
                '
                ' TIME_TIMELINE待ちなどの途中で
                ' 停止された場合。
                '==================================
                If Not scrollTimerRunning Then

                    Exit Sub

                End If


                Call UpdateLabelsInRow(r)


                '==================================
                ' TIME_WAIT
                '
                ' 表示した行の時間だけ待つ。
                '==================================
                If TimeCounterMode = TIME_WAIT Then

                    Do While TimeWaitActive

                        If Not scrollTimerRunning Then

                            Debug.Print _
                                "SetInitialText : " & _
                                "TIME_WAIT 中断"

                            Exit Sub

                        End If

                        Call ProcessTimeWait( _
                            TimeWaitSeconds)

                        If TimeWaitActive Then

                            DoEvents
                            wait 10

                        End If

                    Loop

                End If


            Next r



        '==================================
        ' DISPLAY_ENDROLL
        '
        ' 最初の1行を
        ' 一番下の行へ表示する。
        '==================================
        Case DISPLAY_ENDROLL


            r = ROWS


            '==================================
            ' TIME_TIMELINE
            '==================================
            If TimeCounterMode = TIME_TIMELINE Then

                Do While Not IsTimeLineReady

                    If Not scrollTimerRunning Then

                        Debug.Print _
                            "SetInitialText : " & _
                            "TIME_TIMELINE 中断"

                        Exit Sub

                    End If

                    DoEvents
                    wait 10

                Loop

            End If


            '==================================
            ' 最初の1行を表示
            '==================================
            Debug.Print _
                "Initial EndRoll Row : " & _
                r & _
                " / TextLine : " & _
                TextLine

            Call UpdateLabelsInRow(r)


            '==================================
            ' SCRIPT END
            '==================================
            If Not scrollTimerRunning Then

                Exit Sub

            End If


            '==================================
            ' TIME_WAIT
            '
            ' 最初の1行を表示したあと、
            ' 指定時間だけ待つ。
            '==================================
            If TimeCounterMode = TIME_WAIT Then

                Do While TimeWaitActive

                    If Not scrollTimerRunning Then

                        Debug.Print _
                            "SetInitialText : " & _
                            "TIME_WAIT 中断"

                        Exit Sub

                    End If

                    Call ProcessTimeWait( _
                        TimeWaitSeconds)

                    If TimeWaitActive Then

                        DoEvents
                        wait 10

                    End If

                Loop

            End If


    End Select


    Debug.Print "SetInitialText End"

End Sub

Private Sub StartScrolling()

    Dim r As Long
    Dim c As Long

    Dim InitialFillMode As Boolean
    Dim InitialRow As Long
    Dim InitialRowsUsed As Long

    Debug.Print "StartScrolling"
    Debug.Print "DisplayMode : " & DisplayMode

    '==================================================
    ' 初期表示
    '==================================================

    Call SetInitialText

    '--------------------------------------------------
    ' MESSAGEの場合
    ' 初期画面がまだ埋まっていない場合は
    ' 空いている行へリピート表示する
    '--------------------------------------------------

    If DisplayMode = DISPLAY_MESSAGE Then

        InitialRowsUsed = _
            TextLine - (ScriptStartLine + 1)

        InitialRow = _
            FirstLineRows + InitialRowsUsed

        If InitialRow <= ROWS Then

            InitialFillMode = True

            Debug.Print _
                "Initial Message : 画面未充填" & _
                " / Continue Repeat"

        Else

            InitialFillMode = False

        End If

        DisplayMode = DISPLAY_ENDROLL

        Debug.Print _
            "DisplayMode Change : DISPLAY_ENDROLL"

    End If


    '==================================================
    ' メインループ
    '==================================================

    Do While scrollTimerRunning

        '==================================================
        ' 初期画面充填
        '==================================================

        If InitialFillMode Then

            '--------------------------------------------------
            ' すでにSCRIPT ENDなら、
            ' 表示処理の前にリピートする
            '
            ' ★ 初回の空白行対策
            '--------------------------------------------------

            If TextLine >= ScriptEndLine Then

                If ScriptRepeatMode Then

                    Debug.Print _
                        "INITIAL REPEAT : Final Line" & _
                        " / TextLine=" & TextLine

                    If RepeatWaitSeconds > 0 Then

                        Debug.Print _
                            "INITIAL REPEAT WAIT START : " & _
                            RepeatWaitSeconds

                        RepeatWaitStart = GetTickCount()

                        Do While _
                            GetTickCount() - RepeatWaitStart < _
                            CLng(RepeatWaitSeconds * 1000#)

                            If Not scrollTimerRunning Then
                                Exit Do
                            End If

                            DoEvents
                            wait 10

                        Loop

                        Debug.Print _
                            "INITIAL REPEAT WAIT END"

                    End If

                    If Not scrollTimerRunning Then
                        Exit Do
                    End If

                    '------------------------------------------
                    ' スクリプト先頭へ戻る
                    '------------------------------------------

                    TextLine = ScriptStartLine + 1

                    Debug.Print _
                        "INITIAL REPEAT : Reset TextLine = " & _
                        TextLine

                    '------------------------------------------
                    ' Timelineも先頭から再スタート
                    '------------------------------------------

                    TimeLineStarted = False
                    TimeLineReady = False
                    TimeLineFirstSeconds = 0

                Else

                    ' Repeatしない場合
                    InitialFillMode = False

                    GoTo LoopContinue

                End If

            End If


            '--------------------------------------------------
            ' Timeline待ち
            '--------------------------------------------------

            If TimeCounterMode = TIME_TIMELINE Then

                If Not IsTimeLineReady Then

                    DoEvents
                    wait 10

                    GoTo LoopContinue

                End If

            End If


            '--------------------------------------------------
            ' 次の空行へ表示
            '--------------------------------------------------

            Debug.Print _
                "Initial Fill : Row=" & _
                InitialRow & _
                " / TextLine=" & _
                TextLine

            Call UpdateLabelsInRow(InitialRow)


            '--------------------------------------------------
            ' TIME_WAIT
            '--------------------------------------------------

            If TimeCounterMode = TIME_WAIT Then

                If TimeWaitActive Then

                    Do While TimeWaitActive

                        If Not scrollTimerRunning Then
                            Exit Do
                        End If

                        Call ProcessTimeWait( _
                            TimeWaitSeconds)

                        If TimeWaitActive Then

                            DoEvents
                            wait 10

                        End If

                    Loop

                End If

            End If


            If Not scrollTimerRunning Then
                Exit Do
            End If


            '--------------------------------------------------
            ' 次の空行へ
            '--------------------------------------------------

            InitialRow = InitialRow + 1


            '--------------------------------------------------
            ' 11行すべて埋まった
            '--------------------------------------------------

            If InitialRow > ROWS Then

                InitialFillMode = False

                Debug.Print _
                    "Initial Fill : COMPLETE"

                '--------------------------------------------------
                ' 初期充填中にリピートした場合、
                ' そのRepeatWaitはすでに済んでいる。
                '
                ' 通常スクロール開始時に
                ' もう一度RepeatWaitしないようにする。
                '--------------------------------------------------

                If TextLine >= ScriptEndLine Then

                    TextLine = ScriptStartLine + 1

                    TimeLineStarted = False
                    TimeLineReady = False
                    TimeLineFirstSeconds = 0

                    Debug.Print _
                        "Initial Fill : Reset TextLine = " & _
                        TextLine

                End If

            End If

            GoTo LoopContinue

        End If


        '==================================================
        ' TIME_WAIT
        ' 通常スクロール
        '==================================================

        If TimeCounterMode = TIME_WAIT Then

            If TimeWaitActive Then

                Call ProcessTimeWait(TimeWaitSeconds)

                If TimeWaitActive Then

                    DoEvents
                    wait 10

                    GoTo LoopContinue

                End If

            End If

        End If


        '==================================================
        ' TIME_TIMELINE
        ' 通常スクロール
        '==================================================

        If TimeCounterMode = TIME_TIMELINE Then

            If Not IsTimeLineReady Then

                DoEvents
                wait 10

                GoTo LoopContinue

            End If

        End If


        '==================================================
        ' SCRIPT END
        ' 通常スクロール
        '==================================================

        If TextLine >= ScriptEndLine Then

            If ScriptRepeatMode Then

                Debug.Print _
                    "SCRIPT REPEAT : Final Line" & _
                    " / TextLine=" & TextLine

                If RepeatWaitSeconds > 0 Then

                    Debug.Print _
                        "REPEAT WAIT START : " & _
                        RepeatWaitSeconds

                    RepeatWaitStart = GetTickCount()

                    Do While _
                        GetTickCount() - RepeatWaitStart < _
                        CLng(RepeatWaitSeconds * 1000#)

                        If Not scrollTimerRunning Then
                            Exit Do
                        End If

                        DoEvents
                        wait 10

                    Loop

                    Debug.Print _
                        "REPEAT WAIT END"

                End If

                TextLine = ScriptStartLine + 1

                Debug.Print _
                    "SCRIPT REPEAT : Reset TextLine = " & _
                    TextLine

                TimeLineStarted = False
                TimeLineReady = False
                TimeLineFirstSeconds = 0

                GoTo LoopContinue

            Else

                scrollTimerRunning = False

                Debug.Print _
                    "SCRIPT STOP : Final Line" & _
                    " / TextLine=" & TextLine

                Exit Do

            End If

        End If


        '==================================================
        ' Label内容を1行上へ移動
        '==================================================

        Debug.Print "Scroll Execute"

        For r = 1 To ROWS - 1

            For c = 1 To COLS

                labelControls(r, c).Caption = _
                    labelControls(r + 1, c).Caption

            Next c

        Next r


        '==================================================
        ' 最下段へ次のスクリプトを表示
        '==================================================

        Debug.Print _
            "Update Bottom Row : TextLine = " & _
            TextLine

        Call UpdateLabelsInRow(ROWS)

        Debug.Print _
            "After UpdateLabelsInRow : " & _
            "Row=" & ROWS & _
            " / TextLine=" & _
            TextLine


LoopContinue:

        DoEvents

        If TimeCounterMode <> TIME_WAIT _
           And TimeCounterMode <> TIME_TIMELINE Then

            wait 1000

        End If

    Loop

    Debug.Print "StartScrolling End"

End Sub


'======================================
' 1行分のテキストをLabelへセット
'
' 現在のTextLineを表示し、
' 次に表示する行へTextLineを進める。
'
' Repeat / SCRIPT終了の判断は
' このルーチンでは行わない。
'======================================

Private Sub UpdateLabelsInRow( _
    ByVal RowIndex As Long)

    Dim c As Long

    Dim TextValue As String
    Dim RawTextValue As String

    Dim TextLength As Long
    Dim DisplayLine As Long

    Dim WaitSeconds As Double


    '==================================
    ' 現在のスクリプト行を記録
    '==================================
    DisplayLine = TextLine


    '==================================
    ' SCRIPT END確認
    '
    ' </SCRIPT> は表示しない。
    '
    ' ここへ来た場合は、
    ' 呼び出し側でSCRIPT終了を判断する。
    '==================================
    If TextLine >= ScriptEndLine Then

        Debug.Print _
            "SCRIPT END : TextLine=" & _
            TextLine

        Exit Sub

    End If


    '==================================
    ' 現在行のテキストを取得
    '==================================
    RawTextValue = ScriptText(TextLine)


    '==================================
    ' 表示文字列を解析
    '
    ' TIMEなどのタグや
    ' 行頭タイムカウンターを処理する
    '==================================
    TextValue = ParseEndRollLine(RawTextValue)

    TextLength = Len(TextValue)


    '==================================
    ' 40個のLabelへ
    ' 1文字ずつセット
    '==================================
    For c = 1 To COLS

        If c <= TextLength Then

            labelControls(RowIndex, c).Caption = _
                Mid$(TextValue, c, 1)

        Else

            labelControls(RowIndex, c).Caption = ""

        End If

    Next c


    '==================================
    ' TIME_WAIT
    '
    ' 実際に表示した行の待ち時間を決定
    '==================================
    If TimeCounterMode = TIME_WAIT Then

        '----------------------------------
        ' タイムカウンターがあるか確認
        '
        ' 必ず解析前の
        ' RawTextValueを使用する
        '----------------------------------
        If IsTimeCounter(RawTextValue) Then

            '----------------------------------
            ' 00:00:00 → 秒へ変換
            '----------------------------------
            WaitSeconds = _
                TimeTextToSeconds( _
                    Left$(RawTextValue, 8))

            Debug.Print _
                "TIME_WAIT COUNTER : Line=" & _
                DisplayLine & _
                " / Wait=" & _
                WaitSeconds

        ElseIf Left$(RawTextValue, 6) = "<TIME>" Then

            '----------------------------------
            ' <TIME>タグ内の時刻を使用
            '----------------------------------
            WaitSeconds = _
                TimeTextToSeconds( _
                    Mid$(RawTextValue, 7, 8))

            Debug.Print _
                "TIME_WAIT <TIME> : Line=" & _
                DisplayLine & _
                " / Wait=" & _
                WaitSeconds

        Else

            '----------------------------------
            ' タイムカウンターなし
            ' デフォルト待ち時間を使用
            '----------------------------------
            WaitSeconds = _
                DefaultWaitSeconds

            Debug.Print _
                "TIME_WAIT DEFAULT : Line=" & _
                DisplayLine & _
                " / Wait=" & _
                WaitSeconds

        End If


        '----------------------------------
        ' 待ち時間を保存
        '----------------------------------
        TimeWaitLine = DisplayLine

        TimeWaitSeconds = WaitSeconds


        '----------------------------------
        ' この行を表示した瞬間から
        ' WAITを開始する
        '----------------------------------
        TimeWaitStart = GetTickCount()

        TimeWaitActive = True


        Debug.Print _
            "TIME_WAIT START : Line=" & _
            TimeWaitLine & _
            " / Wait=" & _
            TimeWaitSeconds

    End If


    '==================================
    ' 次のスクリプト行へ進む
    '
    ' 最終データ行を表示した場合は
    ' TextLine = ScriptEndLine となる。
    '
    ' ここではRepeatしない。
    '==================================
    TextLine = TextLine + 1


End Sub
