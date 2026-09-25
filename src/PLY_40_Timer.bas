Attribute VB_Name = "PLY_40_Timer"
Option Explicit

#If VBA7 Then
    ' 全体で使うウェイト（フレーム制御）
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' 起動からの経過時間（ミリ秒）
Public Declare Function GetTickCount Lib "kernel32" () As Long
Sub wait(t As Long)

' ウエイト
sTime = GetTickCount

Do While GetTickCount - sTime < t
 
    Sleep 1
          
Loop

End Sub



'==================================================
' TIME_TIMELINEの次行が表示可能か確認する。
'==================================================
Function IsTimeLineReady() As Boolean

    IsTimeLineReady = True

    If TimeCounterMode <> TIME_TIMELINE Then

        Exit Function

    End If

    ProcessTimeCounter

    If Not TimeLineReady Then

        IsTimeLineReady = False

    End If

End Function

Sub ProcessTimeTimeline( _
    ByVal TargetSeconds As Long)

    Dim ElapsedSeconds As Long
    Dim AdjustedSeconds As Long

    '----------------------------------
    ' タイムライン開始
    '----------------------------------
    If Not TimeLineStarted Then

        TimeLineStart = Now
        TimeLineStarted = True

        '----------------------------------
        ' 1行目の時刻を取得
        '----------------------------------
        
    If TimeCounterDictionary.Count > 0 Then

        Dim FirstLine As Long

        FirstLine = _
            TimeCounterDictionary.Keys()(0)

        TimeLineFirstSeconds = _
            TimeCounterDictionary(FirstLine)

    Else

        TimeLineFirstSeconds = 0

    End If
        

        Debug.Print _
            "Timeline First : " & _
            TimeLineFirstSeconds

    End If

    '----------------------------------
    ' 経過時間
    '----------------------------------
    ElapsedSeconds = _
        DateDiff( _
            "s", _
            TimeLineStart, _
            Now)

    '----------------------------------
    ' TimeLineModeによる補正
    '----------------------------------
    Select Case TimeLineMode

        '----------------------------------
        ' 1行目の時刻を実時間として使用
        '----------------------------------
        Case TIMELINE_FIRST_AS_START

            AdjustedSeconds = _
                TargetSeconds

        '----------------------------------
        ' 1行目の時刻を0秒として使用
        '----------------------------------
        Case TIMELINE_FIRST_AS_ZERO

            AdjustedSeconds = _
                TargetSeconds - _
                TimeLineFirstSeconds

        '----------------------------------
        ' 指定秒数後を開始基準として使用
        '----------------------------------
        Case TIMELINE_START_OFFSET

            AdjustedSeconds = _
                TargetSeconds - _
                TimeLineFirstSeconds + _
                TimeLineStartOffset

    End Select

    '----------------------------------
    ' マイナス防止
    '----------------------------------
    If AdjustedSeconds < 0 Then

        AdjustedSeconds = 0

    End If

    '----------------------------------
    ' デバッグ
    '----------------------------------
    Debug.Print _
        "Elapsed : " & ElapsedSeconds

    Debug.Print _
        "Target : " & TargetSeconds

    Debug.Print _
        "First : " & TimeLineFirstSeconds

    Debug.Print _
        "Adjusted : " & AdjustedSeconds

    '----------------------------------
    ' 到達判定
    '----------------------------------
    If ElapsedSeconds >= AdjustedSeconds Then

        TimeLineReady = True

    Else

        TimeLineReady = False

    End If

End Sub

'==================================================
' 現在行以降のタイムカウンターを取得し、
' TIME_WAIT / TIME_TIMELINEへ振り分ける。
'==================================================
Sub ProcessTimeCounter()

    Dim CurrentTime As Long
    Dim Key As Variant
    Dim NextTimeLine As Long

    '----------------------------------
    ' タイムカウンターを使用しない
    '----------------------------------
    If TimeCounterMode = TIME_IGNORE Then

        Debug.Print "TIME_IGNORE"

        Exit Sub

    End If


    '----------------------------------
    ' Dictionary確認
    '----------------------------------
    If TimeCounterDictionary Is Nothing Then

        Debug.Print "Dictionary"

        Exit Sub

    End If


    '----------------------------------
    ' 現在行以降の
    ' 次のタイムカウンター行を探す
    '----------------------------------
    NextTimeLine = 0

    For Each Key In TimeCounterDictionary.Keys

        If CLng(Key) >= TextLine Then

            If NextTimeLine = 0 Then

                NextTimeLine = CLng(Key)

            ElseIf CLng(Key) < NextTimeLine Then

                NextTimeLine = CLng(Key)
                
                Debug.Print "Next : " & NextTimeLine

            End If

        End If

    Next Key


    '----------------------------------
    ' 次のタイムカウンターがない
    '----------------------------------
    If NextTimeLine = 0 Then

        Debug.Print "No TimeCounter"

        Exit Sub

    End If


    '----------------------------------
    ' タイムカウンター取得
    '----------------------------------
    CurrentTime = _
        TimeCounterDictionary(NextTimeLine)


    If IsScriptEndLine(NextTimeLine) Then

        Debug.Print "SCRIPT END : " & NextTimeLine

    End If


    Debug.Print _
        "TimeCounter : " & _
        TimeCounterMode

    Debug.Print _
        "TimeLine : " & _
        NextTimeLine

    Debug.Print _
        "Seconds : " & _
        CurrentTime


    '----------------------------------
    ' モード別処理
    '----------------------------------
    Select Case TimeCounterMode

        Case TIME_WAIT

            ProcessTimeWait CurrentTime


        Case TIME_TIMELINE

            ProcessTimeTimeline CurrentTime

    End Select

End Sub


'==================================================
' TIME_WAIT
'
' その行を表示してから、
' 指定された秒数が経過するまで待つ
'
' タイムカウンターあり
'   → その行の時間を使用
'
' タイムカウンターなし
'   → デフォルト待ち時間を使用
'==================================================

Sub ProcessTimeWait( _
    ByVal WaitSeconds As Double)

    Dim ElapsedMilliseconds As Long
    Dim WaitMilliseconds As Long


    '==================================
    ' 待ち時間をミリ秒へ変換
    '==================================
    If WaitSeconds < 0 Then
        WaitSeconds = 0
    End If

    WaitMilliseconds = _
        CLng(WaitSeconds * 1000#)


    '==================================
    ' WAIT中
    '==================================
    If TimeWaitActive Then

        ElapsedMilliseconds = _
            GetTickCount() - TimeWaitStart


        Debug.Print _
            "TIME_WAIT : Line=" & _
            TimeWaitLine & _
            " / Elapsed=" & _
            ElapsedMilliseconds & _
            "ms / Wait=" & _
            WaitMilliseconds & "ms"


        '----------------------------------
        ' WAIT完了
        '----------------------------------
        If ElapsedMilliseconds >= WaitMilliseconds Then

            TimeWaitActive = False

            Debug.Print _
                "TIME_WAIT END : Line=" & _
                TimeWaitLine

        End If

        Exit Sub

    End If


    '==================================
    ' 新しいWAITを開始
    '==================================
    TimeWaitStart = GetTickCount()

    TimeWaitActive = True

    Debug.Print _
        "TIME_WAIT START : Line=" & _
        TimeWaitLine & _
        " / Wait=" & _
        WaitMilliseconds & "ms"

End Sub

