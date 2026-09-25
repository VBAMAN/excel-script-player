Attribute VB_Name = "PLY_01_Config"
Option Explicit

'--------------------------------------
' 時刻カウンター表示設定
'--------------------------------------
Public ShowTimeCounter As Boolean

Public Enum ENDROLL_TIME_MODE

    TIME_IGNORE = 0       ' A: カウンターを処理しない
    TIME_WAIT = 1         ' B: 相対時間として待つ
    TIME_TIMELINE = 2     ' C: 開始からの絶対時間

End Enum

Public TimeCounterMode As ENDROLL_TIME_MODE

' タイムカウンターの行番号 → 秒数を保持する。
Public TimeCounterDictionary As Object

' タイムカウンター処理用の現在行。
' ※現在の主要処理では直接使用していないが、状態保持用として残す。
Public TimeCounterLine As Long

Public TimeWaitLine As Long
Public TimeWaitStart As Long
Public TimeWaitActive As Boolean

Public TimeLineReady As Boolean
Public TimeLineStart As Date
Public TimeLineStarted As Boolean


' 直接メモリーへ Runモード
Public ScriptText() As String


'--------------------------------------
' テキスト管理
'--------------------------------------
Public TextLine As Long
Public TextLastLine As Long

'----------------------------------
' TimeLine開始位置の解釈モード
'----------------------------------
Public Enum ENDROLL_TIMELINE_MODE

    '1行目の時刻を実時間として使用
    TIMELINE_FIRST_AS_START = 0

    '1行目の時刻を無視して開始
    TIMELINE_FIRST_AS_ZERO = 1

    '指定秒数後を開始基準として使用
    TIMELINE_START_OFFSET = 2

End Enum

Public TimeLineMode As ENDROLL_TIMELINE_MODE

Public TimeLineStartOffset As Long
Public TimeLineFirstSeconds As Long

' 初期表示時に使用する先頭行数。
Public FirstLineRows As Long

'----------------------------------
' タイムカウンター形式
'
' 8 = "00:00:00"
' 9 = "00:00:00:"
'----------------------------------
Public TimeCounterFormat As Long
Public Const TIME_COUNTER_FORMAT_8 As Long = 8
Public Const TIME_COUNTER_FORMAT_9 As Long = 9

' SCRIPT終了後に先頭へ戻って再生するか。
Public ScriptRepeatMode As Boolean


Public Enum TEXT_DISPLAY_MODE
    DISPLAY_MESSAGE = 0
    DISPLAY_ENDROLL = 1
End Enum

' メッセージ表示 / エンドロール表示の選択。
Public DisplayMode As TEXT_DISPLAY_MODE

Public DefaultWaitSeconds As Double
Public TimeWaitSeconds As Double

'-----------------------------
' タイマー管理用変数
'-----------------------------

Public sTime As Long


' Validateを有効にする設定値。
Public ValidateEnabled As Boolean


Public ScriptStartLine As Long
Public ScriptEndLine As Long

' LABEL 位置用

Public LabelOffsetY As Long
Public LabelOffsetX As Long

Public RepeatWaitSeconds As Double
Public RepeatWaitActive As Boolean
Public RepeatWaitStart As Long


