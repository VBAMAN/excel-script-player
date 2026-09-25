Attribute VB_Name = "PLY_03_Init"
Option Explicit

'========================================
' ENDROLL_03_Init
'
' Player起動時の初期設定
'
' Run / Preview 共通
'========================================


Public Sub Player_Init()

    
    RepeatWaitSeconds = 8

    '----------------------------------
    ' 表示設定
    '----------------------------------
    ShowTimeCounter = True
    
    ShowTime = True
    
    ' DISPLAY_MESSAGE or DISPLAY_ENDROLL

    DisplayMode = DISPLAY_MESSAGE

    FirstLineRows = 1


    '----------------------------------
    ' 時間設定
    '----------------------------------
    
    ' TIME_TIMELINE or TIME_WAIT
    
    TimeCounterMode = TIME_TIMELINE

    DefaultWaitSeconds = 0.6

    TimeLineMode = _
        TIMELINE_FIRST_AS_ZERO


    '----------------------------------
    ' TimeCounter設定
    '----------------------------------
    TimeCounterFormat = _
        TIME_COUNTER_FORMAT_9


    '----------------------------------
    ' リピート設定
    '----------------------------------
    ScriptRepeatMode = False

End Sub

