Attribute VB_Name = "PLY_05_Show"
Option Explicit

' テキストファイル

Public Const SCRIPT_FOLDER As String = "Script"
Public ENDROLL_FILE As String

'==================================================
' Run実行入口
'
' ファイル名を指定して直接再生する。
' Sheet1への読み込み、Validateは行わない。
'==================================================

Sub Test_Run()



Call Player_Run("EndRoll.txt")



End Sub

Public Sub Player_Run( _
    ByVal FileName As String)

    ENDROLL_FILE = FileName

    LabelOffsetY = 20
    LabelOffsetX = 20

    Player_Init

    If Not LoadEndRollText() Then
        Exit Sub
    End If
    
    '----------------------------------
    ' Script範囲設定
    '----------------------------------
    If Not SetScriptRange() Then
        Exit Sub
    End If

    
    '----------------------------------
    ' タイムカウンター作成
    '----------------------------------
    BuildTimeCounterDictionary

    '----------------------------------
    ' 再生開始位置
    '----------------------------------
    TextLine = ScriptStartLine + 1

    UserForm1.Show

End Sub


'==================================================
' プレビュー実行入口
'==================================================
Sub Run_preview()


    ENDROLL_FILE = "EndRoll.txt"

    Call preview


End Sub

Public Sub preview()

    '----------------------------------
    ' ラベル位置
    '----------------------------------
    LabelOffsetY = 20
    LabelOffsetX = 20


    '----------------------------------
    ' Player初期化
    '----------------------------------
    Player_Init


    '----------------------------------
    ' スクリプト読み込み
    '----------------------------------
    If Not LoadEndRollText() Then
        Exit Sub
    End If


    Debug.Print _
        "After Load : " & TextLastLine


    '----------------------------------
    ' Script範囲設定
    '
    ' Playerが使用する
    ' ScriptStartLine / ScriptEndLine
    ' を確定する。
    '----------------------------------
    If Not SetScriptRange() Then

        Debug.Print _
            "SetScriptRange : NG"

        Exit Sub

    End If


    '----------------------------------
    ' SCRIPT構造 Validate
    '----------------------------------
    If Not ValidateScriptStructure() Then

        Debug.Print _
            "ValidateScriptStructure : NG"

        Exit Sub

    End If

    Debug.Print _
        "ValidateScriptStructure : OK"


    '----------------------------------
    ' スクリプト開始位置
    '
    ' ScriptStartLineが確定した後に設定する。
    '----------------------------------
    TextLine = ScriptStartLine + 1


    '----------------------------------
    ' タグ Validate
    '----------------------------------
    If Not ValidateTag() Then

        Debug.Print _
            "ValidateTag : NG"

        Exit Sub

    End If

    Debug.Print _
        "ValidateTag : OK"


    '----------------------------------
    ' TimeCounter Validate
    '----------------------------------
    If Not ValidateTimeCounter() Then

        MsgBox "TimeCounter"

        Debug.Print _
            "ValidateTimeCounter : NG"

        Exit Sub

    End If


    '----------------------------------
    ' TIME Validate
    '----------------------------------
    If Not ValidateTIME() Then

        MsgBox "ValidateTIME"

        Debug.Print _
            "ValidateTIME : NG"

        Exit Sub

    End If


    '----------------------------------
    ' タイムカウンター作成
    '----------------------------------
    BuildTimeCounterDictionary


    '----------------------------------
    ' Timeline Validate
    '----------------------------------
    If TimeCounterMode = TIME_TIMELINE Then

        If Not ValidateTimeline() Then

            MsgBox "ValidateTimeline"

            Debug.Print _
                "ValidateTimeline : NG"

            Exit Sub

        End If

    End If


    '----------------------------------
    ' UserForm表示
    '----------------------------------
    UserForm1.Show

End Sub


'======================================
' テキストファイルをメモリへ読み込む
'======================================

Private Function LoadEndRollText() As Boolean

    Dim FilePath As String
    Dim FileNo As Integer
    Dim LineText As String
    Dim RowNo As Long

    LoadEndRollText = False

    FilePath = ThisWorkbook.Path _
             & "\" _
             & SCRIPT_FOLDER _
             & "\" _
             & ENDROLL_FILE

    '==================================
    ' ファイル存在確認
    '==================================
    If Dir(FilePath) = "" Then

        MsgBox _
            "スクリプトファイルが見つかりません。" _
            & vbCrLf & vbCrLf _
            & FilePath, _
            vbExclamation

        Debug.Print _
            "Load : File Not Found"

        Exit Function

    End If

    '==================================
    ' ファイル読み込み
    '==================================
    On Error GoTo LoadError

    FileNo = FreeFile

    Open FilePath For Input As #FileNo

    ReDim ScriptText(1 To 1)

    RowNo = 1

    Do Until EOF(FileNo)

        Line Input #FileNo, LineText

        ReDim Preserve ScriptText(1 To RowNo)

        ScriptText(RowNo) = LineText

        RowNo = RowNo + 1

    Loop

    Close #FileNo

    TextLastLine = RowNo - 1

    Debug.Print "Load : " & TextLastLine & " lines"
    Debug.Print "Script Start : " & ScriptStartLine
    Debug.Print "Script End   : " & ScriptEndLine

    LoadEndRollText = True
    
    
    ScriptStartLine = 0
    ScriptEndLine = 0

    Exit Function


LoadError:

    On Error Resume Next

    Close #FileNo

    MsgBox _
        "スクリプトファイルを読み込めませんでした。" _
        & vbCrLf & vbCrLf _
        & FilePath _
        & vbCrLf & vbCrLf _
        & "Error " & Err.Number _
        & ": " & Err.Description, _
        vbExclamation

    Debug.Print _
        "Load Error : " & _
        Err.Number & _
        " / " & _
        Err.Description

    LoadEndRollText = False

End Function
