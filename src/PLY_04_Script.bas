Attribute VB_Name = "PLY_04_Script"
Option Explicit

'==================================================
' PLY_04_Script
'
' Scriptの読み込み・Script範囲管理
'
' Validateとは分離する。
'
' RunではValidateを行わないため、
' Playerが使用するScriptStartLine /
' ScriptEndLineをここで確定する。
'==================================================


'==================================================
' Script範囲設定
'
' <SCRIPT>
'    ↓
' 表示対象
'    ↓
' </SCRIPT>
'
' <SCRIPT> / </SCRIPT> は
' 行頭の空白・タブを許可する。
'
' 例
'
'   <SCRIPT>          OK
'    <SCRIPT>         OK
'   [TAB]<SCRIPT>     OK
'
' SCRIPTタグそのものの正当性は
' ValidateScriptStructureで確認する。
'
' このルーチンの役割は、
' Playerが使用する範囲を確定すること。
'==================================================
Public Function SetScriptRange() As Boolean

    Dim LineNo As Long
    Dim SourceText As String

    Dim StartPos As Long
    Dim EndPos As Long


    SetScriptRange = False


    '==================================
    ' 初期化
    '==================================
    ScriptStartLine = 0
    ScriptEndLine = 0

    StartPos = 0
    EndPos = 0


    '==================================
    ' Scriptを走査
    '==================================
    For LineNo = 1 To TextLastLine

        '----------------------------------
        ' 行頭の空白・タブを許可
        '----------------------------------
        SourceText = _
            LTrim$(ScriptText(LineNo))


        '==================================
        ' <SCRIPT>
        '==================================
        If Left$(SourceText, Len("<SCRIPT>")) = _
            "<SCRIPT>" Then

            '----------------------------------
            ' 最初の<SCRIPT>を保存
            '
            ' 複数存在する場合の検証は
            ' ValidateScriptStructureで行う。
            '----------------------------------
            If StartPos = 0 Then

                StartPos = LineNo

            End If

        End If


        '==================================
        ' </SCRIPT>
        '==================================
        If Left$(SourceText, Len("</SCRIPT>")) = _
            "</SCRIPT>" Then

            '----------------------------------
            ' 最初の</SCRIPT>を保存
            '
            ' 複数存在する場合の検証は
            ' ValidateScriptStructureで行う。
            '----------------------------------
            If EndPos = 0 Then

                EndPos = LineNo

            End If

        End If

    Next LineNo


    '==================================
    ' Script開始位置を設定
    '==================================
    ScriptStartLine = StartPos


    '==================================
    ' Script終了位置を設定
    '==================================
    ScriptEndLine = EndPos


    '==================================
    ' 結果確認
    '==================================
    If ScriptStartLine = 0 Then

        Debug.Print _
            "SetScriptRange : <SCRIPT> がありません。"

        Exit Function

    End If


    If ScriptEndLine = 0 Then

        Debug.Print _
            "SetScriptRange : </SCRIPT> がありません。"

        Exit Function

    End If


    '==================================
    ' 順序確認
    '
    ' これはPlayerが使用する範囲を
    ' 確定できるかどうかの最低限の確認。
    '
    ' 詳細な構造検証はValidate側で行う。
    '==================================
    If ScriptStartLine >= ScriptEndLine Then

        Debug.Print _
            "SetScriptRange : " & _
            "<SCRIPT> / </SCRIPT> の順序が不正です。"

        Exit Function

    End If


    '==================================
    ' OK
    '==================================
    Debug.Print _
        "SetScriptRange : OK" & _
        " / Start=" & ScriptStartLine & _
        " / End=" & ScriptEndLine

    Debug.Print _
        "  Display Start : " & _
        ScriptStartLine + 1

    SetScriptRange = True

End Function


Public Function IsScriptEndLine( _
    ByVal LineNo As Long) As Boolean

    Dim SourceText As String

    IsScriptEndLine = False

    SourceText = _
        ScriptText(LineNo)

    If InStr( _
        1, _
        SourceText, _
        "</SCRIPT>", _
        vbTextCompare) > 0 Then

        IsScriptEndLine = True

    End If

End Function
