'--------------------------------------------------
'Standard Software Library For VBScript
'
'ModuleName:    Project01.vbs
'--------------------------------------------------
'version        2015/02/04
'--------------------------------------------------

Option Explicit

'--------------------------------------------------
'■Include Standard Software Library
'--------------------------------------------------
'FileNameには相対アドレスも指定可能
'--------------------------------------------------
'Include ".\Test\..\..\StandardSoftwareLibrary_vbs\StandardSoftwareLibrary.vbs"  
Call Include(".\Lib\StandardSoftwareLibrary.vbs")

Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.GetParentFolderName(WScript.ScriptFullName) _
        + "\" + FileName, 1)
    ExecuteGlobal Stream.ReadAll() 
    Call Stream.Close
End Sub
'--------------------------------------------------

Call Main

Sub Main
Do
    Dim NewProjectName: NewProjectName = InputBox("新しいプロジェクト名を入力してください。")

    If NewProjectName = "" Then
        Call MsgBox("入力がありません。"+ vbCrLf + _
            "処理を終了します。")
        Exit Do
    End If

    Dim NewProjectFolderPath: NewProjectFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, "..\..\..\..\" + NewProjectName)
    If fso.FolderExists(NewProjectFolderPath) Then
        Call MsgBox("プロジェクト名は既に使われています。"+ vbCrLf + _
            "処理を終了します。")
        Exit Do
    End If

    'プロジェクトフォルダ一式コピー
    Call ForceCreateFolder(NewProjectFolderPath)
    Call CopyFolderIgnorePath( _
        AbsoluteFilePath(ScriptFolderPath, "..\..\..\Project01\Source"), _
        PathCombine(Array(NewProjectFolderPath, "Source")), _
        "CreateNewProject.vbs,SupportTool.ini")

    'プロジェクトファイル名の変更
    Call fso.MoveFile( _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01", "Project01.vbs")), _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01", NewProjectName + ".vbs")))
    Call fso.MoveFolder( _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01")), _
        PathCombine(Array(NewProjectFolderPath, "Source", NewProjectName)))

    'Iniファイル設定
    Call CopyFile( _
        PathCombine(Array(ScriptFolderPath, "SupportTool.ini")), _
        PathCombine(Array(NewProjectFolderPath, "Source", "_SupportTool", "SupportTool.ini")))
    Dim IniFilePath: IniFilePath = _
        PathCombine(Array(NewProjectFolderPath, "Source", "_SupportTool", "SupportTool.ini"))
    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)
    Call IniFile.WriteString( _
        "Option", "ProjectName", NewProjectName)
    IniFile.Update
    Set IniFile = Nothing

    'バージョンファイル設置
    Dim VersionTxt: VersionTxt = _
        "◇" + FormatYYYY_MM_DD(Now, "/") + "    ver 1.0.0" + vbCrLf + _
        "・  作成"
    Call SaveTextFile(VersionTxt, _
        PathCombine(Array(NewProjectFolderPath, "version.txt")), _
        "Shift_JIS")

    Call MsgBox("新しいプロジェクト[" + NewProjectName + "]を作成しました。" + vbCrLf + _
        "-----" + vbCrLf + _
        NewProjectFolderPath)

Loop While False
End Sub
