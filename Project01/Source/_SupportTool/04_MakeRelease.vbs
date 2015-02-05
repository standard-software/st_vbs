Option Explicit

Dim ProjectFolderName: ProjectFolderName = _
    "Project01"

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
    Dim MessageText: MessageText = ""

    Dim NowValue: NowValue = Now
    Dim ReleaseFolderPath: ReleaseFolderPath = _
        PathCombine(Array( _
            "..\..\\Release", _
            "Recent", _
            ProjectFolderName))

    Dim SourceFolderPath: SourceFolderPath = _
        "..\..\Source\" + _
        ProjectFolderName
    SourceFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, SourceFolderPath)

    If not fso.FolderExists(SourceFolderPath) Then
        WScript.Echo _
            "コピー元フォルダが見つかりません" + vbCrLF + _
            SourceFolderPath
        Exit Sub
    End If

    'フォルダ再生成コピー
    Call ReCreateCopyFolder(SourceFolderPath, ReleaseFolderPath)

    MessageText = MessageText + _
        fso.GetFileName(SourceFolderPath) + vbCrLf

    'バージョン情報ファイル
    Dim VersionInfoFilePath: VersionInfoFilePath = _
        "..\..\version.txt"
    VersionInfoFilePath = _
        AbsoluteFilePath(ScriptFolderPath, VersionInfoFilePath)
    If fso.FileExists(VersionInfoFilePath) Then
        Call fso.CopyFile( _
            VersionInfoFilePath, _
                IncludeLastPathDelim(ReleaseFolderPath) + _
                fso.GetFileName(VersionInfoFilePath))
    End If

    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)

End Sub

