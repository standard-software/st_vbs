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

'------------------------------
'◇メイン処理
'------------------------------
Call Main

Sub Main
    Dim MessageText: MessageText = ""

    Dim IniFilePath: IniFilePath = _
        PathCombine(Array(ScriptFolderPath, "SupportTool.ini"))

    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)

    '--------------------
    '・設定読込
    '--------------------
    Dim Library_Source_Path: Library_Source_Path = _
        IniFile.ReadString("Option", "LibrarySourcePath", "")

    Dim ProjectName: ProjectName = _
        IniFile.ReadString("Option", "ProjectName", "")
    '--------------------

    Dim NowValue: NowValue = Now
    Dim ReleaseFolderPath: ReleaseFolderPath = _
        PathCombine(Array( _
            "..\..\\Release", _
            "Recent", _
            ProjectName))

    Dim SourceFolderPath: SourceFolderPath = _
        "..\..\Source\" + _
        ProjectName
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

