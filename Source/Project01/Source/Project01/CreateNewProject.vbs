'--------------------------------------------------
'st_vbs
'--------------------------------------------------
'ModuleName:    CreateNewProject.vbs.vbs
'--------------------------------------------------
'Version:       2015/07/24
'--------------------------------------------------

Option Explicit

'--------------------------------------------------
'■Include st.vbs
'--------------------------------------------------
Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.BuildPath( _
            fso.GetParentFolderName(WScript.ScriptFullName), _
            FileName) , 1)
    Call ExecuteGlobal(Stream.ReadAll())
    Call Stream.Close
End Sub
'--------------------------------------------------
Call Include(".\Lib\st.vbs")
'--------------------------------------------------

Call Main

Sub Main
Do
    Dim NewProjectName: NewProjectName = InputBox("新しいプロジェクト名を入力してください。")

    'Test
    'Dim NewProjectName: NewProjectName = "NewProject01"

    If NewProjectName = "" Then
        Call MsgBox("入力がありません。"+ vbCrLf + _
            "処理を終了します。")
        Exit Do
    End If

    Dim NewProjectFolderPath: NewProjectFolderPath = _
        AbsolutePath(ScriptFolderPath, "..\..\..\..\..\" + NewProjectName)
    If fso.FolderExists(NewProjectFolderPath) Then
        Call MsgBox("プロジェクト名は既に使われています。"+ vbCrLf + _
            "処理を終了します。")
        Exit Do
    End If

    'プロジェクトフォルダ一式コピー
    Call ForceCreateFolder(NewProjectFolderPath)
    Call CopyFolderIgnorePath( _
        AbsolutePath(ScriptFolderPath, "..\..\..\Project01\Source"), _
        PathCombine(Array(NewProjectFolderPath, "Source")), _
        "CreateNewProject.vbs,SupportTool.ini", "")

    'プロジェクトファイル名の変更
    Call fso.MoveFile( _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01", "Project01.vbs")), _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01", NewProjectName + ".vbs")))
    Call fso.MoveFolder( _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01")), _
        PathCombine(Array(NewProjectFolderPath, "Source", NewProjectName)))

    'Tools内ファイル名の変更
    Call fso.MoveFile( _
        PathCombine(Array(NewProjectFolderPath, "Source", NewProjectName, "Tools", "SetShortcutLink_Project01.vbs")), _
        PathCombine(Array(NewProjectFolderPath, "Source", NewProjectName, "Tools", "SetShortcutLink_" + NewProjectName + ".vbs")))


    '新規プロジェクトファイルのヘッダー加工
    Dim VbsFilePath: VbsFilePath = PathCombine(Array( _
        NewProjectFolderPath, "Source", NewProjectName, NewProjectName + ".vbs"))
    Dim FileText: FileText = LoadTextFile(VbsFilePath, "SHIFT_JIS")
    FileText = Replace(FileText, "Project01.vbs", NewProjectName + ".vbs")
    FileText = Replace(FileText, "Project01", NewProjectName)
    FileText = Replace(FileText, "YYYY/MM/DD", FormatYYYY_MM_DD(Now, "/"))
    Call SaveTextFile(FileText, VbsFilePath, "SHIFT_JIS")

    'SupportToolフォルダコピー
    Call CopyFolderIgnorePath( _
        AbsolutePath(ScriptFolderPath, "..\..\..\_SupportTool"), _
        PathCombine(Array(NewProjectFolderPath, "Source\_SupportTool")), _
        "Update_HereLib.vbs", "")

    'Iniファイル設定
    Dim IniFilePath: IniFilePath = _
        PathCombine(Array(NewProjectFolderPath, "Source", "_SupportTool", "SupportTool.ini"))
    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)
    Call IniFile.WriteString("Common", "ProjectName", NewProjectName)

    Call IniFile.SectionIdentDelete("Update_HereLib", "LibrarySourceFilePath")
    Call IniFile.SectionIdentDelete("Update_HereLib", "LibraryDestFilePath")

    Call IniFile.WriteString("Update_ProjectLib", "LibrarySourceFilePath01", "..\..\..\StandardSoftwareLibrary_vbs\Source\StandardSoftwareLibrary_vbs\StandardSoftwareLibrary.vbs")
    Call IniFile.SectionIdentDelete("Update_ProjectLib", "LibrarySourceFilePath02")
    Call IniFile.SectionIdentDelete("Update_ProjectLib", "LibrarySourceFilePath03")
    Call IniFile.WriteString("Update_ProjectLib", "LibraryDestFolderPath", "..\" + NewProjectName + "\Lib")

    Call IniFile.WriteString("Update_SupportTool", "SupportToolSourcePath", "..\..\..\StandardSoftwareLibrary_vbs\Source\_SupportTool")

    Call IniFile.WriteString("SourceBackup", "BackupSourceFolderPaths", "..\..\Source")
    Call IniFile.WriteString("SourceBackup", "BackupDestFolderPaths", "..\..\Backup\Source")
    Call IniFile.WriteString("SourceBackup", "BackupFolderLastYYYY_MM_DD", "False")

    Call IniFile.WriteString("MakeRelease", "ReleaseIgnoreFileFolderName", "")
    Call IniFile.WriteString("MakeRelease", "ReleaseIncludeFileFolderPath", "..\..\version.txt")
    Call IniFile.WriteString("MakeRelease", "ScriptEncoderExePath", "C:\Program Files\Windows Script Encoder\screnc.exe")
    Call IniFile.WriteString("MakeRelease", "ScriptEncodeTargets", "..\" + NewProjectName + "\" + NewProjectName + ".vbs")

    Call IniFile.WriteString("ReleaseInstall", "InstallParentFolderPath", "C:\Program Files")
    Call IniFile.WriteString("ReleaseInstall", "InstallOverWriteIgnoreFiles", "*.ini")

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
