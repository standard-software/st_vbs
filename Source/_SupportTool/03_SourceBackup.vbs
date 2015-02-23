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

    Dim NowValue: NowValue = Now
    Dim BackupFolderPath: BackupFolderPath = _
        PathCombine(Array( _
            "..\..\Backup\Source\", _
            FormatYYYY_MM_DD(NowValue, "-") + _
            "_" + _
            FormatHH_MM_SS(NowValue, "-")))

    Call ForceCreateFolder(BackupFolderPath)

    Dim SourceFolderPath: SourceFolderPath = _
        "..\..\Source"
    SourceFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, SourceFolderPath)

    Dim Folders: Folders = Split( _
        FolderPathListTopFolder(SourceFolderPath), vbCrLf)
    Dim Folder
    For Each Folder In Folders
        Call fso.CopyFolder(Folder, _
            PathCombine(Array(BackupFolderPath, _
            fso.GetFileName(Folder))), True)
        MessageText = MessageText + fso.GetFileName(Folder) + vbCrLf
    Next

    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)

End Sub

