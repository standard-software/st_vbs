Option Explicit

Dim ProjectFolderName: ProjectFolderName = _
    "Project01"

Dim ProgramFilesFolderPath: ProgramFilesFolderPath = _
    "C:\Program Files"

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

    Dim InstallFolderPath: InstallFolderPath = _
        PathCombine(Array( _
            ProgramFilesFolderPath, _
            ProjectFolderName))

    'フォルダ再生成コピー
    Call ReCreateCopyFolder(ReleaseFolderPath, InstallFolderPath)

    MessageText = MessageText + _
        fso.GetFileName(InstallFolderPath) + vbCrLf

    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)
End Sub

