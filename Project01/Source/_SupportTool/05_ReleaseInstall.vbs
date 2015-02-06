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
            "..\..\Release", _
            "Recent", _
            ProjectFolderName))
    ReleaseFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, ReleaseFolderPath)

    Dim InstallFolderPath: InstallFolderPath = _
        PathCombine(Array( _
            ProgramFilesFolderPath, _
            ProjectFolderName))

    'フォルダ内ファイルコピー
    Dim FileList: FileList = _
        Split( _
            FilePathListSubFolder(ReleaseFolderPath), vbCrLf)

    Dim OverWrite
    Dim CopyDestFilePath
    Dim I
    For I = 0 To ArrayCount(FileList) - 1
    Do
        OverWrite = True
        '除外ファイル
        Dim IgnoreFiles: IgnoreFiles = "*.ini"
        If MatchText(LCase(FileList(I)), Split(LCase(IgnoreFiles), ",")) Then OverWrite = False

        CopyDestFilePath = _
            IncludeFirstStr( _
                ExcludeFirstStr(FileList(I), ReleaseFolderPath), _
                InstallFolderPath)
        '上書き禁止ならファイルがあったらコピーしない
        If OverWrite = False Then
            If fso.FileExists(CopyDestFilePath) then
                Exit Do
            End If
        End If

        Call ForceCreateFolder(fso.GetParentFolderName(CopyDestFilePath))
        Call fso.CopyFile( _
            FileList(I), CopyDestFilePath, True)
    Loop While False
    Next

    MessageText = MessageText + _
        fso.GetFileName(InstallFolderPath) + vbCrLf

    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)
End Sub

