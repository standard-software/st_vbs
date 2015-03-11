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
    Dim LibrarySourceFilePathList: LibrarySourceFilePathList = ""

    'LibrarySourcePath01/02/03...というIniファイル項目の読み取り
    Dim I: I = 1
    Do While IniFile.SectionIdentExists("Update_ProjectLib", "LibrarySourceFilePath" + LongToStrDigitZero(I, 2))
        LibrarySourceFilePathList = StringCombine(vbCrLf, Array(LibrarySourceFilePathList, _
            IniFile.ReadString("Update_ProjectLib", "LibrarySourceFilePath"  + LongToStrDigitZero(I, 2), "")))
        I = I + 1
    Loop
    If LibrarySourceFilePathList = "" Then
        WScript.Echo _
            "設定が読み取れていません"
        Exit Sub
    End If

    Dim ProjectName: ProjectName = _
        IniFile.ReadString("Common", "ProjectName", "")
    If ProjectName = "" Then
        WScript.Echo _
            "設定が読み取れていません"
        Exit Sub
    End If

    Dim LibraryDestFolderPath: LibraryDestFolderPath = _
        IniFile.ReadString("Update_ProjectLib", "LibraryDestFolderPath", "")
    If LibraryDestFolderPath = "" Then
        WScript.Echo _
            "設定が読み取れていません"
        Exit Sub
    End If
    '--------------------

    Dim DestFolderPath: DestFolderPath = _
        AbsolutePath(ScriptFolderPath, LibraryDestFolderPath)
    If not fso.FolderExists(DestFolderPath) Then
        WScript.Echo _
            "コピー先フォルダが見つかりません" + vbCrLF + _
            fso.GetParentFolderName(DestFolderPath)
        Exit Sub
    End If

    Dim FilePaths: FilePaths = Split(LibrarySourceFilePathList, vbCrLf)
    Dim FilePath
    For I = 0 To ArrayCount(FilePaths) - 1
        FilePath = FilePaths(I)

        Dim SourcePath: SourcePath = _
            AbsolutePath(ScriptFolderPath, FilePath)

        If not fso.FileExists(SourcePath) Then
            WScript.Echo _
                "コピー元ファイルが見つかりません" + vbCrLF + _
                SourcePath
            Exit Sub
        End If

        Call fso.CopyFile(SourcePath, IncludeLastPathDelim(DestFolderPath), True)
        MessageText = MessageText + _
            fso.GetFileName(SourcePath) + ">> " + fso.GetFileName(DestFolderPath) + vbCrLf
    Next

    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)
End Sub