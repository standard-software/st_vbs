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


'-----
'設定


'-----

Call Main

Sub Main
Do
    Dim Args
    Set Args = WScript.Arguments

    If Args.Count = 0 Then
        MsgBox "起動引数が指定されていません。" + vbCrLf + _
            "処理を停止します" 
        Exit Do
    End If

    Dim DownloadPath
    DownloadPath = PathCombine(Array( _
        GetEnvironmentalVariables("PROGRAMDATA"), _
        "StandardSoftware\DownloadFileOpen\Download"))
    Call ForceCreateFolder(DownloadPath)

    Dim CopyFileFlag
    CopyFileFlag = False

    Dim I
    For I = 0 to Args.Count - 1
        Dim FilePath: FilePath = Args(I)

        If UCase(fso.GetExtensionName(FilePath)) = UCase("lnk") then
            'ショートカットファイルの場合はオリジナルファイルを割り当てる
            FilePath = GetShortcutFileLinkPath(FilePath)
        End If

        If fso.FileExists(FilePath) Then

            '同名ファイルの削除処理
            Dim FileArray
            FileArray = Split(ExcludeLastStr( _
                GetFilePathListTopFolder(DownloadPath), vbCrLf), vbCrLf)
            Dim J
            For J = 0 to ArrayLength(FileArray) - 1
                If IsFirstStr(fso.GetFileName(FileArray(J)), _
                    fso.GetBaseName(FilePath)) Then
                    call fso.DeleteFile(FileArray(J), True)
                End IF
            Next

            CopyFileFlag = True
            Dim CopyToPath
            CopyToPath = IncludeLastPathDelim(DownloadPath) + _ 
                fso.GetBaseName(FilePath) + "_" + _
                FormatYYYYMMDDHHMMSS(Now()) + _
                "." + fso.GetExtensionName(FilePath)
            Call fso.CopyFile(FilePath, CopyToPath)

            Do While not fso.FileExists(CopyToPath)
            Loop

            '読み取り専用属性にする
            Dim File
            Set File = fso.GetFile(CopyToPath)
            File.Attributes = File.Attributes or 1

            Call ShellFileOpen(CopyToPath, vbNormalFocus, False)
        End If
    Next

    If CopyFileFlag = False Then
        MsgBox "起動引数にファイルが指定されていません。" + vbCrLf + _
            "処理を停止します" 
        Exit Do
    End If

    Loop While False
End Sub
