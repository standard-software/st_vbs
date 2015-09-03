'--------------------------------------------------
'st.vbs
'--------------------------------------------------
'ModuleName:    SetShortcutLink Module
'FileName:      SetShortcutLink.vbs
'--------------------------------------------------
'Version:       2015/09/01
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

'--------------------------------------------------
Dim ShortcutLinkFile_SetFlag
ShortcutLinkFile_SetFlag = True 
'ShortcutLinkFile_SetFlag = False 

'--------------------------------------------------
'Set Script File Name.
Dim TargetSourceFileName: TargetSourceFileName = _
    "Project01.vbs"

'Set ShortcutLink File Name.
Dim ShortcutLinkFileName: ShortcutLinkFileName = _
    "Project01.lnk"

'StartMenu\Programsの中の場合は プログラムグループを指定するために
'"Project01\Project01.lnk" と指定するのがよいでしょう。
'--------------------------------------------------

Dim ShortcutLinkFilePathList: ShortcutLinkFilePathList = ""
Dim DeleteFolderPathList: DeleteFolderPathList = ""

ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
    fso.BuildPath(SendToFolderPath, ShortcutLinkFileName)))

ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
    fso.BuildPath(DesktopFolderPath, ShortcutLinkFileName)))

'ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
'    fso.BuildPath(StartMenuProgramsFolderPath, "Project01\" + ShortcutLinkFileName)))
'DeleteFolderPathList = StringCombine(vbCrLf, Array(DeleteFolderPathList, _
'    fso.BuildPath(StartMenuProgramsFolderPath, "Project01")))
'
'ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
'    fso.BuildPath(StartMenuFolderPath, ShortcutLinkFileName)))
'
'ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
'    fso.BuildPath(StartUpFolderPath, ShortcutLinkFileName)))
'--------------------------------------------------


Call Main

Sub Main
Do
    Dim TargetSourceFilePath: TargetSourceFilePath = _
        fso.BuildPath( _
            fso.GetParentFolderName(WScript.ScriptFullName), _
            TargetSourceFileName)

    If fso.FileExists(TargetSourceFilePath) = False Then
        Call MsgBox(StringCombine(vbCrLf, Array( _
            "ファイルが存在しません。", _
            TargetSourceFilePath )))
        Exit Do
    End If

    Dim I
    If ShortcutLinkFile_SetFlag = False Then

        Dim DeleteFolderPathArray
        DeleteFolderPathArray = Split(DeleteFolderPathList, vbCrLf)
        For I = 0 To ArrayCount(DeleteFolderPathArray) - 1
        Do
            Dim DeleteFolderPath: DeleteFolderPath = _
                DeleteFolderPathArray(I)
            Dim DeleteFolderParentFolderPath: DeleteFolderParentFolderPath = _
                fso.GetParentFolderName(DeleteFolderPath)

            If FolderHasSubItem(DeleteFolderPath) Then
                Call fso.DeleteFolder(DeleteFolderPath)
            End If

            If fso.FolderExists(DeleteFolderParentFolderPath) Then
                Call ShellFileOpen( _
                    DeleteFolderParentFolderPath, _
                    vbNormalFocus, True)
            End If

        Loop While False
        Next
    End If

    Dim ShortcutLinkFilePathArray
    ShortcutLinkFilePathArray = Split(ShortcutLinkFilePathList, vbCrLf)
    For I = 0 To ArrayCount(ShortcutLinkFilePathArray) - 1
    Do
        Dim ShortcutLinkFilePath: ShortcutLinkFilePath = _
            ShortcutLinkFilePathArray(I)
        Dim ShortcutLinkFileParentFolderPath: ShortcutLinkFileParentFolderPath = _
            fso.GetParentFolderName(ShortcutLinkFilePath)

        If fso.FileExists(ShortcutLinkFilePath) Then
            Call fso.DeleteFile(ShortcutLinkFilePath)
        End If

        If ShortcutLinkFile_SetFlag Then
            Call ForceCreateFolder(ShortcutLinkFileParentFolderPath)

            Call CreateShortcutFile( _
                ShortcutLinkFilePath, _
                TargetSourceFilePath, _
                ScriptProgramFilePath, 2, _
                "")
        End If

        If fso.FolderExists(ShortcutLinkFileParentFolderPath) Then
            Call ShellFileOpen( _
                ShortcutLinkFileParentFolderPath, _
                vbNormalFocus, True)
        End If
    Loop While False
    Next

Loop While False
End Sub

