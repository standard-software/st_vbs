'--------------------------------------------------
'st.vbs
'--------------------------------------------------
'ModuleName:    SetShortcutLink Module
'FileName:      SetShortcutLink.vbs
'--------------------------------------------------
'Version:       2015/08/24
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
Dim TargetSourceFileName: TargetSourceFileName = _
    "Project01.vbs"

Dim ShortcutLinkFileName: ShortcutLinkFileName = _
    "Project01.lnk"
'StartMenu\Programsの中の場合は プログラムグループを指定するために
'"Project01\Project01.lnk" と指定するのがよいでしょう。

Dim SpecialFolderPath
SpecialFolderPath = SendToFolderPath
'SpecialFolderPath = DesktopFolderPath
'SpecialFolderPath = StartMenuFolderPath
'SpecialFolderPath = StartMenuProgramsFolderPath
'SpecialFolderPath = StartUpFolderPath

Dim ShortcutLinkFilePath: ShortcutLinkFilePath = _
    fso.BuildPath( _
        SpecialFolderPath, _
        ShortcutLinkFileName)

Call Main

Sub Main
Do
    Dim TargetSourceFilePath: TargetSourceFilePath = _
        fso.BuildPath( _
            fso.GetParentFolderName(WScript.ScriptFullName), _
            TargetSourceFileName)

    If fso.FileExists(TargetSourceFilePath) = False Then
        Call MsgBox( _
            StringCombine(vbCrLf, _
                Array("ファイルが存在しません。", _
                    TargetSourceFilePath)))
        Exit Do
    End If

    If fso.folderExists(SpecialFolderPath) = False Then
        Call MsgBox( _
            StringCombine(vbCrLf, _
                Array("フォルダが存在しません。", _
                    SpecialFolderPath)))
        Exit Do
    End If

    'Call MsgBox(ShortcutLinkFilePath)

    If fso.FileExists(ShortcutLinkFilePath) Then
        fso.DeleteFile(ShortcutLinkFilePath)
    End If

    Call CreateShortcutFile( _
        ShortcutLinkFilePath, _
        TargetSourceFilePath, _
        ScriptProgramFilePath + ",2", _
        "")

     Call ShellFileOpen( _
        fso.GetParentFolderName(ShortcutLinkFilePath), _
        vbNormalFocus)

Loop While False
End Sub


