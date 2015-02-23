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
        IniFile.ReadString("Option", "LibrarySourcePath01", "")
        
	Dim Library_Dest_Path: Library_Dest_Path = _
    	".\Lib\StandardSoftwareLibrary.vbs"
        
    '--------------------

    Dim SourcePath: SourcePath = _
        AbsoluteFilePath(ScriptFolderPath, Library_Source_Path)
    If not fso.FileExists(SourcePath) Then
        WScript.Echo _
            "コピー元ファイルが見つかりません" + vbCrLF + _
            SourcePath
        Exit Sub
    End If

    Dim DestPath: DestPath = _
        AbsoluteFilePath(ScriptFolderPath, Library_Dest_Path)
        
    Call ForceCreateFolder(fso.GetParentFolderName(DestPath))

    Call fso.CopyFile(SourcePath, DestPath)
    MessageText = SourcePath + vbCrLf + _
        ">> " + DestPath
    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)
End Sub