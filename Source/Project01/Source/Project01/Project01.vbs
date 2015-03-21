'--------------------------------------------------
'Standard Software Library For VBScript
'
'ModuleName:    Project01.vbs
'--------------------------------------------------
'Version:       YYYY/MM/DD
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'■Include Standard Software Library
'--------------------------------------------------
'FileNameには相対アドレスを指定可能
'--------------------------------------------------
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
Do
    Call test
Loop While False
End Sub
