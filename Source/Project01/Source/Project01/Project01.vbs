'--------------------------------------------------
'Project01
'--------------------------------------------------
'ModuleName:    Project01.vbs
'--------------------------------------------------
'Version:       YYYY/MM/DD
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'Å°Include Standard Software Library
'--------------------------------------------------
Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.GetParentFolderName(WScript.ScriptFullName) _
        + "\" + FileName, 1)
    Call ExecuteGlobal(Stream.ReadAll())
    Call Stream.Close
End Sub
'--------------------------------------------------
Call Include(".\Lib\st.vbs")
'--------------------------------------------------

Call Main

Sub Main
Do
    Call test
Loop While False
End Sub
