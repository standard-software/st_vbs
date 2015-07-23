'--------------------------------------------------
'Standard Software Library For Windows VBScript
'
'ModuleName:    TestStandardSoftwareLibrary.vbs
'--------------------------------------------------
'Version:       2015/07/23
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
Call Include(".\st.vbs")
'--------------------------------------------------

Call test

