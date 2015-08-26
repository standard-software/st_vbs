'--------------------------------------------------
'st.vbs
'--------------------------------------------------
'ModuleName:    Test Module
'FileName:      st_test.vbs
'--------------------------------------------------
'Discription:   Standard Software Library For VBScript
'--------------------------------------------------
'Version:       2015/07/23
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'Å°Include st.vbs
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
Call Include(".\st.vbs")
'--------------------------------------------------

Call test


