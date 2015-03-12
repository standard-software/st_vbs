'--------------------------------------------------
'Standard Software Library For VBScript
'
'ModuleName:    CreateNewProject.vbs.vbs
'--------------------------------------------------
'version:       2015/03/10
'--------------------------------------------------

Option Explicit

'--------------------------------------------------
'��Include Standard Software Library
'--------------------------------------------------
'FileName�ɂ͑��΃A�h���X���w��\
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
    Dim NewProjectName: NewProjectName = InputBox("�V�����v���W�F�N�g������͂��Ă��������B")

    If NewProjectName = "" Then
        Call MsgBox("���͂�����܂���B"+ vbCrLf + _
            "�������I�����܂��B")
        Exit Do
    End If

    Dim NewProjectFolderPath: NewProjectFolderPath = _
        AbsolutePath(ScriptFolderPath, "..\..\..\..\..\" + NewProjectName)
    If fso.FolderExists(NewProjectFolderPath) Then
        Call MsgBox("�v���W�F�N�g���͊��Ɏg���Ă��܂��B"+ vbCrLf + _
            "�������I�����܂��B")
        Exit Do
    End If

    '�v���W�F�N�g�t�H���_�ꎮ�R�s�[
    Call ForceCreateFolder(NewProjectFolderPath)
    Call CopyFolderIgnorePath( _
        AbsolutePath(ScriptFolderPath, "..\..\..\Project01\Source"), _
        PathCombine(Array(NewProjectFolderPath, "Source")), _
        "CreateNewProject.vbs,SupportTool.ini")

    '�v���W�F�N�g�t�@�C�����̕ύX
    Call fso.MoveFile( _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01", "Project01.vbs")), _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01", NewProjectName + ".vbs")))
    Call fso.MoveFolder( _
        PathCombine(Array(NewProjectFolderPath, "Source", "Project01")), _
        PathCombine(Array(NewProjectFolderPath, "Source", NewProjectName)))

    '�V�K�v���W�F�N�g�t�@�C���̃w�b�_�[���H
    Dim VbsFilePath: VbsFilePath = PathCombine(Array( _
        NewProjectFolderPath, "Source", NewProjectName, NewProjectName + ".vbs"))
    Dim FileText: FileText = LoadTextFile(VbsFilePath, "SHIFT_JIS")
    FileText = Replace(FileText, "Project01.vbs", NewProjectName + ".vbs")
    FileText = Replace(FileText, "YYYY/MM/DD", FormatYYYY_MM_DD(Now, "/"))
    Call SaveTextFile(FileText, VbsFilePath, "SHIFT_JIS")

    'SupportTool�t�H���_�R�s�[
    Call CopyFolderIgnorePath( _
        AbsolutePath(ScriptFolderPath, "..\..\..\_SupportTool"), _
        PathCombine(Array(NewProjectFolderPath, "Source\_SupportTool")), _
        "Update_HereLib.vbs")

    'Ini�t�@�C���ݒ�
    Dim IniFilePath: IniFilePath = _
        PathCombine(Array(NewProjectFolderPath, "Source", "_SupportTool", "SupportTool.ini"))
    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)
    Call IniFile.WriteString("Common", "ProjectName", NewProjectName)

    Call IniFile.SectionIdentDelete("Update_HereLib", "LibrarySourceFilePath")
    Call IniFile.SectionIdentDelete("Update_HereLib", "LibraryDestFilePath")

    Call IniFile.WriteString("Update_ProjectLib", "LibrarySourceFilePath01", "..\..\..\StandardSoftwareLibrary_vbs\Source\StandardSoftwareLibrary_vbs\StandardSoftwareLibrary.vbs")
    Call IniFile.SectionIdentDelete("Update_ProjectLib", "LibrarySourceFilePath02")
    Call IniFile.SectionIdentDelete("Update_ProjectLib", "LibrarySourceFilePath03")
    Call IniFile.WriteString("Update_ProjectLib", "LibraryDestFolderPath", "..\" + NewProjectName + "\Lib")

    Call IniFile.WriteString("Update_SupportTool", "SupportToolSourcePath", "..\..\..\StandardSoftwareLibrary_vbs\Source\_SupportTool")

    Call IniFile.WriteString("SourceBackup", "BackupSourceFolderPaths", "..\..\Source")
    Call IniFile.WriteString("SourceBackup", "BackupDestFolderPaths", "..\..\Backup\Source")
    Call IniFile.WriteString("SourceBackup", "BackupFolderLastYYYY_MM_DD", "False")

    Call IniFile.WriteString("MakeRelease", "ReleaseIgnoreFileFolderName", "")
    Call IniFile.WriteString("MakeRelease", "ReleaseIncludeFileFolderPath", "..\..\version.txt")
    Call IniFile.WriteString("MakeRelease", "ScriptEncoderExePath", "C:\Program Files\Windows Script Encoder\screnc.exe")
    Call IniFile.WriteString("MakeRelease", "ScriptEncodeTargets", "..\" + NewProjectName + "\" + NewProjectName + ".vbs")

    Call IniFile.WriteString("ReleaseInstall", "InstallParentFolderPath", "C:\Program Files")
    Call IniFile.WriteString("ReleaseInstall", "InstallOverWriteIgnoreFiles", "*.ini")

    IniFile.Update
    Set IniFile = Nothing

    '�o�[�W�����t�@�C���ݒu
    Dim VersionTxt: VersionTxt = _
        "��" + FormatYYYY_MM_DD(Now, "/") + "    ver 1.0.0" + vbCrLf + _
        "�E  �쐬"
    Call SaveTextFile(VersionTxt, _
        PathCombine(Array(NewProjectFolderPath, "version.txt")), _
        "Shift_JIS")

    Call MsgBox("�V�����v���W�F�N�g[" + NewProjectName + "]���쐬���܂����B" + vbCrLf + _
        "-----" + vbCrLf + _
        NewProjectFolderPath)

Loop While False
End Sub