Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")

' �ݒ�
Dim iNumFiles : iNumFiles = 2 '�쐬�t�@�C����
Dim iFileSize : iFileSize = 100 '�쐬�t�@�C���T�C�Y

' main
Dim curFolder : Set curFolder = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
CreateSizedTempFilesInFolder curFolder, iNumFiles, iFileSize
WScript.Quit

'�w��̃t�H���_�ɃT�C�Y���w�肵���e���|�����t�@�C�����쐬����
Sub CreateSizedTempFilesInFolder(objFolder, iNum, iSize)

	Dim i
	For i = 1 To iNum
		CreateSizedFile fso.BuildPath(objFolder.Path,fso.GetTempName()), iSize
	Next
	
End Sub

'�w��T�C�Y�̃t�@�C�����쐬����
Sub CreateSizedFile(strFilePath, iSize)

	'MsgBox "strFilePath" & " " & strFilePath
	Dim strCmnd : strCmnd = "fsutil file createnew """ & strFilePath & """ " & iSize
	'MsgBox strCmnd
	shell.Run strCmnd
	
End Sub
