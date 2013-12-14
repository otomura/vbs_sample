Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

Dim strTargetDirPath : strTargetDirPath = InputBox("�����f�B���N�g�����w�肵�Ă�������")
'Dim strTargetDirPath : strTargetDirPath = fso.GetParentFolderName(WScript.ScriptFullName)

If fso.FolderExists(strTargetDirPath) = False Then
	MsgBox "�f�B���N�g�������݂��܂���"
	WScript.Quit
End If
Dim objTargetDir : Set objTargetDir = fso.GetFolder(strTargetDirPath)

'�ŐV�̃t�@�C��
Dim objRecentFile : Set objRecentFile = Nothing
'�ŐV�̃t�@�C���̍X�V����
Dim dtmMostRecentFile
'��r�ΏۂƂȂ����S�t�@�C����
Dim allFileList
'�����Ώۃt�@�C����
Dim fileListSize : fileListSize = 0
'�����Ώۃt�@�C���ő吔
Dim maxFileListSize : maxFileListSize = 10

'main
findMostRecentFile(objTargetDir)
Msgbox "FileList : " & vbCr & _
allFileList & vbCr & vbCr & _
"Most Recent Modified File : " & vbCrLf &_
objRecentFile.Path & " " & objRecentFile.DateLastModified

'�f�B���N�g�����ł����Ƃ��X�V���t�̐V���� File �I�u�W�F�N�g��Ԃ��֐�
Sub findMostRecentFile(targetDir)

	'txt �g���q�t�@�C�����擾
	Dim file
	For Each file In targetDir.Files
		If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
			allFileList = allFileList & file.Name & " " & file.DateLastModified & vbCrLf
			If file.DateLastModified > dtmMostRecentFile Then
				Set objRecentFile = file
				dtmMostRecentFile = file.DateLastModified
			End If
			fileListSize = fileListSize + 1
			If fileListSize >= maxFileListSize Then
				Exit Sub
			End If
		End If
	Next
	
	Dim folder
	For Each folder IN targetDir.SubFolders
		findMostRecentFile(folder)
		If fileListSize >= maxFileListSize Then
			Exit Sub
		End If
	Next
	
End Sub
