Option Explicit
'�f�B���N�g�����܂邲�ƕʂ̏ꏊ�ɃR�s�[����
'�i�R�s�[���ƃR�s�[���InputBox�Ŏw��j
'# FileSysetmObject �� CopyFolder �g��Ȃ�"

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim objCurDir : Set objCurDir = fso.GetFile(WScript.ScriptFullName)

'�R�s�[���f�B���N�g��
Dim srcDirPath : srcDirPath = InputBox("�R�s�[���f�B���N�g���i��΃p�X�j���w�肵�Ă�������")
If fso.FolderExists(srcDirPath) = False Then 
	MsgBox "�R�s�[���f�B���N�g�������݂��܂���" & srcDirPath
	WScript.Quit
End If

'�R�s�[��f�B���N�g��
Dim srcDestPath : srcDestPath = InputBox("�R�s�[��f�B���N�g���i��΃p�X�j���w�肵�Ă�������")
If fso.FolderExists(srcDestPath) = False Then
	MsgBox "�t�H���_�����݂��Ȃ��̂ō쐬���܂� " & vbCrLf & srcDestPath
	fso.CreateFolder(srcDestPath)
End If

copyDirectory fso.GetFolder(srcDirPath), fso.GetFolder(srcDestPath)

' objSrcDir �̓��e�� objDestDir �ɍċA�I�ɃR�s�[����
' ������ Folder �I�u�W�F�N�g
Sub copyDirectory(objSrcDir, objDestDir)

	'�T�u�f�B���N�g���ɑ΂��鏈��
	Dim subdir
	For Each subdir In objSrcDir.SubFolders
	
		'�R�s�[��t�H���_�̍쐬
		Dim newDirPath : newDirPath= fso.BuildPath(objDestDir.Path, subdir.Name)
		'MsgBox newDirPath
		If fso.FolderExists(newDirPath) = True Then
			'���݂��Ȃ��ꍇ�͂Ȃɂ����Ȃ�
		Else
			'���݂��Ȃ��ꍇ�̓t�H���_�쐬����
			fso.CreateFolder(newDirPath)
		End If
		
		'�ċA�Ăяo��
		copyDirectory subdir, fso.GetFolder(newDirPath)
		
	Next
	
	'�f�B���N�g�����̃t�@�C���ɑ΂��鏈��
	Dim file
	For Each file In objSrcDir.Files
		file.Copy(objDestDir.Path & "\")
	Next
	
End Sub
WScript.Quit
