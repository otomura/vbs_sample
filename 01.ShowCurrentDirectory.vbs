' ShowCurrentDirectory
' �X�N���v�g�̂���f�B���N�g���̒��g��\������

Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

'�X�N���v�g������t�H���_�̃I�u�W�F�N�g���擾
Dim objCurrentDirectory
Set objCurrentDirectory = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))

'showDirectory �֐��Ăяo��
showDirectory objCurrentDirectory

'�f�B���N�g�����󂯎���ăf�B���N�g������
'�f�B���N�g���ƃt�@�C����\������֐�
Sub showDirectory(objDirectory)
	
	Dim strOut

	'�f�B���N�g���̕\��
	Dim folder
	For Each folder In objDirectory.SubFolders
		strOut = strOut & "/" & folder.Name & vbCrLf
	Next
	
	Dim file
	For Each file In objDirectory.Files
		strOut = strOut & file.Name & vbCrLf
	Next
	
	MsgBox strOut
	
End Sub
