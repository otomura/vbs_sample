Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
'MsgBox curDir.Path

'main
ChangeAllFilesNameLowCase curDir

Sub ChangeAllFilesNameLowCase(objDir)

	Dim file
	For Each file In objDir.Files
		'MsgBox file.Name
		
		'���t�@�C���̖��O�͕ς��Ȃ�
		If file.Name <> fso.GetFile(WScript.ScriptFullName).Name Then 
			'���� LCase ���g���ƁA�����̃t�@�C��������A�ƃG���[�ɂȂ邽�߁A
			'�ꎞ�t�@�C�������g��
			'NG : file.Name = LCase(file.Name)
			Dim oldFileName : oldFileName = file.Name
			file.Name = fso.GetTempName()
			file.Name = LCase(oldFileName)
			MsgBox oldFileName & "->" & file.Name
		End If
		
	Next
	
End Sub
