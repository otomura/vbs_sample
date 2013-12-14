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
		
		'自ファイルの名前は変えない
		If file.Name <> fso.GetFile(WScript.ScriptFullName).Name Then 
			'直接 LCase を使うと、同名のファイルがある、とエラーになるため、
			'一時ファイル名を使う
			'NG : file.Name = LCase(file.Name)
			Dim oldFileName : oldFileName = file.Name
			file.Name = fso.GetTempName()
			file.Name = LCase(oldFileName)
			MsgBox oldFileName & "->" & file.Name
		End If
		
	Next
	
End Sub
