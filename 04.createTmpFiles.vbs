Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")

' 設定
Dim iNumFiles : iNumFiles = 2 '作成ファイル数
Dim iFileSize : iFileSize = 100 '作成ファイルサイズ

' main
Dim curFolder : Set curFolder = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
CreateSizedTempFilesInFolder curFolder, iNumFiles, iFileSize
WScript.Quit

'指定のフォルダにサイズを指定したテンポラリファイルを作成する
Sub CreateSizedTempFilesInFolder(objFolder, iNum, iSize)

	Dim i
	For i = 1 To iNum
		CreateSizedFile fso.BuildPath(objFolder.Path,fso.GetTempName()), iSize
	Next
	
End Sub

'指定サイズのファイルを作成する
Sub CreateSizedFile(strFilePath, iSize)

	'MsgBox "strFilePath" & " " & strFilePath
	Dim strCmnd : strCmnd = "fsutil file createnew """ & strFilePath & """ " & iSize
	'MsgBox strCmnd
	shell.Run strCmnd
	
End Sub
