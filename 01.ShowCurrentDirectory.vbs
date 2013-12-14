' ShowCurrentDirectory
' スクリプトのあるディレクトリの中身を表示する

Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

'スクリプトがあるフォルダのオブジェクトを取得
Dim objCurrentDirectory
Set objCurrentDirectory = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))

'showDirectory 関数呼び出し
showDirectory objCurrentDirectory

'ディレクトリを受け取ってディレクトリ内の
'ディレクトリとファイルを表示する関数
Sub showDirectory(objDirectory)
	
	Dim strOut

	'ディレクトリの表示
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
