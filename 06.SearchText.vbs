Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
'MsgBox curDir.Path

'ファイル名
Dim strFileName : strFileName = "sample.txt"

'検索文字列
Dim strSearchWord : strSearchWord = "ディレクトリ"

'ファイルを開く
Dim targetFile : Set targetFile = fso.GetFile(fso.BuildPath(curDir.Path, strFileName))
Dim inStrm : Set inStrm = targetFile.OpenAsTextStream

Dim str
Dim line : line = 1
Do Until inStrm.AtEndOfStream = True
	Dim pos : pos = InStr(inStrm.ReadLine(),strSearchWord)
	If pos > 0 Then
		str = str & "line: " & line & " pos: " & pos & vbCrLf
	End If
	line = line + 1
Loop

If str = "" Then 
	str = "該当なし"
End If

MsgBox "Word: " & strSearchWord & vbCrLf & str
