Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
'MsgBox curDir.Path

'�t�@�C����
Dim strFileName : strFileName = "sample.txt"

'����������
Dim strSearchWord : strSearchWord = "�f�B���N�g��"

'�t�@�C�����J��
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
	str = "�Y���Ȃ�"
End If

MsgBox "Word: " & strSearchWord & vbCrLf & str
