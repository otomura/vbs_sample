Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))

'変換対象文字
Dim strTarget : strTarget = "ディレクトリ"
Dim strReplace : strReplace = "フォルダ"

Dim strTargetFileName : strTargetFileName = "sample.txt"
Dim strTargetFilePath : strTargetFilePath = fso.BuildPath(curDir.Path,strTargetFileName)

If fso.FileExists(strTargetFilePath) = False Then
	MsgBox "ファイルが存在しません : " & strTargetFilePath
	WScript.Quit
End If

Dim outStrm : Set outStrm = fso.CreateTextFile(strTargetFilePath & "_replaced")

Dim targetFile : Set targetFile = fso.GetFile(strTargetFilePath)
Dim inStrm : Set inStrm = targetFile.OpenAsTextStream

Do Until inStrm.AtEndOfStream
	outStrm.WriteLine Replace(inStrm.ReadLine, strTarget, strReplace)
Loop

inStrm.Close()
outStrm.Close()
