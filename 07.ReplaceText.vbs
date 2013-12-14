Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))

'�ϊ��Ώە���
Dim strTarget : strTarget = "�f�B���N�g��"
Dim strReplace : strReplace = "�t�H���_"

Dim strTargetFileName : strTargetFileName = "sample.txt"
Dim strTargetFilePath : strTargetFilePath = fso.BuildPath(curDir.Path,strTargetFileName)

If fso.FileExists(strTargetFilePath) = False Then
	MsgBox "�t�@�C�������݂��܂��� : " & strTargetFilePath
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
