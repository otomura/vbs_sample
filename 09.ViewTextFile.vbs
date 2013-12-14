Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))

Dim strTarget : strTarget = "settings\sample.txt"

Dim objTargetFile : Set objTargetFile = fso.GetFile(fso.BuildPath(curDir.Path, strTarget))

Dim istrm : Set istrm = objTargetFile.OpenAsTextStream

Dim result 
Dim iLineCount : iLineCount = 1
Do While istrm.AtEndOfStream = False
	' 3 Œ…‚Å 0–„‚ß
	Do While Len(iLineCount) < 3
		iLineCount = "0" & iLineCount
	Loop
	
	result = result & iLIneCount & ":" & istrm.ReadLine() & vbCrLf
	iLineCount = iLineCount + 1
	
Loop

WScript.Echo result
