Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim wsh : Set wsh = CreateObject("WScript.Shell")
Dim reg : Set reg = new RegExp
Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))

Dim targetDirName
targetDirName = curDir.Path

Dim targetDir : Set targetDir = fso.GetFolder(targetDirName)

Dim searchString
searchString = InputBox("ŒŸõ•¶Žš—ñ‚ð“ü—Í‚µ‚Ä‚­‚¾‚³‚¢")
If searchString = "" Then
	WScript.Quit
End If

reg.Pattern = searchString
reg.IgnoreCase = True
reg.Global = True

Dim result
Dim file
For Each file In targetDir.Files

	Dim matches : Set matches = reg.Execute(file.Name)

	Dim detail : detail = ""
	Dim match : Set match = Nothing
	For Each match In matches
		detail = detail & match.Value & ":" & match.FirstIndex & " "
	Next
	
	If matches.Count > 0 then
		result = result & file.Name & " " & detail & vbCrLf
	End If
	
Next

WScript.Echo result
