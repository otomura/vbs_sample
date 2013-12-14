'ƒtƒ@ƒCƒ‹Œ‹‡
'ˆø”‚ª 3 ‚Â Œ‹‡Œ³1 Œ‹‡Œ³2 Œ‹‡æ

Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim args : Set args = WScript.Arguments

If args.Count <> 3 Then
	WScript.Echo "Usage : 15.Concatinate.vbs src1 src2 dest"
	WScript.Quit
End If

Dim ostrm : Set ostrm = fso.CreateTextFile("./" & args(2))
DIm istrm1 : Set istrm1 = fso.OpenTextFile("./" & args(0))
DIm istrm2 : Set istrm2 = fso.OpenTextFile("./" & args(1))

Do Until istrm1.AtEndOfStream
	ostrm.WriteLine(istrm1.ReadLine())
Loop

Do Until istrm2.AtEndOfStream
	ostrm.WriteLine(istrm2.ReadLine())
Loop

ostrm.Close

WScript.Quit
