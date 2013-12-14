Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim reg : Set reg = new RegExp

Dim targetFileName : targetFileName = ".\settings\sample.txt"
Dim targetFile : Set targetFile = fso.GetFile(targetFileName)

Dim istrm : Set istrm = targetFile.OpenAsTextStream()
Dim ostrm : Set ostrm = fso.CreateTextFile(".\settings\sample.txt_r",1)

reg.Pattern = "(.*ƒtƒ@ƒCƒ‹.*)\r\n"
reg.Global = true
ostrm.Write(reg.Replace(istrm.ReadAll(),"‚±‚±!! $1 " & vbCrLf ))

istrm.Close()
ostrm.Close()
