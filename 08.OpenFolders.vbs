Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim wsh : Set wsh = CreateObject("WScript.Shell")

Dim curDir : Set curDir = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))

'�ݒ�t�@�C���̓ǂݍ���
Dim settingFileName : settingFileName = "settings\folders.txt"
settingFileName = fso.BuildPath(curDir.Path, settingFileName)

If fso.FileExists(settingFileName) <> true then
	MsgBox "�t�@�C�������݂��܂���" & settingFileName
	WScript.Quit
End If

'Dim folderName()
Redim folderName(0)

Dim istrm : Set istrm = fso.OpenTextFile(settingFileName)

Do Until istrm.AtEndOfStream
	folderName(UBound(folderName)) = istrm.ReadLine()
	Redim Preserve folderName(UBound(folderName) + 1)
Loop

'�t�H���_���J��
Dim line
For line = 0 To UBound(folderName) - 1
	wsh.Run("explorer " + folderName(line))
Next
