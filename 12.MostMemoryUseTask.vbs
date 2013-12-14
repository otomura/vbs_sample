Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim wsh : Set wsh = CreateObject("WScript.Shell")
Dim reg : Set reg = New RegExp

'tasklist �̎��s
Dim objExec :Set objExec = wsh.Exec("tasklist")

'���ʊi�[��
Dim memoryList()
Redim memoryList(0)
Dim processList()
Redim processlist(0)

'���K�\���ɂ��K�v�ȏ��̐؂�o��
'smss.exe                    1260 Console                 0        464 K

reg.Pattern = "([\w\s]+.exe)\s+[0-9]+\s\w+\s+[0-9]+\s+([0-9,]+)\sK"
reg.IgnoreCase = true

Do Until objExec.Stdout.AtEndOfStream
	Dim line : line = objExec.StdOut.ReadLine()
	If reg.Test(line) = true Then
		
		' �v���Z�X���̕ۑ�
		Redim Preserve processList(Ubound(processList) + 1)
		processList(Ubound(processList)) = reg.Replace(line,"$1")
		
		' �������g�p�ʂ̕ۑ�
		Redim Preserve memoryList(Ubound(memoryList) + 1)
		memoryList(Ubound(memoryList)) = reg.Replace(line,"$2")
		
	End If
Loop

'�ő�g�p�ʂ̃C���f�b�N�X�����߂�
Dim maxIndex : maxIndex = 1
Dim maxMemory : maxMemory = 0

Dim i
For i = 1 To Ubound(memoryList)
	
	If IsNumeric(memoryList(i)) <> true Then
		WSCript.Echo "���l�ł͂���܂���" & memoryList(i)
		WScript.Quit
	End If
	
	Dim checkValue : checkValue = CLng(memoryList(i))
	If checkValue > maxMemory Then
		maxMemory = checkValue
		maxIndex = i
	End If
Next

WScript.Echo "�ł��������g�p�ʂ��傫���v���Z�X�� " &  processList(maxIndex) & " (" & memoryList(maxIndex) & "K�g�p)�ł��B"

