Dim targetMonth : targetMonth = InputBox("���ׂ���������͂��Ă�������")

If IsNumeric(targetMonth) = False Then
	WScript.Echo "��������͂��Ă�������(1�`12)"
	WScript.Quit
End If

If CLng(targetMonth) < 1 Or CLng(targetMonth) > 12 Then
	WScript.Echo "��������͂��Ă�������(1�`12)"
	WScript.Quit
End If

WScript.Echo Year(Now) & "�N" & targetMonth & "���̍ŏI����" & Day(DateAdd("d", DateSerial(Year(Now), targetMonth + 1, 1), -1)) & "���ł�"
