Dim targetMonth : targetMonth = InputBox("調べたい月を入力してください")

If IsNumeric(targetMonth) = False Then
	WScript.Echo "数字を入力してください(1〜12)"
	WScript.Quit
End If

If CLng(targetMonth) < 1 Or CLng(targetMonth) > 12 Then
	WScript.Echo "数字を入力してください(1〜12)"
	WScript.Quit
End If

WScript.Echo Year(Now) & "年" & targetMonth & "月の最終日は" & Day(DateAdd("d", DateSerial(Year(Now), targetMonth + 1, 1), -1)) & "日です"
