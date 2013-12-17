Dim targetMonth : targetMonth = InputBox("’²‚×‚½‚¢Œ‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢")

If IsNumeric(targetMonth) = False Then
	WScript.Echo "”š‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢(1`12)"
	WScript.Quit
End If

If CLng(targetMonth) < 1 Or CLng(targetMonth) > 12 Then
	WScript.Echo "”š‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢(1`12)"
	WScript.Quit
End If

WScript.Echo Year(Now) & "”N" & targetMonth & "Œ‚ÌÅI“ú‚Í" & Day(DateAdd("d", DateSerial(Year(Now), targetMonth + 1, 1), -1)) & "“ú‚Å‚·"
