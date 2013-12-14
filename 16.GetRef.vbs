Call main(3, GetRef(InputBox("hoge?bar?")))

Sub hoge(count)
	WScript.Echo "hoge " & count
End Sub

Sub bar(count)
	WScript.Echo "bar " & count
End Sub

Sub main(count, proc)
	proc(count)
End Sub
