Dim strm : Set strm = CreateObject("ADODB.Stream")
With strm
	.CharSet = "utf-8"
	.Open
	.LoadFromFile("./settings/sample_utf.txt")
End With

Dim ostrm : Set ostrm = CreateObject("ADODB.Stream")
With ostrm
	.CharSet = "utf-8"
	.Open
	.WriteText strm.ReadText() & "hogehoge"
	.SaveToFile "./sample_utf_out.txt", 2
End With
