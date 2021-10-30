Sub getworddata()
Dim Wapp As Word.Application
Dim wdoc As Word.Document
ce = 2
Do Until Range("A" & ce) = ""
Set Wapp = CreateObject("word.application")
Set wdoc = Wapp.Documents.Open(Range("A" & ce).Value)
Range("B" & ce).Value = wdoc.Content.Text
wdoc.Close
ce = ce + 1
Wapp.Quit
Loop
End Sub
